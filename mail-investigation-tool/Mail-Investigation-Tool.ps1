function Show-Banner {
    $banner = @"
                                   . .,(#&%##(/,.,,...             ..... ....  ........  ..         
        .  ...                    ,%%%(* ..,,..,*(//.  .        ....  .....  ..........             
  ..    ,   ...                ./%%(**##%%%###(((//##..    ............ ............... .  .        
..  . ..    ....              ,(#//(##%&&&&&%##((((.,((........ .....................               
,....,....  .....            /(///**/(#%&&&&&&%###%(**/(, ..........................                
. ...,..     .,,,,          */**,.       .(###*,.,,,/*(/(,...................,.,,,.....  .          
 .,,.....    ......        ,***,*//      ,/(       ..,,**/,..............,,,.,,,,,,.,......         
 ....,.,,.   ......      . ,,,,//(&&&&&&(#%(*/,..,,. ..,**/,........,,,,.....,/*,,.........         
  ....,*,     ,,,.,.     .,,.. ......./#((#**/####((,.,.,**,........,,,,*/.....,/////....           
,,...,.....    ,,,...    .*,.       ./(#,    ,* ...... ..,*,,,,,,,,,.   /(##/,///**........         
    .......    ,,...,.   .,.   . .,,/##.             , ..,*,,,,,,,,, ,###%%%&&&,,   .  . .          
      .,...     ,.,,*,.  ,...   .  **./%//(#/,.(,   .  ../(****,*** *#####%%%%%%%,                  
.........,,..   ,,*,,,,,,*,./.    *.//.      ..        ..*%@&**,**,*..,#((((    .,,                 
.,,.,,,*,..*.    ***,...,%&#/.      *(#(((//.          ,./#@&@@@(*,      ,/.     (,                .
   ......,.,,,...,.(&&@&&&&&* .                        ..*(%&&&%@(..((((*(%(%#(#&&#(*             ..
 ............**,,,@&&&&&&&##,.                         ..//%%%#((, ,**,,*%(#(*,,*****.             .
       ,,.,/&&*...*&&@&#(%##*.                   .     ..(#%%##( .     ,.  *%//.     .              
    ..  .,##&%#..  /#&%%%#%,. *.                ,      .*./%/*..      *.,,*%&%(,(*...               
        ..#,###*.   /./##%/(///                 *.  .*.,,//**.,          .*.. */,.          .  ..   
"@
    Write-Host $banner -ForegroundColor DarkRed
    Write-Host ""
    Write-Host "                     M A I L   I N V E S T I G A T I O N   T O O L" -ForegroundColor DarkRed
    Write-Host ""
}


# ===============================
# Mail Investigation Tool
# ===============================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===============================
# GLOBAL CONTEXT
# ===============================

$Global:SearchContext = $null

# ===============================
# STATIC CONFIG (INTERNAL TOOL)
# ===============================

$TenantDomain = "yourdomain.com"

# ===============================
# UI HELPERS
# ===============================
$Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(180, 9999)
$Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(180, 50)
function Show-Header {
    param([string]$Section = "")
    Clear-Host
    Show-Banner
    if ($Global:SearchContext) {
        Write-Host "  Search  : $($Global:SearchContext.Name)"                                       -ForegroundColor DarkCyan
        Write-Host "  Mailbox : $($Global:SearchContext.Mailboxes -join ', ')"                       -ForegroundColor DarkCyan
        Write-Host "  Items   : $($Global:SearchContext.Items) ($($Global:SearchContext.SizeMB) MB)" -ForegroundColor DarkCyan
        Write-Host "--------------------------------"                                                -ForegroundColor DarkCyan
    }
    if ($Section) {
        Write-Host "  $Section" -ForegroundColor White
        Write-Host "--------------------------------" -ForegroundColor DarkCyan
    }
    Write-Host ""
}

# ===============================
# CONNECTION
# ===============================

function Connect-Environments {

    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }

    Import-Module ExchangeOnlineManagement

    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }

    if (-not (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue)) {
        Connect-IPPSSession -EnableSearchOnlySession
    }

    Write-Host "  Connections ready." -ForegroundColor Green
    Start-Sleep 1
}

# ===============================
# LOAD EXISTING SEARCH
# ===============================

function Load-ExistingSearch {

    $name   = Read-Host "  Enter existing Compliance Search name"
    $search = Get-ComplianceSearch -Identity $name -ErrorAction Stop

    $Global:SearchContext = [PSCustomObject]@{
        Name      = $search.Name
        Items     = $search.Items
        SizeMB    = [math]::Round($search.Size / 1MB, 2)
        Query     = $search.ContentMatchQuery
        Mailboxes = $search.ExchangeLocation
    }
}

# ===============================
# DISCOVERY
# ===============================

function Invoke-MailDiscovery {

    Show-Header "DISCOVERY"

    $type = (Read-Host "  MAIL or MEETING").ToUpper()
    if ($type -notin @("MAIL","MEETING")) { return }

    # ---------------------------
    # Target mailboxes
    # ---------------------------
    $mailboxes = (Read-Host "  Target mailboxes (comma-separated)") `
        -split ";" |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne "" }

    $mailboxes = @($mailboxes)

    if ($mailboxes.Count -eq 0) {
        Write-Error "No valid target mailboxes provided."
        return
    }

    # ---------------------------
    # Query build
    # ---------------------------
    if ($type -eq "MAIL") {

        $sender  = Read-Host "  Sender"
        $subject = Read-Host "  Subject"
        $date    = Read-Host "  Base date (YYYY-MM-DD)"
        $margin  = [int](Read-Host "  Margin days")

        $start = (Get-Date $date).AddDays(-$margin).ToString("yyyy-MM-dd")
        $end   = (Get-Date $date).AddDays($margin).ToString("yyyy-MM-dd")

        $query = "(From:`"$sender`" AND Subject:`"$subject`" AND Received>=$start AND Received<=$end)"
    }
    else {

        $author  = Read-Host "  Organizer"
        $subject = Read-Host "  Subject (optional)"

        if ($subject) {
            $query = "(ItemClass:IPM.Appointment AND Author:`"$author`" AND Subject:`"$subject`")"
        }
        else {
            $query = "(ItemClass:IPM.Appointment AND Author:`"$author`")"
        }
    }

    # ---------------------------
    # Operator normalization
    # ---------------------------
    $upn = (Get-ConnectionInformation | Select-Object -First 1).UserPrincipalName
    $op  = ($upn.Split("@")[0]).Split(".")[-1].ToLower()

    # ---------------------------
    # Naming (Robust Incremental)
    # ---------------------------
    $today = Get-Date -Format yyyyMMdd
    $base  = "${op}_${today}"
    $seq   = 1

    do {
        $name   = "${base}_$seq"
        $exists = Get-ComplianceSearch -Identity $name -ErrorAction SilentlyContinue
        $seq++
    }
    while ($exists)

    Write-Host ""
    Write-Host "  Generated search name: $name" -ForegroundColor Green

    # ---------------------------
    # Create & run search 
    # ---------------------------
    $params = @{
        Name              = $name
        ExchangeLocation  = $mailboxes
        ContentMatchQuery = $query
    }
    New-ComplianceSearch @params

    Start-ComplianceSearch -Identity $name

    do {
        Start-Sleep 10
        $status = Get-ComplianceSearch -Identity $name
        Write-Host "`r  Status: $($status.Status) | Items: $($status.Items)   " -NoNewline -ForegroundColor DarkGray
    }
    while ($status.Status -in @("Starting","InProgress"))

    Write-Host ""

    # ---------------------------
    # Context
    # ---------------------------
    $Global:SearchContext = [PSCustomObject]@{
        Name      = $name
        Items     = $status.Items
        SizeMB    = [math]::Round($status.Size / 1MB, 2)
        Query     = $query
        Mailboxes = $mailboxes
    }

    Write-Host ""
    Write-Host "  Search ready." -ForegroundColor Green
    Read-Host "  Press Enter to continue"
}

# ===============================
# FORWARDING ANALYSIS (FW / RE DETECTION)
# ===============================

function Invoke-ForwardingAnalysis {

    if (-not $Global:SearchContext) {
        Write-Host "  No search loaded." -ForegroundColor Red
        return
    }

    Show-Header "FORWARDING / REPLY ANALYSIS"

    $days = Read-Host "  Days to look back (default 10)"
    if (-not $days) { $days = 10 }

    $start = (Get-Date).AddDays(-[int]$days)
    $end   = Get-Date

    if ($Global:SearchContext.Query -match 'Subject:"([^"]+)"') {
        $baseSubject = $matches[1]
    }
    else {
        Write-Host "  Unable to extract subject from search query." -ForegroundColor Red
        Read-Host "  Press Enter to continue"
        return
    }

    Write-Host ""
    Write-Host "  Base subject detected: $baseSubject" -ForegroundColor Yellow
    Write-Host ""

    $findings = @()

    foreach ($mailbox in $Global:SearchContext.Mailboxes) {

        Write-Host "  Checking mailbox: $mailbox" -ForegroundColor Cyan

        $trace = Get-MessageTraceV2 `
            -StartDate $start `
            -EndDate $end `
            -SenderAddress $mailbox `
            -ResultSize 5000 |
            Where-Object {
                $_.Subject -and (
                    $_.Subject -like "FW:*${baseSubject}*" -or
                    $_.Subject -like "RE:*${baseSubject}*" -or
                    $_.Subject -like "RE: [External]*${baseSubject}*" -or
                    $_.Subject -like "FW: [External]*${baseSubject}*" -or
                    $_.Subject -like "[External]*${baseSubject}*" -or
                    $_.Subject -like "*${baseSubject}*"
                )
            }

        if ($trace) {
            foreach ($item in $trace) {
                $findings += [PSCustomObject]@{
                    Mailbox  = $mailbox
                    SentDate = $item.Received
                    To       = $item.RecipientAddress
                    Subject  = $item.Subject
                }
            }
        }
        else {
            Write-Host "  No forwarding/reply activity detected for this mailbox." -ForegroundColor Green
        }
    }

    Write-Host ""

    if (-not $findings) {
        Write-Host "  No forwarding or reply propagation detected." -ForegroundColor Green
    }
    else {
        Write-Host "================================" -ForegroundColor Red
        Write-Host "   PROPAGATION DETECTED         " -ForegroundColor Red
        Write-Host "================================" -ForegroundColor Red
        $findings |
            Sort-Object SentDate |
            Format-Table Mailbox, SentDate, To, Subject -AutoSize
    }

    Read-Host "  Press Enter to continue"
}

# ===============================
# PURGE (ITERATIVE HARD DELETE)
# ===============================

function Invoke-IterativePurge {

    if (-not $Global:SearchContext) {
        Write-Host "  No search loaded." -ForegroundColor Red
        return
    }

    $name        = $Global:SearchContext.Name
    $iteration   = 1
    $totalItems  = $Global:SearchContext.Items
    $totalSizeMB = $Global:SearchContext.SizeMB

    do {

        Show-Header "PURGE  >>  Iteration $iteration"

        # 1. Re-run search to get a fresh item count — logic unchanged
        Write-Host "  Re-running compliance search..." -ForegroundColor Gray
        Start-ComplianceSearch -Identity $name | Out-Null

        $status = $null
        do {
            Start-Sleep 10
            $status = Get-ComplianceSearch -Identity $name
            Write-Host "`r  Search status: $($status.Status) | Items: $($status.Items)   " -NoNewline -ForegroundColor DarkGray
        } while ($status.Status -in @("Starting", "InProgress"))

        Write-Host ""

        if ($status.Items -eq 0) {
            Write-Host "  No items remaining. Purge completed successfully." -ForegroundColor Green
            break
        }

        Write-Host "  Items found: $($status.Items). Proceeding with purge..." -ForegroundColor Yellow
        Write-Host ""

        # 2. Remove existing purge action if present — logic unchanged
        $existingAction = Get-ComplianceSearchAction -Identity "${name}_Purge" -ErrorAction SilentlyContinue
        if ($existingAction) {
            Write-Host "  Removing previous purge action..." -ForegroundColor DarkGray
            Remove-ComplianceSearchAction -Identity "${name}_Purge" -Confirm:$false | Out-Null
            Start-Sleep 5
        }

        # 3. First pass: HardDelete — logic unchanged
        Write-Host "  [1/2] HardDelete pass..." -ForegroundColor Yellow

        New-ComplianceSearchAction `
            -SearchName $name `
            -Purge `
            -PurgeType HardDelete `
            -Confirm:$false `
            -Force | Out-Null

        $action = $null
        do {
            Start-Sleep 10
            $action = Get-ComplianceSearchAction -Identity "${name}_Purge" -ErrorAction SilentlyContinue
            if ($action) {
                Write-Host "`r        Status: $($action.Status)   " -NoNewline -ForegroundColor DarkGray
            }
        } while ($action -and $action.Status -in @("Starting", "InProgress"))

        Write-Host "`r        Status: Completed   " -ForegroundColor Green
        Write-Host ""

        # 4. Remove action before creating the second one — logic unchanged
        Remove-ComplianceSearchAction -Identity "${name}_Purge" -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep 5

        # 5. Second pass: no PurgeType — logic unchanged
        Write-Host "  [2/2] Recoverable Items pass..." -ForegroundColor Yellow

        New-ComplianceSearchAction `
            -SearchName $name `
            -Purge `
            -Confirm:$false `
            -Force | Out-Null

        $action = $null
        do {
            Start-Sleep 10
            $action = Get-ComplianceSearchAction -Identity "${name}_Purge" -ErrorAction SilentlyContinue
            if ($action) {
                Write-Host "`r        Status: $($action.Status)   " -NoNewline -ForegroundColor DarkGray
            }
        } while ($action -and $action.Status -in @("Starting", "InProgress"))

        Write-Host "`r        Status: Completed   " -ForegroundColor Green
        Write-Host ""

        $iteration++

    } while ($true) # Loop exits when search returns Items = 0

    # Refresh search context after purge completes — logic unchanged
    $updated = Get-ComplianceSearch -Identity $name
    $Global:SearchContext.Items  = $updated.Items
    $Global:SearchContext.SizeMB = [math]::Round($updated.Size / 1MB, 2)

    # Final summary
    Write-Host "================================" -ForegroundColor Green
    Write-Host "   PURGE COMPLETED              " -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Green
    Write-Host "  Search     : $name"                                          -ForegroundColor White
    Write-Host "  Mailbox    : $($Global:SearchContext.Mailboxes -join ', ')"  -ForegroundColor White
    Write-Host "  Purged     : $totalItems items ($totalSizeMB MB)"            -ForegroundColor White
    Write-Host "  Iterations : $($iteration - 1)"                             -ForegroundColor White
    Write-Host "================================" -ForegroundColor Green
    Write-Host ""
    Read-Host "  Press Enter to continue"
}

# ===============================
# MAIN
# ===============================

Show-Header
Connect-Environments

# ---- PHASE 1: SEARCH CONTEXT ----
do {

    Show-Header "SEARCH CONTEXT"
    Write-Host "  1 - Load existing search"
    Write-Host "  2 - Create new discovery"
    Write-Host "  3 - Exit"
    Write-Host ""

    switch (Read-Host "  Select option") {
        "1" { Load-ExistingSearch }
        "2" { Invoke-MailDiscovery }
        "3" { return }
        default { Write-Host "`n  Invalid option. Please select 1, 2 or 3." -ForegroundColor Yellow ; Start-Sleep 2 }
    }

}
while (-not $Global:SearchContext)

# ---- PHASE 2: INVESTIGATION ----
$exitPhase2 = $false

do {

    Show-Header "INVESTIGATION"
    Write-Host "  1 - Forwarding analysis"
    Write-Host "  2 - Execute purge"
    Write-Host "  3 - Exit"
    Write-Host ""

    switch (Read-Host "  Select option") {
        "1" { Invoke-ForwardingAnalysis }
        "2" { Invoke-IterativePurge }
        "3" { $exitPhase2 = $true }
        default { Write-Host "`n  Invalid option. Please select 1, 2 or 3." -ForegroundColor Yellow ; Start-Sleep 2 }
    }

}
while (-not $exitPhase2)