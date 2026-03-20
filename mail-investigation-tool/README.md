# Mail Investigation Tool

A PowerShell-based interactive tool for **Microsoft Purview** that streamlines the process of investigating and purging emails in Microsoft 365 environments.

Built for **Message & Collaboration Admins** who need to perform content searches, analyze email propagation, and execute compliance purges — all from a single, menu-driven interface.

---

## Features

- **Interactive menu UI** — clean terminal interface with persistent context header
- **Compliance Search** — create new searches or load existing ones by name
- **Mail & Meeting Discovery** — build KQL queries interactively for emails or calendar items via numeric menu
- **CSV Recipient Loading** — load large recipient lists from file; accepts any format (raw email client output, semicolon-separated, one per line) with automatic email extraction via regex
- **Exchange Validation** — mailboxes are validated against Exchange Online before search creation, invalid entries are automatically skipped
- **Forwarding Analysis** — detect FW/RE propagation across target mailboxes using Message Trace
- **Iterative Purge** — automated HardDelete purge loop with dual-pass (HardDelete + Recoverable Items) until Items = 0
- **Auto-naming** — search names are generated automatically based on operator and date, with incremental suffix to avoid conflicts

---

## Prerequisites

- PowerShell 5.1 or later
- Modules:
  - `ExchangeOnlineManagement` (installed automatically if missing)
- Microsoft 365 permissions:
  - **Exchange Online**: `View-Only Recipients` or higher
  - **Security & Compliance**: `Compliance Search` + `Search And Purge` roles

---

## Usage

```powershell
.\Mail-Investigation-Tool.ps1
```

The script will:
1. Auto-install `ExchangeOnlineManagement` if not present
2. Prompt for Exchange Online and IPPS authentication
3. Present an interactive menu to load or create a compliance search
4. Allow forwarding analysis or purge execution on the loaded search

---

## Workflow

```
Launch script
    │
    ├── Load existing search  ──┐
    │                           │
    └── Create new discovery    │
         ├── 1 - Mail           ├──► Investigation menu
         └── 2 - Meeting        │       ├── Forwarding analysis
              │                 │       └── Execute purge
              ├── Enter manually│               └── Iterative loop until Items = 0
              └── Load from CSV │
                   └── Regex extraction + Exchange validation
                                └───────────────────────────────────────────────────
```

---

## Purge Logic

The purge function runs in a loop until the compliance search returns `Items = 0`:

1. Re-runs the compliance search to get a fresh item count
2. Executes a **HardDelete** purge action
3. Executes a second purge pass targeting **Recoverable Items**
4. Repeats until no items remain
5. Displays a final summary with item count, size, and iterations

> ⚠️ Purge actions are **irreversible**. Ensure the search scope and query are correct before executing.

---

## Configuration

At the top of the script, update the static config block before running:

```powershell
$TenantDomain  = "yourdomain.com"
$RecipientsCSV = "$env:USERPROFILE\OneDrive - YourOrg\YourFolder\recipients.csv"
```

**`$TenantDomain`** — set this to your M365 tenant domain.

**`$RecipientsCSV`** — path to the CSV file used when loading recipients from file during discovery. Uses `$env:USERPROFILE` to resolve automatically to the current user's profile, so each operator does not need to change it manually as long as the folder structure after `OneDrive` is consistent across the team.

To use CSV-based recipient loading:
1. Create a file named `recipients.csv` at the configured path
2. Paste the recipient list into it — any format is accepted (one per line, semicolon-separated, raw email client output, etc.)
3. The script extracts and deduplicates all valid email addresses automatically
4. Each address is validated against Exchange Online before the search is created — invalid mailboxes are silently skipped

> **Note:** The forwarding analysis detects subjects matching `FW:` and `RE:` prefixes.
> If your tenant appends a custom tag to external emails (e.g. `[External]`, `[EXT]`),
> add the corresponding pattern to the `Where-Object` block in `Invoke-ForwardingAnalysis`.

---

## Notes

- Tested on **Windows PowerShell 5.1** with **Windows Terminal**
- Console width is automatically set to 180 characters to preserve banner formatting
- Search naming format: `{operator}_{yyyyMMdd}_{seq}` (e.g. `jdoe_20260317_1`)
- Exchange validation uses a RunspacePool (20 concurrent threads) for performance on large recipient lists

---

## Author

**Jonathan Martino** — Systems Engineer / M&C Admin  
[github.com/jonamartino](https://github.com/jonamartino)
