# 🔍 Mail Investigation Tool

A PowerShell-based interactive tool for **Microsoft Purview** that streamlines the process of investigating and purging emails in Microsoft 365 environments.

Built for **Message & Collaboration Admins** who need to perform content searches, analyze email propagation, and execute compliance purges — all from a single, menu-driven interface.

---

## Features

- **Interactive menu UI** — clean terminal interface with persistent context header
- **Compliance Search** — create new searches or load existing ones by name
- **Mail & Meeting Discovery** — build KQL queries interactively for emails or calendar items
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
         ├── MAIL               ├──► Investigation menu
         └── MEETING            │       ├── Forwarding analysis
                                │       └── Execute purge
                                │               └── Iterative loop until Items = 0
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

Before running, set your tenant domain at the top of the script:

```powershell
$TenantDomain = "yourdomain.com"
```

> **Note:** The forwarding analysis detects subjects matching `FW:` and `RE:` prefixes.
> If your tenant appends a custom tag to external emails (e.g. `[External]`, `[EXT]`),
> add the corresponding pattern to the `Where-Object` block in `Invoke-ForwardingAnalysis`.

---

## Notes

- Tested on **Windows PowerShell** with **Windows Terminal**
- Console width is automatically set to 180 characters to preserve banner formatting
- Search naming format: `{operator}_{yyyyMMdd}_{seq}` (e.g. `jdoe_20260317_1`)

---

## Author

**Jonathan Martino** — Systems Engineer / M&C Admin  
[github.com/jonamartino](https://github.com/jonamartino)
