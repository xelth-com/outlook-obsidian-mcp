# outlook-obsidian-mcp

Turn your **desktop Outlook** mailbox into a **searchable Obsidian vault** that LLM agents (opencode, Claude Code) can query through the **Model Context Protocol**.

> "Did anyone write me anything important yesterday?" / "How many times did X mention Y?" — ask your AI agent in plain language; it searches the vault via MCP and answers.

- **Email → Markdown**: one `.md` per mail (Inbox + Sent), with structured frontmatter.
- **Content-addressed attachments (CAS)**: every attachment stored once by SHA-256 (`vault/files/<hash>.<ext>`), duplicates collapsed, linked from the note.
- **MCP bridge**: the Obsidian vault is exposed to agents through `mcp-obsidian` (12 tools), so agents can list, read and full-text-search your mail.
- **Autonomous**: a scheduled task pulls new mail Mon–Fri 09:00–18:00 every 30 min, plus shortly after login. No buttons to press.

---

## How it works

```
 Outlook desktop ──(COM, local)──▶  Outlook-to-Obsidian.ps1  ──▶  Obsidian vault
  (auto-downloads mail)              reads new items, writes .md        │
                                      + extracts attachments (CAS)       │
                                                                         ▼
                                        opencode / Claude Code  ◀── mcp-obsidian (HTTP→Obsidian Local REST API)
                                                  │                        │
                                                  ▼                        ▼
                                          "what's new?"           full-text search over mail + attachments
```

Two independent layers — don't confuse them:

1. **Outlook downloads mail from the server** itself, automatically in the background (Send/Receive). You never press a button.
2. **This project's script does not download anything.** It reads what's already in the local Outlook (items newer than the last run) and copies them into the vault as Markdown.

---

## Prerequisites

| Tool | Why | Install |
|---|---|---|
| **Outlook desktop** (Windows, MAPI/Exchange profile) | source of mail | Microsoft 365 / Office |
| **Obsidian** | the vault + Local REST API | https://obsidian.md |
| **Obsidian "Local REST API" plugin** (Adam Coddington) | exposes vault over HTTPS+token | Community plugins |
| **Node.js** | runs the MCP launcher | https://nodejs.org |
| **uv / uvx** | runs `mcp-obsidian` | `pipx install uv` or https://docs.astral.sh/uv |
| **opencode** and/or **Claude Code** | the agent that queries the vault | their official installers |

Windows only (COM automation). Outlook must be set up with your profile at least once (it caches credentials after that — no password lives in this project).

---

## Setup (new machine)

### 1. Obsidian + Local REST API
1. Create a vault, e.g. `%USERPROFILE%\Documents\Obsidian Vault`.
2. Settings → Community plugins → install **"Local REST API"** → enable.
3. Open its settings, copy the **API key** (the token). Leave the server running (Obsidian can stay minimized in the tray).

### 2. Place the scripts
Put these under `%USERPROFILE%\.config\opencode\` (or anywhere — just keep paths consistent):
- `Outlook-to-Obsidian.ps1`
- `outlook-sync-daily.ps1`
- `obsidian-mcp.js`
- `.secret.json` (copy from `.secret.json.example`, paste your REST API token)
- (for opencode) `opencode.json` (copy from `opencode.json.example`, fill provider keys)

`.secret.json`:
```json
{ "obsidianApiKey": "your-obsidian-local-rest-api-token" }
```

> The `obsidian-mcp.js` launcher injects that token into the env before spawning `uvx mcp-obsidian`, because some hosts don't forward env vars to the child (without it, mcp-obsidian crashes on startup).

### 3. Wire the MCP server into your agent
**opencode** — in `opencode.json`:
```jsonc
"mcp": {
  "obsidian": {
    "type": "local",
    "command": ["node", "C:\\Users\\YOU\\.config\\opencode\\obsidian-mcp.js"],
    "enabled": true,
    "timeout": 30000
  }
}
```
**Claude Code**:
```bash
claude mcp add obsidian --scope user -- node "C:\Users\YOU\.config\opencode\obsidian-mcp.js"
```
Restart the agent; it should report 12 `obsidian_*` tools connected.

### 4. First (full) import
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:USERPROFILE\.config\opencode\Outlook-to-Obsidian.ps1" -Mode Full
```
This exports every mail in **Inbox + Sent** to `vault/Mail/` and every attachment to `vault/files/` (deduplicated by hash). Re-run with `-Force` later to rewrite notes / re-extract attachments.

### 5. Keep it fresh — scheduled task
Register the task (run in an elevated or normal PowerShell as your user):
```powershell
$me = whoami
$script = "$env:USERPROFILE\.config\opencode\outlook-sync-daily.ps1"
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$tPoll  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "09:00"
$rep = New-CimInstance -CimClass (Get-CimClass MSFT_TaskRepetitionPattern -Namespace Root/Microsoft/Windows/TaskScheduler) -ClientOnly -Property @{ Interval = "PT30M"; Duration = "PT9H" }
$tPoll.Repetition = $rep
$tLogon = New-ScheduledTaskTrigger -AtLogOn -User $me; $tLogon.Delay = "PT2M"
$principal = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable `
  -RestartInterval (New-TimeSpan -Minutes 30) -RestartCount 3 `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskName "ObsidianOutlookSync" -Action $action `
  -Trigger @($tPoll,$tLogon) -Principal $principal -Settings $settings -Force
```
Behaviour:
- **Mon–Fri, 09:00–18:00**: runs every 30 min.
- **At login**: runs ~2 min later (gives Outlook time to start + sync).
- **If the PC was off / offline**: runs ASAP once on + online (`StartWhenAvailable` + `RunOnlyIfNetworkAvailable`).
- **On failure** (Outlook unreachable, COM error): retries every 30 min, up to 3 times.

> The task must run **only when you're logged on** (Interactive), because it reads your local Outlook MAPI session. `ObsidianOutlookSync.task.xml` is included for reference, but the commands above are the portable way to recreate it.

---

## Usage

Ask your agent (it will call the MCP tools under the hood):
- *"What important mail did I get yesterday?"*
- *"List everything Lea Park sent about the Inbound Instruction."*
- *"Show me the attachments from the invoice thread last week."*

Manual sync options:
```powershell
# last 7 days only
.\Outlook-to-Obsidian.ps1 -Days 7
# dry run (report counts, write nothing)
.\Outlook-to-Obsidian.ps1 -DryRun -Days 7
# include subfolders of each top folder
.\Outlook-to-Obsidian.ps1 -Recurse
# rewrite existing notes + re-extract attachments (e.g. after enabling CAS)
.\Outlook-to-Obsidian.ps1 -Mode Full -Force
# other folders / custom CAS dir
.\Outlook-to-Obsidian.ps1 -Folders Inbox,Sent,Drafts -CasDir files
```

Logs (one file per day, appended): `%USERPROFILE%\.config\opencode\logs\outlook-sync-YYYYMMDD.log`.

---

## What each note looks like

`vault/Mail/2026-05-04 - [Vasja] - ... - a1b2c3d4.md`
```markdown
---
type: mail
outlook_folder: "\\you\Inbox"
date: 2026-05-04 09:35:15
sender: "Vasja"
sender_email: "..."
to: "..."
subject: "..."
entry_id: "..."
attachments: ["79955c2....png", "c804f48....xlsx"]
tags: [mail]
---
# Subject

**From:** ... **Date:** ... **Folder:** ...

<body text>

## Anhaenge
![[79955c2....png]]            ← renderable: inline preview
*c logo.png*
- [[c804f48....xlsx|report.xlsx]]  ← other types: clickable link
```

---

## Files in this repo

| File | Purpose |
|---|---|
| `Outlook-to-Obsidian.ps1` | main exporter (Full/Incremental, CAS attachments, idempotent) |
| `outlook-sync-daily.ps1` | launcher used by the scheduled task (logs + exit code for retry) |
| `obsidian-mcp.js` | MCP launcher: injects the REST API key, spawns `uvx mcp-obsidian` |
| `ObsidianOutlookSync.task.xml` | exported scheduled task (reference) |
| `opencode.json.example` | sanitized opencode config template |
| `.secret.json.example` | where your Obsidian REST API token goes (gitignored) |

---

## Security

- The Obsidian **Local REST API** binds to **localhost**; the token only grants access to your own vault on your own machine.
- **No credentials are stored in this repo.** Real secrets live in `.secret.json` / `opencode.json`, both gitignored — keep them that way. Double-check `git status` / `git log` before pushing.
- The script talks to Outlook **locally via COM** — no cloud, no app secrets, no mail leaves your machine.

## Limitations

- Windows + desktop Outlook only.
- Outlook must be installed and your profile configured (cached creds). The task runs in your interactive session.
- Can't extract: IRM/S-MIME-encrypted bodies, OLE-embedded objects, or online-only items not yet cached by Exchange.
- Internal Exchange senders expose an X.500 address as `sender_email` (not a clean SMTP) — patchable via `MailItem.Sender.GetExchangeUser().PrimarySmtpAddress`.

## License
MIT — see [LICENSE](LICENSE).
