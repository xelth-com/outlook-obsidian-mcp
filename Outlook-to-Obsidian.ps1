<#
.SYNOPSIS
  Exports desktop Outlook mail items into the Obsidian vault as Markdown,
  with attachments stored in a content-addressable (CAS) folder.

.DESCRIPTION
  Talks to Outlook via COM (New-Object Outlook.Application) and writes one .md
  per mail into <VaultPath>/<MailSubDir>/. Filenames are stable and idempotent:
  "YYYY-MM-DD - [Sender] - Subject.md", with a short EntryID hash appended only
  on collision, so re-runs never duplicate.

  Attachments are extracted into <VaultPath>/<CasDir>/ named by their SHA-256
  hash (e.g. a1b2....pdf). Identical content -> identical filename -> one copy.
  Each mail note references them: renderable types via ![[hash.ext]], others via
  [[hash.ext|original_name]].

  Modes:
    Full         - everything in selected folders (skip already-exported).  [default]
    Incremental  - only items newer than last run's high-water mark (state file).

  Switches:
    -Days N      ad-hoc: only mail received within last N days (does not touch state)
    -Limit N     stop after exporting N new files (testing)
    -Recurse     include subfolders of each top folder
    -DryRun      report counts, write nothing
    -Force       overwrite existing .md (use once to backfill attachments after enabling CAS)
    -CasDir DIR  CAS subfolder inside the vault for attachments (default: files)
#>

[CmdletBinding()]
param(
    [string]$VaultPath  = (Join-Path $env:USERPROFILE "Documents\Obsidian Vault"),
    [string]$MailSubDir = "Mail",
    [string]$CasDir     = "files",
    [string[]]$Folders  = @("Inbox", "Sent"),
    [ValidateSet("Full", "Incremental")] [string]$Mode = "Full",
    [int]$Days   = 0,
    [int]$Limit  = 0,
    [switch]$Recurse,
    [switch]$DryRun,
    [switch]$Force,
    [string]$StateFile = (Join-Path $env:USERPROFILE ".config\opencode\.outlook-sync-state.json")
)

$ErrorActionPreference = "Stop"

# ---------------- IOleMessageFilter (handles RPC_E_CALL_REJECTED) ----------------
if (-not ("OleMessageFilter" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

[ComVisible(true)]
public class OleMessageFilter : IOleMessageFilter {
    private const int SERVERCALL_ISHANDLED = 0;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register(out IOleMessageFilter previous) {
        CoRegisterMessageFilter(new OleMessageFilter(), out previous);
    }
    public static void Restore(IOleMessageFilter previous) {
        IOleMessageFilter dummy;
        CoRegisterMessageFilter(previous, out dummy);
    }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return SERVERCALL_ISHANDLED; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) { return 500; /* ms, then retry */ }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return PENDINGMSG_WAITDEFPROCESS; }
}
"@
}

# --- olDefaultFolders / olObjectClass constants ---
$olFolders = @{ Inbox = 6; Sent = 5; Drafts = 16; Outbox = 4; Deleted = 3 }
$olMail = 43

# extensions Obsidian can render inline via ![[ ]]
$EmbeddableExt = @{
    png=$true; jpg=$true; jpeg=$true; gif=$true; bmp=$true; webp=$true; svg=$true; avif=$true; jfif=$true; tif=$true; tiff=$true; heic=$true; ico=$true
    pdf=$true
    mp3=$true; wav=$true; ogg=$true; m4a=$true; flac=$true; aac=$true; aiff=$true; opus=$true
    mp4=$true; webm=$true; mov=$true; mkv=$true; ogv=$true; m4v=$true
}

# ---------------- helpers ----------------
function ConvertTo-SafeName([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "(kein Betreff)" }
    $s = $s -replace '[<>:"/\\|?*\x00-\x1F]', ' '
    $s = $s -replace '\s+', ' '
    $s = $s.Trim().TrimEnd('.')
    if ($s.Length -gt 90) { $s = $s.Substring(0, 90).TrimEnd() }
    return $s
}

function Get-EntryHash([string]$id) {
    if ([string]::IsNullOrEmpty($id)) { return "" }
    $sha = [System.Security.Cryptography.SHA1]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($id)
    $h = $sha.ComputeHash($bytes)
    return ([BitConverter]::ToString($h[0..3]) -replace '-', '').ToLower()
}

function Get-Ext([string]$name) {
    if ([string]::IsNullOrEmpty($name)) { return "" }
    $idx = $name.LastIndexOf('.')
    if ($idx -lt 1 -or $idx -eq $name.Length - 1) { return "" }
    return $name.Substring($idx + 1).ToLower()
}

function Get-FileHash256([string]$path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = [System.IO.File]::OpenRead($path)
    try { $h = $sha.ComputeHash($fs) } finally { $fs.Dispose() }
    return ([BitConverter]::ToString($h) -replace '-', '').ToLower()
}

function Get-MdBody($item) {
    $body = $item.Body
    if ([string]::IsNullOrWhiteSpace($body)) {
        $html = $item.HTMLBody
        if (-not [string]::IsNullOrWhiteSpace($html)) {
            $body = $html -replace '(?s)<script.*?</script>', ' '
            $body = $body -replace '(?s)<style.*?</style>', ' '
            $body = $body -replace '<[^>]+>', ' '
            $body = $body -replace '&nbsp;', ' '
            $body = $body -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
            $body = $body -replace '\s+', ' '
            $body = $body.Trim()
        }
    }
    if ([string]::IsNullOrWhiteSpace($body)) { return "(kein Textinhalt)" }
    return $body
}

# extracts attachments of a mail item into the CAS dir.
# returns a list of @{ file=<hash.ext>; original=<name>; embedded=<bool> }
function Export-Attachments($item, $casPath) {
    $result = New-Object System.Collections.Generic.List[object]
    $count = 0
    try { $count = [int]$item.Attachments.Count } catch { return ,$result }
    if ($count -le 0) { return ,$result }
    for ($i = 1; $i -le $count; $i++) {
        $tmpFile = $null
        try {
            $att = $item.Attachments.Item($i)
            $orig = $att.FileName
            if ([string]::IsNullOrWhiteSpace($orig)) { $orig = $att.DisplayName }
            if ([string]::IsNullOrWhiteSpace($orig)) { $orig = "attachment.bin" }
            $ext = Get-Ext $orig
            $suffix = if ($ext) { "." + $ext } else { "" }
            $tmpFile = Join-Path $env:TEMP ("obs_att_" + [guid]::NewGuid().ToString("N") + $suffix)
            $att.SaveAsFile($tmpFile)
            if (-not (Test-Path -LiteralPath $tmpFile)) { $script:stats.attachFailed++; continue }

            $hash = Get-FileHash256 $tmpFile
            $casName = if ($ext) { "$hash.$ext" } else { $hash }
            $casFile = Join-Path $casPath $casName
            if (Test-Path -LiteralPath $casFile) {
                $script:stats.attachDedup++
            } else {
                [System.IO.File]::Copy($tmpFile, $casFile, $true)
                $script:stats.attachSaved++
            }
            $alias = $orig -replace '[\|\]\r\n]', '_'
            [void]$result.Add(@{ file = $casName; original = $orig; alias = $alias; embedded = [bool]$EmbeddableExt.ContainsKey($ext) })
        } catch {
            $script:stats.attachFailed++
        } finally {
            if ($tmpFile -and (Test-Path -LiteralPath $tmpFile -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
    return ,$result
}

function Resolve-Folder($ns, [string]$name) {
    $code = $olFolders[$name]
    if ($null -ne $code) { return $ns.GetDefaultFolder($code) }
    # try by name across the root subfolders
    $root = $ns.DefaultStore.GetRootFolder()
    foreach ($f in $root.Folders) {
        if ($f.Name -ieq $name) { return $f }
    }
    return $null
}

function Collect-Items($folder, [bool]$recurse) {
    $col = New-Object System.Collections.Generic.List[object]
    foreach ($i in $folder.Items) { [void]$col.Add($i) }
    if ($recurse) {
        foreach ($sub in $folder.Folders) {
            foreach ($i in (Collect-Items $sub $true)) { [void]$col.Add($i) }
        }
    }
    return ,$col
}

# ---------------- connect ----------------
Write-Host "[sync] connecting to Outlook via COM ..."
$prevFilter = $null
[OleMessageFilter]::Register([ref]$prevFilter)

$outlook = $null
$attempt = 0
while ($null -eq $outlook -and $attempt -lt 5) {
    $attempt++
    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Warning ("[sync] connect attempt {0} failed: {1}" -f $attempt, $_.Exception.Message)
        Start-Sleep -Seconds 2
    }
}
if ($null -eq $outlook) { throw "Could not connect to Outlook after $attempt attempts." }
$ns = $outlook.GetNamespace("MAPI")

$targetFolders = New-Object System.Collections.Generic.List[object]
foreach ($fname in $Folders) {
    $f = Resolve-Folder $ns $fname
    if ($f) {
        [void]$targetFolders.Add($f)
        Write-Host ("[sync] folder: {0}  ({1} items)" -f $f.FolderPath, $f.Items.Count)
    } else {
        Write-Warning "[sync] folder not found: $fname"
    }
}
if ($targetFolders.Count -eq 0) {
    throw "No Outlook folders resolved. Check -Folders."
}

# ---------------- date cutoff ----------------
$since = [datetime]::MinValue
switch -Wildcard ($Mode) {
    "Incremental" {
        if (Test-Path -LiteralPath $StateFile) {
            $st = Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json
            if ($st.lastReceived) { $since = [datetime]$st.lastReceived }
            Write-Host ("[sync] incremental since {0:yyyy-MM-dd HH:mm}" -f $since)
        } else {
            Write-Host "[sync] no state file -> incremental acts as full"
        }
    }
    default {
        if ($Days -gt 0) {
            $since = (Get-Date).AddDays(-$Days)
            Write-Host ("[sync] full, filtered to last {0} days (since {1:yyyy-MM-dd HH:mm})" -f $Days, $since)
        } else {
            Write-Host "[sync] full export (no date filter)"
        }
    }
}

# ---------------- prepare output dirs ----------------
$outDir  = Join-Path $VaultPath $MailSubDir
$casPath = Join-Path $VaultPath $CasDir
if (-not $DryRun) {
    if (-not (Test-Path -LiteralPath $outDir))  { New-Item -ItemType Directory -Path $outDir  -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $casPath)) { New-Item -ItemType Directory -Path $casPath -Force | Out-Null }
}
$enc = New-Object System.Text.UTF8Encoding($true) # UTF-8 BOM, Obsidian-friendly

# ---------------- export loop ----------------
$stats = @{ scanned = 0; nonMail = 0; filtered = 0; existed = 0; written = 0; error = 0; attachSaved = 0; attachDedup = 0; attachFailed = 0 }
$highWater = $since

foreach ($folder in $targetFolders) {
    $items = Collect-Items $folder ([bool]$Recurse)
    foreach ($item in $items) {
        $stats.scanned++
        if ($item.Class -ne $olMail) { $stats.nonMail++; continue }

        $received = $item.ReceivedTime
        if (-not $received) { $received = $item.SentOn }
        if (-not $received) { $received = $item.CreationTime }
        try { $received = [datetime]$received } catch { $received = [datetime]::Now }

        if ($received -gt $highWater) { $highWater = $received }

        if ($received -le $since) { $stats.filtered++; continue }

        $dateStr  = $received.ToString("yyyy-MM-dd")
        $dateIso  = $received.ToString("yyyy-MM-dd HH:mm:ss")
        $sender   = $item.SenderName
        $senderE  = $item.SenderEmailAddress
        $subject  = ConvertTo-SafeName $item.Subject
        $entryId  = $item.EntryID
        $hash     = Get-EntryHash $entryId
        $fpath    = if ($item.Parent) { $item.Parent.FolderPath } else { $folder.FolderPath }

        $base = "{0} - [{1}] - {2}" -f $dateStr, $sender, $subject
        $base = (ConvertTo-SafeName $base)
        # deterministic filename: always append the short entry-id hash so re-runs
        # (including -Force) hit the same path and never leave orphaned old-name files.
        if ($hash) { $name = "$base - $hash.md" } else { $name = "$base.md" }
        $path = Join-Path $outDir $name

        if ((-not $Force) -and (Test-Path -LiteralPath $path)) {
            $stats.existed++; continue
        }

        if ($Limit -gt 0 -and $stats.written -ge $Limit) { continue }

        if ($DryRun) {
            $stats.written++
            Write-Host ("  [dry] {0} | {1}" -f $dateStr, $subject)
            continue
        }

        try {
            $body = Get-MdBody $item
            $atts = Export-Attachments $item $casPath

            if ($atts.Count -gt 0) {
                $attYaml = "[" + ((($atts | ForEach-Object { '"' + $_.file + '"' })) -join ', ') + "]"
            } else {
                $attYaml = "[]"
            }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("---")
            [void]$sb.AppendLine("type: mail")
            [void]$sb.AppendLine("outlook_folder: `"$fpath`"")
            [void]$sb.AppendLine("date: $dateIso")
            [void]$sb.AppendLine("sender: `"$sender`"")
            [void]$sb.AppendLine("sender_email: `"$senderE`"")
            [void]$sb.AppendLine("to: `"$($item.To)`"")
            if ($item.CC) { [void]$sb.AppendLine("cc: `"$($item.CC)`"") }
            [void]$sb.AppendLine("subject: `"$($item.Subject)`"")
            [void]$sb.AppendLine("entry_id: `"$entryId`"")
            [void]$sb.AppendLine("attachments: $attYaml")
            [void]$sb.AppendLine("tags: [mail]")
            [void]$sb.AppendLine("---")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("# " + ($(if ($item.Subject) { $item.Subject } else { "(kein Betreff)" })))
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("**From:** $sender <$senderE>  ")
            [void]$sb.AppendLine("**To:** $($item.To)  ")
            if ($item.CC) { [void]$sb.AppendLine("**CC:** $($item.CC)  ") }
            [void]$sb.AppendLine("**Date:** $dateIso  ")
            [void]$sb.AppendLine("**Folder:** $fpath  ")
            if ($atts.Count -gt 0) { [void]$sb.AppendLine("**Anhaenge:** $($atts.Count)  ") }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine($body)
            if ($atts.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("## Anhaenge")
                foreach ($a in $atts) {
                    if ($a.embedded) {
                        [void]$sb.AppendLine("")
                        [void]$sb.AppendLine("![[$($a.file)]]")
                        [void]$sb.AppendLine("*$($a.alias)*")
                    } else {
                        [void]$sb.AppendLine("- [[$($a.file)|$($a.alias)]]")
                    }
                }
            }
            $md = $sb.ToString()
            [System.IO.File]::WriteAllText($path, $md, $enc)
            $stats.written++
            if (($stats.written % 50) -eq 0) {
                Write-Host ("  ... {0} written" -f $stats.written)
            }
        } catch {
            $stats.error++
            Write-Warning ("[sync] failed: {0} | {1} | {2}" -f $dateStr, $subject, $_.Exception.Message)
        }
    }
}

# ---------------- update state (real runs only) ----------------
if (-not $DryRun) {
    $state = @{ lastRun = (Get-Date).ToString("o"); lastReceived = $highWater.ToString("o") }
    $state | ConvertTo-Json | Set-Content -LiteralPath $StateFile -Encoding UTF8
}

# ---------------- release COM ----------------
foreach ($folder in $targetFolders) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) }
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
[OleMessageFilter]::Restore($prevFilter)

Write-Host ""
Write-Host ("[sync] done. scanned={0} nonMail={1} filteredByDate={2} alreadyPresent={3} written={4} errors={5}" `
    -f $stats.scanned, $stats.nonMail, $stats.filtered, $stats.existed, $stats.written, $stats.error)
Write-Host ("[sync] attachments: saved={0} deduped={1} failed={2}" -f $stats.attachSaved, $stats.attachDedup, $stats.attachFailed)
if (-not $DryRun) {
    Write-Host ("[sync] high-water received = {0:yyyy-MM-dd HH:mm:ss}" -f $highWater)
    Write-Host ("[sync] state -> {0}" -f $StateFile)
}
