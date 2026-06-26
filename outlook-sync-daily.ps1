# Launcher for the daily incremental Outlook -> Obsidian sync.
# Logs each run to .config\opencode\logs\ and exits non-zero on failure
# so Task Scheduler's "restart on failure" can retry.
$ErrorActionPreference = "Stop"

$dir    = Join-Path $env:USERPROFILE ".config\opencode"
$logDir = Join-Path $dir "logs"
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$log = Join-Path $logDir ("outlook-sync-" + (Get-Date -Format "yyyyMMdd") + ".log")

$exit = 0
Start-Transcript -LiteralPath $log -Force | Out-Null
try {
    Write-Host ("=== outlook sync incremental @ " + (Get-Date -Format "o") + " ===")
    & (Join-Path $dir "Outlook-to-Obsidian.ps1") -Mode Incremental
    if (-not $?) { Write-Host ("script reported failure: " + $error[0]); $exit = 1 }
} catch {
    Write-Host ("FATAL: " + $_.Exception.Message)
    $exit = 1
} finally {
    Write-Host ("=== exit code: " + $exit + " ===")
    Stop-Transcript | Out-Null
}
exit $exit
