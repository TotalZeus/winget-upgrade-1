#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Minimal, production-ready WinGet updater for Windows 10/11.
.DESCRIPTION
  - Assumes WinGet/App Installer is already installed.
  - Bulk upgrade first; if bulk returns non-zero, retry per-package.
  - Silent, non-interactive, no popups; logs to C:\Utils\winget.txt.
.PARAMETER LogPath
  Log file path. Default: C:\Utils\winget.txt
.PARAMETER TimeoutSec
  Timeout per winget call in seconds. Default: 1800 (30 min)
.PARAMETER IncludePinned
  Also upgrade pinned packages (off by default).
#>

[CmdletBinding()]
param(
  [string]$LogPath = 'C:\Utils\winget.txt',
  [ValidateRange(60,86400)][int]$TimeoutSec = 1800,
  [switch]$IncludePinned
)

# ---------- logging ----------
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) { New-Item $logDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $LogPath)) { New-Item $LogPath -ItemType File -Force | Out-Null }

function Write-Log {
  param([ValidateSet('INFO','WARN','ERROR')]$Level='INFO',[Parameter(Mandatory)][string]$Message)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  Add-Content -Path $LogPath -Value "$ts [$Level] $Message"
  if ($Level -eq 'ERROR') { Write-Error $Message }
  elseif ($Level -eq 'WARN') { Write-Warning $Message }
}

# ---------- helpers ----------
function Ensure-PathHasWindowsApps {
  $apps = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps'
  if ($apps -and (Test-Path $apps) -and ($env:PATH -notmatch [Regex]::Escape($apps))) { $env:PATH += ";$apps" }
}
function Get-WingetPath {
  $c = Get-Command winget -ErrorAction SilentlyContinue
  if ($c) { return $c.Source }

  Ensure-PathHasWindowsApps

  $c = Get-Command winget -ErrorAction SilentlyContinue
  if ($c) { return $c.Source }

  return $null
}

function Invoke-Winget {
  param([Parameter(Mandatory)][string[]]$Args)
  $exe = Get-WingetPath; if (-not $exe) { throw "WinGet not found. Install 'App Installer' for this user." }

  Write-Log INFO ("winget " + ($Args -join ' '))
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $exe
  $psi.Arguments = ($Args -join ' ')
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError  = $true
  $psi.UseShellExecute        = $false
  $psi.CreateNoWindow         = $true
  $psi.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
  $psi.ErrorDialog            = $false

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  [void]$p.Start()

  if (-not $p.WaitForExit($TimeoutSec * 1000)) {
    try { $p.Kill() } catch { }
    Write-Log WARN "winget timed out after $TimeoutSec seconds."
    return @{ ExitCode = -1; StdOut=''; StdErr='timeout' }
  }

  $out = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  if ($out) { $out -split "`r?`n" | ? {$_} | % { Write-Log INFO $_ } }
  if ($err) { $err -split "`r?`n" | ? {$_} | % { Write-Log WARN $_ } }
  @{ ExitCode = $p.ExitCode; StdOut = $out; StdErr = $err }
}

function Get-UpgradableIds {
  # Parse IDs from: winget upgrade --include-unknown
  $res = Invoke-Winget @('upgrade','--include-unknown','--disable-interactivity')
  $ids = @()
  if ($res.StdOut) {
    foreach ($line in ($res.StdOut -split "`r?`n")) {
      # Match: <Name><2+ spaces><Id><spaces><Installed><spaces><Available>
      if ($line -match '^\s*.+?\s{2,}([A-Za-z0-9\.\-]+)\s+\S+\s+\S+\s*$') { $ids += $Matches[1] }
    }
  }
  $ids | Sort-Object -Unique
}

# ---------- run ----------
try {
  Write-Log INFO '=== START ==='
  Ensure-PathHasWindowsApps

  $wg = Get-WingetPath
  if (-not $wg) { throw "WinGet not found. Install App Installer for this user session." }
  Write-Log INFO "WinGet: $wg"

  # (optional) informational pre-snapshot
  $pre = Get-UpgradableIds
  if ($pre.Count -gt 0) { Write-Log INFO ("Pending: " + ($pre -join ', ')) } else { Write-Log INFO 'No pending updates.' }

  # bulk upgrade (quiet, non-interactive)
  $bulkArgs = @(
    'upgrade','--all',
    '--accept-package-agreements','--accept-source-agreements',
    '--silent','--disable-interactivity','--nowarn',
    '--authentication-mode','silent','--force'
  )
  if ($IncludePinned) { $bulkArgs += '--include-pinned' }
  $bulk = Invoke-Winget $bulkArgs

  # if bulk didn’t fully succeed, try per-package passes on whatever still shows as upgradable
  if ($bulk.ExitCode -ne 0) {
    Write-Log WARN "Bulk upgrade returned exit code $($bulk.ExitCode); retrying per-package."
    $ids = Get-UpgradableIds
    foreach ($id in $ids) {
      $args = @(
        'upgrade','--id',$id,'--exact',
        '--accept-package-agreements','--accept-source-agreements',
        '--silent','--disable-interactivity','--nowarn',
        '--authentication-mode','silent','--force'
      )
      if ($IncludePinned) { $args += '--include-pinned' }
      $r = Invoke-Winget $args
      if ($r.ExitCode -eq 0) { Write-Log INFO "OK: $id" } else { Write-Log WARN ("Fail {0}: {1}" -f $r.ExitCode, $id) }
    }
  }

  # post-snapshot (what’s still pending)
  $post = Get-UpgradableIds
  if ($post.Count -gt 0) { Write-Log WARN ("Still pending: " + ($post -join ', ')) } else { Write-Log INFO 'All up to date.' }

  Write-Log INFO '=== DONE ==='
  exit 0
}
catch {
  Write-Log ERROR ("FATAL: {0}" -f $_.Exception.Message)
  exit 1
}
