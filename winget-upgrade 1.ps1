#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Production-ready WinGet updater for both USER and SYSTEM contexts.

.DESCRIPTION
  - USER context: uses winget.exe (bulk upgrade; if needed, per-ID retry).
  - SYSTEM context: uses Microsoft.WinGet.Client PowerShell module (no CLI dependency).
  - Logs to C:\Utils\winget.txt. No UI. Minimal code.

.PARAMETER LogPath
  Log file path. Default: C:\Utils\winget.txt
.PARAMETER TimeoutSec
  Timeout per winget CLI call (USER path only). Default: 1800
.PARAMETER IncludePinned
  Also upgrade pinned packages (USER path only).
#>

[CmdletBinding()]
param(
  [string]$LogPath = 'C:\Utils\winget.txt',
  [ValidateRange(60,86400)][int]$TimeoutSec = 1800,
  [switch]$IncludePinned
)

# -------- basics --------
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

# -------- logging --------
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) { New-Item $logDir -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $LogPath)) { New-Item $LogPath -ItemType File -Force | Out-Null }
function Write-Log {
  param([ValidateSet('INFO','WARN','ERROR')]$Level='INFO', [Parameter(Mandatory)][string]$Message)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  Add-Content -Path $LogPath -Value "$ts [$Level] $Message"
  if ($Level -eq 'ERROR') { Write-Error $Message } elseif ($Level -eq 'WARN') { Write-Warning $Message }
}

# -------- context helpers --------
function Test-IsSystem {
  [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18'
}
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
  $null
}

# -------- USER path (winget.exe) --------
function Invoke-Winget {
  param([Parameter(Mandatory)][string[]]$Args, [int]$Timeout = 1800)
  $exe = Get-WingetPath; if (-not $exe) { throw "WinGet not found for this user session." }
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

  if (-not $p.WaitForExit($Timeout * 1000)) { try{$p.Kill()}catch{}; Write-Log WARN "winget timed out after $Timeout seconds."; return @{ ExitCode = -1; StdOut=''; StdErr='timeout' } }
  $out = $p.StandardOutput.ReadToEnd()
  $err = $p.StandardError.ReadToEnd()
  if ($out) { $out -split "`r?`n" | ? {$_} | % { Write-Log INFO $_ } }
  if ($err) { $err -split "`r?`n" | ? {$_} | % { Write-Log WARN $_ } }
  @{ ExitCode = $p.ExitCode; StdOut = $out; StdErr = $err }
}
function Get-UpgradableIds {
  $res = Invoke-Winget @('upgrade','--include-unknown','--disable-interactivity') -Timeout $TimeoutSec
  $ids = @()
  if ($res.StdOut) {
    foreach ($line in ($res.StdOut -split "`r?`n")) {
      if ($line -match '^\s*.+?\s{2,}([A-Za-z0-9\.\-]+)\s+\S+\s+\S+\s*$') { $ids += $Matches[1] }
    }
  }
  $ids | Sort-Object -Unique
}
function Run-UserPath {
  # Bulk upgrade; if non-zero, retry per-ID
  $bulkArgs = @(
    'upgrade','--all',
    '--accept-package-agreements','--accept-source-agreements',
    '--silent','--disable-interactivity','--nowarn',
    '--authentication-mode','silent','--force'
  )
  if ($IncludePinned) { $bulkArgs += '--include-pinned' }

  $bulk = Invoke-Winget $bulkArgs -Timeout $TimeoutSec
  if ($bulk.ExitCode -eq 0) { return }

  Write-Log WARN "Bulk upgrade returned $($bulk.ExitCode); retrying per-package."
  $ids = Get-UpgradableIds
  foreach ($id in $ids) {
    $args = @(
      'upgrade','--id',$id,'--exact',
      '--accept-package-agreements','--accept-source-agreements',
      '--silent','--disable-interactivity','--nowarn',
      '--authentication-mode','silent','--force'
    )
    if ($IncludePinned) { $args += '--include-pinned' }
    $r = Invoke-Winget $args -Timeout $TimeoutSec
    if ($r.ExitCode -eq 0) { Write-Log INFO "OK: $id" } else { Write-Log WARN ("Fail {0}: {1}" -f $r.ExitCode, $id) }
  }
}

# -------- SYSTEM path (Microsoft.WinGet.Client module) --------
function Ensure-WinGetClientModule {
  if (Get-Module -ListAvailable -Name Microsoft.WinGet.Client) { return $true }
  try {
    if (-not (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue)) {
      Register-PSRepository -Name 'PSGallery' -SourceLocation 'https://www.powershellgallery.com/api/v2' -InstallationPolicy Trusted
    } else {
      Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    }
    Install-Module Microsoft.WinGet.Client -Force -Scope AllUsers -AllowClobber
    $true
  } catch {
    Write-Log ERROR ("Failed to install Microsoft.WinGet.Client module: {0}" -f $_.Exception.Message)
    $false
  }
}
function Run-SystemPath {
  if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client)) {
    if (-not (Ensure-WinGetClientModule)) {
      throw "SYSTEM context: Microsoft.WinGet.Client module unavailable. Run as user, or allow module install."
    }
  }
  Import-Module Microsoft.WinGet.Client -ErrorAction Stop
  Write-Log INFO ("Using module Microsoft.WinGet.Client v{0}" -f ((Get-Module Microsoft.WinGet.Client).Version))

  # Try a bulk update with the module; if not supported, fall back to per-package.
  $didBulk = $false
  try {
    # Newer module versions: -All is supported
    Update-WinGetPackage -All -AcceptPackageAgreements -IncludeUnknown -ErrorAction Stop | Out-Null
    $didBulk = $true
    Write-Log INFO "Module bulk update completed."
  } catch {
    Write-Log WARN ("Module bulk update not available or failed: {0}" -f $_.Exception.Message)
  }

  # Per-package fallback or second pass
  try {
    $outdated = Get-WinGetPackage -Outdated -IncludeUnknown -ErrorAction Stop
  } catch {
    Write-Log WARN ("Get-WinGetPackage failed: {0}" -f $_.Exception.Message)
    $outdated = @()
  }

  if ($outdated -and $outdated.Count -gt 0) {
    foreach ($p in $outdated) {
      try {
        if ($p.Id) {
          Update-WinGetPackage -Id $p.Id -Exact -AcceptPackageAgreements -IncludeUnknown -ErrorAction Stop | Out-Null
          Write-Log INFO ("OK: {0}" -f $p.Id)
        }
      } catch {
        Write-Log WARN ("Fail: {0} -> {1}" -f $p.Id, $_.Exception.Message)
      }
    }
  } elseif (-not $didBulk) {
    Write-Log INFO "No outdated packages reported by module."
  }
}

# -------- main --------
try {
  Write-Log INFO '=== START ==='
  if (Test-IsSystem) {
    Write-Log INFO 'Context: SYSTEM -> using Microsoft.WinGet.Client module'
    Run-SystemPath
  } else {
    Write-Log INFO 'Context: USER -> using winget.exe'
    $wg = Get-WingetPath; if (-not $wg) { throw "WinGet not found. Install App Installer for this user session." }
    Write-Log INFO "WinGet: $wg"
    Run-UserPath
  }

  # Post status (best-effort)
  try {
    if (Test-IsSystem) {
      $remaining = Get-WinGetPackage -Outdated -IncludeUnknown -ErrorAction SilentlyContinue
      if ($remaining -and $remaining.Count -gt 0) {
        Write-Log WARN ("Still pending: " + (($remaining | Select-Object -ExpandProperty Id) -join ', '))
      } else {
        Write-Log INFO 'All up to date.'
      }
    } else {
      $post = Get-UpgradableIds
      if ($post.Count -gt 0) { Write-Log WARN ("Still pending: " + ($post -join ', ')) } else { Write-Log INFO 'All up to date.' }
    }
  } catch {
    Write-Log WARN ("Post-check failed: {0}" -f $_.Exception.Message)
  }

  Write-Log INFO '=== DONE ==='
  exit 0
}
catch {
  Write-Log ERROR ("FATAL: {0}" -f $_.Exception.Message)
  exit 1
}
