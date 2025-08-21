#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Windows 10/11 workstations: silently upgrade all packages with winget (no downloads, no process killing).

.DESCRIPTION
  - Assumes WinGet (App Installer) is already present for the user.
  - Heals only the 'winget' source and quietly removes 'msstore' (common policy block) to reduce noise.
  - Runs a bulk upgrade first, then a per-package second pass if needed (parses text output).
  - No UI popups, bounded execution via timeouts, robust logging with rotation, and single-instance guard.

.PARAMETER LogPath
  Path to log file. Default: C:\Utils\winget.txt
.PARAMETER MaxLogSizeMB
  Rotate the log when it exceeds this size (MB). Default: 10
.PARAMETER WingetTimeoutSec
  Max seconds to allow a winget command before it is killed. Default: 1800 (30 min)
.PARAMETER IncludePinned
  Include pinned packages in bulk/per-ID upgrades (off by default).
#>

[CmdletBinding()]
param(
  [ValidateNotNullOrEmpty()][string]$LogPath      = 'C:\Utils\winget.txt',
  [ValidateRange(1,200)]    [int]   $MaxLogSizeMB = 10,
  [ValidateRange(60,86400)] [int]   $WingetTimeoutSec = 1800,
  [switch]$IncludePinned
)

# -------------------- Strict env --------------------
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13

# -------------------- Logging -----------------------
function Initialize-Logging {
  $dir = Split-Path -Path $LogPath -Parent
  if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
  if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType File -Force | Out-Null }
}
function Write-Log {
  param([ValidateSet('INFO','WARN','ERROR')]$Level='INFO', [Parameter(Mandatory)][string]$Message)
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  Add-Content -Path $LogPath -Value ("{0} [{1}] {2}" -f $ts, $Level, $Message)
  switch ($Level) {
    'ERROR' { Write-Error   $Message }
    'WARN'  { Write-Warning $Message }
    default { Write-Verbose $Message }
  }
}
function Rotate-Log {
  try {
    if (Test-Path $LogPath) {
      $sizeMB = (Get-Item $LogPath).Length / 1MB
      if ($sizeMB -ge $MaxLogSizeMB) {
        $bak = "{0}.{1}.bak" -f $LogPath, (Get-Date).ToString('yyyyMMddHHmmss')
        Move-Item -Path $LogPath -Destination $bak -Force
        Write-Log INFO ("Log rotated to {0}" -f $bak)
      }
    }
  } catch { Write-Log WARN ("Log rotation failed: {0}" -f $_.Exception.Message) }
}
function Get-Count { param($Value) if ($null -eq $Value) { return 0 } try { return ($Value | Measure-Object).Count } catch { return 0 } }

# -------------------- Single-instance guard ----------
$mutex = $null
try {
  $mutex = New-Object System.Threading.Mutex($false, "Global\WingetUpdateMutex")
  if (-not $mutex.WaitOne(0)) {
    Initialize-Logging; Write-Log WARN "Another instance is running. Exiting."; exit 0
  }
} catch {
  Initialize-Logging; Write-Log WARN "Mutex init failed; proceeding without single-instance guard. Error: $($_.Exception.Message)"
}

# -------------------- Winget resolution --------------
function Ensure-PathHasWindowsApps {
  $apps = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps'
  if ($apps -and (Test-Path $apps) -and ($env:PATH -notmatch [Regex]::Escape($apps))) {
    $env:PATH = "$env:PATH;$apps"
    Write-Log INFO ("Added to PATH for current process: {0}" -f $apps)
  }
}
function Get-WingetPath {
  $cmd = Get-Command winget -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  Ensure-PathHasWindowsApps
  $cmd = Get-Command winget -ErrorAction SilentlyContinue
  return ($cmd?.Source)
}

# -------------------- Winget runner ------------------
function Invoke-Winget {
  param([Parameter(Mandatory)][string[]]$Arguments)
  $exe = Get-WingetPath
  if (-not $exe) { throw "WinGet not available in this user session. Ensure 'App Installer' is installed/registered." }

  Write-Log INFO ("Executing: {0} {1}" -f $exe, ($Arguments -join ' '))
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName              = $exe
  $psi.Arguments             = ($Arguments -join ' ')
  $psi.RedirectStandardOutput= $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute       = $false
  $psi.CreateNoWindow        = $true
  $psi.WindowStyle           = [System.Diagnostics.ProcessWindowStyle]::Hidden
  $psi.ErrorDialog           = $false

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi
  [void]$proc.Start()

  if (-not $proc.WaitForExit($WingetTimeoutSec * 1000)) {
    try { $proc.Kill() } catch { }
    $out,$err = $proc.StandardOutput.ReadToEnd(), $proc.StandardError.ReadToEnd()
    if ($out) { foreach ($l in ($out -split "`r?`n")) { if ($l) { Write-Log INFO $l } } }
    if ($err) { foreach ($l in ($err -split "`r?`n")) { if ($l) { Write-Log WARN $l } } }
    return [pscustomobject]@{ ExitCode = -9999; StdOut = $out; StdErr = "winget timed out after $WingetTimeoutSec seconds." }
  }

  $stdout = $proc.StandardOutput.ReadToEnd()
  $stderr = $proc.StandardError.ReadToEnd()
  if ($stdout) { foreach ($l in ($stdout -split "`r?`n")) { if ($l) { Write-Log INFO $l } } }
  if ($stderr) { foreach ($l in ($stderr -split "`r?`n")) { if ($l) { Write-Log WARN $l } } }

  [pscustomobject]@{ ExitCode = $proc.ExitCode; StdOut = $stdout; StdErr = $stderr }
}

# -------------------- Source healing -----------------
function Heal-WingetSources {
  Write-Log INFO "Healing WinGet sources…"
  $u = Invoke-Winget -Arguments @('source','update','--name','winget','--disable-interactivity')
  if ($u.ExitCode -ne 0) {
    Write-Log WARN ("'winget' source update exited with {0}. Attempting 'source reset --force' then retry." -f $u.ExitCode)
    [void](Invoke-Winget -Arguments @('source','reset','--force','--disable-interactivity'))
    $u = Invoke-Winget -Arguments @('source','update','--name','winget','--disable-interactivity')
    if ($u.ExitCode -ne 0) { Write-Log WARN ("'winget' source still returned {0}. Continuing." -f $u.ExitCode) }
  }
  # Quietly remove msstore to avoid policy warnings; ignore failures
  [void](Invoke-Winget -Arguments @('source','remove','msstore','--disable-interactivity'))
}

# -------------------- Upgrade list parsing (text) ----
function Get-UpgradableIds {
  # Parse IDs from the standard table emitted by:
  # winget upgrade --source winget --include-unknown
  $res = Invoke-Winget -Arguments @('upgrade','--source','winget','--include-unknown','--disable-interactivity')
  $ids = @()
  if ($res.StdOut) {
    foreach ($line in ($res.StdOut -split "`r?`n")) {
      # Match: <Name><2+ spaces><Id><spaces><Installed><spaces><Available>
      if ($line -match '^\s*(?<name>.+?)\s{2,}(?<id>[A-Za-z0-9\.\-]+)\s+(?<installed>\S+)\s+(?<available>\S+)\s*$') {
        $ids += $Matches['id']
      }
    }
  }
  @($ids | Sort-Object -Unique)
}

# -------------------- Main --------------------------
try {
  Initialize-Logging
  Rotate-Log
  Write-Log INFO "=== PHASE: PREPARE ==="

  Ensure-PathHasWindowsApps

  Write-Log INFO "=== PHASE: VERIFY WinGet PRESENCE ==="
  $wgPath = Get-WingetPath
  if (-not $wgPath) { throw "WinGet not found. Install 'App Installer' from Microsoft Store or your software catalog." }
  Write-Log INFO ("WinGet found at: {0}" -f $wgPath)

  try {
    $ver = Invoke-Winget -Arguments @('--version')
    if ($ver.StdOut) { Write-Log INFO ("WinGet version: {0}" -f ($ver.StdOut.Trim())) }
  } catch { Write-Log WARN "Could not read winget version." }

  Write-Log INFO "=== PHASE: HEAL SOURCES ==="
  Heal-WingetSources

  Write-Log INFO "=== PHASE: PRE-UPGRADE SNAPSHOT ==="
  $preIds = @(Get-UpgradableIds)
  if ((Get-Count $preIds) -gt 0) {
    Write-Log INFO ("Upgrades detected (IDs): {0}" -f ($preIds -join ', '))
  } else {
    Write-Log INFO "No upgrades detected in pre-snapshot."
  }

  Write-Log INFO "=== PHASE: UPGRADE PACKAGES (BULK) ==="
  $bulkArgs = @(
    'upgrade','--all',
    '--source','winget',
    '--accept-package-agreements','--accept-source-agreements',
    '--silent','--disable-interactivity','--nowarn',
    '--authentication-mode','silent','--force'
  )
  if ($IncludePinned) { $bulkArgs += '--include-pinned' }
  $bulk = Invoke-Winget -Arguments $bulkArgs

  $doSecondPass = $false
  switch ($bulk.ExitCode) {
    0               { Write-Log INFO  "Bulk upgrade completed successfully." }
    -1978335189     { Write-Log INFO  "No applicable updates were found." }
    -1978335188     { Write-Log WARN  "Bulk upgrade reported partial failures (-1978335188); attempting per-package second pass."
                      $doSecondPass = $true }
    -9999           { Write-Log WARN  "Bulk upgrade timed out after $WingetTimeoutSec seconds; attempting per-package second pass."
                      $doSecondPass = $true }
    default         { Write-Log WARN  ("Bulk upgrade exit code {0}; attempting per-package second pass." -f $bulk.ExitCode)
                      $doSecondPass = $true }
  }

  if ($doSecondPass) {
    Write-Log INFO "=== PHASE: UPGRADE PACKAGES (SECOND PASS, PER-PACKAGE) ==="
    $pendingIds = @(Get-UpgradableIds)
    if ((Get-Count $pendingIds) -eq 0) {
      Write-Log INFO "Second pass: nothing pending."
    } else {
      foreach ($id in $pendingIds) {
        Write-Log INFO ("Second pass: upgrading {0}" -f $id)
        $args = @(
          'upgrade','--id', $id,'--exact',
          '--source','winget',
          '--accept-package-agreements','--accept-source-agreements',
          '--silent','--disable-interactivity','--nowarn',
          '--authentication-mode','silent','--force'
        )
        if ($IncludePinned) { $args += '--include-pinned' }
        $r = Invoke-Winget -Arguments $args
        if ($r.ExitCode -eq 0 -or $r.ExitCode -eq -1978335189) {
          Write-Log INFO ("Second pass: {0} completed." -f $id)
        } elseif ($r.ExitCode -eq -9999) {
          Write-Log WARN ("Second pass: {0} timed out after {1}s" -f $id, $WingetTimeoutSec)
        } else {
          Write-Log WARN ("Second pass: {0} failed with exit code {1}" -f $id, $r.ExitCode)
        }
      }
    }
  }

  Write-Log INFO "=== PHASE: POST-UPGRADE SNAPSHOT ==="
  $postIds = @(Get-UpgradableIds)
  if ((Get-Count $postIds) -gt 0) {
    Write-Log WARN ("Still pending after run: {0}" -f ($postIds -join ', '))
  } else {
    Write-Log INFO "No pending upgrades after run."
  }

  $nativeLogDir = "$env:LOCALAPPDATA\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir"
  Write-Log INFO ("Native WinGet logs folder: {0}" -f $nativeLogDir)

  Write-Log INFO "=== COMPLETED SUCCESSFULLY ==="
  exit 0
}
catch {
  Write-Log ERROR ("FATAL: {0}" -f $_.Exception.Message)
  exit 1
}
finally {
  if ($mutex) { try { $mutex.ReleaseMutex() | Out-Null } catch { } $mutex.Dispose() }
}
