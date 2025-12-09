<#
.SYNOPSIS
  Outlook cleanup + logging script for "The information store could not be opened" error.

.DESCRIPTION
  - Creates detailed logs.
  - Kills Outlook process.
  - Runs outlook.exe /resetnavpane and /cleanviews.
  - Backups OST/PST files in common locations (and optionally deletes/renames OST).
  - Optionally disables COM Add-ins by setting LoadBehavior=0 in registry (user scope).
  - Collects Windows Application event logs & WER related to Outlook.

.PARAMETER LogPath
  Directory to store log files. Default: $env:USERPROFILE\OutlookCleanupLogs\<timestamp>

.PARAMETER BackupPath
  Where to backup OST/PST files. Defaults under LogPath\Backups

.PARAMETER DryRun
  If set, actions will only be reported and not performed (no deletions or registry changes).

.PARAMETER DeleteOrRenameOST
  None | Delete | Rename  (Rename will append .old to file)

.PARAMETER DisableAddins
  If specified, will set LoadBehavior to 0 for COM add-ins under HKCU registry keys.

.EXAMPLE
  .\OutlookCleanupWithLogs.ps1 -DryRun

#>

param(
    [string]$LogPath = "$env:USERPROFILE\OutlookCleanupLogs\$((Get-Date).ToString('yyyyMMdd_HHmmss'))",
    [string]$BackupPath = "",
    [switch]$DryRun = $true,
    [ValidateSet('None','Delete','Rename')]
    [string]$DeleteOrRenameOST = 'None',
    [switch]$DisableAddins = $false
)

function Write-Log {
    param($Message, [string]$Level='INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] $Message"
    $line | Tee-Object -FilePath $global:LogFile -Append
    Write-Host $line
}

# Init directories
if(-not $BackupPath) { $BackupPath = Join-Path -Path $LogPath -ChildPath 'Backups' }
New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
$global:LogFile = Join-Path -Path $LogPath -ChildPath 'OutlookCleanup.log'

Write-Log "Starting Outlook cleanup script. DryRun = $DryRun"
Write-Log "LogPath = $LogPath"
Write-Log "BackupPath = $BackupPath"
Write-Log "DeleteOrRenameOST = $DeleteOrRenameOST"
Write-Log "DisableAddins = $DisableAddins"

# Helper: run a command and capture output
function Run-Proc {
    param($File, $Args)
    try {
        Write-Log "Running: $File $Args"
        $proc = Start-Process -FilePath $File -ArgumentList $Args -NoNewWindow -PassThru -WindowStyle Hidden -ErrorAction Stop
        $proc.WaitForExit(30000) # wait 30s
        Write-Log "Process exited. ExitCode = $($proc.ExitCode)"
        return $proc.ExitCode
    } catch {
        Write-Log "Failed to run process: $_" 'ERROR'
        return $null
    }
}

# 1) Collect environment info
Write-Log "Collecting environment info..."
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber |
    Out-String | Write-Log

# Outlook executable detection (try user PATH, Program Files, Program Files (x86))
$possiblePaths = @()
# if Office click-to-run typical locations
$possiblePaths += "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE"
$possiblePaths += "$env:ProgramFiles(x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
$possiblePaths += "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE"
$possiblePaths += "$env:ProgramFiles(x86)\Microsoft Office\Office16\OUTLOOK.EXE"
$possiblePaths += "$env:ProgramFiles\Microsoft Office\root\Office15\OUTLOOK.EXE"
$possiblePaths += "$env:ProgramFiles(x86)\Microsoft Office\root\Office15\OUTLOOK.EXE"

$OutlookExe = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if(-not $OutlookExe) {
    # fallback to calling 'outlook.exe' (rely on PATH)
    $OutlookExe = 'outlook.exe'
    Write-Log "Outlook executable not found in known locations; will call 'outlook.exe' (requires PATH)."
} else {
    Write-Log "Found Outlook executable: $OutlookExe"
}

# 2) Kill Outlook if running
$ol = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if($ol) {
    Write-Log "Outlook processes found (count=$($ol.Count)). Will stop them."
    if(-not $DryRun) {
        $ol | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Write-Log "Stopped Outlook PID $($_.Id)" }
            catch { Write-Log "Failed to stop PID $($_.Id): $_" 'ERROR' }
        }
    } else {
        Write-Log "DryRun: would have stopped Outlook processes."
    }
} else {
    Write-Log "No running Outlook processes found."
}

# 3) Run /resetnavpane and /cleanviews
$cmds = @("/resetnavpane","/cleanviews")
foreach($c in $cmds) {
    if($DryRun) {
        Write-Log "DryRun: would run: $OutlookExe $c"
    } else {
        $rc = Run-Proc -File $OutlookExe -Args $c
        if($rc -eq $null) { Write-Log "Command $c may have failed to start." 'WARN' }
    }
}

# 4) Locate OST/PST files (common locations)
$ostLocations = @()
$ostLocations += "$env:LOCALAPPDATA\Microsoft\Outlook"
$ostLocations += "$env:USERPROFILE\Documents"  # PST might be here
$ostLocations = $ostLocations | Where-Object { Test-Path $_ } | Sort-Object -Unique
Write-Log "Searching OST/PST files in: $($ostLocations -join ', ')"

$foundFiles = @()
foreach($loc in $ostLocations) {
    try {
        $files = Get-ChildItem -Path $loc -Recurse -Include *.ost,*.pst -ErrorAction SilentlyContinue
        if($files) {
            $foundFiles += $files
        }
    } catch { Write-Log "Error scanning $loc : $_" 'WARN' }
}
if($foundFiles.Count -eq 0) {
    Write-Log "No OST/PST files found in common locations." 'WARN'
} else {
    Write-Log "Found $($foundFiles.Count) OST/PST files."
    foreach($f in $foundFiles) {
        Write-Log " - $($f.FullName) (Size: $([math]::Round($f.Length/1MB,2)) MB)"
        # Backup
        $destination = Join-Path -Path $BackupPath -ChildPath $f.Name
        $i = 1
        $origDest = $destination
        while(Test-Path $destination) {
            $destination = "$($origDest).$i"
            $i++
        }
        if($DryRun) {
            Write-Log "DryRun: would copy $($f.FullName) -> $destination"
        } else {
            try {
                Copy-Item -Path $f.FullName -Destination $destination -Force -ErrorAction Stop
                Write-Log "Backed up to $destination"
            } catch {
                Write-Log "Failed to backup $($f.FullName) -> $destination : $_" 'ERROR'
            }
        }

        # Delete or rename if requested and if file is .ost
        if(($DeleteOrRenameOST -ne 'None') -and ($f.Extension -ieq '.ost')) {
            if($DryRun) {
                Write-Log "DryRun: would $DeleteOrRenameOST OST file $($f.FullName)"
            } else {
                try {
                    if($DeleteOrRenameOST -eq 'Delete') {
                        Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                        Write-Log "Deleted OST $($f.FullName)"
                    } else {
                        $newName = $f.FullName + '.old'
                        Rename-Item -Path $f.FullName -NewName $newName -ErrorAction Stop
                        Write-Log "Renamed OST to $newName"
                    }
                } catch {
                    Write-Log "Failed to $DeleteOrRenameOST OST $($f.FullName) : $_" 'ERROR'
                }
            }
        }
    }
}

# 5) Collect Event Logs (Outlook/Office related) - last 7 days
Write-Log "Collecting Application event logs for Outlook/Office (last 7 days)..."
try {
    $now = Get-Date
    $since = $now.AddDays(-7)
    $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$since} -ErrorAction Stop |
                Where-Object {
                    ($_.ProviderName -match 'Outlook') -or
                    ($_.Message -match 'Outlook') -or
                    ($_.ProviderName -match 'Office') -or
                    ($_.Message -match 'Office')
                }
    $eventsPath = Join-Path -Path $LogPath -ChildPath 'Application_OutlookEvents.evtx'
    if($events) {
        # Export as text
        $events | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message |
            Out-File -FilePath (Join-Path $LogPath 'Application_OutlookEvents.txt') -Encoding UTF8
        Write-Log "Saved Application events to Application_OutlookEvents.txt"
    } else {
        Write-Log "No Outlook/Office-related application events found in last 7 days."
    }
} catch {
    Write-Log "Failed to collect Application event logs: $_" 'ERROR'
}

# 6) Collect Windows Error Reporting (WER) entries for Outlook
Write-Log "Collecting WER (Windows Error Reporting) for Outlook..."
try {
    $wer = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$since} -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match 'Faulting application name: OUTLOOK.EXE|Outlook.exe' }
    if($wer) {
        $wer | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message |
            Out-File -FilePath (Join-Path $LogPath 'WER_Outlook.txt') -Encoding UTF8
        Write-Log "Saved WER results to WER_Outlook.txt"
    } else {
        Write-Log "No WER entries for Outlook found."
    }
} catch {
    Write-Log "Failed to search WER entries: $_" 'WARN'
}

# 7) Optionally disable user COM add-ins (HKCU)
if($DisableAddins) {
    Write-Log "Attempting to enumerate and disable user COM add-ins under HKCU..."
    $addinKeyPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins"
    if(Test-Path $addinKeyPath) {
        $addinKeys = Get-ChildItem -Path $addinKeyPath -ErrorAction SilentlyContinue
        if($addinKeys.Count -eq 0) {
            Write-Log "No add-in keys found under $addinKeyPath"
        } else {
            foreach($k in $addinKeys) {
                Write-Log "Found add-in key: $($k.Name)"
                $lbPath = Join-Path $k.PSPath 'LoadBehavior'
                if($DryRun) {
                    Write-Log "DryRun: Would set LoadBehavior=0 for $($k.Name)"
                } else {
                    try {
                        Set-ItemProperty -Path $k.PSPath -Name LoadBehavior -Value 0 -ErrorAction Stop
                        Write-Log "Set LoadBehavior=0 for $($k.Name)"
                    } catch {
                        Write-Log "Failed to set LoadBehavior for $($k.Name) : $_" 'ERROR'
                    }
                }
            }
        }
    } else {
        Write-Log "Addin registry path not found: $addinKeyPath" 'WARN'
    }
}

# 8) Optionally collect Office repair guidance info and sfc
Write-Log "Collecting system file check (sfc) scan info (quick) - will only run in non-dry mode."
if(-not $DryRun) {
    try {
        Write-Log "Running sfc /scannow (this may take time) ..."
        # Run sfc quietly and capture result
        $sfcProc = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -PassThru
        Write-Log "sfc started (PID $($sfcProc.Id)). Please wait for it to complete manually or check later."
    } catch {
        Write-Log "Failed to start sfc: $_" 'WARN'
    }
} else {
    Write-Log "DryRun: skipping sfc /scannow"
}

# 9) Summary + Next steps
Write-Log "Cleanup script finished. Summary:"
Write-Log "  - Logs + backups: $LogPath"
Write-Log "  - Backups: $BackupPath"
if($DryRun) {
    Write-Log "DryRun was ON: no destructive changes were made. Re-run without -DryRun to apply changes." 'INFO'
} else {
    Write-Log "Non-dry run completed."
}

Write-Log "Recommended next steps:"
Write-Log "  1) Try launching Outlook now. If it still fails, run: `outlook.exe /safe` and test account access." 
Write-Log "  2) Check Application_OutlookEvents.txt and WER_Outlook.txt for specific crash messages (DLLs, faulting module names)." 
Write-Log "  3) If add-ins were disabled and Outlook opens, re-enable add-ins one-by-one to identify culprit." 
Write-Log "  4) If OST was renamed/deleted, opening Outlook will recreate the OST and resync. For PST backups, restore if needed."

# Print final location to console
Write-Host ""
Write-Host "=== OUTLOOK CLEANUP COMPLETE ==="
Write-Host "Logs and backups saved to: $LogPath"
Write-Host "If DryRun was used, re-run with -DryRun:$false to perform changes."
