# ==========================
# Fix-OutlookStartup.ps1
# ==========================

$LogPath = "$env:LOCALAPPDATA\OutlookFixLogs"
if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Force -Path $LogPath | Out-Null }
$LogFile = "$LogPath\OutlookStartupFix_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

Function Write-Log {
    param([string]$Message)
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$Timestamp : $Message" | Out-File -Append -FilePath $LogFile
}

Write-Log "===== Starting Outlook Startup Troubleshooting ====="

# ---------------------------
# STEP 1: Check if Outlook is Running
# ---------------------------
Write-Log "Checking if Outlook is currently running..."
$proc = Get-Process outlook -ErrorAction SilentlyContinue

if ($proc) {
    Write-Log "Outlook is running. Attempting to stop Outlook..."
    Stop-Process -Name Outlook -Force
    Write-Log "Outlook terminated."
} else {
    Write-Log "Outlook is not running."
}

# ---------------------------
# STEP 2: Reset Navigation Pane
# ---------------------------
Write-Log "Resetting Outlook Navigation Pane..."
try {
    outlook.exe /resetnavpane
    Write-Log "Navigation Pane reset executed successfully."
} catch {
    Write-Log "Failed to reset Navigation Pane: $_"
}

# ---------------------------
# STEP 3: Clear Outlook Cache (OST, Autocomplete, Temp files)
# ---------------------------
$OutlookCachePath = "$env:LOCALAPPDATA\Microsoft\Outlook"
Write-Log "Outlook Cache Path: $OutlookCachePath"

if (Test-Path $OutlookCachePath) {
    Write-Log "Clearing OST & temporary cache files..."
    Get-ChildItem $OutlookCachePath -Include *.ost, *.dat, Stream_Autocomplete* -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                Remove-Item $_.FullName -Force
                Write-Log "Deleted: $($_.FullName)"
            } catch {
                Write-Log "Failed to delete $($_.FullName) :: $_"
            }
        }
    Write-Log "Cache cleanup completed."
} else {
    Write-Log "Cache directory does not exist."
}

# ---------------------------
# STEP 4: Repair PST if present using SCANPST
# ---------------------------
$ScanPstPaths = @(
    "$env:PROGRAMFILES\Microsoft Office\root\Office16\SCANPST.EXE",
    "$env:PROGRAMFILES(x86)\Microsoft Office\root\Office16\SCANPST.EXE"
)

$ScanTool = $ScanPstPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($ScanTool) {
    Write-Log "SCANPST found at: $ScanTool"
    $PSTFiles = Get-ChildItem "$env:USERPROFILE\Documents\Outlook Files" -Filter *.pst -ErrorAction SilentlyContinue

    foreach ($pst in $PSTFiles) {
        Write-Log "Running SCANPST on: $($pst.FullName)"
        Start-Process -FilePath $ScanTool -ArgumentList "`"$($pst.FullName)`"" -Wait
    }
} else {
    Write-Log "SCANPST not found – skipping PST repair."
}

# ---------------------------
# STEP 5: Reset Outlook Profile Registry Keys (Safe)
# ---------------------------
Write-Log "Resetting Profile registry keys under HKCU..."
$ProfileKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles",
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Search"
)

foreach ($key in $ProfileKeys) {
    if (Test-Path $key) {
        try {
            Remove-Item $key -Recurse -Force
            Write-Log "Deleted registry key: $key"
        } catch {
            Write-Log "Failed to delete registry key $key :: $_"
        }
    } else {
        Write-Log "Registry key not found: $key"
    }
}

# ---------------------------
# STEP 6: Clear Credential Manager Entries for Outlook/Office
# ---------------------------
Write-Log "Clearing cached credentials for Outlook..."
try {
    cmdkey /list | findstr /I "MicrosoftOffice" | ForEach-Object {
        $target = ($_ -split ':')[1].Trim()
        cmdkey /delete:$target
        Write-Log "Deleted Credential: $target"
    }
} catch {
    Write-Log "Failed to clear credentials: $_"
}

# ---------------------------
# STEP 7: Outlook Safe Mode Test Launch
# ---------------------------
Write-Log "Performing Safe Mode test launch..."
try {
    Start-Process "outlook.exe" -ArgumentList "/safe"
    Write-Log "Outlook Safe Mode launched for diagnostic test."
} catch {
    Write-Log "Failed to launch Outlook in Safe Mode: $_"
}

Write-Log "===== Outlook Troubleshooting Completed ====="
Write-Output "Logs are saved at: $LogFile"
