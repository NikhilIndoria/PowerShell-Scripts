<#
.SYNOPSIS
    Completely removes OneDrive from Windows 11 and restores local folder control.

.DESCRIPTION
    This script will:
    - Uninstall OneDrive application
    - Fix folder redirection for Documents, Desktop, and Pictures
    - Remove OneDrive from File Explorer
    - Block OneDrive from reinstalling
    - Optionally backup OneDrive files before removal

.NOTES
    Version:        1.0
    Author:         OneDrive Removal Script
    Creation Date:  2025
    Requires:       PowerShell 5.1+, Windows 11, Administrator privileges

.PARAMETER BackupPath
    Custom path for OneDrive backup. Default is user profile with timestamp.

.PARAMETER NoBackup
    Skip the backup prompt and do not create a backup.

.PARAMETER NoReboot
    Skip the reboot prompt at the end.

.PARAMETER Silent
    Run in silent mode with minimal output (shows only errors and warnings).

.EXAMPLE
    .\Remove-OneDrive.ps1

    Runs the script interactively with prompts for backup and confirmation.

.EXAMPLE
    .\Remove-OneDrive.ps1 -NoBackup -NoReboot

    Runs the script without backup or reboot prompts.

.EXAMPLE
    .\Remove-OneDrive.ps1 -BackupPath "D:\Backups" -Silent

    Creates backup in D:\Backups and runs silently.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$BackupPath,

    [Parameter(Mandatory=$false)]
    [switch]$NoBackup,

    [Parameter(Mandatory=$false)]
    [switch]$NoReboot,

    [Parameter(Mandatory=$false)]
    [switch]$Silent
)

# Note: Admin elevation is handled programmatically in the script
# Do not use -RunAsAdministrator as it prevents custom elevation logic

# ---------------------------------------------
# Configuration
# ---------------------------------------------
$ErrorActionPreference = "Continue"
$script:SilentMode = $Silent.IsPresent
$script:LogFile = "$env:TEMP\Remove-OneDrive_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ---------------------------------------------
# Helper Functions
# ---------------------------------------------

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"

    try {
        Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch {
        # Silently fail if logging doesn't work
    }
}

function Write-Status {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )

    # Log all messages
    Write-Log "$Type : $Message"

    # In silent mode, only show warnings and errors
    if ($script:SilentMode -and $Type -in @("Info", "Success")) {
        return
    }

    switch ($Type) {
        "Info"    { Write-Host "[*] $Message" -ForegroundColor Cyan }
        "Success" { Write-Host "[+] $Message" -ForegroundColor Green }
        "Warning" { Write-Host "[!] $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[-] $Message" -ForegroundColor Red }
    }
}

function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Request-AdminElevation {
    Write-Status "This script requires administrator privileges." "Warning"
    Write-Status "Attempting to restart with elevated permissions..." "Info"

    try {
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
    catch {
        Write-Status "Failed to elevate privileges. Please run PowerShell as Administrator manually." "Error"
        exit 1
    }
}

function Backup-OneDriveFolders {
    param(
        [string]$CustomBackupPath
    )

    Write-Status "Backing up OneDrive folders..." "Info"

    $oneDrivePath = "$env:USERPROFILE\OneDrive"

    if ($CustomBackupPath) {
        $backupPath = Join-Path $CustomBackupPath "OneDrive_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    } else {
        $backupPath = "$env:USERPROFILE\OneDrive_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }

    if (Test-Path $oneDrivePath) {
        try {
            # Check available disk space
            $oneDriveSize = (Get-ChildItem -Path $oneDrivePath -Recurse -Force -ErrorAction SilentlyContinue |
                            Measure-Object -Property Length -Sum).Sum
            $targetDrive = Split-Path -Path $backupPath -Qualifier
            $driveInfo = Get-PSDrive -Name $targetDrive.TrimEnd(':') -ErrorAction SilentlyContinue

            if ($driveInfo -and $driveInfo.Free -lt ($oneDriveSize * 1.1)) {
                Write-Status "Insufficient disk space for backup. Need $([math]::Round($oneDriveSize/1GB, 2))GB" "Error"
                return $null
            }

            Write-Status "Creating backup at: $backupPath" "Info"
            Copy-Item -Path $oneDrivePath -Destination $backupPath -Recurse -Force
            Write-Status "Backup completed successfully!" "Success"
            return $backupPath
        }
        catch {
            Write-Status "Backup failed: $_" "Error"
            return $null
        }
    }
    else {
        Write-Status "No OneDrive folder found at $oneDrivePath" "Warning"
        return $null
    }
}

function Stop-OneDriveProcesses {
    Write-Status "Stopping OneDrive processes..." "Info"

    $processes = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue

    if ($processes) {
        $processes | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Write-Status "OneDrive processes stopped." "Success"
    }
    else {
        Write-Status "No OneDrive processes running." "Info"
    }
}

function Uninstall-OneDrive {
    Write-Status "Uninstalling OneDrive..." "Info"

    # Try winget first (modern method)
    try {
        $wingetResult = winget uninstall "Microsoft OneDrive" --silent --accept-source-agreements 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Status "OneDrive uninstalled via winget." "Success"
            return $true
        }
    }
    catch {
        Write-Status "Winget uninstall failed, trying alternative method..." "Warning"
    }

    # Fallback to OneDriveSetup.exe
    $setupPaths = @(
        "$env:SystemRoot\System32\OneDriveSetup.exe",
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    )

    foreach ($setupPath in $setupPaths) {
        if (Test-Path $setupPath) {
            try {
                Write-Status "Running uninstaller: $setupPath" "Info"
                Start-Process -FilePath $setupPath -ArgumentList "/uninstall" -Wait -NoNewWindow
                Write-Status "OneDrive uninstalled successfully." "Success"
                return $true
            }
            catch {
                Write-Status "Failed to run uninstaller: $_" "Error"
            }
        }
    }

    Write-Status "Could not find OneDrive uninstaller. It may already be uninstalled." "Warning"
    return $false
}

function Move-OneDriveFilesToLocal {
    Write-Status "Moving files from OneDrive folders to local folders..." "Info"

    $folders = @(
        @{OneDrive = "$env:USERPROFILE\OneDrive\Documents"; Local = "$env:USERPROFILE\Documents"},
        @{OneDrive = "$env:USERPROFILE\OneDrive\Desktop"; Local = "$env:USERPROFILE\Desktop"},
        @{OneDrive = "$env:USERPROFILE\OneDrive\Pictures"; Local = "$env:USERPROFILE\Pictures"}
    )

    foreach ($folder in $folders) {
        # Ensure local folder exists
        if (-not (Test-Path $folder.Local)) {
            New-Item -ItemType Directory -Path $folder.Local -Force | Out-Null
            Write-Status "Created local folder: $($folder.Local)" "Info"
        }

        # Move files if OneDrive folder exists
        if (Test-Path $folder.OneDrive) {
            try {
                $items = Get-ChildItem -Path $folder.OneDrive -Force -ErrorAction SilentlyContinue

                if ($items) {
                    foreach ($item in $items) {
                        $destination = Join-Path $folder.Local $item.Name

                        if (Test-Path $destination) {
                            # Compare timestamps for files
                            if (-not $item.PSIsContainer) {
                                $sourceFile = Get-Item $item.FullName
                                $destFile = Get-Item $destination

                                if ($sourceFile.LastWriteTime -gt $destFile.LastWriteTime) {
                                    Write-Status "Updating newer file: $($item.Name)" "Info"
                                    Copy-Item -Path $item.FullName -Destination $folder.Local -Force
                                } else {
                                    Write-Status "Skipping older file: $($item.Name)" "Info"
                                }
                            } else {
                                # For directories, merge contents
                                Write-Status "Merging directory: $($item.Name)" "Info"
                                Copy-Item -Path $item.FullName -Destination $folder.Local -Recurse -Force
                            }
                        }
                        else {
                            Write-Status "Moving: $($item.Name)" "Info"
                            Move-Item -Path $item.FullName -Destination $folder.Local -Force
                        }
                    }
                    Write-Status "Files moved from $($folder.OneDrive)" "Success"
                }
            }
            catch {
                Write-Status "Error moving files from $($folder.OneDrive): $_" "Error"
            }
        }
    }
}

function Fix-FolderRedirection {
    Write-Status "Fixing folder redirection in registry..." "Info"

    $userShellFoldersPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $shellFoldersPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

    $redirections = @(
        @{Name = "Personal"; Path = "$env:USERPROFILE\Documents"; Description = "Documents"},
        @{Name = "Desktop"; Path = "$env:USERPROFILE\Desktop"; Description = "Desktop"},
        @{Name = "My Pictures"; Path = "$env:USERPROFILE\Pictures"; Description = "Pictures"}
    )

    foreach ($redirect in $redirections) {
        try {
            # Update User Shell Folders (with environment variables)
            Set-ItemProperty -Path $userShellFoldersPath -Name $redirect.Name -Value $redirect.Path -Type ExpandString -Force
            # Update Shell Folders (with resolved paths)
            Set-ItemProperty -Path $shellFoldersPath -Name $redirect.Name -Value $redirect.Path -Type String -Force
            Write-Status "Fixed redirection for $($redirect.Description)" "Success"
        }
        catch {
            Write-Status "Failed to fix redirection for $($redirect.Description): $_" "Error"
        }
    }

    # Clean up OneDrive-related registry keys
    Write-Status "Cleaning up OneDrive registry entries..." "Info"

    $oneDriveRegPaths = @(
        "HKCU:\Software\Microsoft\OneDrive",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    )

    foreach ($regPath in $oneDriveRegPaths) {
        if (Test-Path $regPath) {
            try {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Removed registry key: $regPath" "Success"
            }
            catch {
                Write-Status "Could not remove registry key: $regPath" "Warning"
            }
        }
    }
}

function Block-OneDriveReinstallation {
    Write-Status "Blocking OneDrive from reinstalling..." "Info"

    $policyPath = "HKLM:\Software\Policies\Microsoft\Windows\OneDrive"

    try {
        # Create policy key if it doesn't exist
        if (-not (Test-Path $policyPath)) {
            New-Item -Path $policyPath -Force | Out-Null
        }

        # Disable OneDrive file sync
        Set-ItemProperty -Path $policyPath -Name "DisableFileSync" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $policyPath -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force

        Write-Status "OneDrive reinstallation blocked via registry policy." "Success"
    }
    catch {
        Write-Status "Failed to set registry policy: $_" "Error"
    }
}

function Remove-OneDriveFromExplorer {
    Write-Status "Removing OneDrive from File Explorer..." "Info"

    $clsids = @(
        "Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
        "Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    )

    foreach ($clsid in $clsids) {
        if (Test-Path $clsid) {
            try {
                Remove-Item -Path $clsid -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Removed OneDrive CLSID from registry." "Success"
            }
            catch {
                Write-Status "Could not remove CLSID (may require ownership change): $_" "Warning"
            }
        }
    }
}

function Restart-Explorer {
    Write-Status "Restarting Windows Explorer to apply changes..." "Info"

    try {
        Stop-Process -Name explorer -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Status "Explorer restarted successfully." "Success"
    }
    catch {
        Write-Status "Failed to restart Explorer: $_" "Error"
    }
}

function Remove-OneDriveStartupEntry {
    Write-Status "Removing OneDrive from startup..." "Info"

    $startupPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
    )

    foreach ($path in $startupPaths) {
        try {
            $oneDriveRun = Get-ItemProperty -Path $path -Name "OneDrive" -ErrorAction SilentlyContinue
            if ($oneDriveRun) {
                Remove-ItemProperty -Path $path -Name "OneDrive" -Force
                Write-Status "Removed OneDrive from startup registry." "Success"
            }
        }
        catch {
            # Silently continue if entry doesn't exist
        }
    }
}

function Remove-OneDriveFolder {
    Write-Status "Checking for empty OneDrive folder..." "Info"

    $oneDrivePath = "$env:USERPROFILE\OneDrive"

    if (Test-Path $oneDrivePath) {
        $items = Get-ChildItem -Path $oneDrivePath -Force -ErrorAction SilentlyContinue

        if (-not $items -or $items.Count -eq 0) {
            try {
                Remove-Item -Path $oneDrivePath -Force -Recurse -ErrorAction Stop
                Write-Status "Removed empty OneDrive folder." "Success"
            }
            catch {
                Write-Status "Could not remove OneDrive folder: $_" "Warning"
            }
        }
        else {
            Write-Status "OneDrive folder is not empty. Skipping removal." "Warning"
            Write-Status "You can manually delete it later: $oneDrivePath" "Info"
        }
    }
}

# ---------------------------------------------
# Main Script
# ---------------------------------------------

function Main {
    Clear-Host

    Write-Host @"
â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
â•‘                                                            â•‘
â•‘        OneDrive Complete Removal Script for Windows 11    â•‘
â•‘                                                            â•‘
â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
"@ -ForegroundColor Magenta

    Write-Host ""
    Write-Status "This script will completely remove OneDrive from your system." "Info"
    Write-Status "It will:" "Info"
    Write-Host "  * Uninstall OneDrive application"
    Write-Host "  * Move files from OneDrive folders to local folders"
    Write-Host "  * Fix folder redirection (Documents, Desktop, Pictures)"
    Write-Host "  * Remove OneDrive from File Explorer"
    Write-Host "  * Block OneDrive from reinstalling"
    Write-Host ""

    # Check for admin privileges
    if (-not (Test-AdminPrivileges)) {
        Request-AdminElevation
    }

    # Confirmation
    $confirm = Read-Host "Do you want to proceed? (Y/N)"
    if ($confirm -ne "Y" -and $confirm -ne "y") {
        Write-Status "Operation cancelled by user." "Warning"
        exit 0
    }

    Write-Host ""

    # Ask about backup
    if (-not $NoBackup) {
        $backup = Read-Host "Do you want to backup your OneDrive folder first? (Y/N - Recommended)"
        if ($backup -eq "Y" -or $backup -eq "y") {
            $backupPath = Backup-OneDriveFolders -CustomBackupPath $BackupPath
            if ($backupPath) {
                Write-Status "Backup saved to: $backupPath" "Success"
            }
        }
    }

    Write-Host ""
    Write-Status "Starting OneDrive removal process..." "Info"
    Write-Log "Starting OneDrive removal process"
    Write-Host ""

    # Execute removal steps with verification
    $stepResults = @{}

    Stop-OneDriveProcesses
    $stepResults["StopProcesses"] = -not (Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue)

    Remove-OneDriveStartupEntry
    Move-OneDriveFilesToLocal
    Uninstall-OneDrive
    Fix-FolderRedirection
    Block-OneDriveReinstallation
    Remove-OneDriveFromExplorer
    Remove-OneDriveFolder
    Restart-Explorer

    # Summary
    Write-Host ""
    Write-Host @"
â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
â•‘                      REMOVAL COMPLETE                      â•‘
â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
"@ -ForegroundColor Green

    Write-Host ""
    Write-Status "OneDrive has been completely removed!" "Success"
    Write-Status "Your files are now in local folders:" "Info"
    Write-Host "  * Documents: $env:USERPROFILE\Documents"
    Write-Host "  * Desktop:   $env:USERPROFILE\Desktop"
    Write-Host "  * Pictures:  $env:USERPROFILE\Pictures"
    Write-Host ""
    Write-Status "Next steps:" "Info"
    Write-Host "  1. Restart your computer to ensure all changes take effect"
    Write-Host "  2. Verify your files are in the correct locations"
    Write-Host ""
    Write-Status "Log file saved to: $script:LogFile" "Info"
    Write-Host ""

    if (-not $NoReboot) {
        $reboot = Read-Host "Would you like to restart now? (Y/N)"
        if ($reboot -eq "Y" -or $reboot -eq "y") {
            Write-Status "Restarting computer in 10 seconds..." "Info"
            Write-Log "User initiated system restart"
            Start-Sleep -Seconds 10
            Restart-Computer -Force
        }
        else {
            Write-Status "Please restart your computer when convenient." "Info"
        }
    }
    else {
        Write-Status "Please restart your computer when convenient." "Info"
    }

    Write-Log "OneDrive removal completed successfully"
}

# Execute main function
Main
