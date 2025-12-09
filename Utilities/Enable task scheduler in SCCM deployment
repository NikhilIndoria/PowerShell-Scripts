<#
.SYNOPSIS
    Enables and configures Task Scheduler service for SCCM application deployment.

.DESCRIPTION
    This script ensures the Task Scheduler service is enabled, started, and configured
    for automatic startup. It includes error handling and logging for SCCM deployments.

.NOTES
    Author: SCCM Admin
    Date: 2025
    Compatible with: Windows 7/8/10/11, Server 2012+
    SCCM Exit Codes: 0 = Success, 1 = Failure
#>

# Variables
$ServiceName = "Schedule"
$LogPath = "$env:SystemRoot\Temp\TaskScheduler_Enable.log"
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Function to write log entries
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $LogEntry = "$Timestamp - [$Level] - $Message"
    Add-Content -Path $LogPath -Value $LogEntry
    Write-Host $LogEntry
}

# Start logging
Write-Log "========================================"
Write-Log "Task Scheduler Enable Script Started"
Write-Log "========================================"

try {
    # Check if running with admin privileges
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $IsAdmin) {
        Write-Log "Script must be run with administrative privileges" "ERROR"
        exit 1
    }
    
    Write-Log "Running with administrative privileges"

    # Get Task Scheduler service
    $TaskSchedulerService = Get-Service -Name $ServiceName -ErrorAction Stop
    Write-Log "Current Task Scheduler Status: $($TaskSchedulerService.Status)"
    Write-Log "Current Startup Type: $($TaskSchedulerService.StartType)"

    # Set service to Automatic startup
    Write-Log "Setting Task Scheduler to Automatic startup..."
    Set-Service -Name $ServiceName -StartupType Automatic -ErrorAction Stop
    Write-Log "Task Scheduler startup type set to Automatic" "SUCCESS"

    # Start the service if it's not running
    if ($TaskSchedulerService.Status -ne "Running") {
        Write-Log "Starting Task Scheduler service..."
        Start-Service -Name $ServiceName -ErrorAction Stop
        
        # Wait for service to start (max 30 seconds)
        $Timeout = 30
        $Timer = [Diagnostics.Stopwatch]::StartNew()
        
        while ((Get-Service -Name $ServiceName).Status -ne "Running" -and $Timer.Elapsed.TotalSeconds -lt $Timeout) {
            Start-Sleep -Seconds 1
        }
        
        $Timer.Stop()
        
        $FinalStatus = (Get-Service -Name $ServiceName).Status
        if ($FinalStatus -eq "Running") {
            Write-Log "Task Scheduler service started successfully" "SUCCESS"
        } else {
            Write-Log "Task Scheduler service failed to start within timeout period" "ERROR"
            exit 1
        }
    } else {
        Write-Log "Task Scheduler service is already running" "SUCCESS"
    }

    # Verify service is responsive
    Write-Log "Verifying Task Scheduler is responsive..."
    $TaskSchedulerCOM = New-Object -ComObject Schedule.Service
    $TaskSchedulerCOM.Connect()
    Write-Log "Task Scheduler COM interface is responsive" "SUCCESS"

    # Final verification
    $FinalService = Get-Service -Name $ServiceName
    Write-Log "========================================"
    Write-Log "Final Status: $($FinalService.Status)"
    Write-Log "Final Startup Type: $($FinalService.StartType)"
    Write-Log "========================================"
    Write-Log "Task Scheduler enabled successfully" "SUCCESS"
    
    exit 0

} catch {
    Write-Log "Error occurred: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
