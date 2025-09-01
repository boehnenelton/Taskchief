<#
.SYNOPSIS
    A menu-driven PowerShell script to create and manage scheduled jobs.
    This script allows a user to run other PowerShell scripts as an administrator after a specified time interval.
.DESCRIPTION
    This script provides a user interface to create, list, and delete scheduled tasks.
.NOTES
    Author: Elton Boehnen
    Author Email: Boehnenemail2024@gmail.com
    GitHub: github.com/boehnenelton
.VERSION
    1-0-26 (Revised for Robustness)
#>

#region Script Metadata & Configuration
$Script_Name = "TaskChief"
$Script_Version = "1-0-26" # Updated version
$Script_Description = "A menu-driven PowerShell script to create and manage scheduled jobs."
$Script_Type = "System_Script"
$Human_Friendly_Name = "Task Chief"
$Author = "Elton Boehnen"
$Author_Email = "Boehnenemail2024@gmail.com"
$Script_GUID = ""
#endregion

#region Path Definitions
$Script_Framework_Path = "C:\Users\Boehn\Cloud\OneDrive\Scripting_Framework"
$Log_Root_Path = "$Script_Framework_Path\Global_Logs\$Script_Name"
$Backup_Root_Path = "$Script_Framework_Path\Automated-Self_Backup"
$Script_Index_Path = "$Script_Framework_Path\Script_Index.csv"
$Data_Root_Path = "$PSScriptRoot\data\$Script_Name"
$MetaData_Path = "$Data_Root_Path\$Script_Name.meta"
$DebugFlagPath = "$PSScriptRoot\BEFlag.DEBUG"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$TranscriptPath = "$PSScriptRoot\Last_Transcript_$Timestamp.txt"
$JobDataPath = "$Data_Root_Path\timers"
$JobLogPath = "$JobDataPath\scheduled_tasks.json"

if (-not (Test-Path $Data_Root_Path)) { New-Item -ItemType Directory -Path $Data_Root_Path -Force | Out-Null }
if (-not (Test-Path $JobDataPath)) { New-Item -ItemType Directory -Path $JobDataPath -Force | Out-Null }
#endregion

#region COMPLIANCY FEATURES
[bool]$compliant_features = $true

if ($compliant_features) {
    try {
        Start-Transcript -Path $TranscriptPath -Append -Force | Out-Null
    } catch {
        Write-Host "Error starting transcript: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Error starting transcript: $($_.Exception.Message)"
    }

    Function Write-Log {
        Param (
            [string]$Message
        )
        try {
            if (-not (Test-Path $Log_Root_Path)) {
                New-Item -Path $Log_Root_Path -ItemType Directory -Force | Out-Null
            }
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] $Message"
            $logFilePath = Join-Path -Path $Log_Root_Path -ChildPath "$Script_Name.log"
            $logMessage | Out-File -FilePath $logFilePath -Append -Encoding utf8
            if (Test-Path $DebugFlagPath) {
                Write-Host $logMessage
            }
        } catch {
            Write-Host "Error logging: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Function Get-Script-Meta {
        if (Test-Path $MetaData_Path) {
            return (Get-Content -Path $MetaData_Path -Raw | ConvertFrom-Json)
        }
        return $null
    }

    Function Save-Script-Meta {
        Param (
            [PSCustomObject]$MetaArray
        )
        try {
            $MetaArray | ConvertTo-Json -Depth 5 | Out-File -FilePath $MetaData_Path -Encoding utf8 -Force
        } catch {
            Write-Log -Message "Error saving metadata: $($_.Exception.Message)"
        }
    }

    Function Get-Indexed-Version {
        if (Test-Path $Script_Index_Path) {
            $indexData = Import-Csv -Path $Script_Index_Path
            $scriptEntry = $indexData | Where-Object { $_.Script_Name -eq $Script_Name }
            if ($scriptEntry) {
                return $scriptEntry.Script_Version
            }
        }
        return $null
    }

    Function Update-Script-Index {
        Param (
            [PSCustomObject]$MetaArray
        )
        try {
            $indexData = if (Test-Path $Script_Index_Path) {
                Import-Csv -Path $Script_Index_Path
            } else {
                @()
            }
            
            $existingEntry = $indexData | Where-Object { $_.Script_Name -eq $Script_Name }
            
            $scriptFileName = if ($MyInvocation.MyCommand.Path) { (Get-Item -Path $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue).Name } else { "Unknown" }
            $scriptPath = if ($MyInvocation.MyCommand.Path) { (Get-Item -Path $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue).DirectoryName } else { "Unknown" }
            
            if ($existingEntry) {
                $indexData = $indexData | Where-Object { $_.Script_Name -ne $Script_Name }
                $indexData += [PSCustomObject]@{
                    Script_Name = $MetaArray.Script_Name
                    Script_Version = $MetaArray.Script_Version
                    Script_Description = $MetaArray.Script_Description
                    Script_FileName = $scriptFileName
                    Script_Path = $scriptPath
                    Script_GUID = $MetaArray.Script_GUID
                    Last_Version = $existingEntry.Script_Version
                }
            } else {
                $indexData += [PSCustomObject]@{
                    Script_Name = $MetaArray.Script_Name
                    Script_Version = $MetaArray.Script_Version
                    Script_Description = $MetaArray.Script_Description
                    Script_FileName = $scriptFileName
                    Script_Path = $scriptPath
                    Script_GUID = $MetaArray.Script_GUID
                    Last_Version = "N/A"
                }
            }
            
            $indexData | Export-Csv -Path $Script_Index_Path -NoTypeInformation -Force
            Write-Log -Message "Successfully updated script index."
        } catch {
            Write-Log -Message "Error updating script index: $($_.Exception.Message)"
        }
    }
    
    Function Backup-Script {
        Param (
            [string]$currentVersion,
            [string]$indexedVersion
        )
        
        if ($currentVersion -ne $indexedVersion) {
            try {
                if (-not $MyInvocation.MyCommand.Path) {
                    Write-Log -Message "Cannot create backup: Script path is null."
                    return
                }
                if (-not (Test-Path $Backup_Root_Path)) {
                    New-Item -ItemType Directory -Path $Backup_Root_Path -Force | Out-Null
                }
                $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
                $backupFileName = "$Human_Friendly_Name-($indexedVersion)-($timestamp).ps1"
                $backupPath = Join-Path -Path $Backup_Root_Path -ChildPath $backupFileName
                Copy-Item -Path $MyInvocation.MyCommand.Path -Destination $backupPath -Force -ErrorAction SilentlyContinue
                Write-Log -Message "Backup created: $backupFileName"
            } catch {
                Write-Log -Message "Error creating backup: $($_.Exception.Message)"
            }
        }
    }
    
    $scriptMeta = Get-Script-Meta
    if (-not $scriptMeta) {
        Write-Log -Message "First run detected, creating metadata."
        $Script_GUID = [guid]::NewGuid().ToString()
        $scriptMeta = [PSCustomObject]@{
            Script_Name = $Script_Name
            Script_Version = $Script_Version
            Script_Description = $Script_Description
            Script_Type = $Script_Type
            Human_Friendly_Name = $Human_Friendly_Name
            Script_GUID = $Script_GUID
            Author = $Author
            Author_Email = $Author_Email
        }
        Save-Script-Meta -MetaArray $scriptMeta
    } else {
        $Script_GUID = $scriptMeta.Script_GUID
    }
    
    $indexedVersion = Get-Indexed-Version
    if ($indexedVersion -ne $Script_Version) {
        Backup-Script -currentVersion $Script_Version -indexedVersion $indexedVersion
        Update-Script-Index -MetaArray $scriptMeta
    }
}
#endregion COMPLIANCY FEATURES

#region Job Log Management
$JobDataPath = "$PSScriptRoot\data\$Script_Name\timers"
$JobLogPath = "$JobDataPath\scheduled_tasks.json"

Function Read-Jobs {
    if (Test-Path $JobLogPath) {
        try {
            $jsonContent = Get-Content -Path $JobLogPath -Raw | ConvertFrom-Json
            if ($jsonContent.PSObject.Properties.Name -contains "BEJson_Scheduled_Tasks") {
                $tasks = $jsonContent.BEJson_Scheduled_Tasks.Indexed_Tasks_Data
                if ($null -eq $tasks) {
                    Write-Log -Message "Jobs file is empty or invalid, returning empty array."
                    return @()
                }
                $validTasks = $tasks | Where-Object { $_.JobName }
                Write-Log -Message "Read jobs: $($validTasks | ConvertTo-Json -Compress)"
                return ,@($validTasks)
            } else {
                Write-Log -Message "Invalid JSON structure: Missing BEJson_Scheduled_Tasks."
                return @()
            }
        } catch {
            Write-Log -Message "Error reading or parsing jobs from JSON: $($_.Exception.Message)"
            Write-Host "Error reading jobs file, starting with empty list." -ForegroundColor Yellow
            return @()
        }
    }
    Write-Log -Message "Jobs file does not exist, returning empty array."
    return @()
}

Function Write-Jobs {
    Param (
        [array]$Jobs
    )
    try {
        if (-not (Test-Path $JobDataPath)) {
            New-Item -Path $JobDataPath -ItemType Directory -Force | Out-Null
        }
        $jsonOutput = [PSCustomObject]@{
            "BEJson_Scheduled_Tasks" = @{
                "MetaData" = @{
                    "CreationDate" = (Get-Date).ToString('M-d-yy')
                    "CreationTime" = (Get-Date).ToString('HHmmss')
                    "Format_Version" = "V1-0-0"
                    "Data_Type" = "Scheduled_Tasks"
                    "Format" = "BEJson"
                    "Variant" = "Compatibility"
                }
                "Indexed_Tasks_Data" = @($Jobs)
            }
        }
        $jsonOutput | ConvertTo-Json -Depth 5 | Out-File -FilePath $JobLogPath -Encoding utf8 -Force
        Write-Log -Message "Successfully wrote jobs to JSON: $($Jobs | ConvertTo-Json -Compress)"
    } catch {
        Write-Log -Message "Error writing jobs to JSON: $($_.Exception.Message)"
        Write-Host "Error saving jobs file." -ForegroundColor Red
    }
}
#endregion

#region Cleanup Old Tasks
Function Cleanup-OldTasks {
    try {
        $jobs = Read-Jobs
        $currentTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "PS-Timer-Job-*" }
        Write-Log -Message "Found $($currentTasks.Count) tasks: $($currentTasks.TaskName -join ', ')"
        foreach ($task in $currentTasks) {
            if ($jobs.JobName -notcontains $task.TaskName) {
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log -Message "Cleaned up orphaned task: $($task.TaskName)"
            }
        }
    } catch {
        Write-Log -Message "Error cleaning up old tasks: $($_.Exception.Message)"
    }
}
#endregion

#region Main Menu Logic
Function Show-Menu {
    Clear-Host
    Write-Host "========================================="
    Write-Host "       PowerShell Timer Menu"
    Write-Host "========================================="
    Write-Host "1. Create a new timed script run"
    Write-Host "2. List and manage scheduled timers"
    Write-Host "3. Search for and delete scheduled tasks"
    Write-Host "4. Quit"
    Write-Host "Q. Quit (alternate)"
    Write-Host "========================================="
}

Function Create-TimedJob {
    try {
        # Verify running as admin
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Write-Host "This script must be run as Administrator to create scheduled tasks." -ForegroundColor Red
            Write-Log -Message "Script not running as Administrator, cannot create task."
            return
        }

        # Check Task Scheduler service
        $service = Get-Service -Name Schedule -ErrorAction SilentlyContinue
        if ($service.Status -ne 'Running') {
            Write-Host "Task Scheduler service is not running. Please start it." -ForegroundColor Red
            Write-Log -Message "Task Scheduler service not running: $($service.Status)"
            return
        }

        # Check user session state
        $sessionState = (qwinsta | Where-Object { $_ -like "*$env:USERNAME*" }).Trim() -split '\s+' | Select-Object -Index 3
        Write-Log -Message "User session state for $env:USERNAME: $sessionState"

        Add-Type -AssemblyName System.Windows.Forms
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.InitialDirectory = $PSScriptRoot
        $fileDialog.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
        
        Write-Host "Opening file selection dialog..."
        $dialogResult = $fileDialog.ShowDialog()

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $scriptToRun = $fileDialog.FileName
            Write-Host "Selected script: $scriptToRun"
        } else {
            Write-Host "File selection canceled. Returning to main menu." -ForegroundColor Yellow
            return
        }

        # Verify script exists
        if (-not (Test-Path $scriptToRun)) {
            Write-Host "Script file does not exist: $scriptToRun" -ForegroundColor Red
            Write-Log -Message "Script file does not exist: $scriptToRun"
            return
        }

        # Test script execution with timeout
        Write-Host "Testing script execution..." -ForegroundColor Cyan
        try {
            $job = Start-Job -ScriptBlock { & powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -NoExit -File $args[0] } -ArgumentList $scriptToRun
            $job | Wait-Job -Timeout 10 | Out-Null
            if ($job.State -eq 'Running') {
                Write-Host "Warning: Script test timed out after 10 seconds. Script may have an interactive loop." -ForegroundColor Yellow
                Write-Log -Message "Test execution of $scriptToRun timed out after 10 seconds."
                $job | Stop-Job
            } else {
                $testOutput = $job | Receive-Job
                Write-Log -Message "Test execution of $scriptToRun succeeded: $testOutput"
            }
            $job | Remove-Job
        } catch {
            Write-Host "Warning: Test execution of $scriptToRun failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Log -Message "Test execution of $scriptToRun failed: $($_.Exception.Message)"
        }

        Write-Host ""
        Write-Host "Select a timer interval:"
        Write-Host "  0) 1 Minute"
        Write-Host "  1) 10 Minutes"
        Write-Host "  2) 30 Minutes"
        Write-Host "  3) 1 Hour"
        Write-Host "  4) 3 Hours"
        Write-Host "  5) 6 Hours"
        Write-Host "  6) 12 Hours"

        $intervalChoice = Read-Host "Choice"
        $intervalMinutes = switch ($intervalChoice) {
            "0" { 1 }
            "1" { 10 }
            "2" { 30 }
            "3" { 60 }
            "4" { 180 }
            "5" { 360 }
            "6" { 720 }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                return
            }
        }
        
        $jobName = "PS-Timer-Job-$(Get-Random)"
        $delay = New-TimeSpan -Minutes $intervalMinutes
        $runTime = (Get-Date).Add($delay)

        $trigger = New-ScheduledTaskTrigger -Once -At $runTime
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -NoExit -File `"$scriptToRun`""
        
        # Try interactive first, fallback to SYSTEM
        try {
            $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)
            $registeredTask = Register-ScheduledTask -TaskName $jobName -Trigger $trigger -Action $action -Principal $principal -Settings $settings -Force
            Write-Log -Message "Created task '$jobName' as $env:USERNAME with LogonType Interactive"
        } catch {
            Write-Log -Message "Failed to create task as $env:USERNAME: $($_.Exception.Message). Falling back to SYSTEM."
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)
            $registeredTask = Register-ScheduledTask -TaskName $jobName -Trigger $trigger -Action $action -Principal $principal -Settings $settings -Force
            Write-Log -Message "Created task '$jobName' as SYSTEM"
        }
        
        # Verify task creation
        $task = Get-ScheduledTask -TaskName $jobName -ErrorAction SilentlyContinue
        if ($task) {
            Write-Log -Message "Created task '$jobName' to run at $runTime with state $($task.State)"
            Write-Host "Task '$jobName' scheduled to run at $runTime. To test, open Task Scheduler and run '$jobName' manually." -ForegroundColor Cyan
        } else {
            Write-Log -Message "Failed to verify task '$jobName' creation."
            Write-Host "Failed to verify task creation." -ForegroundColor Red
            return
        }

        $jobs = Read-Jobs
        if ($null -eq $jobs -or $jobs -isnot [array]) {
            Write-Log -Message "Jobs is null or not an array, initializing as empty array."
            $jobs = @()
        }
        $newJob = [PSCustomObject]@{
            JobName = $jobName
            ScriptPath = $scriptToRun
            CreationDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Interval = "$intervalMinutes Minutes"
        }
        
        $jobs += $newJob
        Write-Jobs -Jobs $jobs
        Write-Log -Message "Successfully created scheduled job '$jobName' to run '$scriptToRun' in $intervalMinutes minutes."

        Write-Host "Successfully created scheduled job '$jobName' to run '$scriptToRun' in $intervalMinutes minutes." -ForegroundColor Green
    } catch {
        Write-Host "An error occurred while creating the timed job: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Error creating timed job: $($_.Exception.Message)"
    }
}

Function List-And-Manage-Jobs {
    try {
        $jobs = Read-Jobs
        if ($jobs.Count -eq 0) {
            Write-Host "No scheduled timers found in the tracking file." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "Scheduled Timers:"
        Write-Host "-----------------"
        for ($i = 0; $i -lt $jobs.Count; $i++) {
            Write-Host "$($i + 1). Job Name: $($jobs[$i].JobName)"
            Write-Host "    Script Path: $($jobs[$i].ScriptPath)"
            Write-Host "    Interval: $($jobs[$i].Interval)"
            Write-Host "    Created: $($jobs[$i].CreationDate)"
            Write-Host ""
        }

        Write-Host "Enter the number of the job to delete, or '0' to return to the main menu."
        $choice = Read-Host "Choice"
        if ($choice -eq "0") {
            return
        }

        # IMPROVEMENT: Better input validation
        [int]$index = 0
        if (-not [int]::TryParse($choice, [ref]$index) -or ($index - 1) -lt 0 -or ($index - 1) -ge $jobs.Count) {
            Write-Host "Invalid choice. Please enter a number from the list." -ForegroundColor Red
            return
        }

        $jobToDelete = $jobs[$index - 1]

        # IMPROVEMENT: Add a confirmation prompt
        $confirmation = Read-Host "Are you sure you want to delete '$($jobToDelete.JobName)'? (y/n)"
        if ($confirmation.Trim().ToLower() -ne 'y') {
            Write-Host "Deletion canceled." -ForegroundColor Yellow
            return
        }
        
        # IMPROVEMENT: Check if the task still exists on the system before trying to delete
        $taskExists = Get-ScheduledTask -TaskName $jobToDelete.JobName -ErrorAction SilentlyContinue
        
        if ($taskExists) {
            try {
                Unregister-ScheduledTask -TaskName $jobToDelete.JobName -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully removed scheduled task '$($jobToDelete.JobName)' from the system." -ForegroundColor Green
                Write-Log -Message "Successfully unregistered scheduled task '$($jobToDelete.JobName)'."
            } catch {
                Write-Host "Error removing task from system: $($_.Exception.Message)" -ForegroundColor Red
                Write-Log -Message "Error unregistering task '$($jobToDelete.JobName)': $($_.Exception.Message)"
                return # Stop if system removal fails
            }
        } else {
            Write-Host "Task '$($jobToDelete.JobName)' not found on the system. It may have been deleted manually." -ForegroundColor Yellow
            Write-Log -Message "Task '$($jobToDelete.JobName)' not found on system during deletion attempt. Assuming it's already gone."
        }
        
        # Now that the system task is handled, update the local tracking file
        $jobs = $jobs | Where-Object { $_.JobName -ne $jobToDelete.JobName }
        Write-Jobs -Jobs $jobs
        Write-Host "Removed '$($jobToDelete.JobName)' from local tracking file."
        Write-Log -Message "Removed job '$($jobToDelete.JobName)' from JSON tracking file."
            
    } catch {
        Write-Host "An error occurred while managing jobs: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Error managing jobs: $($_.Exception.Message)"
    }
}

Function Search-And-Delete-Job {
    try {
        $query = Read-Host "Enter a search query for the task name"
        
        $matchingTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*$query*" }
        
        if (-not $matchingTasks) {
            Write-Host "No scheduled tasks found matching '$query'." -ForegroundColor Yellow
            return
        }

        Write-Host ""
        Write-Host "Matching Scheduled Tasks:"
        Write-Host "-------------------------"
        for ($i = 0; $i -lt $matchingTasks.Count; $i++) {
            Write-Host "$($i + 1). Task Name: $($matchingTasks[$i].TaskName)"
            Write-Host "    State: $($matchingTasks[$i].State)"
            Write-Host ""
        }
        
        Write-Host "Enter the number of the task to delete, or '0' to return to the main menu."
        $choice = Read-Host "Choice"
        if ($choice -eq "0") {
            return
        }
        
        # IMPROVEMENT: Better input validation
        [int]$index = 0
        if (-not [int]::TryParse($choice, [ref]$index) -or ($index - 1) -lt 0 -or ($index - 1) -ge $matchingTasks.Count) {
            Write-Host "Invalid choice. Please enter a number from the list." -ForegroundColor Red
            return
        }

        $taskToDelete = $matchingTasks[$index - 1]

        # IMPROVEMENT: Add a confirmation prompt
        $confirmation = Read-Host "Are you sure you want to delete '$($taskToDelete.TaskName)'? (y/n)"
        if ($confirmation.Trim().ToLower() -ne 'y') {
            Write-Host "Deletion canceled." -ForegroundColor Yellow
            return
        }
            
        try {
            # IMPROVEMENT: Use try/catch for robust error handling
            Unregister-ScheduledTask -TaskName $taskToDelete.TaskName -Confirm:$false -ErrorAction Stop
            Write-Host "Scheduled task '$($taskToDelete.TaskName)' has been deleted from the system." -ForegroundColor Green
            Write-Log -Message "Deleted scheduled task '$($taskToDelete.TaskName)'."
        } catch {
            Write-Host "Error removing task from system: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log -Message "Error unregistering task '$($taskToDelete.TaskName)': $($_.Exception.Message)"
            return # Exit if system deletion fails
        }
        
        # If system deletion was successful, update the local tracking file
        $jobs = Read-Jobs
        $updatedJobs = $jobs | Where-Object { $_.JobName -ne $taskToDelete.TaskName }
        
        if ($updatedJobs.Count -lt $jobs.Count) {
            Write-Jobs -Jobs $updatedJobs
            Write-Host "Removed '$($taskToDelete.TaskName)' from local tracking file."
            Write-Log -Message "Removed '$($taskToDelete.TaskName)' from local tracking file."
        }

    } catch {
        Write-Host "An error occurred while searching for tasks: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Error searching for and deleting tasks: $($_.Exception.Message)"
    }
}

# Run cleanup on startup
Cleanup-OldTasks

# Menu loop with explicit exit condition
$continue = $true
while ($continue) {
    Show-Menu
    $menuChoice = (Read-Host "Select an option").Trim().ToLower()
    switch ($menuChoice) {
        "1" { Create-TimedJob }
        "2" { List-And-Manage-Jobs }
        "3" { Search-And-Delete-Job }
        "4" { $continue = $false }
        "q" { $continue = $false }
        default { Write-Host "Invalid option. Please try again." -ForegroundColor Red }
    }
    if ($continue) {
        Write-Host "`nPress any key to continue..."
        $null = [System.Console]::ReadKey($true)
    }
}

#region Cleanup
try {
    Stop-Transcript -ErrorAction SilentlyContinue
    if (Test-Path $DebugFlagPath) {
        Remove-Item $DebugFlagPath -Force -ErrorAction SilentlyContinue
    }
} catch {
    Write-Host "Error during cleanup: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion