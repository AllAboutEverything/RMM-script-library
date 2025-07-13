<#
.SYNOPSIS
    This script enables centralized deployment and scheduling of the DATEV Update Automation "DUPupd" via an RMM.
.DESCRIPTION
    1. Downloads DUPupd files from a server.
    2. Updates the "config.json" file with a license ID, ensuring the license ID is not publicly accessible on the server.
    3. Sets the system variable "DUPHOST" if it does not exist, extracted from the server's domain. This variable is the customer number visible on the website to identify the server's owner.
    4. Checks if the scheduled task named "START-DUPupd" exists, otherwise warns in the RMM; creation must be done by the admin user.
    5. Verifies if Autologon is enabled, otherwise warns in the RMM; Autologon must use the DATEV admin user.
    6. Sets the update time, passed as a parameter.
    7. Logs off users before the update on All-in-One servers.
.EXAMPLE
    Specify a date and time in the format 'DD.MM.YYYY HH:mm' via the RMM command line. If no update is desired, set a past date.
.NOTES
    Add your file server URL and license ID in line 30 and 32.
    Author: Samuel Keidel
    Last Modified: 05.06.2025
#>

# Parameter block to accept the new date and time for scheduling the update
param (
    [string]$NewDateTime  # Format: 'DD.MM.YYYY HH:mm'
)

# Variables 
$TaskName = "START-DUPupd"
$TaskFolder = "\GLSH\"
$targetFolder = "C:\admglsh\DUPupd"
$zipFilePath = "$targetFolder\DUPupd.zip"
$httpUrl = "http://<YourFileServer>.de/DUPupd/DUPupd.zip"
$configFilePath = "$targetFolder\config.json"
$licenseId = "<YourLicenseID>"
$autoLogonEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon").AutoAdminLogon -eq "1"
$IsAIOregPath = "HKLM:\Software\DUPupd\IsAIO" 

function Convert-DateTime {
    param (
        [string]$NewDateTime
    )

    try {
        $script:NewDateTimeParsed = [DateTime]::ParseExact($NewDateTime, "dd.MM.yyyy HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "invalid date/time format"
        exit 1001
    }
}

function IsAIO {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IsAIOregPath
    )
    try {
        New-Item -Path $IsAIOregPath -Force | Out-Null
        $serverName = $env:COMPUTERNAME
        $collections = Get-RDSessionCollection -ErrorAction SilentlyContinue | Select-Object -ExpandProperty CollectionName
        $isAIO = $false
        foreach ($collection in $collections) {
            $rdpHosts = Get-RDSessionHost -CollectionName $collection -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SessionHost
            foreach ($hostserver in $rdpHosts) {
                $hostShortName = $hostserver.Split('.')[0]
                if ($hostShortName -eq $serverName) {
                    $isAIO = $true
                    break
                }
            }
        }
        New-ItemProperty -Path $IsAIOregPath -Name "IsAIO" -Value $isAIO -PropertyType "String" -Force | Out-Null
        return $isAIO
    }
    catch {
        Write-Output "Error checking RDP Collection. $($_.Exception.Message)"
        New-ItemProperty -Path $IsAIOregPath -Name "IsAIO" -Value $false -PropertyType "String" -Force | Out-Null
        return $false
    }
}

function Get-TaskInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder,
        [Parameter(Mandatory=$true)]
        [string]$TaskName
    )
    $result = @{
        Exists = $false
        StartBoundary = $null
    }
    
    try {
        $task = Get-ScheduledTask -TaskPath $TaskFolder -TaskName $TaskName -ErrorAction Stop
        $result.Exists = $true
        $currentTrigger = ($task.Triggers | Where-Object { $_.StartBoundary -ne $null })
        if ($currentTrigger) {
            $result.StartBoundary = [DateTime]::Parse($currentTrigger.StartBoundary)
        }
        else {
            Write-Output "Kein Trigger in der Aufgabe '$TaskName' gefunden."
        }
    }
    catch {
        Write-Output "Aufgabe '$TaskName' nicht gefunden."
    }
    return $result
}

function Register-LogoffNonAdminsBeforeUpdate {
    param (
        [Parameter(Mandatory=$true)]
        [datetime]$NewDateTimeParsed,
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder
    )
    $LogOffDateTimeParsed = $NewDateTimeParsed.AddMinutes(-1)
    $LogoffNonAdminsBeforeUpdate = "LogoffNonAdminsBeforeUpdate"
    $TaskFullPath = "$TaskFolder$LogoffNonAdminsBeforeUpdate"
    $taskInfo = Get-TaskInfo -TaskFolder $TaskFolder -TaskName $LogoffNonAdminsBeforeUpdate
    $currentStartBoundary = $taskInfo.StartBoundary
    if ($currentStartBoundary -ne $LogOffDateTimeParsed) {
        try {
            $trigger = New-ScheduledTaskTrigger -Once -At $LogOffDateTimeParsed
            $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-Command "quser | Select-Object -Skip 1 | ForEach-Object { if ($_ -notmatch '' admin'') { $sessionId = ($_ -split ''\s+'')[2]; logoff $sessionId } }"'
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $LogoffNonAdminsBeforeUpdate -TaskPath $TaskFolder -Action $action -Trigger $trigger -Principal $principal -Force
            Write-Output "Neue Aufgabe '$LogoffNonAdminsBeforeUpdate' wurde erstellt und läuft um $LogOffDateTimeParsed."
        }
        catch {
            Write-Output "Fehler beim Erstellen der geplanten Aufgabe '$LogoffNonAdminsBeforeUpdate': $($_.Exception.Message)"
            exit 1001
        }
    }
    else {
        # Write-Output "Aufgabe '$LogoffNonAdminsBeforeUpdate' ist bereits für $LogOffDateTimeParsed geplant."
    }
}

function Update-DatevScheduledTask {
    param (
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder,
        [Parameter(Mandatory=$true)]
        [string]$TaskName,
        [Parameter(Mandatory=$true)]
        [datetime]$NewDateTimeParsed,
        [Parameter(Mandatory=$true)]
        [bool]$taskAvailable,
        [Parameter(Mandatory=$true)]
        [bool]$autoLogonEnabled
    )

    if ($autoLogonEnabled) {
        try {
            $trigger = New-ScheduledTaskTrigger -Once -At $NewDateTimeParsed
            $taskDefinition = Get-ScheduledTask -TaskPath $TaskFolder -TaskName $TaskName
            $taskDefinition.Triggers = @($trigger)
            Set-ScheduledTask -TaskPath $TaskFolder -TaskName $TaskName -Trigger $taskDefinition.Triggers
            Write-Output "Aufgabe $TaskName wurde auf $NewDateTimeParsed aktualisiert."
        }
        catch {
            Write-Output "Error: Fehler beim Aktualisieren der geplanten Aufgabe $TaskName."
            Write-Output $_.Exception.Message
            Assert-DUPupdPrerequisites -NewDateTimeParsed $NewDateTimeParsed -TaskFolder $TaskFolder -taskAvailable $taskAvailable -autoLogonEnabled $autoLogonEnabled
        }
        Assert-DUPupdPrerequisites -NewDateTimeParsed $NewDateTimeParsed -TaskFolder $TaskFolder -taskAvailable $taskAvailable -autoLogonEnabled $autoLogonEnabled
    }
    else {
        Assert-DUPupdPrerequisites -NewDateTimeParsed $NewDateTimeParsed -TaskFolder $TaskFolder -taskAvailable $taskAvailable -autoLogonEnabled $autoLogonEnabled
    }
}
function Get-HTTPLastModified {
    param (
        [string]$url
    )

    try {
        $request = [System.Net.HttpWebRequest]::Create($url)
        $request.Method = "HEAD"
        $response = $request.GetResponse()
        
        $lastModified = $response.LastModified
        $response.Close()
        
        if ($lastModified) {
            # Write-Host "LastModified: $lastModified" # This output is only distracting in LMS
            return $lastModified
        }
        else {
            Write-Warning "no LastModified in the header."
            return $null
        }
    }
    catch {
        Write-Error "Error getting LastModified: $_"
        return $null
    }
}

function Test-UpdateRequired {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ZipFilePath,
        [Parameter(Mandatory=$true)]
        [string]$TargetFolder,
        [Parameter(Mandatory=$true)]
        [datetime]$ServerLastModified
    )

    try {
        if (-not (Test-Path -Path $TargetFolder) -or 
            -not (Test-Path -Path $ZipFilePath) -or 
            ($ServerLastModified -gt (Get-Item $ZipFilePath -ErrorAction Stop).LastWriteTime)) {
            return $true
        }
        return $false
    }
    catch {
        Write-Output "Error checking if DUPupd files need to be downloaded or updated: $_"
        return $false
    }
}

function Update-DatevFiles {
    param (
        [string]$HttpUrl,
        [string]$TargetFolder,
        [string]$ConfigFilePath,
        [string]$LicenseId,
        [datetime]$LastModified,
        [string]$ZipFilePath
    )

    try {
        New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null
        Invoke-WebRequest -Uri $HttpUrl -OutFile $ZipFilePath -ErrorAction Stop
        Expand-Archive -Path $ZipFilePath -DestinationPath $TargetFolder -Force
        Write-Output "Files successfully downloaded and extracted to '$testFolder'."
    }
    catch {
        Write-Output "Error: Could not create folder '$TargetFolder' or download files."
        Write-Output $_.Exception.Message
        exit 1001
    }

    if (Test-Path -Path $ConfigFilePath) {
        try {
            $config = Get-Content -Path $ConfigFilePath | ConvertFrom-Json
            $config.licenseid = $LicenseId
            $config | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigFilePath
            Write-Output "License ID updated in '$ConfigFilePath'."
        }
        catch {
            Write-Output "Error: Could not update the license ID in '$ConfigFilePath'."
            Write-Output $_.Exception.Message
            exit 1001
        }
    }
    else {
        Write-Output "Error: Config file '$ConfigFilePath' not found."
        exit 1001
    }
}

function Test-AdminLoggedIn {
    try {
        $users = quser #| Select-Object -Skip 1
        
        foreach ($user in $users) {
            if ($user -match "admin") {
                Write-Output "Admin found."
                return $true
            }
        }
        Write-Output "No Admin user is currently logged in."
        return $false
    }
    catch {
        Write-Output "Error: looking for logged in users $($_.Exception.Message)"
        return $false
    }
}

function Register-RestartIfNoAdmin {
    param (
        [Parameter(Mandatory=$true)]
        [datetime]$NewDateTimeParsed,  # Time for scheduling restart
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder,           # Folder path for the scheduled task
        [Parameter(Mandatory=$true)]
        [bool]$taskAvailable,          # Indicates if the task exists
        [Parameter(Mandatory=$true)]
        [bool]$autoLogonEnabled       # Indicates if AutoLogon is enabled
    )

    $currentDate = Get-Date
    if ($currentDate.Date -eq $NewDateTimeParsed.Date) {
        $RestartServer = "RestartServer"
        $TaskFullPath = "$TaskFolder$RestartServer"

        # Check if task exists with today's trigger
        $taskInfo = Get-TaskInfo -TaskFolder $TaskFolder -TaskName $RestartServer
        if ($taskInfo.Exists -and $taskInfo.StartBoundary -and ($taskInfo.StartBoundary -eq $NewDateTimeParsed.AddMinutes(-4))) {
            Write-Output "Task '$TaskFullPath' is already scheduled for today at $($taskInfo.StartBoundary)."
            return
        }

        $isAdminLoggedIn = Test-AdminLoggedIn
        if (-not $isAdminLoggedIn) {
            # Restart time: 4 minutes before NewDateTimeParsed
            $RestartServerDateTime = $NewDateTimeParsed.AddMinutes(-4)

            # Create or update task
            Register-RestartServer -RestartServerDateTime $RestartServerDateTime `
                                  -TaskFolder $TaskFolder `
                                  -taskAvailable $taskAvailable `
                                  -autoLogonEnabled $autoLogonEnabled
            Write-Output "Restart task '$TaskFullPath' scheduled for $RestartServerDateTime."
        }
        else {
            Write-Output "Admin is logged in. Skipping restart task registration."
        }
    }
    else {
        # Write-Output "Current date does not match NewDateTimeParsed date. Skipping restart task registration." # This output is only distracting in LMS
    }
}

function Register-RestartServer {
    param (
        [Parameter(Mandatory=$true)]
        [datetime]$RestartServerDateTime,
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder,
        [Parameter(Mandatory=$true)]
        [bool]$taskAvailable,
        [Parameter(Mandatory=$true)]
        [bool]$autoLogonEnabled
    )

    $RestartServer = "RestartServer"
    $TaskFullPath = "$TaskFolder$RestartServer"

    try {
        $trigger = New-ScheduledTaskTrigger -Once -At $RestartServerDateTime
        $action = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /t 120 /c 'Server startet in 2 Minuten neu. Bitte jetzt abmelden."
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        Register-ScheduledTask -TaskName $RestartServer -TaskPath $TaskFolder -Action $action -Trigger $trigger -Principal $principal -Force
        Write-Output "Task '$TaskFullPath' created successfully for restart at $RestartServerDateTime."
    }
    catch {
        Write-Output "Error creating scheduled task '$TaskFullPath'. Please ensure an admin is logged into the server."
        Write-Output $_.Exception.Message
        exit 1001
    }
}

function Ensure-DUPHOSTVariable {
    $duphostSet = Test-Path Env:\DUPHOST

    if (-not $duphostSet) {
        try {
            $domainName = (Get-WmiObject Win32_ComputerSystem).Domain
            if ($domainName -match "\d{5}") {
                $duphostValue = "[$($matches[0])]"
                [System.Environment]::SetEnvironmentVariable("DUPHOST", $duphostValue, [System.EnvironmentVariableTarget]::Machine)
                $duphostSet = $true
                Write-Output "System variable 'DUPHOST' set."
            } else {
                # [System.Environment]::SetEnvironmentVariable("DUPHOST", $domainName, [System.EnvironmentVariableTarget]::Machine)
                $duphostSet = $false
                Write-Output "Error: Domain name does not contain 5 digits."
            }
        } catch {
            Write-Output "Error: Could not set the 'DUPHOST' system variable."
            Write-Output $_.Exception.Message
            $duphostSet = $false
        }
    }
    else {
        Write-Output "System variable 'DUPHOST' exists."
        $duphostSet = $true
    }

    return $duphostSet
}

function Assert-DUPupdPrerequisites {
    param (       
        [Parameter(Mandatory=$true)]
        [datetime]$NewDateTimeParsed,  # Time for scheduling restart
        [Parameter(Mandatory=$true)]
        [string]$TaskFolder,           # Folder path for the scheduled task
        [Parameter(Mandatory=$true)]
        [bool]$taskAvailable,          # Indicates if the task exists
        [Parameter(Mandatory=$true)]
        [bool]$autoLogonEnabled        # Indicates if AutoLogon is enabled
    )

    $duphostSet = Ensure-DUPHOSTVariable
    Register-RestartIfNoAdmin -NewDateTimeParsed $NewDateTimeParsed `
                              -TaskFolder $TaskFolder `
                              -taskAvailable $taskAvailable `
                              -autoLogonEnabled $autoLogonEnabled

    if (-not ($taskAvailable -and $autoLogonEnabled -and $duphostSet)) {
        if (-not $duphostSet) {
            Write-Output "System variable 'DUPHOST' is missing and could not be set automatically. Please set it manually."
        }

        if (-not $taskAvailable) {
            Write-Output "The Windows DUPupd task does not exist. Create the task by running the C:\admglsh\DUPupd\CreateDUPupdTask.ps1 script as an admin user on the server. The execution must be explicitly performed as Administrator, e.g., via Total Commander."
        }

        if (-not $autoLogonEnabled) {
            Write-Output "AutoLogon is not enabled. Enable AutoLogon on the server to use DUPupd."
        }
        Write-Output "exit 1001"
        exit 1001
    } else {
        Write-Output "DATEV update is scheduled."
        exit 0
    }
}

# Main script execution
try {
    Convert-DateTime -NewDateTime $NewDateTime
}
catch {
    Write-Output "Error: Invalid date/time format. Please use 'DD.MM.YYYY HH:mm'."
    exit 1001
}

# Check if DUPupd files need to be downloaded or updated
try {
    $lastModified = Get-HTTPLastModified -url $httpUrl
    if ($lastModified -and (Test-UpdateRequired -ZipFilePath $ZipFilePath -TargetFolder $targetFolder -ServerLastModified $lastModified)) {
        Update-DatevFiles -HttpUrl $httpUrl -TargetFolder $targetFolder -ConfigFilePath $configFilePath -LicenseId $licenseId -LastModified $lastModified -ZipFilePath $zipFilePath
    }
}
catch {
    Write-Output "Error checking if DUPupd files need to be downloaded or updated: $_"
}

# Check if server is an AIO server and register logoff task for non-admin users if necessary
try {
    if (Test-Path -Path $IsAIOregPath) {
        $IsAIO = (Get-ItemProperty -Path $IsAIOregPath -Name "IsAIO" -ErrorAction SilentlyContinue).IsAIO -eq "True"
        if ($IsAIO) {
            # Write-Output "This is an AIO server. Registering logoff task for non-admin users before update."
            Register-LogoffNonAdminsBeforeUpdate -NewDateTimeParsed $NewDateTimeParsed -TaskFolder $TaskFolder
        }
    }
    else {
        Write-Output "Checking if this is an AIO server."
        IsAIO -IsAIOregPath $IsAIOregPath
    }
}
catch {
    Write-Output "Error while AIO handling."
}

# Check if the scheduled DUPupd task exists and update it if necessary
try {
    $taskInfo = Get-TaskInfo -TaskFolder $TaskFolder -TaskName $TaskName
    $taskAvailable = $taskInfo.Exists
    $currentStartBoundary = $taskInfo.StartBoundary
    if ($currentStartBoundary -ne $NewDateTimeParsed) {
        try {
        Update-DatevScheduledTask -TaskFolder $TaskFolder -TaskName $TaskName -NewDateTimeParsed $NewDateTimeParsed -taskAvailable $taskAvailable -autoLogonEnabled $autoLogonEnabled
        }
        catch {
            Write-Output "Error updating the scheduled task: $($_.Exception.Message)"
        }
    }
}
catch {
    Write-Output "Error checking or updating the scheduled task."
}

# Ensure prerequisites are met before exiting
Assert-DUPupdPrerequisites -NewDateTimeParsed $NewDateTimeParsed -TaskFolder $TaskFolder -taskAvailable $taskAvailable -autoLogonEnabled $autoLogonEnabled
