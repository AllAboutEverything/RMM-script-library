# RMM script to monitor which software is installed and alert if something new is installed
# The script ignores version numbers of installed software, so the script does not alert every time a software updates

# The script will alert for up to this many days about new software
param (
    [int]$Days = 7  # Default to 7 days if no argument is provided
)

# Define the file paths to store the list of already installed software and the new software log
$installedSoftwareFilePath = "C:\ProgramData\InstalledSoftwareList.json"
$newSoftwareLogPath = "C:\ProgramData\NewSoftwareLog.json"

# Function to get the list of installed software
function Get-InstalledSoftware {
    $softwareList = Get-WmiObject -Class Win32_Product | Select-Object -Property Name
    return $softwareList
}

# Function to load previously saved software list from file
function Load-PreviousSoftwareList {
    if (Test-Path -Path $installedSoftwareFilePath) {
        $previousSoftwareList = Get-Content -Path $installedSoftwareFilePath | ConvertFrom-Json
        return $previousSoftwareList
    } else {
        return @() # Return an empty array if the file doesn't exist
    }
}

# Function to load new software log
function Load-NewSoftwareLog {
    if (Test-Path -Path $newSoftwareLogPath) {
        $newSoftwareLog = Get-Content -Path $newSoftwareLogPath | ConvertFrom-Json
        return $newSoftwareLog
    } else {
        return @() # Return an empty array if the file doesn't exist
    }
}

# Function to save new software log to file
function Save-NewSoftwareLog {
    param (
        [Parameter(Mandatory=$true)]
        [array]$newSoftwareLog
    )
    $newSoftwareLog | ConvertTo-Json | Set-Content -Path $newSoftwareLogPath
}

# Function to save the current software list to file
function Save-CurrentSoftwareList {
    param (
        [Parameter(Mandatory=$true)]
        [array]$softwareList
    )
    $softwareList | ConvertTo-Json | Set-Content -Path $installedSoftwareFilePath
}

# Function to compare software lists and generate alert if new software is found
function Compare-SoftwareLists {
    param (
        [Parameter(Mandatory=$true)]
        [array]$currentList,
        [Parameter(Mandatory=$true)]
        [array]$previousList
    )
    # Regular expression to remove all numbers and optional dots/hyphens/spaces from software names, to ignore version numbers
    $versionPattern = "[\d\.\-\s]"

    # Strip version numbers from software names for comparison
    $previousNames = $previousList | ForEach-Object {     
        if ($_.Name -ne $null -and $_.Name -ne "null") {
            ($_).Name -replace $versionPattern, '' 
        } else {
            "null"
        }
    }

    $newSoftware = @()
    foreach ($software in $currentList) {
        $softwareName = if ($software.Name -ne $null -and $software.Name -ne "null") { $software.Name -replace $versionPattern, '' } else { "null" }
        if ($softwareName -notin $previousNames) {
            $newSoftware += $software
        }
    }

    return $newSoftware
}

# Main script logic
$currentSoftwareList = Get-InstalledSoftware
$previousSoftwareList = Load-PreviousSoftwareList
$newSoftwareLog = Load-NewSoftwareLog

# Check if this is the first run by seeing if previous list is empty
if ($previousSoftwareList.Count -eq 0) {
    # Save the current list and exit
    Save-CurrentSoftwareList -softwareList $currentSoftwareList
    Write-Output "First run: Saved the current list of installed software."
    exit 0
}

# Compare software lists to find newly installed software
$newSoftware = Compare-SoftwareLists -currentList $currentSoftwareList -previousList $previousSoftwareList

# Process new software found
if ($newSoftware) {
    $currentDate = Get-Date

    # Add new software to the log
    foreach ($software in $newSoftware) {
        $newSoftwareLog | Out-File $newSoftwareLogPath -Encoding utf8
        @{
            Name = $software.Name
            Date = $currentDate
        } | ConvertTo-Json -Depth 4 | Out-File -Append $newSoftwareLogPath -Encoding utf8
    }

    # Display the new software found
    $alertMessage = "New software installed since the last run:`n" + ($newSoftware | Format-Table -AutoSize | Out-String)
    Write-Output $alertMessage
}

# Clean up log: keep only entries from the last 7 days
if ($newSoftwareLog) {
    $thresholdDate = (Get-Date).AddDays(-$Days)
    $newSoftwareLog = $newSoftwareLog | Where-Object { $_.Date -ge $thresholdDate }
    if ($newSoftwareLog) {
        Save-NewSoftwareLog -newSoftwareLog $newSoftwareLog
    } else {
        Clear-Content -Path $newSoftwareLogPath
    }
} elseif (-not (Test-Path -Path $newSoftwareLogPath)) {
    # Save an empty log if the log file does not exist
    New-Item -Path $newSoftwareLogPath -ItemType File > $null
}

# Save the current list for future comparisons
Save-CurrentSoftwareList -softwareList $currentSoftwareList

if ($newSoftware) {
    exit 1001
} elseif ($newSoftwareLog) {
    $alertMessage = "New software installed within the last $Days days:`n" + ($newSoftwareLog | Format-Table -AutoSize | Out-String)
    Write-Output $alertMessage
    exit 1001
} else {
    Write-Output "No new software detected."
    exit 0
}
