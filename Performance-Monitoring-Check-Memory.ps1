# The "Performance Monitoring Check - Memory" from n-able focuses on memory paging.
# With this script, we want to see the percentage usage of the *physical* RAM as shown in Task Manager.
# Last edit 10.03.2025 Keidel, Check if DATEV SQL Server is running and change threshold, as we are fine with the SQL not being limited

param (
    [int]$Threshold = 95
)

# Check if DATEV SQL Server is running and change threshold, as we are fine with the SQL not being limited
$datevSqlServerRunning = Get-Process -Name "sqlservr" -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "*DATEV*" }
if ($datevSqlServerRunning) {
    $Threshold = 98
    Write-Host "DATEV SQL running, threshold set to $Threshold%"
}

# Get the total physical memory and available memory
$totalMemory = (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory
$availableMemory = (Get-CimInstance -ClassName Win32_OperatingSystem).FreePhysicalMemory

# Convert memory values to GB for better readability
$totalMemoryGB = [math]::Round($totalMemory / 1GB, 2)
$availableMemoryGB = [math]::Round($availableMemory / 1MB, 2)

# Calculate the used memory and its percentage
$usedMemoryGB = $totalMemoryGB - $availableMemoryGB
$usedMemoryPercentage = [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 2)

# Get commit (virtual) memory details
$os = Get-CimInstance Win32_OperatingSystem
$committedUsed = $os.TotalVirtualMemorySize - $os.FreeVirtualMemory
$committedTotal = $os.TotalVirtualMemorySize

$committedUsedGB = [math]::Round($committedUsed / 1MB, 2)
$committedTotalGB = [math]::Round($committedTotal / 1MB, 2)

# Output the memory details
Write-Host "Total Physical Memory: $totalMemoryGB GB"
Write-Host "Available Memory: $availableMemoryGB GB"
Write-Host "Used Memory: $usedMemoryGB GB ($usedMemoryPercentage%)"
Write-Host "Committed Memory: $committedUsedGB GB / $committedTotalGB GB"

# Compare the used memory percentage to the threshold
if ($usedMemoryPercentage -ge $Threshold) {
    Write-Host "WARNING: Memory usage is above or equal to the threshold of $Threshold%."

    if (-not $datevSqlServerRunning) {
        # Get Top 5 Processes by Memory Usage
        Write-Host "`nTop 5 Memory-Consuming Processes:"
        Get-Process | Select ProcessName, WorkingSet64 | 
        Group-Object -Property ProcessName | 
        Select @{n='Memory (GB)';e={ "{0:N2}" -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1GB) }},
            @{n='Process Count';e={ (($_.Group | Measure-Object ProcessName).Count) }},
            @{n='Name';e={ $_.Name }} | 
        Sort-Object -Property "Memory (GB)" -Descending | 
        Select -First 5 | Format-Table -AutoSize
    }

    exit 1001
} else {
    Write-Host "Memory usage is below the threshold of $Threshold%."
    exit 0
}