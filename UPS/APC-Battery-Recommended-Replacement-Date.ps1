# TSÜ Script to alert if the APC - UPS - Battery Recommended Replacement Date has been reached
# created by Samuel Keidel
# LastChangeDate: 02.10.2023

# pass DNS Name or IP of the UPS with a parameter
param([string]$HostName = 'APC')

# n-sight RMM expects script return results as follows
$CHECK_PASSED = 0
$CHECK_FAILED = 1001

# SNMP OID of the APC - UPS - Battery Recommended Replacement Date		
$oid = ".1.3.6.1.4.1.318.1.1.1.2.2.21.0"

# Retrieve the SNMP value using $SNMP.Get
$SNMP = New-Object -ComObject olePrn.OleSNMP
$SNMP.Open($HostName, "public",2,1000)
$snmpResult = $SNMP.Get("$oid")

# Check if the SNMP query was successful
if ($snmpResult -ne $null) {

    # Convert the SNMP date string to a DateTime object
    $replaceDate = [datetime]::ParseExact($snmpResult, "MM/dd/yyyy", $null)

    # Calculate the alert date (30 days before the replace date)
    $alertDate = $replaceDate.AddDays(-30)

    # Get the current date and time
    $currentDate = Get-Date

    # Compare the alert date with the current date
    if ($alertDate -ge $currentDate) {
        Write-Host "$snmpResult month/day/year Battery Recommended Replacement Date"
        Exit $CHECK_PASSED
    } else {
        Write-Host "$snmpResult month/day/year The Battery Recommended Replacement Date has been reached or exceeded."
        Exit $CHECK_FAILED
    }
} else {
    Write-Host "Failed to retrieve SNMP data. Check your SNMP settings."
    Exit $CHECK_FAILED
}