# Check for monitoring running MailStore instances
# Expected instances are saved in a text file
# New instances are automatically added to the text file
# Instances can be excluded via parameter

param (
    [string[]]$Exclude = @()
)

$storagePath = "C:\ProgramData\MailStore"
$instanceFile = Join-Path $storagePath "instances_lms.txt"

if (-not (Test-Path $storagePath)) {
    New-Item -Path $storagePath -ItemType Directory -Force | Out-Null
}

# Read MailStore processes using aliases
$processes = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -eq "MailStoreServer_x64.exe" -and $_.CommandLine -match '/cloud-instance-url-alias:'
}

# Extract currently running aliases
$currentAliases = @()
foreach ($proc in $processes) {
    if ($proc.CommandLine -match '/cloud-instance-url-alias:\s*"?(?<alias>[^"\s]+)"?') {
        $currentAliases += $matches['alias']
    }
}

# Load or initialize expected aliases
$expectedAliases = @()
if (Test-Path $instanceFile) {
    $expectedAliases = Get-Content $instanceFile | Where-Object { $_ -ne "" }
}

# Add new aliases and update the file
$mergedAliases = ($expectedAliases + $currentAliases) | Sort-Object -Unique
$mergedAliases | Set-Content -Path $instanceFile -Encoding UTF8

# Remove excluded aliases from the checks
$effectiveExpectedAliases = $expectedAliases | Where-Object { $_ -notin $Exclude }

# Detect missing instances
$missing = $effectiveExpectedAliases | Where-Object { $_ -notin $currentAliases }

# Result
if ($missing.Count -eq 0) {
    Write-Host ""
    Write-Host "Alle erwarteten MailStore Instanzen laufen."
    Write-Host ""
    foreach ($alias in $effectiveExpectedAliases) {
        Write-Host "$alias"
    }
    exit 0
} else {
    Write-Host "Es fehlen folgende MailStore Instanzen:"
    foreach ($alias in $missing) {
        Write-Host "- $alias"
    }
    Write-Host "`nErklaerung:"
    Write-Host "Es fehlt eine zuvor erfasste MailStore Instanz von: $($missing -join ', ')"
    Write-Host "Bitte pruefe, ob die Instanz gestartet werden muss oder ob sie gewollt nicht laeuft, z. B. weil der Kunde unser Mailstore Produkt abbestellt hat."
    Write-Host "Sollte dies der Fall sein, kann die Instanz per Parameter mit dem Aliasnamen beim Skriptaufruf ausgeschlossen werden:"
    Write-Host "-Exclude alias1, alias2"
    exit 1001
}
