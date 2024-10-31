# ===========================
# Checking Prerequisites
# ===========================

Write-Host -f Magenta "=== Checking Prerequisites ==="

# Required minimum PowerShell version for PnP.PowerShell
$requiredVersion = [Version]"7.4.0"
$currentVersion = $PSVersionTable.PSVersion

# Check if the current version meets or exceeds the required version
if ($currentVersion -lt $requiredVersion) {
    Write-Host -f Red "PowerShell version $($currentVersion) is installed. Version 7.4 or higher is required."
    Exit
}

Write-Host -f Green "PowerShell version $($currentVersion) is installed, which meets the requirement."

# Check PnP.PowerShell is installed, if not, install it.
If(-not (Get-Module PnP.PowerShell -ListAvailable)){
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
}

Write-Host -f Green "All prerequisities met."

# ===========================
# Get and Define Settings
# ===========================

Write-Host -f Magenta "=== Load AppSettings.json ==="

$SettingsFilePath = "./appsettings.json"

# XML or PNP
$ExtensionType = "xml"
$OutputFolder = "data"

Try {
    # Load Settings from Json
    $Settings = Get-Content -Path $SettingsFilePath | ConvertFrom-Json
}
Catch {
    Write-Host -f Red "Error reading appsettings.json. Ensure it exists." $_.Exception.Message
    Exit
}

# Must be a Multi-Tenant App Registration
$ClientId = $Settings.ClientId
Write-Host -f Cyan "Client ID: $ClientId"

# Site to Import to
$DestSiteURL = $Settings.ImportSettings.DestinationSiteURL
Write-Host -f Cyan "DestSiteURL: $DestSiteURL"

# Local file path
$ImportFilePath = $Settings.ImportSettings.ImportFilePath
Write-Host -f Cyan "ImportFileName: $ImportFilePath"

Write-Host -f Green "Successfully loaded appsettings.json."

# ===========================
# Importing Lists and Data
# ===========================

Write-Host -f Magenta "=== Import Lists and Data ==="

Write-Host -f Cyan "Prompting for login..."

Connect-PnPOnline -Url $DestSiteURL -Interactive -ClientId $ClientId
$Context  = Get-PnPContext

If ($Context) {
    Write-Host -f Green "Logged in Successfully."
}

Function ImportLists {
    param ($FilePath)

    Write-Host -f Cyan "Starting List import..."

    if (Test-Path -Path $FilePath) {
        Write-Host -f Green "Template exists at $FilePath."
    } else {
        Write-Host -f Red "Template does not exist at $FilePath. Check file exists and verify appsettings.json."
        Exit
    }

    Try {
        Invoke-PnPSiteTemplate -Path $FilePath
    }
    Catch {
        Write-Host -f Red "Error importing Lists from template." $_.Exception.Message
        Exit
    }

    Write-Host -f Green "All Lists have been successfully imported to '$DestSiteURL'"
}

ImportLists $ImportFilePath

Disconnect-PnPOnline

Write-Host -f Magenta "=== Import Finished ==="