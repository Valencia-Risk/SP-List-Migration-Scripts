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

# Site to Export from
$SourceSiteURL = $Settings.ExportSettings.SourceSiteURL
Write-Host -f Cyan "SourceSiteURL: $SourceSiteURL"

# True or False
$IncludeData = $Settings.ExportSettings.IncludeData
Write-Host -f Cyan "IncludeData: $IncludeData"

# Array of List Names to Export
$ListsToExport = $Settings.ExportSettings.ListsToExport
Write-Host -f Cyan "ListsToExport: $ListsToExport"

# File name excluding the path and extension
$ExportFileName = $Settings.ExportSettings.ExportFileName
Write-Host -f Cyan "ExportFileName: $ExportFileName"

Write-Host -f Green "Successfully loaded appsettings.json."

# ===========================
# Exporting Lists and Data
# ===========================

Write-Host -f Magenta "=== Export Lists and Data ==="

Write-Host -f Cyan "Prompting for login..."

Connect-PnPOnline -Url $SourceSiteURL -Interactive -ClientId $ClientId
$Context  = Get-PnPContext

If ($Context) {
    Write-Host -f Green "Logged in Successfully."
}

Function ExportLists {
    param ($Lists, $FileName, $IncludeData)

    Write-Host -f Cyan "Starting List export..."

    Try {
        $TemplatePath = "$OutputFolder/$FileName.$ExtensionType"

        Write-Host -f Cyan "Starting Schema export for Lists provided..."
        Get-PnPSiteTemplate -Out $TemplatePath -Handlers Lists -ListsToExtract $Lists
        Write-Host -f Green "Lists Schema successfully extracted."

        If ($IncludeData) {
            Write-Host -f Cyan "Starting data export for Lists provided..."
            foreach ($List in $Lists) {
                Write-Host -f Cyan "Exporting Data for '$List'."
                Add-PnPDataRowsToSiteTemplate -Path $TemplatePath -List $List
                Write-Host -f Green "Data for '$List' successfully exported."
            }
            Write-Host -f Green "Successfully exported all Lists data."           
        }
    }
    Catch {
        Write-Host -f Red "Error saving Lists as a template." $_.Exception.Message
        Exit
    }

    Write-Host -f Green "All Lists have been successfully exported to '$TemplatePath'"
}

ExportLists $ListsToExport $ExportFileName $IncludeData

Disconnect-PnPOnline

Write-Host -f Magenta "=== Export Finished ==="