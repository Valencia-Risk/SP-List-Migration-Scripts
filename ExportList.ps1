# ===========================
# Checking Prerequisites
# ===========================

# Required minimum PowerShell version for PnP.PowerShell
$requiredVersion = [Version]"7.4.0"
$currentVersion = $PSVersionTable.PSVersion

# Check if the current version meets or exceeds the required version
if ($currentVersion -lt $requiredVersion) {
    Write-Output -f Red "PowerShell version $($currentVersion) is installed. Version 7.4 or higher is required."
    Exit
}

Write-Output -f Green "PowerShell version $($currentVersion) is installed, which meets the requirement."

# Check PnP.PowerShell is installed, if not, install it.
If(-not (Get-Module PnP.PowerShell -ListAvailable)){
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
}

# ===========================
# Get and Define Settings
# ===========================

$SettingsFilePath = "./appsettings.json"

# XML or PNP
$ExtensionType = "xml"

$OutputFolder = "data"

Try {
    # Load Settings from Json
    $settings = Get-Content -Path $SettingsFilePath | ConvertFrom-Json
}
Catch {
    write-host -f Red "Error reading appsettings.json. Ensure it exists." $_.Exception.Message
    Exit
}

# Must be a Multi-Tenant App Registration
$ClientId = $Settings.ClientId
write-host -f Yellow "Client ID: $ClientId"

# Site to Export from
$SourceSiteURL = $Settings.ExportSettings.SourceSiteURL
write-host -f Yellow "SourceSiteURL: $SourceSiteURL"

# True or False
$IncludeData = $Settings.ExportSettings.IncludeData
write-host -f Yellow "IncludeData: $IncludeData"

# Array of List Names to Export
$ListsToExport = $Settings.ExportSettings.ListsToExport
write-host -f Yellow "ListsToExport: $ListsToExport"

# File name excluding the path and extension
$ExportFileName = $Settings.ExportSettings.ExportFileName
write-host -f Yellow "ExportFileName: $ExportFileName"

write-host -f Green "Successfully loaded appsettings.json"

# ===========================
# Exporting Lists and Data
# ===========================

Connect-PnPOnline -Url $SourceSiteURL -Interactive -ClientId "2407e12d-6a66-44b2-b9af-f207b7f0ea7d"
$Context  = Get-PnPContext

If ($Context) {
    Write-Host -f Green "Logged in Successfully."
}

Function ExportLists {
    param ($Lists, $FileName, $IncludeData)

    Try {
        $TemplatePath = "$OutputFolder/$FileName.$ExtensionType"

        Write-Host -f Yellow "Starting Schema export for Lists provided."
        Get-PnPSiteTemplate -Out $TemplatePath -Handlers Lists -ListsToExtract $Lists
        Write-Host -f Green "Lists Schema successfully extracted."

        If ($IncludeData) {
            Write-Host -f Yellow "Starting Data export for Lists provided."
            foreach ($List in $Lists) {
                Write-Host -f Yellow "Exporting Data for '$List'."
                Add-PnPDataRowsToSiteTemplate -Path $TemplatePath -List $List
                Write-Host -f Green "Data for '$List' successfully exported."
            }
            Write-Host -f Green "All Lists Data has been successfully exported."
        }
    }
    Catch {
        write-host -f Red "Error saving List '$ListName' as template!" $_.Exception.Message
    }
}

ExportLists $ListsToExport $ExportFileName $IncludeData

Write-Host -f Green "All Lists have been successfully exported to '$OutputFolder/$ExportFileName.$ExtensionType'"