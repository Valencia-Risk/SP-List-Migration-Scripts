# SharePoint List Migration Scripts
Simple PowerShell Scripts to Import/Export Lists using PnP Provisioning's Site Templates. Useful for Tenant migrations.
These use the modern SharePoint solution for templating, Site Templates. Instead of the old SharePoint method of App Templates (.stp file).

Data can be optionally included. 

## Configuration

### Prerequisites

* Ensure you have PowerShell version 7.4.0 or higher installed. You can check your current version by running the following command in your PowerShell terminal:
  ```powershell
  $PSVersionTable.PSVersion
  ```
  If your version is lower than 7.4.0, you will need to update PowerShell.

* Install the PnP.PowerShell module if it is not already installed. You can do this by running the following command in your PowerShell terminal:
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser -Force
  ```

These steps are necessary to ensure the scripts in this repository run correctly. `ExportList.ps1` will include checks for these as well.

### Entra ID App Registration

* Set up a multi-tenant application in Entra ID.
* Configure the following delegated API permissions for the application:
  * `Graph - Sites.FullControl.All`
  * `SharePoint - AllSites.FullControl`
* Ensure the `ClientId` in the `appsettings.json` file matches the application registration's client ID.
* A client-secret is not needed as it's Multi-Tenant and uses Delegated permissions.

These steps are necessary to ensure the application has the required permissions to access and manage SharePoint sites and lists.

### Setting up appsettings.json

1. Copy the `appsettings.template.json` file and rename it to `appsettings.json`.
2. Open the `appsettings.json` file and update the values as needed.

## Running the Scripts

To run the scripts, use the following command in your PowerShell terminal:

### Export SharePoint List
The script will use the settings defined in the `appsettings.json` file. Therefore, ensure the correct List Names and Source URL are configured in `appsettings.json`.

```powershell
./ExportList.ps1
```
