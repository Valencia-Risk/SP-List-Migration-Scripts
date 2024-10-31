# SharePoint List Migration Scripts
Simple PowerShell Scripts to Import/Export Lists using PnP Provisioning's Site Templates. Useful for Tenant migrations.
PnP Provisioning uses the modern SharePoint solution for templating, Site Templates, instead of the old SharePoint method of App Templates (.stp file).

Data can be optionally exported using the `IncludeData` export setting in `appsettings.json`.

## Features

- [x] Export an array of Lists and Data (optional) from **one** SharePoint Site.
- [x] Import Lists and Data (if included) to **one** SharePoint Site on the same Tenant or another Tenant.

### Needs Developing

- [ ] Combined export-to-import automation for migrating lists.
- [ ] Rename lists on import/export.

## Important Note on Importing

**TLDR:** Destination's existing rows will **not** be replaced or merged. New or missing rows **will** be added to an existing List.

Imported lists will be applied on top of any Lists on the destination that share the same name, but it will preserve rows that existed before based on Unique column requirements. For example, ID 23 exists on Destination already, but the template also includes ID 23. As ID is **unique**, the new (from template) row violates this rule and is dropped, not merged.

However, any new rows that don't violate **unique** rules will be added to an existing List.

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

This step is only required on one Tenant as we use a Multi-Tenant setup. Meaning if I'm doing a full migration, I can just set up the App Registration in tenant 1 and I'll be asked for approval/consent when I sign into tenant 2.

*If you've been provided a Client ID from someone else, like Valencia Risk, skip the Entra ID App Registration and enter the provided ID into the appsettings.json.*

* Set up a multi-tenant application in Entra ID.
* Configure the following delegated API permissions for the application:
  * `Graph - Sites.FullControl.All`
  * `SharePoint - AllSites.FullControl`
* Grant Admin Approval for all API permissions.
* Ensure the `ClientId` in the `appsettings.json` file matches the application registration's client ID. See next section.
* A client-secret is not needed as it's Multi-Tenant and uses Delegated permissions. Meaning we get the token and tenant via the User sign in.

These steps are necessary to ensure the application has the required permissions to access and manage SharePoint sites and lists.

### Setting up appsettings.json

1. Copy the `appsettings.template.json` file and rename it to `appsettings.json`.
2. Open the `appsettings.json` file and update the values as needed.
    - If you wish to only use one script, skip the other specific settings. E.g. You only want to Import, then don't fill out the ExportSettings.

## Running the Scripts

To run the scripts, use the following command in your PowerShell terminal:

### Export SharePoint Lists
The script will use the settings defined in the `appsettings.json` file. Therefore, ensure the correct **List Names** and **Source URL** are configured in `appsettings.json`.

```powershell
./ExportList.ps1
```

### Import SharePoint Lists
The script will use the settings defined in the `appsettings.json` file. Therefore, ensure the correct **Destination URL** and **ImportFilePath** are configured in `appsettings.json`.

This will use a PnP Template, either .xml or .pnp file types. Ensure the template exists at the location specified in `appsettings.json`.

```powershell
./ImportList.ps1
```
