# Microsoft 365 Apps packager for Intune

A script and workflow for creating a Microsoft Intune package for the Microsoft 365 Apps.

## Scripts

* `New-Microsoft365AppsPackage.ps1` - Creates and imports a Microsoft 365 Apps package into Intune via GitHub Actions or from a local copy of this repository
* `Create-Win32App.ps1` imports the intunewin package into the target Intune tenant, using `App.json` as the template. Called by `New-Microsoft365AppsPackage.ps1`

### Usage via Administrator Sign-in

Use `New-Microsoft365AppsPackage.ps1` by authenticating with an Intune Administrator account before running the script. Run `Connect-MSIntuneGraph` to authenticate with administrator credentials using a sign-in window or device login URL.

```powershell
Connect-MSIntuneGraph -TenantID "lab.stealthpuppy.com"
$params = @{
    Path             = "E:\project\m365Apps"
    ConfigurationFile = "E:\project\m365Apps\configs\O365ProPlus.xml"
    Channel          = "Current"
    CompanyName      = "stealthpuppy"
    TenantId         = "6cdd8179-23e5-43d1-8517-b6276a8d3189"
    Import           = $true 
}
.\New-Microsoft365AppsPackage.ps1 @params
```

### Usage via App Registration

Create a new package by passing credentials to an Azure AD app registration that has rights to import applications into Microsoft Intune. This approach can be modified for use within a pipeline:

```powershell
$params = @{
    Path             = "E:\project\m365Apps"
    ConfigurationFile = "E:\project\m365Apps\configs\O365ProPlus.xml"
    Channel          = "MonthlyEnterprise"
    CompanyName      = "stealthpuppy"
    TenantId         = "6cdd8179-23e5-43d1-8517-b6276a8d3189"
    ClientId         = "60912c81-37e8-4c94-8cd6-b8b90a475c0e"
    ClientSecret     = "<secret>"
    Import           = $true 
}
.\New-Microsoft365AppsPackage.ps1 @params
```

The app registration requires the following API permissions:

| API / Permissions name | Type | Description | Admin consent required |
|:--|:--|:--|:--|
| DeviceManagementApps.ReadAll | Application | Read Microsoft Intune apps | Yes |
| DeviceManagementApps.ReadWriteAll | Application | Read and write Microsoft Intune apps | Yes |

## Import Package Workflow

Requires the following secrets on the repo:

* `TENANT_ID` - tenant ID used by `import-package.yml`
* `CLIENT_ID` - app registration client ID used by `import-package.yml` to authenticate to the target tenent
* `CLIENT_SECRET` - password used by `import-package.yml` to authenticate to the target tenent

## Update Binaries Workflow

This repository includes copies of the following binaries and support files that are automatically kept updated with the latest versions:

* [Microsoft 365 Apps / Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117) (`setup.exe`) - the key installer required to install, configure and uninstall the Microsoft 365 Apps
* [Microsoft Win32 Content Prep Tool](https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool) - the tool that converts Win32 applications into the intunewin package format
* [PSAppDeployToolkit](https://psappdeploytoolkit.com/) - the install is managed with the PowerShell App Deployment Toolkit

If you have cloned this repository, ensure that you synchronise changes to update binaries to the latest version releases.
