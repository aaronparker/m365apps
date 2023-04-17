# Microsoft 365 Apps packager for Intune

A script and workflow for creating a Microsoft Intune package for the Microsoft 365 Apps.

## Scripts

* `New-Microsoft365AppsPackage.ps1` - Creates and imports a Microsoft 365 Apps package into Intune via GitHub Actions or from a local copy of this repository
* `Create-Win32App.ps1` imports the intunewin package into the target Intune tenant, using `App.json` as the template. Called by `New-Microsoft365AppsPackage.ps1`

### Usage

Use `New-Microsoft365AppsPackage.ps1` by authenticating with an Intune Administrator account before running the script:

```powershell
Connect-MSIntuneGraph -TenantID "lab.stealthpuppy.com"
$params = @{
    Path             = E:\project\m365Apps
    ConfigurationFile = E:\project\m365Apps\configs\O365ProPlus.xml
    Channel          = Current
    CompanyName      = stealthpuppy   
    TenantId         = 6cdd8179-23e5-43d1-8517-b6276a8d3189
    Import           = $true 
}
.\New-Microsoft365AppsPackage.ps1 @params
```

Create a new package by passing credentials to an Azure AD app registration that has rights to import applications into Microsoft Intune. This approach can be modified for use within a pipeline:

```powershell
$params = @{
    Path             = E:\project\m365Apps
    ConfigurationFile = E:\project\m365Apps\configs\O365ProPlus.xml
    Channel          = MonthlyEnterprise
    CompanyName      = stealthpuppy   
    TenantId         = 6cdd8179-23e5-43d1-8517-b6276a8d3189
    ClientId         = 60912c81-37e8-4c94-8cd6-b8b90a475c0e
    ClientSecret     = <secret>
    Import           = $true 
}
.\New-Microsoft365AppsPackage.ps1 @params
```

## Workflow

Requires the following secrets on the repo:

* `TENANT_ID` - tenant ID used by `import-package.yml`
* `CLIENT_ID` - app registration client ID used by `import-package.yml` to authenticate to the target tenent
* `CLIENT_SECRET` - password used by `import-package.yml` to authenticate to the target tenent
