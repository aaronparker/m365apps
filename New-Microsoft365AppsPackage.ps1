#Requires -PSEdition Desktop
#Requires -Modules Evergreen, MSAL.PS, IntuneWin32App
using namespace System.Management.Automation
<#
    .SYNOPSIS
        Create the Intune package for the Microsoft 365 Apps and imported into an Intune tenant.

    .DESCRIPTION
        Uses a specified configuration.xml to create an intunewin package for the Microsoft 365 Apps with the PSAppDeployToolkit.

    .PARAMETER Path
        Path to the top level directory of the m365apps repository on a local Windows machine.

    .PARAMETER ConfigurationFile
        Full path to the Microsoft 365 Apps package configuration file.

    .PARAMETER Channel
        A supported Microsoft 365 Apps release channel.

    .PARAMETER CompanyName
        Company name to include in the configuration.xml.

    .PARAMETER TenantId
        The tenant id (GUID) of the target Azure AD tenant.

    .PARAMETER ClientId
        The client id (GUID) of the target Azure AD app registration.

    .PARAMETER ClientSecret
        Client secret used to authenticate against the app registration.

    .PARAMETER Import
        Switch parameter to specify that the the package should be imported into the Microsoft Intune tenant.

    .EXAMPLE
        Connect-MSIntuneGraph -TenantID "lab.stealthpuppy.com"
        $params = @{
            Path              = E:\project\m365Apps
            ConfigurationFile  = E:\project\m365Apps\configs\O365ProPlus.xml
            Channel           = Current
            CompanyName       = stealthpuppy
            TenantId          = 6cdd8179-23e5-43d1-8517-b6276a8d3189
            Import            = $true
        }
        .\New-Microsoft365AppsPackage.ps1 @params

    .EXAMPLE
        $params = @{
            Path              = E:\project\m365Apps
            ConfigurationFile  = E:\project\m365Apps\configs\O365ProPlus.xml
            Channel           = Current
            CompanyName       = stealthpuppy
            TenantId          = 6cdd8179-23e5-43d1-8517-b6276a8d3189
            ClientId          = 60912c81-37e8-4c94-8cd6-b8b90a475c0e
            ClientSecret      = <secret>
            Import            = $true
        }
        .\New-Microsoft365AppsPackage.ps1 @params

    .NOTES
        Author: Aaron Parker
        Twitter: @stealthpuppy
#>
[CmdletBinding(SupportsShouldProcess = $false)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to the top level directory of the repository.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ if (Test-Path -Path $_ -PathType "Container") { $true } else { throw "Path not found: '$_'" } })]
    [System.String] $Path = $PSScriptRoot,

    [Parameter(Mandatory = $true, HelpMessage = "Path to the Microsoft 365 Apps package configuration file.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ if (Test-Path -Path $_ -PathType "Leaf") { $true } else { throw "Path not found: '$_'" } })]
    [ValidateScript({ if ((Get-Item -Path $_).Extension -eq ".xml") { $true } else { throw "File not not an XML file: '$_'" } })]
    [System.String] $ConfigurationFile,

    [Parameter(Mandatory = $false, HelpMessage = "Microsoft 365 Apps release channel.")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("BetaChannel", "CurrentPreview", "Current", "MonthlyEnterprise", "PerpetualVL2021", "SemiAnnualPreview", "SemiAnnual", "PerpetualVL2019")]
    [System.String] $Channel = "MonthlyEnterprise",

    [Parameter(Mandatory = $false, HelpMessage = "Company name to include in the configuration.xml.")]
    [ValidateNotNullOrEmpty()]
    [System.String] $CompanyName = "stealthpuppy",

    [Parameter(Mandatory = $true, HelpMessage = "The tenant id (GUID) of the target Azure AD tenant.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ $ObjectGuid = [System.Guid]::empty; if ([System.Guid]::TryParse($_, [System.Management.Automation.PSReference]$ObjectGuid)) { $true } else { throw "$_ is not a GUID" } })]
    [System.String] $TenantId,

    [Parameter(Mandatory = $false, HelpMessage = "The client id (GUID) of the target Azure AD app registration.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ $ObjectGuid = [System.Guid]::empty; if ([System.Guid]::TryParse($_, [System.Management.Automation.PSReference]$ObjectGuid)) { $true } else { throw "$_ is not a GUID" } })]
    [System.String] $ClientId,

    [Parameter(Mandatory = $false, HelpMessage = "Client secret used to authenticate against the app registration.")]
    [ValidateNotNullOrEmpty()]
    [System.String] $ClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Import the package into Microsoft Intune.")]
    [System.Management.Automation.SwitchParameter] $Import
)

begin {
    function Write-Msg ($Msg) {
        $Message = [HostInformationMessage]@{
            Message         = "[$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')]"
            ForegroundColor = "Black"
            BackgroundColor = "DarkCyan"
            NoNewline       = $true
        }
        $params = @{
            MessageData       = $Message
            InformationAction = "Continue"
            Tags              = "Microsoft365"
        }
        Write-Information @params
        $params = @{
            MessageData       = " $Msg"
            InformationAction = "Continue"
            Tags              = "Microsoft365"
        }
        Write-Information @params
    }

    # Validate that the input file is XML
    [System.Xml.XmlDocument]$Xml = Get-Content -Path $ConfigurationFile -ErrorAction "Stop"

    # Unblock all files in the repo
    Write-Msg -Msg "Unblock exe files in $Path."
    Get-ChildItem -Path $Path -Recurse -Include "*.exe" | Unblock-File

    # Validate required files exist
    Write-Msg -Msg "Validate required files exist."
    @(
        "$Path\configs\Uninstall-Microsoft365Apps.xml",
        "$Path\intunewin\IntuneWinAppUtil.exe",
        "$Path\m365\setup.exe",
        "$Path\icons\Microsoft365.png",
        "$Path\scripts\App.json",
        "$Path\scripts\Create-Win32App.ps1",
        "$Path\PSAppDeployToolkit\Toolkit\Deploy-Application.exe",
        "$Path\PSAppDeployToolkit\Toolkit\Deploy-Application.exe.config",
        "$Path\PSAppDeployToolkit\Toolkit\Deploy-Application.ps1",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitBanner.png",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitConfig.xml",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitExtensions.ps1",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitHelp.ps1",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitLogo.ico",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitLogo.png",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitMain.cs",
        "$Path\PSAppDeployToolkit\Toolkit\AppDeployToolkit\AppDeployToolkitMain.ps1",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrub03.vbs",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrub07.vbs",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrub10.vbs",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrubc2r.vbs",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrub_O15msi.vbs",
        "$Path\PSAppDeployToolkit\Toolkit\SupportFiles\OffScrub_O16msi.vbs"
    ) | ForEach-Object { if (-not (Test-Path -Path $_)) { throw [System.IO.FileNotFoundException]::New("File not found: $_") } }
}

process {
    #region Create working directories; Copy files for the package
    try {
        # Set output directory and ensure it is empty
        $OutputPath = "$Path\package"
        if ((Get-ChildItem -Path $OutputPath -Recurse -File).Count -gt 0) {
            Write-Warning -Message "'$OutputPath' is not empty. Remove path and try again."
            return
        }

        # Create the package directory structure
        Write-Msg -Msg "Create new package structure."
        Write-Msg -Msg "Using package path: $OutputPath"
        New-Item -Path "$OutputPath\source" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        New-Item -Path "$OutputPath\output" -ItemType "Directory" -ErrorAction "SilentlyContinue"

        # Copy the PSAppDeployToolkit files to the package source
        # Copy the customised Deploy-Application.ps1 to the package source
        Write-Msg -Msg "Copy PSAppDeployToolkit to: $OutputPath\source."
        Copy-Item -Path "$Path\PSAppDeployToolkit\Toolkit\*" -Destination "$OutputPath\source" -Recurse
        New-Item -Path "$OutputPath\source\Files" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        Write-Msg -Msg "Copy Deploy-Application.ps1 to: $OutputPath\source\Deploy-Application.ps1."
        Copy-Item -Path "Deploy-Application.ps1" -Destination "$OutputPath\source\Deploy-Application.ps1" -Force

        # Copy the configuration files and setup.exe to the package source
        Write-Msg -Msg "Copy configuration files and setup.exe to package source."
        Copy-Item -Path $ConfigurationFile -Destination "$OutputPath\source\Files\Install-Microsoft365Apps.xml"
        Copy-Item -Path "$Path\configs\Uninstall-Microsoft365Apps.xml" -Destination "$OutputPath\source\Files\Uninstall-Microsoft365Apps.xml"
        Copy-Item -Path "$Path\m365\setup.exe" -Destination "$OutputPath\source\Files\setup.exe"
    }
    catch {
        throw $_
    }
    #endregion

    #region Update the configuration.xml
    try {
        $InstallXml = "$OutputPath\source\Files\Install-Microsoft365Apps.xml"
        Write-Msg -Msg "Read configuration file: $InstallXml."
        [System.Xml.XmlDocument]$Xml = Get-Content -Path $InstallXml

        Write-Msg -Msg "Set Microsoft 365 Apps channel to: $Channel."
        $Xml.Configuration.Add.Channel = $Channel

        Write-Msg -Msg "Set tenant id to: $TenantId."
        $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "TenantId"))
        $Xml.Configuration.Property[$Index].Value = $TenantId

        Write-Msg -Msg "Set company name to: $CompanyName."
        $Xml.Configuration.AppSettings.Setup.Value = $CompanyName

        Write-Msg -Msg "Save configuration xml to: $InstallXml."
        $Xml.Save($InstallXml)
    }
    catch {
        throw $_
    }
    #endregion

    #region Create the intunewin package
    Write-Msg -Msg "Create intunewin package in: $Path\output."
    $params = @{
        SourceFolder         = "$OutputPath\source"
        SetupFile            = "Files\setup.exe"
        OutputFolder         = "$OutputPath\output"
        Force                = $true
        IntuneWinAppUtilPath = "$Path\intunewin\IntuneWinAppUtil.exe"
    }
    New-IntuneWin32AppPackage @params
    #endregion

    # Save a copy of the modified configuration file to the output folder for reference
    $OutputXml = "$OutputPath\output\$(Split-Path -Path $ConfigurationFile -Leaf)"
    Write-Msg -Msg "Saved configuration file to: $OutputXml."
    $Xml.Save($OutputXml)

    #region Create a new App.json for the package & update based on the setup.exe version & configuration.xml
    try {
        $SetupVersion = (Get-Item -Path "$Path\m365\setup.exe").VersionInfo.FileVersion
        Write-Msg -Msg "Using setup.exe version: $SetupVersion."

        Write-Msg -Msg "Copy App.json to: $OutputPath\output\m365apps.json."
        Copy-Item -Path "$Path\scripts\App.json" -Destination "$OutputPath\output\m365apps.json"

        Write-Msg -Msg "Get content from: $OutputPath\output\m365apps.json."
        $Manifest = Get-Content -Path "$OutputPath\output\m365apps.json" | ConvertFrom-Json

        Write-Msg -Msg "Using setup.exe version: $SetupVersion."
        $Manifest.PackageInformation.Version = $SetupVersion

        Write-Msg -Msg "Read configuration xml file: $InstallXml."
        [System.Xml.XmlDocument]$Xml = Get-Content -Path $InstallXml

        # Update package display name
        [System.String] $ProductID = ""
        switch ($Xml.Configuration.Add.Product.ID) {
            "O365ProPlusRetail" {
                $ProductID += "Microsoft 365 apps for enterprise, "
            }
            "O365BusinessRetail" {
                $ProductID += "Microsoft 365 apps for business, "
            }
            "VisioProRetail" {
                $ProductID += "Visio, "
            }
            "ProjectProRetail" {
                $ProductID += "Project, "
            }
            "AccessRuntimeRetail" {
                $ProductID += "Access Runtime, "
            }
        }
        [System.String] $DisplayName = "$ProductID$($Xml.Configuration.Add.Channel)"
        if ($Xml.Configuration.Add.OfficeClientEdition -eq "64") { $DisplayName = "$DisplayName, x64" }
        if ($Xml.Configuration.Add.OfficeClientEdition -eq "32") { $DisplayName = "$DisplayName, x86" }
        Write-Msg -Msg "Package display name: $DisplayName."
        $Manifest.Information.DisplayName = $DisplayName

        # Set the PSPackageFactory GUID to the GUID in the configuration.xml
        # This allows us to track the app via the configuration ID once imported into Intune
        $Manifest.Information.PSPackageFactoryGuid = $Xml.Configuration.ID

        # Update icon location
        Write-Msg -Msg "Using icon location: $Path\icons\Microsoft365.png."
        $Manifest.PackageInformation.IconFile = "$Path\icons\Microsoft365.png"

        # Update package description
        $Description = "$($xml.Configuration.Info.Description)`n`n**This package will uninstall previous versions of Microsoft Office**. Uses setup.exe $SetupVersion. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
        Write-Msg -Msg "Package description: $Description."
        $Manifest.Information.Description = $Description

        # Read the product Ids from the XML, order in alphabetical order, update ProductReleaseIds value in JSON
        $ProductReleaseIDs = ($Xml.Configuration.Add.Product.ID | Sort-Object) -join ","
        $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "ProductReleaseIds"))
        Write-Msg -Msg "Update registry ProductReleaseIds detection rule: $ProductReleaseIDs."
        $Manifest.DetectionRule[$Index].Value = $ProductReleaseIDs

        # Update the registry VersionToReport version number detection rule
        Remove-Variable -Name "Index" -ErrorAction "SilentlyContinue"
        $ChannelVersion = Get-EvergreenApp -Name "Microsoft365Apps" | Where-Object { $_.Channel -eq $Channel }
        $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "VersionToReport"))
        Write-Msg -Msg "Update registry VersionToReport detection rule: $($ChannelVersion.Version)."
        $Manifest.DetectionRule[$Index].Value = $ChannelVersion.Version

        # Update the registry SharedComputerLicensing detection rule
        Remove-Variable -Name "Index" -ErrorAction "SilentlyContinue"
        $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "SharedComputerLicensing"))
        $Value = ($Xml.Configuration.Property | Where-Object { $_.Name -eq "SharedComputerLicensing" }).Value
        Write-Msg -Msg "Update registry SharedComputerLicensing detection rule: $Value."
        $Manifest.DetectionRule[$Index].Value = $Value

        # Output details back to the JSON file
        Write-Msg -Msg "Write updated App.json details back to: $OutputPath\output\m365apps.json."
        $Manifest | ConvertTo-Json | Out-File -FilePath "$OutputPath\output\m365apps.json" -Force
    }
    catch {
        throw $_
    }
    #endregion

    #region Authn to the Microsoft Graph
    if ($PSBoundParameters.ContainsKey("ClientId")) {
        $params = @{
            TenantId     = $TenantId
            ClientId     = $ClientId
            ClientSecret = $ClientSecret
        }
        Write-Msg -Msg "Authenticate to tenant: $TenantId."
        [Void](Connect-MSIntuneGraph @params)
    }
    #endregion

    #region Lets see if this application is already in Intune and needs to be updated
    Write-Msg -Msg "Retrieve existing Microsoft 365 Apps in Intune"
    Remove-Variable -Name "ExistingApp" -ErrorAction "SilentlyContinue"
    $ExistingApp = Get-IntuneWin32App | `
        Select-Object -Property * -ExcludeProperty "largeIcon" | `
        Where-Object { $_.notes -match "PSPackageFactory" } | `
        Where-Object { ($_.notes | ConvertFrom-Json -ErrorAction "SilentlyContinue").Guid -eq $Manifest.Information.PSPackageFactoryGuid } | `
        Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
        Select-Object -First 1

    # Determine whether the new package should be imported
    if ($null -eq $ExistingApp) {
        Write-Msg -Msg "Import new application: '$($Manifest.Information.DisplayName), $($ExistingApp.displayVersion)'"
        $UpdateApp = $true
    }
    elseif ([System.String]::IsNullOrEmpty($ExistingApp.displayVersion)) {
        Write-Msg -Msg "Found matching app but `displayVersion` is null: '$($ExistingApp.displayName)'"
        $UpdateApp = $false
    }
    elseif ($Manifest.PackageInformation.Version -le $ExistingApp.displayVersion) {
        Write-Msg -Msg "Existing Intune app version is current: '$($ExistingApp.displayName), $($ExistingApp.displayVersion)'"
        $UpdateApp = $false
    }
    elseif ($Manifest.PackageInformation.Version -gt $ExistingApp.displayVersion) {
        Write-Msg -Msg "Import application version: '$($Manifest.Information.DisplayName), $($ExistingApp.displayVersion)'"
        $UpdateApp = $true
    }
    #endregion

    if ($UpdateApp -eq $true -or $Force -eq $true) {
        if ($Import -eq $true) {
            #region Authn if authn parameters are passed; Import package into Intune
            Write-Msg -Msg "-Import specified. Importing package into tenant."

            # Get the package file
            $PackageFile = Get-ChildItem -Path "$OutputPath\output" -Recurse -Include "setup.intunewin"
            if ($null -eq $PackageFile) { throw [System.IO.FileNotFoundException]::New("Intunewin package file not found.") }

            # Launch script to import the package
            Write-Msg -Msg "Create package with: $Path\scripts\Create-Win32App.ps1."
            $params = @{
                Json        = "$OutputPath\output\m365apps.json"
                PackageFile = $PackageFile.FullName
            }
            $ImportedApp = & "$Path\scripts\Create-Win32App.ps1" @params | Select-Object -Property * -ExcludeProperty "largeIcon"
            Write-Msg -Msg "Package import complete."
            #endregion

            #region Add supersedence for existing packages
            Write-Msg -Msg "Retrieve existing Microsoft 365 Apps in Intune"
            $Supersedence = Get-IntuneWin32App | `
                Where-Object { $_.id -ne $ImportedApp.id } | `
                Where-Object { $_.notes -match "PSPackageFactory" } | `
                Where-Object { ($_.notes | ConvertFrom-Json -ErrorAction "SilentlyContinue").Guid -eq $Manifest.Information.PSPackageFactoryGuid } | `
                Select-Object -Property * -ExcludeProperty "largeIcon" | `
                Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
                ForEach-Object { New-IntuneWin32AppSupersedence -ID $_.id -SupersedenceType "Update" }
            if ($null -ne $Supersedence) {
                Add-IntuneWin32AppSupersedence -ID $ImportedApp.id -Supersedence $Supersedence
            }
            #endregion

            # Output imported application details
            $ImportedApp
        }
    }
}

end {
}
