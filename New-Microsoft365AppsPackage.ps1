#Requires -PSEdition Desktop
#Requires -Modules Evergreen, MSAL.PS, PSIntuneAuth, AzureAD, IntuneWin32App, Microsoft.Graph.Intune
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
            ConfigurationFile = E:\project\m365Apps\configs\O365ProPlus.xml
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
    [ValidateScript({ Test-Path -Path $_ })]
    [System.String] $Path = $PSScriptRoot,

    [Parameter(Mandatory = $true, HelpMessage = "Path to the Microsoft 365 Apps package configuration file.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -Path $_ })]
    [ValidateScript({ (Get-Item -Path $_).Extension -eq ".xml" })]
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
    [System.String] $TenantId,

    [Parameter(Mandatory = $false, HelpMessage = "The client id (GUID) of the target Azure AD app registration.")]
    [ValidateNotNullOrEmpty()]
    [System.String] $ClientId,

    [Parameter(Mandatory = $false, HelpMessage = "Client secret used to authenticate against the app registration.")]
    [ValidateNotNullOrEmpty()]
    [System.String] $ClientSecret,

    [Parameter(Mandatory = $false, HelpMessage = "Import the package into Microsoft Intune.")]
    [System.Management.Automation.SwitchParameter] $Import
)

begin {
    function Write-Msg ($Msg) {
        $params = @{
            MessageData       = "[$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')] $Msg"
            InformationAction = "Continue"
            Tags              = "Microsoft365"
        }
        Write-Information @params
    }

    try {
        # Validate that the input file is XML
        [System.Xml.XmlDocument]$Xml = Get-Content -Path $ConfigurationFile
    }
    catch {
        throw [System.Xml.XmlException]::New("Failed to read as XML: $ConfigurationFile")
    }

    # Unblock all files in the repo
    Write-Msg -Msg "Unblock exe files in $Path."
    Get-ChildItem -Path $Path -Recurse -Include "*.exe" | Unblock-File

    # Validate required files exist
    Write-Msg -Msg "Validate required files exist."
    @(
        "$Path\configs\Uninstall-Microsoft365Apps.xml",
        "$Path\intunewin\IntuneWinAppUtil.exe",
        "$Path\m365\setup.exe",
        "$Path\scripts\App.json",
        "$Path\PSAppDeployToolkit\Toolkit\Deploy-Application.exe",
        "$Path\icons\Microsoft365.png"
    ) | ForEach-Object { if (-not (Test-Path -Path $_)) { throw [System.IO.FileNotFoundException]::New("File not found: $_") } }
}

process {
    #region Create working directories; Copy files for the package
    try {
        Write-Msg -Msg "Create new package structure."
        $OutputPath = "$Path\package"
        Write-Msg -Msg "Using package path: $OutputPath"
        New-Item -Path "$OutputPath\source" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        New-Item -Path "$OutputPath\output" -ItemType "Directory" -ErrorAction "SilentlyContinue"

        Write-Msg -Msg "Copy PSAppDeployToolkit to: $OutputPath\source."
        Copy-Item -Path "$Path\PSAppDeployToolkit\Toolkit\*" -Destination "$OutputPath\source" -Recurse
        New-Item -Path "$OutputPath\source\Files" -ItemType "Directory" -ErrorAction "SilentlyContinue"

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
        $AppJson = Get-Content -Path "$OutputPath\output\m365apps.json" | ConvertFrom-Json

        Write-Msg -Msg "Using setup.exe version: $SetupVersion."
        $AppJson.PackageInformation.Version = $SetupVersion

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
        $AppJson.Information.DisplayName = $DisplayName

        # Update icon location
        Write-Msg -Msg "Using icon location: $Path\icons\Microsoft365.png."
        $AppJson.PackageInformation.IconFile = "$Path\icons\Microsoft365.png"

        # Update package description
        $Description = "$($xml.Configuration.Info.Description)`n`n**This package will uninstall previous versions of Microsoft Office**. Uses setup.exe $SetupVersion. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
        Write-Msg -Msg "Package description: $Description."
        $AppJson.Information.Description = $Description

        # Read the product Ids from the XML, order in alphabetical order, update ProductReleaseIds value in JSON
        $ProductReleaseIDs = ($Xml.Configuration.Add.Product.ID | Sort-Object) -join ","
        $Index = $AppJson.DetectionRule.IndexOf($($AppJson.DetectionRule -cmatch "ProductReleaseIds"))
        Write-Msg -Msg "Update registry ProductReleaseIds detection rule: $ProductReleaseIDs."
        $AppJson.DetectionRule[$Index].Value = $ProductReleaseIDs

        # Update the registry VersionToReport version number detection rule
        Remove-Variable -Name "Index" -ErrorAction "SilentlyContinue"
        $ChannelVersion = Get-EvergreenApp -Name "Microsoft365Apps" | Where-Object { $_.Channel -eq $Channel }
        $Index = $AppJson.DetectionRule.IndexOf($($AppJson.DetectionRule -cmatch "VersionToReport"))
        Write-Msg -Msg "Update registry VersionToReport detection rule: $($ChannelVersion.Version)."
        $AppJson.DetectionRule[$Index].Value = $ChannelVersion.Version

        # Update the registry SharedComputerLicensing detection rule
        Remove-Variable -Name "Index" -ErrorAction "SilentlyContinue"
        $Index = $AppJson.DetectionRule.IndexOf($($AppJson.DetectionRule -cmatch "SharedComputerLicensing"))
        $Value = ($Xml.Configuration.Property | Where-Object { $_.Name -eq "SharedComputerLicensing" }).Value
        Write-Msg -Msg "Update registry SharedComputerLicensing detection rule: $Value."
        $AppJson.DetectionRule[$Index].Value = $Value

        # Output details back to the JSON file
        Write-Msg -Msg "Write updated App.json details back to: $OutputPath\output\m365apps.json."
        $AppJson | ConvertTo-Json | Out-File -FilePath "$OutputPath\output\m365apps.json" -Force
    }
    catch {
        throw $_
    }
    #endregion

    #region Authn if authn parameters are passed; Import package into Intune
    if ($PSBoundParameters.ContainsKey("Import")) {
        Write-Msg -Msg "-Import specified. Importing package into tenant."

        # Get the package file
        $PackageFile = Get-ChildItem -Path "$OutputPath\output" -Recurse -Include "setup.intunewin"
        if ($null -eq $PackageFile) { throw [System.IO.FileNotFoundException]::New("Intunewin package file not found.") }

        if ($PSBoundParameters.ContainsKey("ClientId")) {
            $params = @{
                TenantId     = $TenantId
                ClientId     = $ClientId
                ClientSecret = $ClientSecret
            }
            Write-Msg -Msg "Authenticate to tenant: $TenantId."
            [Void](Connect-MSIntuneGraph @params)
        }

        # Launch script to import the package
        Write-Msg -Msg "Create package with: $Path\scripts\Create-Win32App.ps1."
        $params = @{
            Json        = "$OutputPath\output\m365apps.json"
            PackageFile = $PackageFile.FullName
        }
        & "$Path\scripts\Create-Win32App.ps1" @params | Select-Object -Property * -ExcludeProperty "largeIcon"
    }
    #endregion
}

end {
}
