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
        Path to the Microsoft 365 Apps package configuration file.

    .PARAMETER Channel
        Microsoft 365 Apps release channel.

    .PARAMETER CompanyName
        Company name to include in the configuration.xml.

    .PARAMETER TenantId
        The tenant id (GUID) of the target Azure AD tenant.

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
    [Parameter(Mandatory = $true, HelpMessage = "Path to the top level directory of the repository.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -Path $_ })]
    [System.String] $Path,

    [Parameter(Mandatory = $true, HelpMessage = "Path to the Microsoft 365 Apps package configuration file.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Test-Path -Path $_ })]
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

    [System.String] $ClientId,
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
        "$Path\PSAppDeployToolkit\Toolkit\Deploy-Application.exe"
    ) | ForEach-Object { if (-not (Test-Path -Path $_)) { throw [System.IO.FileNotFoundException]::New("File not found: $_") } }
}

process {
    #region Create working directories; Copy files for the package
    try {
        Write-Msg -Msg "Create directories."
        New-Item -Path "$Path\output" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        New-Item -Path "$Path\PSAppDeployToolkit\Toolkit\Files" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        Write-Msg -Msg "Copy configuration files."
        Copy-Item -Path $ConfigurationFile -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\Install-Microsoft365Apps.xml"
        Copy-Item -Path "$Path\configs\Uninstall-Microsoft365Apps.xml" -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\Uninstall-Microsoft365Apps.xml"
        Copy-Item -Path "$Path\m365\setup.exe" -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\setup.exe"
    }
    catch {
        throw $_
    }
    #endregion

    #region Update the configuration.xml
    try {
        $InstallXml = "$Path\PSAppDeployToolkit\Toolkit\Files\Install-Microsoft365Apps.xml"
        Write-Msg -Msg "Read configuration file: $InstallXml."
        [System.Xml.XmlDocument]$Xml = Get-Content -Path $InstallXml
        Write-Msg -Msg "Set Microsoft 365 Apps channel to: $Channel."
        $Xml.Configuration.Add.Channel = $Channel
        $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "TenantId"))
        Write-Msg -Msg "Set tenant id to: $TenantId."
        $Xml.Configuration.Property[$Index] = $TenantId
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
        SourceFolder         = "$Path\PSAppDeployToolkit\Toolkit"
        SetupFile            = "Files\setup.exe"
        OutputFolder         = "$Path\output"
        Force                = $true
        IntuneWinAppUtilPath = "$Path\intunewin\IntuneWinAppUtil.exe"
    }
    New-IntuneWin32AppPackage @params
    #endregion

    #region Create a new App.json for the package & update based on the setup.exe version & configuration.xml
    try {
        $SetupVersion = (Get-Item -Path "$Path\m365\setup.exe").VersionInfo.FileVersion
        Write-Msg -Msg "Using setup.exe version: $SetupVersion."
        Write-Msg -Msg "Copy App.json to: $Path\output\m365apps.json."
        Copy-Item -Path "$Path\scripts\App.json" -Destination "$Path\output\m365apps.json"

        Write-Msg -Msg "Get content from: $Path\output\m365apps.json."
        $AppJson = Get-Content -Path "$Path\output\m365apps.json" | ConvertFrom-Json
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
        $AppJson.Information.DisplayName = $DisplayName
        Write-Msg -Msg "Package display name: $DisplayName."

        # Update package description
        $Description = "$($AppJson.Information.Description) Uses setup.exe $SetupVersion. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
        Write-Msg -Msg "Package description: $Description."
        $AppJson.Information.Description = $Description

        # Read the product Ids from the XML, order in alphabetical order, update value in JSON
        $ProductReleaseIDs = ($Xml.Configuration.Add.Product.ID | Sort-Object) -join ","
        $Index = $AppJson.DetectionRule.IndexOf($($AppJson.DetectionRule -cmatch "ProductReleaseIds"))
        Write-Msg -Msg "Update product release Ids for registry detection rule: $ProductReleaseIDs."
        $AppJson.DetectionRule[$Index].Value = $ProductReleaseIDs
        
        # Update the registry version number detection rule
        Remove-Variable -Name "Index" -ErrorAction "SilentlyContinue"
        $ChannelVersion = Get-EvergreenApp -Name "Microsoft365Apps" | Where-Object { $_.Channel -eq $Channel }
        $Index = $AppJson.DetectionRule.IndexOf($($AppJson.DetectionRule -cmatch "VersionToReport"))
        Write-Msg -Msg "Update channel version number for registry detection rule: $($ChannelVersion.Version)."
        $AppJson.DetectionRule[$Index].Value = $ChannelVersion.Version

        # Output details back to the JSON file
        Write-Msg -Msg "Write updated App.json details back to: $Path\output\m365apps.json."
        $AppJson | ConvertTo-Json | Out-File -FilePath "$Path\output\m365apps.json" -Force
    }
    catch {
        throw $_
    }
    #endregion
    
    #region Authn if authn parameters are passed; Import package into Intune
    if ($PSBoundParameters.ContainsKey("Import")) {
        Write-Msg -Msg "-Import specified. Importing package into tenant."

        # Get the package file
        $PackageFile = Get-ChildItem -Path "$Path\output" -Recurse -Include "setup.intunewin"
        if ($null -eq $PackageFile) { throw [System.IO.FileNotFoundException]::New("Intunewin package file not found.") }

        if ($PSBoundParameters.ContainsKey("ClientId")) {
            $params = @{
                TenantId     = $TenantId
                ClientId     = $ClientId
                ClientSecret = $ClientSecret
            }
            Write-Msg -Msg "Authenticate to tenant: $TenantId."
            $global:AuthToken = Connect-MSIntuneGraph @params
        }

        # Launch script to import the package
        Write-Msg -Msg "Create package with: $Path\scripts\Create-Win32App.ps1."
        $params = @{
            Json        = "$Path\output\m365apps.json"
            PackageFile = $PackageFile.FullName
        }
        & "$Path\scripts\Create-Win32App.ps1" @params
    }
    #endregion
}

end {
}
