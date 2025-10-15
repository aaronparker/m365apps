#Requires -PSEdition Desktop
#Requires -Modules Evergreen, MSAL.PS, IntuneWin32App, PSAppDeployToolkit
<#
    .SYNOPSIS
        Create the Intune package for the Microsoft 365 Apps and imported into an Intune tenant.

    .DESCRIPTION
        Uses a specified configuration.xml to create an intunewin package for the Microsoft 365 Apps with the PSAppDeployToolkit.

    .PARAMETER Path
        Path to the top level directory of the m365apps repository on a local Windows machine.

    .PARAMETER Destination
        Path where the package will be created. Defaults to a 'package' directory under $Path.

    .PARAMETER ConfigurationFile
        Full path to the Microsoft 365 Apps package configuration file.

    .PARAMETER Channel
        A supported Microsoft 365 Apps release channel.

    .PARAMETER CompanyName
        Company name - this is used in the configuration.xml.

    .PARAMETER TenantId
        The tenant id (GUID) of the target Entra ID tenant - this is used in the configuration.xml.

    .PARAMETER UsePsadt
        Wrap the Microsoft 365 Apps installer with the PowerShell App Deployment Toolkit.

    .PARAMETER SkipImport
        Switch parameter to specify that the the package should not be imported into the Microsoft Intune tenant.

    .EXAMPLE
        Connect-MSIntuneGraph -TenantID "lab.stealthpuppy.com"
        $params = @{
            Path              = "E:\projects\m365Apps"
            ConfigurationFile = "E:\projects\m365Apps\configs\O365ProPlus.xml"
            Channel           = "Current"
            CompanyName       = "stealthpuppy"
            UsePsadt          = $true
            TenantId          = 6cdd8179-23e5-43d1-8517-b6276a8d3189
            SkipImport        = $false
        }
        .\New-Microsoft365AppsPackage.ps1 @params

    .EXAMPLE
        $params = @{
            Path              = "E:\projects\m365Apps"
            ConfigurationFile = "E:\projects\m365Apps\configs\O365ProPlusVisioProRetailProjectProRetail.xml"
            Channel           = "Current"
            CompanyName       = "stealthpuppy"
            TenantId          = "6cdd8179-23e5-43d1-8517-b6276a8d3189"
            SkipImport        = $false
        }
        .\New-Microsoft365AppsPackage.ps1 @params

    .NOTES
        Author: Aaron Parker
        Bluesky: @stealthpuppy.com
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to the top level directory of the repository.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ if (Test-Path -Path $_ -PathType "Container") { $true } else { throw "Path not found: '$_'" } })]
    [System.String] $Path = $PSScriptRoot,

    [Parameter(Mandatory = $false, HelpMessage = "Path where the package will be created.")]
    [ValidateNotNullOrEmpty()]
    [System.String] $Destination = "$Path\package",

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

    [Parameter(Mandatory = $true, HelpMessage = "The tenant id (GUID) of the target Entra ID tenant.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ $ObjectGuid = [System.Guid]::empty; if ([System.Guid]::TryParse($_, [System.Management.Automation.PSReference]$ObjectGuid)) { $true } else { throw "$_ is not a GUID" } })]
    [System.String] $TenantId,

    [Parameter(Mandatory = $false, HelpMessage = "Wrap the Microsoft 365 Apps installer with the PowerShell App Deployment Toolkit.")]
    [System.Management.Automation.SwitchParameter] $UsePsadt,

    [Parameter(Mandatory = $false, HelpMessage = "Validate package creation without executing.")]
    [System.Management.Automation.SwitchParameter] $ValidateOnly,

    [Parameter(Mandatory = $false, HelpMessage = "Skip Intune import operations.")]
    [System.Management.Automation.SwitchParameter] $SkipImport,

    [Parameter(Mandatory = $false, HelpMessage = "Force import even if the same version already exists.")]
    [System.Management.Automation.SwitchParameter] $Force
)

begin {
    # Configure the environment
    $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    $ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

    # Import modules
    Import-Module -Name "$PSScriptRoot\Microsoft365AppsPackage.psm1" -Force

    # Validate prerequisites
    Test-PackagePrerequisites -Path $Path -ConfigurationFile $ConfigurationFile -TenantId $TenantId

    # Validate that the input file is XML and read it
    Write-Msg -Msg "Read configuration file: $ConfigurationFile."
    $Xml = Import-XmlFile -FilePath $ConfigurationFile

    # Unblock all files in the repo
    Write-Msg -Msg "Unblock exe files in $Path."
    Get-ChildItem -Path $Path -Recurse -Include "*.exe" | Unblock-File
}

process {
    # If ValidateOnly is specified, run validation and exit
    if ($ValidateOnly) {
        $validationResults = Invoke-PackageValidation -Path $Path -Destination $Destination -ConfigurationFile $ConfigurationFile -Channel $Channel -CompanyName $CompanyName -TenantId $TenantId -UsePsadt $UsePsadt.IsPresent
        return $validationResults
    }

    #region Initialize package structure and copy files
    Invoke-WithErrorHandling -Operation "Initialize package structure" -ScriptBlock {
        Initialize-PackageStructure -Destination $Destination -Path $Path -ConfigurationFile $ConfigurationFile -UsePsadt $UsePsadt.IsPresent
    }
    #endregion

    #region Update the configuration.xml
    $xml = Invoke-WithErrorHandling -Operation "Update configuration" -ScriptBlock {
        $configPath = if ($UsePsadt) { "$Destination\source\Files\Install-Microsoft365Apps.xml" } else { "$Destination\source\Install-Microsoft365Apps.xml" }
        Update-M365Configuration -ConfigurationPath $configPath -Channel $Channel -TenantId $TenantId -CompanyName $CompanyName
    }
    #endregion

    #region Create the intunewin package
    Invoke-WithErrorHandling -Operation "Create intunewin package" -ScriptBlock {
        Write-Msg -Msg "Create intunewin package in: $Destination\output."
        $params = @{
            SourceFolder         = "$Destination\source"
            SetupFile            = if ($UsePsadt) { "Files\setup.exe" } else { "setup.exe" }
            OutputFolder         = "$Destination\output"
            Force                = $true
            IntuneWinAppUtilPath = "$Path\intunewin\IntuneWinAppUtil.exe"
        }
        New-IntuneWin32AppPackage @params
    }
    #endregion

    # Save a copy of the modified configuration file to the output folder for reference
    $OutputXml = "$Destination\output\$(Split-Path -Path $ConfigurationFile -Leaf)"
    Write-Msg -Msg "Saved configuration file to: $OutputXml."
    $Xml.Save($OutputXml)

    #region Create and update package manifest
    $manifest = Invoke-WithErrorHandling -Operation "Create package manifest" -ScriptBlock {
        New-PackageManifest -Xml $xml -Destination $Destination -Path $Path -ConfigurationFile $ConfigurationFile -Channel $Channel -UsePsadt $UsePsadt.IsPresent
    }
    #endregion

    #region Check for existing application and determine if update is needed
    Write-Msg -Msg "Retrieve existing Microsoft 365 Apps packages from Intune"
    $ExistingApp = Get-M365AppsFromIntune -PackageId $manifest.Information.PSPackageFactoryGuid | `
        Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
        Select-Object -First 1

    $UpdateApp = Test-ShouldUpdateApp -Manifest $manifest -ExistingApp $ExistingApp -Force $Force.IsPresent
    #endregion

    if ($UpdateApp -and -not $SkipImport) {
        Invoke-WithErrorHandling -Operation "Import package to Intune" -ScriptBlock {
            Write-Msg -Msg "-Import specified. Importing package into tenant."

            # Get the package file
            $PackageFile = Get-ChildItem -Path "$Destination\output" -Recurse -Include "setup.intunewin"
            if ($null -eq $PackageFile) {
                throw [System.IO.FileNotFoundException]::New("Intunewin package file not found.")
            }

            # Use the integrated function to import the package
            Write-Msg -Msg "Create package using integrated function."
            $params = @{
                Json        = "$Destination\output\m365apps.json"
                PackageFile = $PackageFile.FullName
            }
            $ImportedApp = New-IntuneWin32AppFromManifest @params | Select-Object -Property * -ExcludeProperty "largeIcon"
            Write-Msg -Msg "Package import complete."

            #region Add supersedence for existing packages
            Write-Msg -Msg "Retrieve Microsoft 365 Apps packages from Intune for supersedence"
            $Supersedence = Get-M365AppsFromIntune -PackageId $manifest.Information.PSPackageFactoryGuid | `
                Where-Object { $_.id -ne $ImportedApp.id } | `
                Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
                ForEach-Object { New-IntuneWin32AppSupersedence -ID $_.id -SupersedenceType "Update" }
            if ($null -ne $Supersedence) {
                Add-IntuneWin32AppSupersedence -ID $ImportedApp.id -Supersedence $Supersedence
            }
            #endregion

            # Output imported application details
            $ImportedApp
        }
    } elseif ($SkipImport) {
        Write-Msg -Msg "Skipping Intune import due to -SkipImport parameter."
    }
}

end {
}
