using namespace System.Management.Automation

# Configure the environment
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

function Write-Msg ($Msg) {
    $InformationPreference = [System.Management.Automation.ActionPreference]::continue
    $DateTime = $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')
    if ($IsCoreCLR) {
        $params = @{
            MessageData = "$($PSStyle.Foreground.Cyan)[$DateTime]$($PSStyle.Reset) $Msg"
            Tags        = "Microsoft365"
        }
        Write-Information @params
    }
    else {
        $Message = [HostInformationMessage]@{
            Message         = "[$DateTime]"
            ForegroundColor = "Cyan"
            NoNewline       = $true
        }
        $params = @{
            MessageData = $Message
            Tags        = "Microsoft365"
        }
        Write-Information @params
        $params = @{
            MessageData = " $Msg"
            Tags        = "Microsoft365"
        }
        Write-Information @params
    }
}

function Import-XmlFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.String]$FilePath
    )

    if (Test-Path -Path $FilePath -PathType "Leaf") {
        try {
            $xmlContent = [System.Xml.XmlDocument](Get-Content -Path $FilePath)
            return $xmlContent
        }
        catch {
            Write-Error -Message "Failed to read XML file: $_"
            return $null
        }
    }
    else {
        Write-Error -Message "File not found: $FilePath"
        return $null
    }
}

function Test-RequiredFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.String]$Path
    )
    @(
        "$Path\configs\Uninstall-Microsoft365Apps.xml",
        "$Path\intunewin\IntuneWinAppUtil.exe",
        "$Path\m365\setup.exe",
        "$Path\icons\Microsoft365.png",
        "$Path\scripts\App.json",
        "$Path\scrub\OffScrub03.vbs",
        "$Path\scrub\OffScrub07.vbs",
        "$Path\scrub\OffScrub10.vbs",
        "$Path\scrub\OffScrubc2r.vbs",
        "$Path\scrub\OffScrub_O15msi.vbs",
        "$Path\scrub\OffScrub_O16msi.vbs"
    ) | ForEach-Object { if (-not (Test-Path -Path $_)) { throw [System.IO.FileNotFoundException]::New("File not found: $_") } }
}

function Test-Destination {
    param(
        [Parameter(Mandatory = $true)]
        [System.String]$Path
    )

    if (-not (Test-Path -Path $Path -PathType "Container")) {
        throw "'$Path' does not exist or is not a directory."
    }
    if ((Get-ChildItem -Path $Path -Recurse -File).Count -gt 0) {
        throw "'$Path' is not empty. Remove path and try again."
    }
}

function Get-M365AppsFromIntune {
    [CmdletBinding(SupportsShouldProcess = $false)]
    param (
        [System.String[]] $PackageId
    )

    begin {
        $DisplayNamePattern = "^Microsoft 365 Apps*"
        $NotesPattern = '^{"CreatedBy":"PSPackageFactory","Guid":.*}$'

        # Get the existing Win32 applications from Intune
        Write-Verbose -Message "Retrieving existing Win32 applications from Intune."
        $ExistingIntuneApps = Get-IntuneWin32App | `
            Where-Object { $_.displayName -match $DisplayNamePattern -and $_.notes -match $NotesPattern } | `
            Select-Object -Property * -ExcludeProperty "largeIcon"
        if ($ExistingIntuneApps.Count -gt 0) {
            Write-Verbose -Message "Found $($ExistingIntuneApps.Count) existing Microsoft 365 Apps packages in Intune."
        }
    }

    process {
        foreach ($Id in $PackageId) {
            Write-Verbose -Message "Filtering existing applications to match VcList PackageId."
            foreach ($Application in $ExistingIntuneApps) {
                if (($Application.notes | ConvertFrom-Json -ErrorAction "Stop").Guid -in $Id) {

                    # Add the packageId to the application object for easier reference
                    $Application | Add-Member -MemberType "NoteProperty" -Name "packageId" -Value $($Application.notes | ConvertFrom-Json -ErrorAction "Stop").Guid -Force
                    Write-Output -InputObject $Application
                }
            }
        }
    }
}

function Get-RequiredM365AppsUpdatesFromIntune {
    [CmdletBinding(SupportsShouldProcess = $false)]
    param (
        [System.String[]] $PackageId
    )

    begin {
        # Get the existing VcRedist Win32 applications from Intune
        $ExistingIntuneApps = Get-VcRedistAppsFromIntune -VcList $VcList
    }

    process {
        foreach ($Application in $ExistingIntuneApps) {
            $VcRedist = $PackageId | Where-Object { $_.PackageId -eq $Application.packageId }
            if ($null -eq $VcRedist) {
                Write-Verbose -Message "No matching VcRedist found for application with ID: $($Application.Id). Skipping."
                continue
            }
            else {
                $Update = $false
                if ([System.Version]$VcRedist.Version -gt [System.Version]$Application.displayVersion) {
                    $Update = $true
                    Write-Verbose -Message "Update required for $($Application.displayName): $($VcRedist.Version) > $($Application.displayVersion)."
                }
                $Object = [PSCustomObject]@{
                    "AppId"          = $Application.Id
                    "IntuneVersion"  = $Application.displayVersion
                    "UpdateVersion"  = $VcRedist.Version
                    "UpdateRequired" = $Update
                }
                Write-Output -InputObject $Object
            }
        }
    }
}

function Test-PackagePrerequisites {
    <#
    .SYNOPSIS
        Validates all prerequisites for package creation.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFile,

        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [string]$ClientId
    )

    Write-Msg -Msg "Validating package prerequisites."

    # Test paths
    if (-not (Test-Path -Path $Path -PathType Container)) {
        throw "Path not found: '$Path'"
    }

    if (-not (Test-Path -Path $ConfigurationFile -PathType Leaf)) {
        throw "Configuration file not found: '$ConfigurationFile'"
    }

    if ((Get-Item -Path $ConfigurationFile).Extension -ne ".xml") {
        throw "Configuration file is not an XML file: '$ConfigurationFile'"
    }

    # Test GUID format
    $guidTest = [System.Guid]::empty
    if (-not [System.Guid]::TryParse($TenantId, [ref]$guidTest)) {
        throw "TenantId is not a valid GUID: '$TenantId'"
    }

    if ($ClientId -and -not [System.Guid]::TryParse($ClientId, [ref]$guidTest)) {
        throw "ClientId is not a valid GUID: '$ClientId'"
    }

    # Test required files
    Test-RequiredFiles -Path $Path
    Write-Msg -Msg "Prerequisites validation completed successfully."
}

function Initialize-PackageStructure {
    <#
    .SYNOPSIS
        Creates the package directory structure and copies required files.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFile,

        [Parameter(Mandatory = $false)]
        [bool]$UsePsadt = $false
    )

    Write-Msg -Msg "Initializing package structure."

    # Set output directory and ensure it is empty
    Test-Destination -Path $Destination

    # Create the package directory structure
    Write-Msg -Msg "Create new package structure at: $Destination"
    New-Item -Path "$Destination\source" -ItemType "Directory" -ErrorAction "SilentlyContinue" | Out-Null
    New-Item -Path "$Destination\output" -ItemType "Directory" -ErrorAction "SilentlyContinue" | Out-Null

    if ($UsePsadt) {
        Copy-PsadtFiles -Destination $Destination -Path $Path -ConfigurationFile $ConfigurationFile
    }
    else {
        Copy-StandardFiles -Destination $Destination -Path $Path -ConfigurationFile $ConfigurationFile
    }

    Write-Msg -Msg "Package structure initialization completed."
    return $Destination
}

function Copy-PsadtFiles {
    <#
    .SYNOPSIS
        Copies files for PSADT package structure.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFile
    )

    # Create a PSADT template
    Write-Msg -Msg "Create PSADT template"
    New-ADTTemplate -Destination "$Env:TEMP\psadt" -Force
    $PsAdtSource = Get-ChildItem -Path "$Env:TEMP\psadt" -Directory -Filter "PSAppDeployToolkit*"
    Copy-Item -Path "$($PsAdtSource.FullName)\*" -Destination "$Destination\source" -Recurse -Force

    # Copy the PSAppDeployToolkit files to the package source
    Write-Msg -Msg "Copy Office scrub scripts to: $Destination\source\SupportFiles."
    Copy-Item -Path "$Path\scrub\*" -Destination "$Destination\source\SupportFiles" -Recurse
    New-Item -Path "$Destination\source\Files" -ItemType "Directory" -ErrorAction "SilentlyContinue" | Out-Null
    Write-Msg -Msg "Copy Invoke-AppDeployToolkit.ps1 to: $Destination\source\Invoke-AppDeployToolkit.ps1."
    Copy-Item -Path "$Path\scripts\Invoke-AppDeployToolkit.ps1" -Destination "$Destination\source\Invoke-AppDeployToolkit.ps1" -Force

    # Copy the configuration files and setup.exe to the package source
    Write-Msg -Msg "Copy configuration files and setup.exe to package source."
    Copy-Item -Path $ConfigurationFile -Destination "$Destination\source\Files\Install-Microsoft365Apps.xml"
    Copy-Item -Path "$Path\configs\Uninstall-Microsoft365Apps.xml" -Destination "$Destination\source\Files\Uninstall-Microsoft365Apps.xml"
    Copy-Item -Path "$Path\m365\setup.exe" -Destination "$Destination\source\Files\setup.exe"
}

function Copy-StandardFiles {
    <#
    .SYNOPSIS
        Copies files for standard package structure.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFile
    )

    # Copy the configuration files and setup.exe to the package source
    Write-Msg -Msg "Copy configuration files and setup.exe to package source."
    Copy-Item -Path $ConfigurationFile -Destination "$Destination\source\Install-Microsoft365Apps.xml"
    Copy-Item -Path "$Path\configs\Uninstall-Microsoft365Apps.xml" -Destination "$Destination\source\Uninstall-Microsoft365Apps.xml"
    Copy-Item -Path "$Path\m365\setup.exe" -Destination "$Destination\source\setup.exe"
}

function Update-M365Configuration {
    <#
    .SYNOPSIS
        Updates the Microsoft 365 Apps configuration XML file.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigurationPath,

        [Parameter(Mandatory = $true)]
        [string]$Channel,

        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$CompanyName
    )

    Write-Msg -Msg "Updating configuration file: $ConfigurationPath"
    [System.Xml.XmlDocument]$Xml = Get-Content -Path $ConfigurationPath

    Write-Msg -Msg "Set Microsoft 365 Apps channel to: $Channel."
    $Xml.Configuration.Add.Channel = $Channel

    Write-Msg -Msg "Set tenant id to: $TenantId."
    $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "TenantId"))
    $Xml.Configuration.Property[$Index].Value = $TenantId

    Write-Msg -Msg "Set company name to: $CompanyName."
    $Xml.Configuration.AppSettings.Setup.Value = $CompanyName

    Write-Msg -Msg "Save configuration xml to: $ConfigurationPath."
    $Xml.Save($ConfigurationPath)

    return $Xml
}

function New-PackageManifest {
    <#
    .SYNOPSIS
        Creates and updates the package manifest (App.json).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ConfigurationFile,

        [Parameter(Mandatory = $true)]
        [string]$Channel,

        [Parameter(Mandatory = $false)]
        [bool]$UsePsadt = $false
    )

    Write-Msg -Msg "Creating package manifest."

    $SetupVersion = (Get-Item -Path "$Path\m365\setup.exe").VersionInfo.FileVersion
    Write-Msg -Msg "Using setup.exe version: $SetupVersion."

    # Copy appropriate App.json template
    $sourceJson = if ($UsePsadt) { "$Path\scripts\App.json" } else { "$Path\scripts\AppNoPsadt.json" }
    Copy-Item -Path $sourceJson -Destination "$Destination\output\m365apps.json"

    # Load and update manifest
    $Manifest = Get-Content -Path "$Destination\output\m365apps.json" | ConvertFrom-Json
    $Manifest.PackageInformation.Version = $SetupVersion

    # Update package display name
    $DisplayName = Get-PackageDisplayName -Xml $Xml
    Write-Msg -Msg "Package display name: $DisplayName."
    $Manifest.Information.DisplayName = $DisplayName

    # Set the PSPackageFactory GUID
    $Manifest.Information.PSPackageFactoryGuid = $Xml.Configuration.ID

    # Update icon location
    $Manifest.PackageInformation.IconFile = "$Path\icons\Microsoft365.png"

    # Update package description
    $Description = "$($xml.Configuration.Info.Description)`n`n**This package uses the PSAppDeployToolkit and will uninstall previous versions of Microsoft Office**. Uses setup.exe $SetupVersion. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
    $Manifest.Information.Description = $Description

    # Update detection rules
    Update-DetectionRules -Manifest $Manifest -Xml $Xml -Channel $Channel

    # Save updated manifest
    $Manifest | ConvertTo-Json | Out-File -FilePath "$Destination\output\m365apps.json" -Force
    return $Manifest
}

function Get-PackageDisplayName {
    <#
    .SYNOPSIS
        Generates the package display name based on configuration.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Xml
    )

    [System.String] $ProductID = ""
    switch ($Xml.Configuration.Add.Product.ID) {
        "O365ProPlusRetail" { $ProductID += "Microsoft 365 Apps for enterprise, " }
        "O365BusinessRetail" { $ProductID += "Microsoft 365 Apps for business, " }
        "VisioProRetail" { $ProductID += "Visio, " }
        "ProjectProRetail" { $ProductID += "Project, " }
        "AccessRuntimeRetail" { $ProductID += "Access Runtime, " }
    }

    [System.String] $DisplayName = "$ProductID$($Xml.Configuration.Add.Channel)"
    if ($Xml.Configuration.Add.OfficeClientEdition -eq "64") { $DisplayName = "$DisplayName, x64" }
    if ($Xml.Configuration.Add.OfficeClientEdition -eq "32") { $DisplayName = "$DisplayName, x86" }
    return $DisplayName
}

function Update-DetectionRules {
    <#
    .SYNOPSIS
        Updates the detection rules in the manifest.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Manifest,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$Channel
    )

    # Update ProductReleaseIds detection rule
    $ProductReleaseIDs = ($Xml.Configuration.Add.Product.ID | Sort-Object) -join ","
    $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "ProductReleaseIds"))
    Write-Msg -Msg "Update registry ProductReleaseIds detection rule: $ProductReleaseIDs."
    $Manifest.DetectionRule[$Index].Value = $ProductReleaseIDs

    # Update VersionToReport detection rule
    $ChannelVersion = Get-EvergreenApp -Name "Microsoft365Apps" | Where-Object { $_.Channel -eq $Channel }
    $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "VersionToReport"))
    Write-Msg -Msg "Update registry VersionToReport detection rule: $($ChannelVersion.Version)."
    $Manifest.DetectionRule[$Index].Value = $ChannelVersion.Version

    # Update SharedComputerLicensing detection rule
    $Index = $Manifest.DetectionRule.IndexOf($($Manifest.DetectionRule -cmatch "SharedComputerLicensing"))
    $Value = ($Xml.Configuration.Property | Where-Object { $_.Name -eq "SharedComputerLicensing" }).Value
    Write-Msg -Msg "Update registry SharedComputerLicensing detection rule: $Value."
    $Manifest.DetectionRule[$Index].Value = $Value
}

function Connect-IntuneService {
    <#
    .SYNOPSIS
        Connects to Intune service if credentials are provided.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [string]$ClientId,

        [Parameter(Mandatory = $false)]
        [string]$ClientSecret
    )

    if ($ClientId -and $ClientSecret) {
        $params = @{
            TenantId     = $TenantId
            ClientId     = $ClientId
            ClientSecret = $ClientSecret
        }
        Write-Msg -Msg "Authenticate to tenant: $TenantId."
        [Void](Connect-MSIntuneGraph @params)
        return $true
    }
    return $false
}

function Test-ShouldUpdateApp {
    <#
    .SYNOPSIS
        Determines if an application should be updated based on version comparison.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Manifest,

        [Parameter(Mandatory = $false)]
        [PSObject]$ExistingApp,

        [Parameter(Mandatory = $false)]
        [bool]$Force = $false
    )

    if ($Force) {
        Write-Msg -Msg "Force parameter specified. Package will be imported."
        return $true
    }

    if ($null -eq $ExistingApp) {
        Write-Msg -Msg "Import new application: '$($Manifest.Information.DisplayName)'"
        return $true
    }

    if ([System.String]::IsNullOrEmpty($ExistingApp.displayVersion)) {
        Write-Msg -Msg "Found matching app but displayVersion is null: '$($ExistingApp.displayName)'"
        return $false
    }

    if ($Manifest.PackageInformation.Version -le $ExistingApp.displayVersion) {
        Write-Msg -Msg "Existing Intune app version is current: '$($ExistingApp.displayName), $($ExistingApp.displayVersion)'"
        return $false
    }

    if ($Manifest.PackageInformation.Version -gt $ExistingApp.displayVersion) {
        Write-Msg -Msg "Import application version: '$($Manifest.Information.DisplayName), $($ExistingApp.displayVersion)'"
        return $true
    }

    return $false
}

function Invoke-WithErrorHandling {
    <#
    .SYNOPSIS
        Executes a script block with centralized error handling.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    try {
        Write-Msg -Msg "Starting: $Operation"
        $result = & $ScriptBlock
        Write-Msg -Msg "Completed: $Operation"
        return $result
    }
    catch {
        Write-Error "Failed during $Operation`: $_"
        throw
    }
}

# Configuration classes for simplified parameters
class PackageConfig {
    [string]$Path
    [string]$ConfigurationFile
    [string]$Channel
    [string]$CompanyName
    [bool]$UsePsadt
    [string]$Destination
}

class AuthConfig {
    [string]$TenantId
    [string]$ClientId
    [string]$ClientSecret
}

function New-Microsoft365AppsPackageSimplified {
    <#
    .SYNOPSIS
        Simplified entry point for creating Microsoft 365 Apps packages.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PackageConfig]$PackageConfig,

        [Parameter(Mandatory = $true)]
        [AuthConfig]$AuthConfig,

        [Parameter(Mandatory = $false)]
        [bool]$Import = $false,

        [Parameter(Mandatory = $false)]
        [bool]$Force = $false
    )

    # Single validation call
    Test-PackagePrerequisites -Path $PackageConfig.Path -ConfigurationFile $PackageConfig.ConfigurationFile -TenantId $AuthConfig.TenantId -ClientId $AuthConfig.ClientId

    # Read initial XML
    $xml = Import-XmlFile -FilePath $PackageConfig.ConfigurationFile

    # Unblock files
    Get-ChildItem -Path $PackageConfig.Path -Recurse -Include "*.exe" | Unblock-File

    # Simplified workflow
    Invoke-WithErrorHandling -Operation "Initialize package structure" -ScriptBlock {
        Initialize-PackageStructure -Destination $PackageConfig.Destination -Path $PackageConfig.Path -ConfigurationFile $PackageConfig.ConfigurationFile -UsePsadt $PackageConfig.UsePsadt
    }

    $xml = Invoke-WithErrorHandling -Operation "Update configuration" -ScriptBlock {
        $configPath = if ($PackageConfig.UsePsadt) { "$($PackageConfig.Destination)\source\Files\Install-Microsoft365Apps.xml" } else { "$($PackageConfig.Destination)\source\Install-Microsoft365Apps.xml" }
        Update-M365Configuration -ConfigurationPath $configPath -Channel $PackageConfig.Channel -TenantId $AuthConfig.TenantId -CompanyName $PackageConfig.CompanyName
    }

    Invoke-WithErrorHandling -Operation "Create intunewin package" -ScriptBlock {
        $params = @{
            SourceFolder         = "$($PackageConfig.Destination)\source"
            SetupFile            = if ($PackageConfig.UsePsadt) { "Files\setup.exe" } else { "setup.exe" }
            OutputFolder         = "$($PackageConfig.Destination)\output"
            Force                = $true
            IntuneWinAppUtilPath = "$($PackageConfig.Path)\intunewin\IntuneWinAppUtil.exe"
        }
        New-IntuneWin32AppPackage @params
    }

    # Save configuration file copy
    $OutputXml = "$($PackageConfig.Destination)\output\$(Split-Path -Path $PackageConfig.ConfigurationFile -Leaf)"
    $xml.Save($OutputXml)

    $manifest = Invoke-WithErrorHandling -Operation "Create package manifest" -ScriptBlock {
        New-PackageManifest -Xml $xml -Destination $PackageConfig.Destination -Path $PackageConfig.Path -ConfigurationFile $PackageConfig.ConfigurationFile -Channel $PackageConfig.Channel -UsePsadt $PackageConfig.UsePsadt
    }

    if ($Import) {
        Connect-IntuneService -TenantId $AuthConfig.TenantId -ClientId $AuthConfig.ClientId -ClientSecret $AuthConfig.ClientSecret

        # Get existing app
        $ExistingApp = Get-IntuneWin32App | `
            Select-Object -Property * -ExcludeProperty "largeIcon" | `
            Where-Object { $_.notes -match "PSPackageFactory" } | `
            Where-Object { ($_.notes | ConvertFrom-Json -ErrorAction "SilentlyContinue").Guid -eq $manifest.Information.PSPackageFactoryGuid } | `
            Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
            Select-Object -First 1

        $UpdateApp = Test-ShouldUpdateApp -Manifest $manifest -ExistingApp $ExistingApp -Force $Force

        if ($UpdateApp) {
            return Invoke-WithErrorHandling -Operation "Import package to Intune" -ScriptBlock {
                $PackageFile = Get-ChildItem -Path "$($PackageConfig.Destination)\output" -Recurse -Include "setup.intunewin"
                if ($null -eq $PackageFile) {
                    throw [System.IO.FileNotFoundException]::New("Intunewin package file not found.")
                }

                $params = @{
                    Json        = "$($PackageConfig.Destination)\output\m365apps.json"
                    PackageFile = $PackageFile.FullName
                }
                $ImportedApp = New-IntuneWin32AppFromManifest @params | Select-Object -Property * -ExcludeProperty "largeIcon"

                # Add supersedence
                $Supersedence = Get-IntuneWin32App | `
                    Where-Object { $_.id -ne $ImportedApp.id } | `
                    Where-Object { $_.notes -match "PSPackageFactory" } | `
                    Where-Object { ($_.notes | ConvertFrom-Json -ErrorAction "SilentlyContinue").Guid -eq $manifest.Information.PSPackageFactoryGuid } | `
                    Select-Object -Property * -ExcludeProperty "largeIcon" | `
                    Sort-Object -Property @{ Expression = { [System.Version]$_.displayVersion }; Descending = $true } -ErrorAction "SilentlyContinue" | `
                    ForEach-Object { New-IntuneWin32AppSupersedence -ID $_.id -SupersedenceType "Update" }

                if ($null -ne $Supersedence) {
                    Add-IntuneWin32AppSupersedence -ID $ImportedApp.id -Supersedence $Supersedence
                }

                return $ImportedApp
            }
        }
    }

    return $manifest
}

function New-IntuneWin32AppFromManifest {
    <#
    .SYNOPSIS
        Create a Win32 app in Microsoft Intune based on input from app manifest file.

    .DESCRIPTION
        Create a Win32 app in Microsoft Intune based on input from app manifest file.
        This function is based on the Create-Win32App.ps1 script but integrated as a module function.

    .PARAMETER Json
        Specify the application JSON definition file path.

    .PARAMETER PackageFile
        Specify the application package file path.

    .PARAMETER ScriptsFolder
        Specify the scripts folder path. Defaults to PSScriptRoot\Scripts when used in script context.

    .PARAMETER Validate
        Specify to validate manifest file configuration.

    .EXAMPLE
        New-IntuneWin32AppFromManifest -Json "C:\temp\app.json" -PackageFile "C:\temp\app.intunewin"

    .NOTES
        This function replaces the Create-Win32App.ps1 script for integration into the module.
        Original script author: Nickolaj Andersen (@NickolajA)
        Updated for module integration: Aaron Parker (@stealthpuppy)
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Specify the application JSON definition.")]
        [ValidateNotNullOrEmpty()]
        [System.String] $Json,

        [Parameter(Mandatory = $true, HelpMessage = "Specify the application package path.")]
        [ValidateNotNullOrEmpty()]
        [System.String] $PackageFile,

        [Parameter(Mandatory = $false, HelpMessage = "Specify the scripts folder path.")]
        [ValidateNotNullOrEmpty()]
        [System.String] $ScriptsFolder,

        [Parameter(Mandatory = $false, HelpMessage = "Specify to validate manifest file configuration.")]
        [System.Management.Automation.SwitchParameter] $Validate
    )

    begin {
        Write-Msg -Msg "Starting Win32 app creation from manifest: $Json"
    }

    process {
        # Read app data from JSON manifest
        $AppData = Get-Content -Path $Json | ConvertFrom-Json

        # Set default scripts folder if not provided
        if ([System.String]::IsNullOrEmpty($ScriptsFolder)) {
            $ScriptsFolder = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($Json), "Scripts")
        }

        # Icon file - download the file, if the property is a URL
        if ($AppData.PackageInformation.IconFile -match "^http") {
            $OutFile = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($Json), $(Split-Path -Path $AppData.PackageInformation.IconFile -Leaf))
            $params = @{
                Uri             = $AppData.PackageInformation.IconFile
                OutFile         = $OutFile
                UseBasicParsing = $true
            }
            Invoke-WebRequest @params
            $AppIconFile = $OutFile
        }
        else {
            $AppIconFile = $AppData.PackageInformation.IconFile
        }

        # Create default requirement rule
        $params = @{
            Architecture                   = $AppData.RequirementRule.Architecture
            MinimumSupportedWindowsRelease = $AppData.RequirementRule.MinimumRequiredOperatingSystem
        }
        $RequirementRule = New-IntuneWin32AppRequirementRule @params

        # Create additional custom requirement rules
        $CustomRequirementRuleCount = ($AppData.CustomRequirementRule | Measure-Object).Count
        if ($CustomRequirementRuleCount -ge 1) {
            $RequirementRules = New-Object -TypeName "System.Collections.ArrayList"
            foreach ($RequirementRuleItem in $AppData.CustomRequirementRule) {
                switch ($RequirementRuleItem.Type) {
                    "File" {
                        switch ($RequirementRuleItem.DetectionMethod) {
                            "Existence" {
                                # Create a custom file based requirement rule
                                $RequirementRuleArgs = @{
                                    "Existence"            = $true
                                    "Path"                 = $RequirementRuleItem.Path
                                    "FileOrFolder"         = $RequirementRuleItem.FileOrFolder
                                    "DetectionType"        = $RequirementRuleItem.DetectionType
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "DateModified" {
                                # Create a custom file based requirement rule
                                $RequirementRuleArgs = @{
                                    "DateModified"         = $true
                                    "Path"                 = $RequirementRuleItem.Path
                                    "FileOrFolder"         = $RequirementRuleItem.FileOrFolder
                                    "Operator"             = $RequirementRuleItem.Operator
                                    "DateTimeValue"        = $RequirementRuleItem.DateTimeValue
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "DateCreated" {
                                # Create a custom file based requirement rule
                                $RequirementRuleArgs = @{
                                    "DateCreated"          = $true
                                    "Path"                 = $RequirementRuleItem.Path
                                    "FileOrFolder"         = $RequirementRuleItem.FileOrFolder
                                    "Operator"             = $RequirementRuleItem.Operator
                                    "DateTimeValue"        = $RequirementRuleItem.DateTimeValue
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "Version" {
                                # Create a custom file based requirement rule
                                $RequirementRuleArgs = @{
                                    "Version"              = $true
                                    "Path"                 = $RequirementRuleItem.Path
                                    "FileOrFolder"         = $RequirementRuleItem.FileOrFolder
                                    "Operator"             = $RequirementRuleItem.Operator
                                    "VersionValue"         = $RequirementRuleItem.VersionValue
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "Size" {
                                # Create a custom file based requirement rule
                                $RequirementRuleArgs = @{
                                    "Size"                 = $true
                                    "Path"                 = $RequirementRuleItem.Path
                                    "FileOrFolder"         = $RequirementRuleItem.FileOrFolder
                                    "Operator"             = $RequirementRuleItem.Operator
                                    "SizeInMBValue"        = $RequirementRuleItem.SizeInMBValue
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                        }

                        # Create file based requirement rule
                        $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                    }
                    "Registry" {
                        switch ($RequirementRuleItem.DetectionMethod) {
                            "Existence" {
                                # Create a custom registry based requirement rule
                                $RequirementRuleArgs = @{
                                    "Existence"            = $true
                                    "KeyPath"              = $RequirementRuleItem.KeyPath
                                    "ValueName"            = $RequirementRuleItem.ValueName
                                    "DetectionType"        = $RequirementRuleItem.DetectionType
                                    "Check32BitOn64System" = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "StringComparison" {
                                # Create a custom registry based requirement rule
                                $RequirementRuleArgs = @{
                                    "StringComparison"         = $true
                                    "KeyPath"                  = $RequirementRuleItem.KeyPath
                                    "ValueName"                = $RequirementRuleItem.ValueName
                                    "StringComparisonOperator" = $RequirementRuleItem.Operator
                                    "StringComparisonValue"    = $RequirementRuleItem.Value
                                    "Check32BitOn64System"     = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "VersionComparison" {
                                # Create a custom registry based requirement rule
                                $RequirementRuleArgs = @{
                                    "VersionComparison"         = $true
                                    "KeyPath"                   = $RequirementRuleItem.KeyPath
                                    "ValueName"                 = $RequirementRuleItem.ValueName
                                    "VersionComparisonOperator" = $RequirementRuleItem.Operator
                                    "VersionComparisonValue"    = $RequirementRuleItem.Value
                                    "Check32BitOn64System"      = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                            "IntegerComparison" {
                                # Create a custom registry based requirement rule
                                $RequirementRuleArgs = @{
                                    "IntegerComparison"         = $true
                                    "KeyPath"                   = $RequirementRuleItem.KeyPath
                                    "ValueName"                 = $RequirementRuleItem.ValueName
                                    "IntegerComparisonOperator" = $RequirementRuleItem.Operator
                                    "IntegerComparisonValue"    = $RequirementRuleItem.Value
                                    "Check32BitOn64System"      = [System.Convert]::ToBoolean($RequirementRuleItem.Check32BitOn64System)
                                }
                            }
                        }

                        # Create registry based requirement rule
                        $CustomRequirementRule = New-IntuneWin32AppRequirementRuleRegistry @RequirementRuleArgs
                    }
                    "Script" {
                        switch ($RequirementRuleItem.DetectionMethod) {
                            "StringOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "StringOutputDataType"     = $true
                                    "ScriptFile"               = (Join-Path -Path $ScriptsFolder -ChildPath $RequirementRuleItem.ScriptFile)
                                    "ScriptContext"            = $RequirementRuleItem.ScriptContext
                                    "StringComparisonOperator" = $RequirementRuleItem.Operator
                                    "StringValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"     = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"    = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                            "IntegerOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "IntegerOutputDataType"     = $true
                                    "ScriptFile"                = $RequirementRuleItem.ScriptFile
                                    "ScriptContext"             = $RequirementRuleItem.ScriptContext
                                    "IntegerComparisonOperator" = $RequirementRuleItem.Operator
                                    "IntegerValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"      = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"     = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                            "BooleanOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "BooleanOutputDataType"     = $true
                                    "ScriptFile"                = $RequirementRuleItem.ScriptFile
                                    "ScriptContext"             = $RequirementRuleItem.ScriptContext
                                    "BooleanComparisonOperator" = $RequirementRuleItem.Operator
                                    "BooleanValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"      = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"     = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                            "DateTimeOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "DateTimeOutputDataType"     = $true
                                    "ScriptFile"                 = $RequirementRuleItem.ScriptFile
                                    "ScriptContext"              = $RequirementRuleItem.ScriptContext
                                    "DateTimeComparisonOperator" = $RequirementRuleItem.Operator
                                    "DateTimeValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"       = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"      = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                            "FloatOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "FloatOutputDataType"     = $true
                                    "ScriptFile"              = $RequirementRuleItem.ScriptFile
                                    "ScriptContext"           = $RequirementRuleItem.ScriptContext
                                    "FloatComparisonOperator" = $RequirementRuleItem.Operator
                                    "FloatValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"    = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"   = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                            "VersionOutput" {
                                # Create a custom script based requirement rule
                                $RequirementRuleArgs = @{
                                    "VersionOutputDataType"     = $true
                                    "ScriptFile"                = $RequirementRuleItem.ScriptFile
                                    "ScriptContext"             = $RequirementRuleItem.ScriptContext
                                    "VersionComparisonOperator" = $RequirementRuleItem.Operator
                                    "VersionValue"              = $RequirementRuleItem.Value
                                    "RunAs32BitOn64System"      = [System.Convert]::ToBoolean($RequirementRuleItem.RunAs32BitOn64System)
                                    "EnforceSignatureCheck"     = [System.Convert]::ToBoolean($RequirementRuleItem.EnforceSignatureCheck)
                                }
                            }
                        }

                        # Create script based requirement rule
                        $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                    }
                }

                # Add requirement rule to list
                $RequirementRules.Add($CustomRequirementRule) | Out-Null
            }
        }

        # Create an array for multiple detection rules if required
        if ($AppData.DetectionRule.Count -gt 1) {
            if ("Script" -in $AppData.DetectionRule.Type) {
                # When a Script detection rule is used, other detection rules cannot be used as well. This should be handled within the module itself by the Add-IntuneWin32App function
            }
        }

        # Create detection rules
        $DetectionRules = New-Object -TypeName "System.Collections.ArrayList"
        foreach ($DetectionRuleItem in $AppData.DetectionRule) {
            switch ($DetectionRuleItem.Type) {
                "MSI" {
                    # Create a MSI installation based detection rule
                    $DetectionRuleArgs = @{
                        "ProductCode"            = $DetectionRuleItem.ProductCode
                        "ProductVersionOperator" = $DetectionRuleItem.ProductVersionOperator
                    }
                    if (-not([System.String]::IsNullOrEmpty($DetectionRuleItem.ProductVersion))) {
                        $DetectionRuleArgs.Add("ProductVersion", $DetectionRuleItem.ProductVersion)
                    }

                    # Create MSI based detection rule
                    $DetectionRule = New-IntuneWin32AppDetectionRuleMSI @DetectionRuleArgs
                }
                "Script" {
                    # Create a PowerShell script based detection rule
                    $DetectionRuleArgs = @{
                        "ScriptFile"            = (Join-Path -Path $ScriptsFolder -ChildPath $DetectionRuleItem.ScriptFile)
                        "EnforceSignatureCheck" = [System.Convert]::ToBoolean($DetectionRuleItem.EnforceSignatureCheck)
                        "RunAs32Bit"            = [System.Convert]::ToBoolean($DetectionRuleItem.RunAs32Bit)
                    }

                    # Create script based detection rule
                    $DetectionRule = New-IntuneWin32AppDetectionRuleScript @DetectionRuleArgs
                }
                "Registry" {
                    switch ($DetectionRuleItem.DetectionMethod) {
                        "Existence" {
                            # Construct registry existence detection rule parameters
                            $DetectionRuleArgs = @{
                                "Existence"            = $true
                                "KeyPath"              = $DetectionRuleItem.KeyPath
                                "DetectionType"        = $DetectionRuleItem.DetectionType
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                            if (-not([System.String]::IsNullOrEmpty($DetectionRuleItem.ValueName))) {
                                $DetectionRuleArgs.Add("ValueName", $DetectionRuleItem.ValueName)
                            }
                        }
                        "VersionComparison" {
                            # Construct registry version comparison detection rule parameters
                            $DetectionRuleArgs = @{
                                "VersionComparison"         = $true
                                "KeyPath"                   = $DetectionRuleItem.KeyPath
                                "ValueName"                 = $DetectionRuleItem.ValueName
                                "VersionComparisonOperator" = $DetectionRuleItem.Operator
                                "VersionComparisonValue"    = $DetectionRuleItem.Value
                                "Check32BitOn64System"      = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "StringComparison" {
                            # Construct registry string comparison detection rule parameters
                            $DetectionRuleArgs = @{
                                "StringComparison"         = $true
                                "KeyPath"                  = $DetectionRuleItem.KeyPath
                                "ValueName"                = $DetectionRuleItem.ValueName
                                "StringComparisonOperator" = $DetectionRuleItem.Operator
                                "StringComparisonValue"    = $DetectionRuleItem.Value
                                "Check32BitOn64System"     = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "IntegerComparison" {
                            # Construct registry integer comparison detection rule parameters
                            $DetectionRuleArgs = @{
                                "IntegerComparison"         = $true
                                "KeyPath"                   = $DetectionRuleItem.KeyPath
                                "ValueName"                 = $DetectionRuleItem.ValueName
                                "IntegerComparisonOperator" = $DetectionRuleItem.Operator
                                "IntegerComparisonValue"    = $DetectionRuleItem.Value
                                "Check32BitOn64System"      = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                    }

                    # Create registry based detection rule
                    $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry @DetectionRuleArgs
                }
                "File" {
                    switch ($DetectionRuleItem.DetectionMethod) {
                        "Existence" {
                            # Create a custom file based requirement rule
                            $DetectionRuleArgs = @{
                                "Existence"            = $true
                                "Path"                 = $DetectionRuleItem.Path
                                "FileOrFolder"         = $DetectionRuleItem.FileOrFolder
                                "DetectionType"        = $DetectionRuleItem.DetectionType
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "DateModified" {
                            # Create a custom file based requirement rule
                            $DetectionRuleArgs = @{
                                "DateModified"         = $true
                                "Path"                 = $DetectionRuleItem.Path
                                "FileOrFolder"         = $DetectionRuleItem.FileOrFolder
                                "Operator"             = $DetectionRuleItem.Operator
                                "DateTimeValue"        = $DetectionRuleItem.DateTimeValue
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "DateCreated" {
                            # Create a custom file based requirement rule
                            $DetectionRuleArgs = @{
                                "DateCreated"          = $true
                                "Path"                 = $DetectionRuleItem.Path
                                "FileOrFolder"         = $DetectionRuleItem.FileOrFolder
                                "Operator"             = $DetectionRuleItem.Operator
                                "DateTimeValue"        = $DetectionRuleItem.DateTimeValue
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "Version" {
                            # Create a custom file based requirement rule
                            $DetectionRuleArgs = @{
                                "Version"              = $true
                                "Path"                 = $DetectionRuleItem.Path
                                "FileOrFolder"         = $DetectionRuleItem.FileOrFolder
                                "Operator"             = $DetectionRuleItem.Operator
                                "VersionValue"         = $DetectionRuleItem.VersionValue
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                        "Size" {
                            # Create a custom file based requirement rule
                            $DetectionRuleArgs = @{
                                "Size"                 = $true
                                "Path"                 = $DetectionRuleItem.Path
                                "FileOrFolder"         = $DetectionRuleItem.FileOrFolder
                                "Operator"             = $DetectionRuleItem.Operator
                                "SizeInMBValue"        = $DetectionRuleItem.SizeInMBValue
                                "Check32BitOn64System" = [System.Convert]::ToBoolean($DetectionRuleItem.Check32BitOn64System)
                            }
                        }
                    }

                    # Create file based detection rule
                    $DetectionRule = New-IntuneWin32AppDetectionRuleFile @DetectionRuleArgs
                }
            }

            # Add detection rule to list
            $DetectionRules.Add($DetectionRule) | Out-Null
        }

        # Add icon
        if (Test-Path -Path $AppIconFile) {
            $Icon = New-IntuneWin32AppIcon -FilePath $AppIconFile
        }

        # Create a Notes property with identifying information
        $Notes = [PSCustomObject] @{
            "CreatedBy" = "PSPackageFactory"
            "Guid"      = $AppData.Information.PSPackageFactoryGuid
            "Date"      = $(Get-Date -Format "yyyy-MM-dd")
        } | ConvertTo-Json -Compress

        # Construct a table of default parameters for Win32 app
        $Win32AppArgs = @{
            "FilePath"                 = $PackageFile
            "DisplayName"              = $AppData.Information.DisplayName
            "Description"              = $AppData.Information.Description
            "AppVersion"               = $AppData.PackageInformation.Version
            "Notes"                    = $Notes
            "Publisher"                = $AppData.Information.Publisher
            "Developer"                = $AppData.Information.Publisher
            "InformationURL"           = $AppData.Information.InformationURL
            "PrivacyURL"               = $AppData.Information.PrivacyURL
            "CompanyPortalFeaturedApp" = $false
            "InstallExperience"        = $AppData.Program.InstallExperience
            "RestartBehavior"          = $AppData.Program.DeviceRestartBehavior
            "DetectionRule"            = $DetectionRules
            "RequirementRule"          = $RequirementRule
            #"UseAzCopy"                = $true
        }

        # Dynamically add additional parameters for Win32 app
        if ($null -ne $RequirementRules) {
            $Win32AppArgs.Add("AdditionalRequirementRule", $RequirementRules)
        }
        if (Test-Path -Path $AppIconFile) {
            $Win32AppArgs.Add("Icon", $Icon)
        }
        if (-not([System.String]::IsNullOrEmpty($AppData.Information.Categories))) {
            $Win32AppArgs.Add("CategoryName", $AppData.Information.Categories)
        }
        if (-not([System.String]::IsNullOrEmpty($AppData.Program.InstallCommand))) {
            $Win32AppArgs.Add("InstallCommandLine", $AppData.Program.InstallCommand)
        }
        if (-not([System.String]::IsNullOrEmpty($AppData.Program.UninstallCommand))) {
            $Win32AppArgs.Add("UninstallCommandLine", $AppData.Program.UninstallCommand)
        }

        if ($PSBoundParameters["Validate"]) {
            if (-not([System.String]::IsNullOrEmpty($Win32AppArgs["Icon"]))) {
                # Redact icon Base64 code for better visibility in validate context
                $Win32AppArgs["Icon"] = $Win32AppArgs["Icon"].SubString(0, 20) + "... //redacted for validation context//"
            }

            # Output manifest configuration
            $Win32AppArgs | ConvertTo-Json
        }
        else {
            # Create Win32 app
            Write-Msg -Msg "Creating Win32 app in Intune: $($AppData.Information.DisplayName)"
            $Win32App = Add-IntuneWin32App @Win32AppArgs

            # Add assignments
            if ($AppData.Assignments.Count -ge 1) {
                Write-Msg -Msg "Adding assignments to Win32 app"
                foreach ($Assignment in $AppData.Assignments) {

                    # Construct the assignment arguments
                    $AssignmentArgs = @{
                        "ID"           = $Win32App.id
                        "Intent"       = $Assignment.Intent
                        "Notification" = $Assignment.Notification
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.DeliveryOptimizationPriority))) {
                        $AssignmentArgs.Add("DeliveryOptimizationPriority", $Assignment.DeliveryOptimizationPriority)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.EnableRestartGracePeriod))) {
                        $AssignmentArgs.Add("EnableRestartGracePeriod", $Assignment.EnableRestartGracePeriod)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.RestartGracePeriod))) {
                        $AssignmentArgs.Add("RestartGracePeriod", $Assignment.RestartGracePeriod)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.RestartCountDownDisplay))) {
                        $AssignmentArgs.Add("RestartCountDownDisplay", $Assignment.RestartCountDownDisplay)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.RestartNotificationSnooze))) {
                        $AssignmentArgs.Add("RestartNotificationSnooze", $Assignment.RestartNotificationSnooze)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.AvailableTime))) {
                        $AssignmentArgs.Add("AvailableTime", $Assignment.AvailableTime)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.DeadlineTime))) {
                        $AssignmentArgs.Add("DeadlineTime", $Assignment.DeadlineTime)
                    }
                    if (-not([System.String]::IsNullOrEmpty($Assignment.UseLocalTime))) {
                        $AssignmentArgs.Add("UseLocalTime", $Assignment.UseLocalTime)
                    }

                    switch ($Assignment.Type) {
                        "AllDevices" {
                            [void](Add-IntuneWin32AppAssignmentAllDevices @AssignmentArgs)
                        }
                        "AllUsers" {
                            [void](Add-IntuneWin32AppAssignmentAllUsers @AssignmentArgs)
                        }
                        "Group" {
                            $AssignmentArgs.Add("GroupID", $Assignment.GroupID)
                            [void](Add-IntuneWin32AppAssignmentGroup @AssignmentArgs -Include)
                        }
                    }
                }
            }

            # Output application package details
            Write-Msg -Msg "Win32 app created successfully: $($Win32App.displayName)"
            Get-IntuneWin32App -ID $Win32App.id
        }
    }

    end {
        Write-Msg -Msg "Completed Win32 app creation from manifest"
    }
}
