# Configure the environment
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$InformationPreference = [System.Management.Automation.ActionPreference]::Continue
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
using namespace System.Management.Automation

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
        "$Path\scripts\Create-Win32App.ps1",
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
    <#
    #>
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
                $ImportedApp = & "$($PackageConfig.Path)\scripts\Create-Win32App.ps1" @params | Select-Object -Property * -ExcludeProperty "largeIcon"

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
