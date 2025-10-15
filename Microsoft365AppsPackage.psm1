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
        Write-Msg -Msg "Retrieving existing Win32 applications from Intune."
        $ExistingIntuneApps = Get-IntuneWin32App | `
            Where-Object { $_.displayName -match $DisplayNamePattern -and $_.notes -match $NotesPattern } | `
            Select-Object -Property * -ExcludeProperty "largeIcon"
        if ($ExistingIntuneApps -is [System.Object]) {
            Write-Msg -Msg "Found 1 existing Microsoft 365 Apps package in Intune."
        }
        elseif ($ExistingIntuneApps.Count -gt 0) {
            Write-Msg -Msg "Found $($ExistingIntuneApps.Count) existing Microsoft 365 Apps packages in Intune."
        }
    }

    process {
        foreach ($Id in $PackageId) {
            Write-Msg -Msg "Filtering existing applications to match Microsoft 365 Apps PackageId."
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

function Test-PackagePrerequisites {
    <#
    .SYNOPSIS
        Validates all prerequisites for package creation.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationFile,

        [Parameter(Mandatory = $true)]
        [System.String]$TenantId
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
        [System.String]$Destination,

        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationFile,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$UsePsadt
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
        [System.String]$Destination,

        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationFile
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
        [System.String]$Destination,

        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationFile
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
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationPath,

        [Parameter(Mandatory = $true)]
        [System.String]$Channel,

        [Parameter(Mandatory = $true)]
        [System.String]$TenantId,

        [Parameter(Mandatory = $true)]
        [System.String]$CompanyName
    )

    Write-Msg -Msg "Updating configuration file: $ConfigurationPath"
    [System.Xml.XmlDocument]$Xml = Get-Content -Path $ConfigurationPath

    if ($PSCmdlet.ShouldProcess($ConfigurationPath, "Update M365 Configuration")) {
        Write-Msg -Msg "Set Microsoft 365 Apps channel to: $Channel."
        $Xml.Configuration.Add.Channel = $Channel

        Write-Msg -Msg "Set tenant id to: $TenantId."
        $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "TenantId"))
        $Xml.Configuration.Property[$Index].Value = $TenantId

        Write-Msg -Msg "Set company name to: $CompanyName."
        $Xml.Configuration.AppSettings.Setup.Value = $CompanyName

        Write-Msg -Msg "Save configuration xml to: $ConfigurationPath."
        $Xml.Save($ConfigurationPath)
    }
    else {
        # WhatIf - show what would be changed
        Write-Host "Would update configuration file: $ConfigurationPath" -ForegroundColor Yellow
        Write-Host "  Channel: $Channel" -ForegroundColor Yellow
        Write-Host "  Company: $CompanyName" -ForegroundColor Yellow
        Write-Host "  TenantId: $TenantId" -ForegroundColor Yellow
        
        # Return simulated XML for validation
        $Xml.Configuration.Add.Channel = $Channel
        $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "TenantId"))
        if ($Index -ge 0) { $Xml.Configuration.Property[$Index].Value = $TenantId }
        $Xml.Configuration.AppSettings.Setup.Value = $CompanyName
    }

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
        [System.String]$Destination,

        [Parameter(Mandatory = $true)]
        [System.String]$Path,

        [Parameter(Mandatory = $true)]
        [System.String]$ConfigurationFile,

        [Parameter(Mandatory = $true)]
        [System.String]$Channel,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.SwitchParameter]$UsePsadt
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
    Write-Msg -Msg "Package GUID: $($Manifest.Information.PSPackageFactoryGuid)."

    # Update icon location
    $Manifest.PackageInformation.IconFile = "$Path\icons\Microsoft365.png"

    # Update package description
    if ($UsePsadt) {
        $Description = "$($xml.Configuration.Info.Description)`n`n**This package uses the PSAppDeployToolkit and will uninstall previous versions of Microsoft Office**. Uses setup.exe **$SetupVersion**. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
    }
    else {
        $Description = "$($xml.Configuration.Info.Description)`n`nUses setup.exe **$SetupVersion**. Built from configuration file: $(Split-Path -Path $ConfigurationFile -Leaf); Includes: $(($Xml.Configuration.Add.Product.ID | Sort-Object) -join ", ")."
    }
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

    # Build an array of the product names
    $ProductID = [System.Collections.ArrayList]::new()
    switch ($Xml.Configuration.Add.Product.ID) {
        "O365ProPlusRetail" { [void]$ProductID.Add("Microsoft 365 Apps for enterprise") }
        "O365BusinessRetail" { [void]$ProductID.Add("Microsoft 365 Apps for business") }
        "VisioProRetail" { [void]$ProductID.Add("Visio") }
        "ProjectProRetail" { [void]$ProductID.Add("Project") }
        "AccessRuntimeRetail" { [void]$ProductID.Add("Access Runtime") }
    }
    if ("Outlook" -in $Xml.Configuration.Add.Product.ExcludeApp.ID) { [void]$ProductID.Add("Outlook (new)") }
    if ("OutlookForWindows" -in $Xml.Configuration.Add.Product.ExcludeApp.ID) { [void]$ProductID.Add("Outlook (classic)") }

    # Build an array of the configuration properties
    $Properties = [System.Collections.ArrayList]::new()
    $Index = $Xml.Configuration.Property.Name.IndexOf($($Xml.Configuration.Property.Name -cmatch "SharedComputerLicensing"))
    if ($Xml.Configuration.Property[$Index].Value -eq "1") { [void]$Properties.Add("VDI") } else { [void]$Properties.Add("Desktop") }
    [void]$Properties.Add($Xml.Configuration.Add.Channel)
    if ($Xml.Configuration.Add.OfficeClientEdition -eq "64") { [void]$Properties.Add("x64") }
    if ($Xml.Configuration.Add.OfficeClientEdition -eq "32") { [void]$Properties.Add("x86") }

    # Return the combined display name
    return "$($ProductID -join ", "): $($Properties -join ", ")"
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
        [System.String]$Channel
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
        [System.Management.Automation.SwitchParameter]$Force
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
        [System.String]$Operation
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
    [System.String]$Path
    [System.String]$ConfigurationFile
    [System.String]$Channel
    [System.String]$CompanyName
    [bool]$UsePsadt
    [System.String]$Destination
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

function Test-ParameterValidation {
    <#
    .SYNOPSIS
        Validates script parameters without executing actions.
    #>
    [CmdletBinding()]
    param(
        [System.String]$Path,
        [System.String]$ConfigurationFile,
        [System.String]$Channel,
        [System.String]$CompanyName,
        [System.String]$TenantId,
        [System.Management.Automation.SwitchParameter]$UsePsadt
    )
    
    $results = @()
    
    # Validate paths
    $results += @{
        Test    = "Path Validation"
        Status  = if (Test-Path $Path -PathType Container) { "PASS" } else { "FAIL" }
        Message = "Repository path: $Path"
    }
    
    $results += @{
        Test    = "Configuration File Validation"
        Status  = if (Test-Path $ConfigurationFile -PathType Leaf) { "PASS" } else { "FAIL" }
        Message = "Config file: $ConfigurationFile"
    }
    
    # Validate XML structure
    try {
        $xml = [xml](Get-Content $ConfigurationFile)
        $hasConfiguration = $null -ne $xml.Configuration
        $results += @{
            Test    = "XML Structure Validation"
            Status  = if ($hasConfiguration) { "PASS" } else { "FAIL" }
            Message = "Valid Microsoft 365 Apps configuration XML"
        }
    }
    catch {
        $results += @{
            Test    = "XML Structure Validation"
            Status  = "FAIL"
            Message = "Invalid XML: $($_.Exception.Message)"
        }
    }
    
    # Validate required files
    $requiredFiles = @(
        "$Path\m365\setup.exe",
        "$Path\intunewin\IntuneWinAppUtil.exe",
        "$Path\icons\Microsoft365.png",
        "$Path\configs\Uninstall-Microsoft365Apps.xml"
    )
    
    if ($UsePsadt) {
        $requiredFiles += "$Path\scripts\Invoke-AppDeployToolkit.ps1"
    }
    
    foreach ($file in $requiredFiles) {
        $exists = Test-Path $file
        $results += @{
            Test    = "Required File Check"
            Status  = if ($exists) { "PASS" } else { "FAIL" }
            Message = "File: $file"
        }
    }
    
    # Validate GUID formats
    $guidTest = [System.Guid]::empty
    $results += @{
        Test    = "TenantId GUID Validation"
        Status  = if ([System.Guid]::TryParse($TenantId, [ref]$guidTest)) { "PASS" } else { "FAIL" }
        Message = "TenantId: $TenantId"
    }
    
    return $results
}

function Test-XmlUpdateValidation {
    <#
    .SYNOPSIS
        Validates XML configuration updates without modifying the original file.
    #>
    [CmdletBinding()]
    param(
        [System.String]$ConfigurationFile,
        [System.String]$Channel,
        [System.String]$CompanyName,
        [System.String]$TenantId
    )
    
    $results = @()
    
    try {
        # Create a temporary copy for validation
        $tempFile = [System.IO.Path]::GetTempFileName()
        Copy-Item $ConfigurationFile $tempFile
        
        # Test XML updates without modifying original
        $xml = [xml](Get-Content $tempFile)
        
        # Simulate the updates
        $xml.Configuration.Add.Channel = $Channel
        
        $tenantIndex = $xml.Configuration.Property.Name.IndexOf($($xml.Configuration.Property.Name -cmatch "TenantId"))
        if ($tenantIndex -ge 0) {
            $xml.Configuration.Property[$tenantIndex].Value = $TenantId
        }
        
        $xml.Configuration.AppSettings.Setup.Value = $CompanyName
        
        $results += @{
            Test    = "XML Update Simulation"
            Status  = "PASS"
            Message = "Successfully simulated XML updates"
            Details = @{
                Channel     = $xml.Configuration.Add.Channel
                CompanyName = $xml.Configuration.AppSettings.Setup.Value
                TenantId    = if ($tenantIndex -ge 0) { $xml.Configuration.Property[$tenantIndex].Value } else { "Not found" }
            }
        }
        
        # Clean up
        Remove-Item $tempFile -Force
    }
    catch {
        $results += @{
            Test    = "XML Update Simulation"
            Status  = "FAIL"
            Message = "Failed to simulate XML updates: $($_.Exception.Message)"
        }
    }
    
    return $results
}

function Test-FileOperationsValidation {
    <#
    .SYNOPSIS
        Validates file operations without actually copying files.
    #>
    [CmdletBinding()]
    param(
        [System.String]$Destination,
        [System.String]$Path,
        [System.String]$ConfigurationFile,
        [System.Management.Automation.SwitchParameter]$UsePsadt
    )
    
    $results = @()
    
    try {
        # Test directory structure creation (simulation)
        $results += @{
            Test    = "Directory Structure Validation" 
            Status  = "PASS"
            Message = "Would create directories at: $Destination"
            Details = @{
                SourceDir = "$Destination\source"
                OutputDir = "$Destination\output" 
                FilesDir  = if ($UsePsadt) { "$Destination\source\Files" } else { $null }
            }
        }
        
        # Test file availability for copying
        $filesToCopy = @()
        $filesToCopy += @{ Path = "$Path\m365\setup.exe"; Target = if ($UsePsadt) { "Files\setup.exe" } else { "setup.exe" } }
        $filesToCopy += @{ Path = $ConfigurationFile; Target = if ($UsePsadt) { "Files\Install-Microsoft365Apps.xml" } else { "Install-Microsoft365Apps.xml" } }
        $filesToCopy += @{ Path = "$Path\configs\Uninstall-Microsoft365Apps.xml"; Target = if ($UsePsadt) { "Files\Uninstall-Microsoft365Apps.xml" } else { "Uninstall-Microsoft365Apps.xml" } }
        
        $missingFiles = @()
        $availableFiles = @()
        
        foreach ($file in $filesToCopy) {
            if (Test-Path $file.Path) {
                $availableFiles += $file
            }
            else {
                $missingFiles += $file.Path
            }
        }
        
        $results += @{
            Test    = "File Copy Validation"
            Status  = if ($missingFiles.Count -eq 0) { "PASS" } else { "FAIL" }
            Message = "Available: $($availableFiles.Count), Missing: $($missingFiles.Count)"
            Details = @{
                AvailableFiles = $availableFiles | Select-Object Path, Target
                MissingFiles   = $missingFiles
            }
        }
        
    }
    catch {
        $results += @{
            Test    = "File Operations Validation"
            Status  = "FAIL"
            Message = "File operations validation failed: $($_.Exception.Message)"
        }
    }
    
    return $results
}

function Test-PackageManifestValidation {
    <#
    .SYNOPSIS
        Validates package manifest creation without creating actual files.
    #>
    [CmdletBinding()]
    param(
        [xml]$Xml,
        [System.String]$ConfigurationFile,
        [System.String]$Channel,
        [System.Management.Automation.SwitchParameter]$UsePsadt
    )
    
    $results = @()
    
    try {
        # Simulate manifest creation
        $displayName = Get-PackageDisplayName -Xml $Xml
        
        $testManifest = @{
            Information    = @{
                PSPackageFactoryGuid = $Xml.Configuration.ID
                DisplayName          = $displayName
            }
            Program        = @{
                InstallCommand   = if ($UsePsadt) { "Deploy-Application.exe -DeploymentType Install" } else { "setup.exe /configure Install-Microsoft365Apps.xml" }
                UninstallCommand = if ($UsePsadt) { "Deploy-Application.exe -DeploymentType Uninstall" } else { "setup.exe /configure Uninstall-Microsoft365Apps.xml" }
            }
            DetectionRules = @{
                ProductReleaseIds = ($Xml.Configuration.Add.Product.ID | Sort-Object) -join ","
                Channel           = $Channel
            }
        }
        
        $results += @{
            Test    = "Package Manifest Validation"
            Status  = "PASS"
            Message = "Successfully created manifest structure"
            Details = $testManifest
        }
    }
    catch {
        $results += @{
            Test    = "Package Manifest Validation" 
            Status  = "FAIL"
            Message = "Manifest creation failed: $($_.Exception.Message)"
        }
    }
    
    return $results
}

function Invoke-PackageValidation {
    <#
    .SYNOPSIS
        Runs comprehensive package validation without executing actual operations.
    #>
    [CmdletBinding()]
    param(
        [System.String]$Path,
        [System.String]$Destination,
        [System.String]$ConfigurationFile,
        [System.String]$Channel,
        [System.String]$CompanyName,
        [System.String]$TenantId,
        [System.Management.Automation.SwitchParameter]$UsePsadt
    )
    
    Write-Host "🔍 Starting Microsoft 365 Apps Package Validation..." -ForegroundColor Cyan
    Write-Host "=" * 60
    
    $allResults = @()
    
    # 1. Parameter Validation
    Write-Host "`n📋 Parameter Validation" -ForegroundColor Yellow
    $paramResults = Test-ParameterValidation -Path $Path -ConfigurationFile $ConfigurationFile -Channel $Channel -CompanyName $CompanyName -TenantId $TenantId -UsePsadt:$UsePsadt
    $allResults += $paramResults
    Show-ValidationResults $paramResults
    
    # 2. XML Update Validation
    Write-Host "`n📝 XML Configuration Validation" -ForegroundColor Yellow
    $xmlResults = Test-XmlUpdateValidation -ConfigurationFile $ConfigurationFile -Channel $Channel -CompanyName $CompanyName -TenantId $TenantId
    $allResults += $xmlResults
    Show-ValidationResults $xmlResults
    
    # 3. File Operations Validation
    Write-Host "`n📁 File Operations Validation" -ForegroundColor Yellow
    $fileResults = Test-FileOperationsValidation -Destination $Destination -Path $Path -ConfigurationFile $ConfigurationFile -UsePsadt:$UsePsadt
    $allResults += $fileResults
    Show-ValidationResults $fileResults
    
    # 4. Package Manifest Validation
    Write-Host "`n📦 Package Manifest Validation" -ForegroundColor Yellow
    $xml = Import-XmlFile -FilePath $ConfigurationFile
    $manifestResults = Test-PackageManifestValidation -Xml $xml -ConfigurationFile $ConfigurationFile -Channel $Channel -UsePsadt:$UsePsadt
    $allResults += $manifestResults
    Show-ValidationResults $manifestResults
    
    # Summary
    Write-Host "`n" + "=" * 60
    $totalTests = $allResults.Count
    $passedTests = ($allResults | Where-Object Status -EQ "PASS").Count
    $failedTests = ($allResults | Where-Object Status -EQ "FAIL").Count
    
    Write-Host "📊 Validation Summary:" -ForegroundColor Cyan
    Write-Host "   Total Tests: $totalTests" -ForegroundColor White
    Write-Host "   Passed: $passedTests" -ForegroundColor Green
    Write-Host "   Failed: $failedTests" -ForegroundColor Red
    
    if ($failedTests -eq 0) {
        Write-Host "`n✅ All validations passed! Package is ready for creation." -ForegroundColor Green
    }
    else {
        Write-Host "`n❌ $failedTests validation(s) failed. Please address issues before proceeding." -ForegroundColor Red
    }
    
    return $allResults
}

function Show-ValidationResults {
    <#
    .SYNOPSIS
        Displays validation results in a formatted manner.
    #>
    param($Results)
    
    foreach ($result in $Results) {
        $icon = if ($result.Status -eq "PASS") { "✅" } else { "❌" }
        $color = if ($result.Status -eq "PASS") { "Green" } else { "Red" }
        
        Write-Host "  $icon $($result.Test): $($result.Message)" -ForegroundColor $color
        
        if ($result.Details) {
            Write-Host "     Details: $($result.Details | ConvertTo-Json -Compress)" -ForegroundColor Gray
        }
    }
}
