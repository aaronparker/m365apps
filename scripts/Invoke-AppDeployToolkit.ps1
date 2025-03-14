<#

.SYNOPSIS
PSAppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION
- The script is provided as a template to perform an install, uninstall, or repair of an application(s).
- The script either performs an "Install", "Uninstall", or "Repair" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script imports the PSAppDeployToolkit module which contains the logic and functions required to install or uninstall an application.

PSAppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2025 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham, Muhammad Mashwani, Mitch Richters, Dan Gough).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType
The type of deployment to perform.

.PARAMETER DeployMode
Specifies whether the installation should be run in Interactive (shows dialogs), Silent (no dialogs), or NonInteractive (dialogs without prompts) mode.

NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru
Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode
Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging
Disables logging to file for the script.

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeployMode Silent

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -AllowRebootPassThru

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeploymentType Uninstall

.EXAMPLE
Invoke-AppDeployToolkit.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None. This script does not generate any output.

.NOTES
Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Invoke-AppDeployToolkit.ps1, and Invoke-AppDeployToolkit.exe
- 69000 - 69999: Recommended for user customized exit codes in Invoke-AppDeployToolkit.ps1
- 70000 - 79999: Recommended for user customized exit codes in PSAppDeployToolkit.Extensions module.

.LINK
https://psappdeploytoolkit.com

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [PSDefaultValue(Help = 'Install', Value = 'Install')]
    [System.String]$DeploymentType,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [PSDefaultValue(Help = 'Interactive', Value = 'Interactive')]
    [System.String]$DeployMode,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$AllowRebootPassThru,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$TerminalServerMode,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.SwitchParameter]$DisableLogging
)


##================================================
## MARK: Variables
##================================================

$adtSession = @{
    # App variables.
    AppVendor                   = 'Microsoft'
    AppName                     = '365 Apps'
    AppVersion                  = ''
    AppArch                     = 'x64'
    AppLang                     = 'EN'
    AppRevision                 = '01'
    AppSuccessExitCodes         = @(0)
    AppRebootExitCodes          = @(1641, 3010)
    AppScriptVersion            = '1.0.0'
    AppScriptDate               = '2025-03-13'
    AppScriptAuthor             = '<author name>'

    # Install Titles (Only set here to override defaults set by the toolkit).
    InstallName                 = ''
    InstallTitle                = ''

    # Script variables.
    DeployAppScriptFriendlyName = $MyInvocation.MyCommand.Name
    DeployAppScriptVersion      = '4.0.6'
    DeployAppScriptParameters   = $PSBoundParameters
}

function Install-ADTDeployment {
    ##*===============================================
    ##* PRE-INSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
    Show-ADTInstallationWelcome -CloseProcesses 'iexplore' -AllowDefer  -DeferTimes 3 -CheckDiskSpace  -PersistPrompt 

    ## Show Progress Message (with the default message)
    Show-ADTInstallationProgress 

    ## <Perform Pre-Installation tasks here>
    ## Remove Office 2013  MSI installations
    if (Test-Path -Path "$envProgramFilesX86\Microsoft Office\Office15", "$envProgramFiles\Microsoft Office\Office15") {
        Show-ADTInstallationProgress -StatusMessage "Uninstalling Microsoft Office 2013"
        Write-ADTLogEntry -Message "Microsoft Office 2013 was detected. Uninstalling..."
        Start-ADTProcess -FilePath "CScript.exe" -ArgumentList "`"$($adtSession.DirSupportFiles)\OffScrub_O15msi.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes 1, 2, 3
    }

    ## Remove Office 2016 MSI installations
    if (Test-Path -Path "$envProgramFilesX86\Microsoft Office\Office16", "$envProgramFiles\Microsoft Office\Office16") {
        Show-ADTInstallationProgress -StatusMessage "Uninstalling Microsoft Office 2016"
        Write-ADTLogEntry -Message "Microsoft Office 2016 was detected. Uninstalling..."
        Start-ADTProcess -FilePath "CScript.exe" -ArgumentList "`"$($adtSession.DirSupportFiles)\OffScrub_O16msi.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes 1, 2, 3
    }

    ##*===============================================
    ##* INSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## Handle Zero-Config MSI Installations
    if ($adtSession.UseDefaultMsi) {
        [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $adtSession.DefaultMsiFile }; If ($defaultMstFile) {
            $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
        }
        Start-ADTMsiProcess ; If ($defaultMspFiles) {
            $defaultMspFiles | ForEach-Object { Start-ADTMsiProcess -Action 'Patch' -FilePath $_ }
        }
    }

    ## <Perform Installation tasks here>
    # Install Microsoft 365 Apps for Enterprise with content from the Office CDN
    try {
        Write-ADTLogEntry -Message "Find Install-Microsoft365Apps.xml in $($adtSession.DirFiles)"
        $XmlFile = Get-ChildItem -Path $adtSession.DirFiles -Recurse -Include "Install-Microsoft365Apps.xml"
        Write-ADTLogEntry -Message "Found: $($XmlFile.FullName)"
        $XmlDocument = New-Object -TypeName "System.Xml.XmlDocument"
        $XmlDocument.Load($XmlFile.FullName)
        $Msg = "Installing: $($XmlDocument.Configuration.Info.Description) Channel: $($XmlDocument.Configuration.Add.Channel)."
    }
    catch {
        Write-ADTLogEntry -Message "Error: $($_.Exception.Message)"
        $Msg = "Installing the Microsoft 365 Apps"
    }
    Show-ADTInstallationProgress -StatusMessage $Msg
    Start-ADTProcess -FilePath "setup.exe" -ArgumentList "/configure $($XmlFile.FullName)"

    ##*===============================================
    ##* POST-INSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>

    ## Display a message at the end of the install
    If (-not $adtSession.UseDefaultMsi) {
        Show-ADTInstallationPrompt -Message 'Installation complete.' -ButtonRightText 'OK' -Icon Information -NoWait 
    }
}

function Uninstall-ADTDeployment {
    ##*===============================================
    ##* PRE-UNINSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
    Show-ADTInstallationWelcome -CloseProcesses 'iexplore' -CloseProcessesCountdown 60

    ## Show Progress Message (with the default message)
    Show-ADTInstallationProgress 

    ## <Perform Pre-Uninstallation tasks here>


    ##*===============================================
    ##* UNINSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## Handle Zero-Config MSI Uninstallations
    If ($adtSession.UseDefaultMsi) {
        [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $adtSession.DefaultMsiFile }; If ($defaultMstFile) {
            $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
        }
        Start-ADTMsiProcess 
    }

    ## <Perform Uninstallation tasks here>
    $Msg = "Uninstalling the Microsoft 365 Apps"
    Show-ADTInstallationProgress -StatusMessage $Msg
    Start-ADTProcess -FilePath "setup.exe" -ArgumentList "/configure Uninstall-Microsoft365Apps.xml"

    ##*===============================================
    ##* POST-UNINSTALLATION
    ##*===============================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Uninstallation tasks here>


}

function Repair-ADTDeployment {
    ##*===============================================
    ##* PRE-REPAIR
    ##*===============================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
    Show-ADTInstallationWelcome -CloseProcesses 'iexplore' -CloseProcessesCountdown 60

    ## Show Progress Message (with the default message)
    Show-ADTInstallationProgress 

    ## <Perform Pre-Repair tasks here>

    ##*===============================================
    ##* REPAIR
    ##*===============================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## Handle Zero-Config MSI Repairs
    if ($adtSession.UseDefaultMsi) {
        [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $adtSession.DefaultMsiFile; }; If ($defaultMstFile) {
            $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
        }
        Start-ADTMsiProcess 
    }
    ## <Perform Repair tasks here>
    try {
        Write-ADTLogEntry -Message "Find Install-Microsoft365Apps.xml in $($adtSession.DirFiles)"
        $XmlFile = Get-ChildItem -Path $adtSession.DirFiles -Recurse -Include "Install-Microsoft365Apps.xml"
        Write-ADTLogEntry -Message "Found: $($XmlFile.FullName)"
        $XmlDocument = New-Object -TypeName "System.Xml.XmlDocument"
        $XmlDocument.Load($XmlFile.FullName)
        $Msg = "Reinstalling: $($XmlDocument.Configuration.Info.Description) Channel: $($XmlDocument.Configuration.Add.Channel)."
    }
    catch {
        Write-ADTLogEntry -Message "Error: $($_.Exception.Message)"
        $Msg = "Reinstalling the Microsoft 365 Apps"
    }
    Show-ADTInstallationProgress -StatusMessage $Msg
    Start-ADTProcess -FilePath "setup.exe" -ArgumentList "/configure $($XmlFile.FullName)"

    ##*===============================================
    ##* POST-REPAIR
    ##*===============================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Repair tasks here>


}


##================================================
## MARK: Initialization
##================================================

# Set strict error handling across entire operation.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

# Import the module and instantiate a new session.
try {
    $moduleName = if ([System.IO.File]::Exists("$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1")) {
        Get-ChildItem -LiteralPath $PSScriptRoot\PSAppDeployToolkit -Recurse -File | Unblock-File -ErrorAction Ignore
        "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
    }
    else {
        'PSAppDeployToolkit'
    }
    Import-Module -FullyQualifiedName @{ ModuleName = $moduleName; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.0.6' } -Force
    try {
        $iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation
        $adtSession = Open-ADTSession -SessionState $ExecutionContext.SessionState @adtSession @iadtParams -PassThru
    }
    catch {
        Remove-Module -Name PSAppDeployToolkit* -Force
        throw
    }
}
catch {
    $Host.UI.WriteErrorLine((Out-String -InputObject $_ -Width ([System.Int32]::MaxValue)))
    exit 60008
}


##================================================
## MARK: Invocation
##================================================

try {
    Get-Item -Path $PSScriptRoot\PSAppDeployToolkit.* | & {
        process {
            Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
            Import-Module -Name $_.FullName -Force
        }
    }
    & "$($adtSession.DeploymentType)-ADTDeployment"
    Close-ADTSession
}
catch {
    Write-ADTLogEntry -Message ($mainErrorMessage = Resolve-ADTErrorRecord -ErrorRecord $_) -Severity 3
    Show-ADTDialogBox -Text $mainErrorMessage -Icon Stop | Out-Null
    Close-ADTSession -ExitCode 60001
}
finally {
    Remove-Module -Name PSAppDeployToolkit* -Force
}
