#Requires -PSEdition Desktop
#Requires -Modules MSAL.PS, PSIntuneAuth, AzureAD, IntuneWin32App, Microsoft.Graph.Intune
<#
    .SYNOPSIS
        Create the Intune package for the Microsoft 365 Apps and imported into an Intune tenant.

    .DESCRIPTION


    .PARAMETER Path
        Path to the top level directory of the repository

    .EXAMPLE


    .NOTES
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

    [System.String] $TenantId,
    [System.String] $ClientId,
    [System.String] $ClientSecret,
    [System.Management.Automation.SwitchParameter] $Import
)

begin {
    # Unblock all files in the repo
    Get-ChildItem -Path $Path -Recurse -Include "*.exe" | Unblock-File

    # Validate required files exist
    @(
        "$Path\configs\Uninstall-Microsoft365Apps.xml",
        "$Path\intunewin\IntuneWinAppUtil.exe",
        "$Path\m365\setup.exe",
        "$Path\scripts\App.json"
    ) | ForEach-Object { if (-not (Test-Path -Path $_)) { throw [System.IO.FileNotFoundException]::New("File not found: $_") } }
}

process {

    #region Create working directories; Copy files for the package
    try {
        New-Item -Path "$Path\output" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        New-Item -Path "$Path\PSAppDeployToolkit\Toolkit\Files" -ItemType "Directory" -ErrorAction "SilentlyContinue"
        Copy-Item -Path "$Path\configs\$ConfigurationFile" -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\Install-Microsoft365Apps.xml"
        Copy-Item -Path "$Path\m365\setup.exe" -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\setup.exe"
        Copy-Item -Path "$Path\configs\Uninstall-Microsoft365Apps.xml" -Destination "$Path\PSAppDeployToolkit\Toolkit\Files\Uninstall-Microsoft365Apps.xml"
    }
    catch {
        throw $_
    }
    #endregion

    #region Create the intunewin package
    $params = @{
        SourceFolder         = "$Path\PSAppDeployToolkit\Toolkit"
        SetupFile            = "$Path\PSAppDeployToolkit\Toolkit\Files\setup.exe"
        OutputFolder         = "$Path\output"
        Force                = $true
        IntuneWinAppUtilPath = "$Path\intunewin\IntuneWinAppUtil.exe"
    }
    New-IntuneWin32AppPackage @params
    #endregion

    #region Create a new App.json for the package & update based on the setup.exe version & configuration.xml
    try {
        $SetupVersion = (Get-Item -Path "$Path\m365\setup.exe").VersionInfo.FileVersion
        Copy-Item -Path "$Path\scripts\App.json" -Destination "$Path\scripts\Temp.json"
        $AppJson = Get-Content -Path "$Path\scripts\Temp.json" | ConvertFrom-Json
        $AppJson.PackageInformation.Version = $SetupVersion
        [System.Xml.XmlDocument]$Xml = Get-Content -Path $ConfigurationFile
        [System.String] $ProductID = ""
        switch ($Xml.Configuration.Add.Product.ID) {
            "O365ProPlusRetail" {
                $ProductID += "Microsoft 365 apps for enterprise, "
            }
            "O365BusinessRetail" {
                $ProductID += "Microsoft 365 apps for business, "
            }
            "VisioProRetail" {
                $ProductID += "Visio Plan 2, "
            }
            "ProjectProRetail" {
                $ProductID += "Project Online Desktop Client, "
            }
            "AccessRuntimeRetail" {
                $ProductID += "Office 365 Access Runtime, "
            }
        }
        [System.String] $DisplayName = "$ProductID$($Xml.Configuration.Add.Channel)"
        if ($Xml.Configuration.Add.OfficeClientEdition -eq "64") { $DisplayName = "$DisplayName, x64" }
        if ($Xml.Configuration.Add.OfficeClientEdition -eq "32") { $DisplayName = "$DisplayName, x86" }
        $AppJson.Information.DisplayName = $DisplayName
        $AppJson.Information.Description = $Xml.Configuration.Info.Description
        $AppJson | ConvertTo-Json | Out-File -FilePath "$Path\scripts\Temp.json" -Force
    }
    catch {
        throw $_
    }
    #endregion
    
    #region Authn if authn parameters are passed; Import package into Intune
    if ($PSBoundParameters.Contains("Import")) {
        if ($PSBoundParameters.Contains("TenantId")) {
            $params = @{
                TenantId     = $TenantId
                ClientId     = $ClientId
                ClientSecret = $ClientSecret
            }
            $global:AuthToken = Connect-MSIntuneGraph @params
        }

        # Launch script to import the package
        $params = @{
            Json         = "$Path\scripts\Temp.json"
            PackageFile  = $PackageFile
            SetupVersion = $SetupVersion
        }
        & "$Path\scripts\Create-Win32App.ps1" @params
    }
    #endregion
}

end {
}
