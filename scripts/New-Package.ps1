<#
    .SYNOPSIS
        Updates the App.json based on the configuration.xml
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [System.String] $ConfigurationFile,
    [System.String] $Version,
    [System.String] $Path = $env:GITHUB_WORKSPACE
)

Copy-Item -Path "$Path\scripts\App.json" -Destination "$Path\scripts\Temp.json"

$AppJson = Get-Content -Path "$Path\scripts\Temp.json" | ConvertFrom-Json
$AppJson.PackageInformation.Version = $Version

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
