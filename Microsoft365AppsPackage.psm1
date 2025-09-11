# Functions
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
    param {
        [Parameter(Mandatory = $true)]
        [System.String]$FilePath
    }

    if (Test-Path -Path $FilePath -PathType "Leaf") {
        try {
            $xmlContent = [System.Xml.XmlDocument](Get-Content -Path $FilePath)
            return $xmlContent
        }
        catch {
            Write-Error "Failed to read XML file: $_"
            return $null
        }
    }
    else {
        Write-Error "File not found: $FilePath"
        return $null
    }
}

function Test-RequiredFiles {
    param {
        [Parameter(Mandatory = $true)]
        [System.String[]]$FilePath
    }

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
    param {
        [Parameter(Mandatory = $true)]
        [System.String]$Path
    }

    if (-not (Test-Path -Path $Path -PathType "Container")) {
        throw "'$Path' does not exist or is not a directory."
        return
    }
    if ((Get-ChildItem -Path $Path -Recurse -File).Count -gt 0) {
        throw "'$Path' is not empty. Remove path and try again."
        return
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

