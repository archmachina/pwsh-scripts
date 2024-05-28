<#
.SYNOPSIS
    Simple script to perform Windows Update patching

.DESCRIPTION
    Author: Jesse Reichman
    Link: https://github.com/archmachina/pwsh-scripts/blob/main/src/patch_system.ps1

    This script will attempt to perform updates for a Windows machine using
    Windows Update, along with configurable parameters provided from a JSON
    configuration file.

    The script can be configured to install only patches older than a particular age
    threshold and reboot the machine, if a reboot is required.

.EXAMPLE
    Example configuration file:

    {
        "UpdateCab": true,
        "UseOfflineScan": true,
        "LogFile": "C:\\_patching\\patching_log.txt",
        "AgeThreshold": 14,
        "CanReboot": true
    }

    Example call to the patch script:

    .\patch_system.ps1 -ConfigFile patch_config.json -DryRun

    This will perform a DryRun (No download, install or reboot) against the configuration.

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$ConfigFile,

    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [switch]$DryRun = $false
)

# Global settings
Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

try {
    $PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText
} catch {}

# Global Variables
$script:LogFile = $null
$script:LogMessageUseInfo = $null

Function LogMessage
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Message
    )

    process
    {
        # Check if we can use Write-Information
        if ($null -eq $script:LogMessageUseInfo)
        {
            $script:LogMessageUseInfo = $false
            try
            {
                Write-Information "test" 6>&1 | Out-Null
                $script:LogMessageUseInfo = $true
            } catch {}
        }

        $dateFmt = ([DateTime]::Now.ToString("o"))
        $Message = "${dateFmt}: $Message"

        # Write out the actual message
        if (![string]::IsNullOrEmpty($LogFile))
        {
            $Message | Out-File -Append $LogFile
        } elseif ($script:LogMessageUseInfo)
        {
            Write-Information $Message
        } else {
            Write-Host $Message
        }
    }
}

Function Remove-UpdatePackageService
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        $Manager,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    Process
    {
        $services = $manager.Services | Where-Object { $_.Name -eq $Name } | ForEach-Object { $_ }
        $services | ForEach-Object {
            $service = $_

            ("Removing service with ID: " + $service.ServiceID)
            $manager.RemoveService($service.ServiceID)
        }
    }
}

Function Update-CabFile
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$CabDownloadUri,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$CabFileLocation
    )

    process
    {
        #
        # Determine modification time for the local cab file
        $cabModificationTime = $null
        try {
            "Getting current cab file modification time"
            $cabModificationTime = (Get-Item $CabFileLocation).LastWriteTimeUtc
        } catch {
            "Could not get current cab file modification time: $_"
        }

        "Cab Modification Time: $cabModificationTime"

        #
        # Determine modification time for the remote cab file
        "Retrieving URL modification time"
        $params = @{
            UseBasicParsing = $true
            Method = "Head"
            Uri = $CabDownloadUri
        }

        $headers = Invoke-WebRequest @params

        $urlModificationTime = $null
        try {
            $urlModificationTime = [DateTime]::Parse($headers.Headers["Last-Modified"])
        } catch {
            "Failed to parse Last-Modified header as DateTime: $_"
        }

        "Url Modification Time: $urlModificationTime"

        #
        # Download the file, if it's newer than what we have locally or we don't have valid modification
        # time information
        if ($null -eq $cabModificationTime -or $null -eq $urlModificationTime -or $urlModificationTime -gt $cabModificationTime)
        {
            "wsusscn2.cab file needs updating. Downloading."
            $params = @{
                UseBasicParsing = $true
                Method = "Get"
                Uri = $CabDownloadUri
                OutFile = ("{0}.new" -f $CabFileLocation)
            }

            # Download to temporary location
            Remove-Item -Force ("{0}.new" -f $CabFileLocation) -EA Ignore
            Invoke-WebRequest @params

            # Move the temporary location to the actual place for the scn2 cab file
            Move-Item ("{0}.new" -f $CabFileLocation) $CabFileLocation -Force

            "Successfully downloaded wsusscn2.cab"
        } else {
            "Local file is newer than the Uri. Not downloading."
        }
    }
}

Function Invoke-PatchingRun
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$OfflineServiceName = "Offline Sync Service",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$PatchDir = "C:\_patching",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$CabDownloadUri = "https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab",

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [bool]$UpdateCab = $true,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [bool]$UseOfflineScan = $false,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [int]$FreeSpaceMinMB = 4096,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [int]$AgeThreshold = 14,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [bool]$CanReboot = $false,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [int]$RebootDelaySec = 300
    )

    process
    {
        # Create the patching directory
        New-Item -Type Directory $PatchDir -EA Ignore
        if (!(Test-Path -PathType Container $PatchDir))
        {
            Write-Error "Could not create patch directory ($PatchDir) or is not a directory"
        }

        # Check for minimum free space on system drive
        $freeMb = (Get-PSDrive "C").Free / 1024 / 1024
        if ($freeMb -lt $FreeSpaceMinMB)
        {
            Write-Error "Insufficient free space on system drive"
        }

        # Check for minimum free space on patching drive
        $drive = (Get-Item $PatchDir).PSDrive
        $freeMb = (Get-PSDrive $drive).Free / 1024 / 1024
        if ($freeMb -lt $FreeSpaceMinMB)
        {
            Write-Error "Insufficient free space on patching drive"
        }

        # Create the ServiceManager object
        "Creating ServiceManager object"
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        # Remove any existing instances of the package service with this name
        "Removing any preexisting service registrations for `"$OfflineServiceName`""
        Remove-UpdatePackageService -Manager $manager -Name $OfflineServiceName

        # Create an update searcher
        "Creating update session using new scan service"
        $session = New-Object -ComObject Microsoft.Update.session
        $searcher = $session.CreateUpdateSearcher()

        # if we're doing an offline scan, the cab file may need updating and the
        # searcher needs to be configured to use the cab file.
        if ($UseOfflineScan)
        {
            $cabFile = [System.IO.Path]::Combine($PatchDir, "wsusscn2.cab")

            # Only update the cab file if specified in the configuration
            if ($UpdateCab)
            {
                Update-CabFile -CabDownloadUri $CabDownloadUri -CabFileLocation $cabFile
            }

            # Whether the update was done or not, make sure the cab file exists in the
            # patch dir
            if (!(Test-Path -PathType Leaf $cabFile))
            {
                Write-Error "Offline scan required, but wsusscn2.cab file not present"
            }

            # Get modification time for the cab file to include in any output
            $cabModificationTime = (Get-item $cabFile).LastWriteTimeUtc
            ("Cab file modification time: " + $cabModificationTime)

            # Update the searcher to use the cab file
            "Updating searcher to use the local cab file"
            $service = $manager.AddScanPackageService($OfflineServiceName, $cabFile, 1)
            $searcher.ServerSelection = 3
            $searcher.ServiceID = $service.ServiceID
        }

        # Capture anything not installed
        "Search for packages that are not installed"
        $result = $searcher.Search("IsInstalled=0")

        # Report on any updates found
        $count = ($result.Updates | Measure-Object).Count
        "Found $count updates not installed"
        $result.Updates |
            Format-List -Property Title, Description, RebootRequired, MsrcSeverity, LastDeploymentChangeTime |
            Out-String

        $updates = New-Object -ComObject Microsoft.Update.UpdateColl
        $AgeThreshold = [Math]::Abs($AgeThreshold)
        $result.Updates |
            Where-Object { $_.LastDeploymentChangeTime -lt ([DateTime]::Now.AddDays(-$AgeThreshold))} |
            ForEach-Object { $updates.Add($_) }

        # Report on applicable updates
        $count = ($updates | Measure-Object).Count
        "Found $count updates meeting age threshold ($AgeThreshold days)"
        $updates |
            Format-List -Property Title, Description, RebootRequired, MsrcSeverity, LastDeploymentChangeTime |
            Out-String

        # Stop here if we're just reporting
        if ($DryRun)
        {
            "DryRun - No download or install"
        } elseif (($updates | Measure-Object).Count -gt 0) {
            # Download any packages
            "Starting download of updates"
            $downloader = New-Object -ComObject Microsoft.Update.Downloader
            $downloader.Updates = $updates
            $downloader.Download()

            # Install any packages
            "Starting installation of updates"
            $installer = New-Object -ComObject Microsoft.Update.Installer
            $installer.ForceQuiet = $true
            $installer.Updates = $updates
            $installer.Install()

            "Update installation completed"
        } else {
            "No updates to install"
        }

        # Remove any instances of the package service with this name
        "Removing the registered scan service"
        Remove-UpdatePackageService -Manager $manager -Name $OfflineServiceName

        # Check if a reboot is required
        $systemInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        $rebootRequired = $systemInfo.RebootRequired
        "Reboot required: $rebootRequired"
        "CanReboot: $CanReboot"
        if ($CanReboot -and $rebootRequired -and -not $DryRun)
        {
            "Scheduling reboot in $RebootDelaySec seconds"
            shutdown -r -f -t "$RebootDelaySec"
        }
    }
}

# Read config file in to a dictionary
$config = @{}
if (![string]::IsNullOrEmpty($ConfigFile))
{
    $configObj = Get-Content -Encoding UTF8 $ConfigFile | ConvertFrom-Json
    $configObj | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" } | ForEach-Object {
        $property = $_.Name

        $config[$property] = $configObj.$property
    }

    # Extract LogFile destination
    if ($config.ContainsKey("LogFile"))
    {
        $script:LogFile = $config["LogFile"] | Out-String -NoNewline
        $config.remove("LogFile")
    }
}

# Truncate log file
if (![string]::IsNullOrEmpty($LogFile))
{
    $content = ""
    if (Test-Path $LogFile)
    {
        $content = Get-Content -Encoding UTF8 $LogFile | Select-Object -Last 2000
    }
    $content | Out-File $LogFile -Encoding UTF8
}

# Start patching process
& {
    try {
        "Starting patching"
        Invoke-PatchingRun @config *>&1
        "Finished"
    } catch {
        LogMessage "Patch apply failed: $_"
    }
} | ForEach-Object {
    LogMessage ($_ | Out-String -NoNewline)
}
