#Requires -RunAsAdministrator

# https://docs.microsoft.com/en-us/windows/terminal/
# https://github.com/microsoft/terminal/releases

Clear-Host

if ($psISE)
{
	exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Installing the latest NuGet
if (-not (Test-Path -Path "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\*\Microsoft.PackageManagement.NuGetProvider.dll"))
{
	Write-Verbose -Message "Installing NuGet" -Verbose

	Install-PackageProvider -Name NuGet -Force
}
if ($null -eq (Get-PackageProvider -ListAvailable | Where-Object -FilterScript {$_.Name -ceq "NuGet"} -ErrorAction Ignore))
{
	Write-Verbose -Message "Installing NuGet" -Verbose

	Install-PackageProvider -Name NuGet -Force
}

$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"

# Install the latest PSResourceGet version
# https://www.powershellgallery.com/packages/PSResourceGet
# https://github.com/PowerShell/PSResourceGet
if ($null -eq (Get-Module -Name PSResourceGet -ListAvailable -ErrorAction Ignore))
{
	# Get latest PackageManagement version
	# https://www.powershellgallery.com/packages/PackageManagement
	# https://github.com/oneget/oneget

	$Parameters = @{
		Uri             = "https://raw.githubusercontent.com/OneGet/oneget/WIP/src/Microsoft.PowerShell.PackageManagement/PackageManagement.psd1"
		OutFile         = "$DownloadsFolder\PackageManagement.psd1"
		UseBasicParsing = $true
		Verbose         = $true
	}
	Invoke-WebRequest @Parameters

	$LatestPackageManagementVersion = (Import-PowerShellDataFile -Path "$DownloadsFolder\PackageManagement.psd1").ModuleVersion

	Remove-Item -Path "$DownloadsFolder\PackageManagement.psd1" -Force
	# If PackageManagement doesn't exist or its' version is lower than the latest one
	$CurrentPackageManagementVersion = ((Get-Module -Name PackageManagement -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()

	if (($null -eq (Get-Module -Name PackageManagement -ListAvailable -ErrorAction Ignore)) -or ([System.Version]$CurrentPackageManagementVersion -lt [System.Version]$LatestPackageManagementVersion))
	{
		Write-Verbose -Message "PackageManagement module doesn't exist" -Verbose
		Write-Verbose -Message "Installing PackageManagement $($LatestPackageManagementVersion)" -Verbose

		# Download nupkg archive to expand it and install
		$DownloadFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
		$Parameters = @{
			Uri             = "https://psg-prod-eastus.azureedge.net/packages/packagemanagement.$($LatestPackageManagementVersion).nupkg"
			OutFile         = "$DownloadFolder\packagemanagement.nupkg"
			UseBasicParsing = $true
		}
		Invoke-RestMethod @Parameters

		Unblock-File -Path "$DownloadFolder\packagemanagement.nupkg"

		Rename-Item -Path "$DownloadFolder\packagemanagement.nupkg" -NewName "packagemanagement.zip" -Force

		# Expanding
		function ExtractZIPFolder
		{
			[CmdletBinding()]
			param
			(
				[string]
				$Source,

				[string]
				$Destination,

				[string[]]
				$Folders,

				[string[]]
				$Files
			)

			Add-Type -Assembly System.IO.Compression.FileSystem

			$ZIP = [IO.Compression.ZipFile]::OpenRead($Source)
			$ZIP.Entries | Where-Object -FilterScript {($_.FullName -like "$($Folders)/*.*") -or ($Files -contains $_.FullName)} | ForEach-Object -Process {
			$File = Join-Path -Path $Destination -ChildPath $_.FullName
				$Parent = Split-Path -Path $File -Parent

				if (-not (Test-Path -Path $Parent))
				{
					New-Item -Path $Parent -Type Directory -Force
				}

				[IO.Compression.ZipFileExtensions]::ExtractToFile($_, $File, $true)
			}

			$ZIP.Dispose()
		}
		$DownloadFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
		$Parameters = @{
			Source      = "$DownloadFolder\packagemanagement.zip"
			Destination = "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement"
			Folders     = "fullclr"
			Files       = @("PackageManagement.format.ps1xml", "PackageManagement.psd1", "PackageManagement.psm1", "PackageManagement.Resources.psd1", "PackageProviderFunctions.psm1")
		}
		ExtractZIPFolder @Parameters

		Get-ChildItem -Path $env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\fullclr -Force | Move-Item -Destination $env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement -Force
		Remove-Item -Path $env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\fullclr -Force
	}

	$LatestPSResourceGetVersion = "2.2.5"
	Write-Verbose -Message "PSResourceGet module doesn't exist" -Verbose
	Write-Verbose -Message "Installing PSResourceGet $($LatestPSResourceGetVersion)" -Verbose

	# Download nupkg archive to expand it and install
	$DownloadFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
	$Parameters = @{
		Uri             = "https://psg-prod-eastus.azureedge.net/packages/PSResourceGet.$($LatestPSResourceGetVersion).nupkg"
		OutFile         = "$DownloadFolder\PSResourceGet.nupkg"
		UseBasicParsing = $true
	}
	Invoke-RestMethod @Parameters

	Unblock-File -Path "$DownloadFolder\PSResourceGet.nupkg"

	Rename-Item -Path "$DownloadFolder\PSResourceGet.nupkg" -NewName "PSResourceGet.zip" -Force

	# Expanding
	function ExtractZIPFolder
	{
		[CmdletBinding()]
		param
		(
			[string]
			$Source,

			[string]
			$Destination,

			[string[]]
			$Folders,

			[string[]]
			$Files
		)

		Add-Type -Assembly System.IO.Compression.FileSystem

		$ZIP = [IO.Compression.ZipFile]::OpenRead($Source)
		$ZIP.Entries | Where-Object -FilterScript {($_.FullName -like "$($Folders)/*.*") -or ($Files -contains $_.FullName)} | ForEach-Object -Process {
		$File = Join-Path -Path $Destination -ChildPath $_.FullName
			$Parent = Split-Path -Path $File -Parent

			if (-not (Test-Path -Path $Parent))
			{
				New-Item -Path $Parent -Type Directory -Force
			}

			[IO.Compression.ZipFileExtensions]::ExtractToFile($_, $File, $true)
		}

		$ZIP.Dispose()
	}
	$DownloadFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"
	$Parameters = @{
		Source      = "$DownloadFolder\PSResourceGet.zip"
		Destination = "$env:ProgramFiles\WindowsPowerShell\Modules\PSResourceGet"
		Folders     = "en-US"
		Files       = @("PSResourceGet.psd1", "PSGet.Format.ps1xml", "PSGet.Resource.psd1", "PSModule.psm1")
	}
	ExtractZIPFolder @Parameters

	# We cannot import PSResourceGet without PackageManagement. So it has to be installed first
	Import-Module PSResourceGet -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PSResourceGet & PackageManagement installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PSResourceGet & PackageManagement installed. Restart the PowerShell session, and re-run the script" -Verbose
	}

	break
}
else
{
	$CurrentPowerShellGetVersion = ((Get-Module -Name PSResourceGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
}

$CurrentStablePSResourceGetVersion = "2.2.5"
if ([System.Version]$CurrentPowerShellGetVersion -lt [System.Version]$CurrentStablePSResourceGetVersion)
{
	Write-Verbose -Message "Installing PSResourceGet $($CurrentStablePSResourceGetVersion)" -Verbose

	Install-Module -Name PSResourceGet -Force
	Install-Module -Name PackageManagement -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PSResourceGet $($CurrentStablePSResourceGetVersion) & PackageManagement installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PSResourceGet $($CurrentStablePSResourceGetVersion) & PackageManagement installed. Restart the PowerShell session, and re-run the script" -Verbose
	}

	break
}

<#
$Parameters = @{
	Uri             = "https://raw.githubusercontent.com/PowerShell/PSResourceGet/master/src/Microsoft.PowerShell.PSResourceGet.psd1"
	OutFile         = "$DownloadsFolder\PSResourceGet.psd1"
	UseBasicParsing = $true
	Verbose         = $true
}
Invoke-WebRequest @Parameters

$LatestPSResourceGetVersion = (Import-PowerShellDataFile -Path "$DownloadsFolder\PSResourceGet.psd1").ModuleVersion

Remove-Item -Path "$DownloadsFolder\PSResourceGet.psd1" -Force

if ([System.Version]$CurrentPowerShellGetVersion -lt [System.Version]$LatestPSResourceGetVersion)
{
	Write-Verbose -Message "Installing PSResourceGet $($LatestPSResourceGetVersion)" -Verbose

	# We cannot install the preview build immediately due to the default 1.0.0.1 build doesn't support -AllowPrerelease
	Install-Module -Name PSResourceGet -AllowPrerelease -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PSResourceGet $($LatestPSResourceGetVersion) installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PSResourceGet $($LatestPSResourceGetVersion) installed. Restart the PowerShell session, and re-run the script" -Verbose
	}

	break
}
#>

# Installing the latest PSReadLine
# https://github.com/PowerShell/PSReadLine/releases
# https://www.powershellgallery.com/packages/PSReadLine
$Parameters = @{
	Uri            = "https://api.github.com/repos/PowerShell/PSReadLine/releases/latest"
	UseBasicParsing = $true
}
$LatestPSReadLineVersion = (Invoke-RestMethod @Parameters).tag_name.Replace("v", "") | Select-Object -First 1

if ($null -eq (Get-Module -Name PSReadline -ListAvailable -ErrorAction Ignore))
{
	Write-Verbose -Message "PSReadline module doesn't exist" -Verbose
	Write-Verbose -Message "Installing PSReadline $($LatestPSReadLineVersion)" -Verbose

	Install-Module -Name PSReadline -Force
	Import-Module -Name PSReadline -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PSReadline $($LatestPSReadLineVersion) installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PSReadline $($LatestPSReadLineVersion) installed. Restart the PowerShell session, and re-run the script" -Verbose
	}

	break
}
else
{
	$CurrentPSReadlineVersion = ((Get-Module -Name PSReadline -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
}

# Installing the latest PSReadLine
if ([System.Version]$CurrentPSReadlineVersion -lt [System.Version]$LatestPSReadLineVersion)
{
	Write-Verbose -Message "Installing PSReadLine $($LatestPSReadLineVersion)" -Verbose

	Install-Module -Name PSReadline -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PSReadline $($LatestPSReadLineVersion) installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PSReadline $($LatestPSReadLineVersion) installed. Restart the PowerShell session, and re-run the script" -Verbose
	}

	break
}

if ([System.Version]$CurrentPSReadlineVersion -eq [System.Version]$LatestPSReadLineVersion)
{
	Write-Verbose -Message "Removing old PSReadLine modules" -Verbose

	# Removing all PSReadLine folders except the latest and the default ones
	Get-Childitem -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PSReadLine" -Force | Where-Object -FilterScript {$_.Name -ne $LatestPSReadLineVersion} | Remove-Item -Recurse -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "All is done. Close this tab and open a new Windows Terminal tab" -Verbose
	}
	else
	{
		Write-Verbose -Message "All is done. Restart the PowerShell session" -Verbose
	}
}
