#Requires -RunAsAdministrator

# https://docs.microsoft.com/en-us/windows/terminal/
# https://github.com/microsoft/terminal/releases

Clear-Host

if ($psISE)
{
	exit
}

Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Installing the latest NuGet
if (-not (Test-Path -Path "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\*\Microsoft.PackageManagement.NuGetProvider.dll"))
{
	Write-Verbose -Message "Installing NuGet" -Verbose

	Install-PackageProvider -Name NuGet -Force
}
if (-not (Get-PackageProvider -ListAvailable | Where-Object -FilterScript {$_.Name -eq "NuGet"} -ErrorAction Ignore))
{
	Write-Verbose -Message "Installing NuGet" -Verbose

	Install-PackageProvider -Name NuGet -Force
}

$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"

#region PackageManagement
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
$LatestPackageManagementVersion = [System.Version](Import-PowerShellDataFile -Path "$DownloadsFolder\PackageManagement.psd1").ModuleVersion

Remove-Item -Path "$DownloadsFolder\PackageManagement.psd1" -Force

$CurrentPackageManagementVersion = ((Get-Module -Name PackageManagement -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()

if (-not (Get-Module -Name PackageManagement -ListAvailable -ErrorAction Ignore))
{
	Write-Verbose -Message "Install PackageManagement manually" -Verbose
	Write-Verbose -Message "https://www.powershellgallery.com/packages/PackageManagement" -Verbose
	Write-Verbose -Message "https://learn.microsoft.com/en-us/powershell/gallery/how-to/working-with-packages/manual-download#installing-powershell-modules-from-a-nuget-package" -Verbose

	exit
}

if ([System.Version]$CurrentPackageManagementVersion -lt $LatestPackageManagementVersion)
{
	Write-Verbose -Message "Installing PackageManagement $($LatestPackageManagementVersion)" -Verbose

	Install-Module -Name PackageManagement -RequiredVersion $LatestPackageManagementVersion -Force

	$PackageManagementVersion = ((Get-Module -Name PackageManagement -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PackageManagement -RequiredVersion $PackageManagementVersion -Force
}
else
{
	Write-Verbose -Message "PackageManagement module installed. Importing..." -Verbose

	$PackageManagementVersion = ((Get-Module -Name PackageManagement -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PackageManagement -RequiredVersion $PackageManagementVersion -Force
}
#endregion PackageManagement

#region PowerShellGet
# Install latest PowerShellGet version
# https://www.powershellgallery.com/packages/PowerShellGet
# https://github.com/PowerShell/PowerShellGet
if ((-not (Get-Module -Name PowerShellGet -ListAvailable -ErrorAction Ignore)))
{
	Write-Verbose -Message "Install PowerShellGet manually" -Verbose
	Write-Verbose -Message "https://www.powershellgallery.com/packages/PowerShellGet" -Verbose
	Write-Verbose -Message "https://learn.microsoft.com/en-us/powershell/gallery/how-to/working-with-packages/manual-download#installing-powershell-modules-from-a-nuget-package" -Verbose

	exit
}
else
{
	Write-Verbose -Message "PowerShellGet module installed. Importing..." -Verbose

	$PowerShellGetVersion = ((Get-Module -Name PowerShellGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PowerShellGet -RequiredVersion $LatestPowerShellGetVersion -Force
}

$CurrentPowerShellGetVersion = [System.Version]((Get-Module -Name PowerShellGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
if ($CurrentPowerShellGetVersion -lt [System.Version]"2.2.5")
{
	Write-Verbose -Message "Installing PowerShellGet $($LatestPowerShellGetModuleVersion)" -Verbose

	Install-Module -Name PowerShellGet -Force

	$PowerShellGetVersion = ((Get-Module -Name PowerShellGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PowerShellGet -RequiredVersion $PowerShellGetVersion -Force

	$PackageManagementVersion = ((Get-Module -Name PackageManagement -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PackageManagement -RequiredVersion $PackageManagementVersion -Force

	if ($env:WT_SESSION)
	{
		Write-Verbose -Message "PowerShellGet $($PowerShellGetModuleVersion) & PackageManagement installed. Close this tab and open a new Windows Terminal tab, and re-run the script" -Verbose
	}
	else
	{
		Write-Verbose -Message "PowerShellGet $PowerShellGetModuleVersion) & PackageManagement installed. Restart the PowerShell session, and re-run the script" -Verbose
	}
}
#endregion PowerShellGet

# Install the latest PSResourceGet version
# https://www.powershellgallery.com/packages/Microsoft.PowerShell.PSResourceGet
# https://github.com/PowerShell/PSResourceGet
$Parameters = @{
	Uri             = "https://api.github.com/repos/PowerShell/PSResourceGet/releases/latest"
	UseBasicParsing = $true
	Verbose         = $true
}
$LatestPSResourceGetVersion = (Invoke-RestMethod @Parameters).tag_name.Replace("v", "")

if (-not (Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable -ErrorAction Ignore))
{
	Write-Verbose -Message "Installing PSResourceGet $($LatestPSResourceGetVersion)" -Verbose

	$PowerShellGetVersion = ((Get-Module -Name PowerShellGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
	Import-Module -Name PowerShellGet -RequiredVersion $PowerShellGetVersion -Force

	Install-Module -Name Microsoft.PowerShell.PSResourceGet -Force

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
else
{
	$CurrentPSResourceGetVersion = [System.Version]((Get-Module -Name Microsoft.PowerShell.PSResourceGet -ListAvailable).Version | Measure-Object -Maximum).Maximum.ToString()
}

if ($CurrentPSResourceGetVersion -lt $LatestPSResourceGetVersion)
{
	Write-Verbose -Message "Installing PSResourceGet $($LatestPSResourceGetVersion)" -Verbose

	Install-Module -Name Microsoft.PowerShell.PSResourceGet -Force

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
else
{
	Write-Verbose -Message "PSResourceGet $($LatestPSResourceGetVersion) already installed" -Verbose
}

Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Installing the latest PSReadLine
# https://github.com/PowerShell/PSReadLine/releases
# https://www.powershellgallery.com/packages/PSReadLine
$Parameters = @{
	Uri            = "https://api.github.com/repos/PowerShell/PSReadLine/releases/latest"
	UseBasicParsing = $true
}
$LatestPSReadLineVersion = [System.Version]((Invoke-RestMethod @Parameters).tag_name.Replace("v", "") | Select-Object -First 1)

if (-not (Get-Module -Name PSReadline -ListAvailable -ErrorAction Ignore))
{
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
if ([System.Version]$CurrentPSReadlineVersion -lt $LatestPSReadLineVersion)
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
