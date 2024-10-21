# https://github.com/ryanoasis/nerd-fonts

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Progress bar can significantly impact cmdlet performance
# https://github.com/PowerShell/PowerShell/issues/2138
$ProgressPreference = "SilentlyContinue"

$Parameters = @{
	Uri             = "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
	UseBasicParsing = $true
	Verbose         = $true
}
$URL = ((Invoke-RestMethod @Parameters).assets | Where-Object -FilterScript {$_.browser_download_url -like "*CascadiaCode*zip"}).browser_download_url

$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"

$Parameters = @{
	Uri             = $URL
	OutFile         = "$DownloadsFolder\CascadiaCode.zip"
	UseBasicParsing = $true
	Verbose         = $true
}
Invoke-WebRequest @Parameters

$Parameters = @{
	Path            = "$DownloadsFolder\CascadiaCode.zip"
	DestinationPath = "$DownloadsFolder\CascadiaCode"
	Force           = $true
	Verbose         = $true
}
Expand-Archive @Parameters

Get-ChildItem -Path "$DownloadsFolder\CascadiaCode" -Recurse -Force | Unblock-File

if (-not ("System.Drawing.Text.PrivateFontCollection" -as [type]))
{
	Add-Type -AssemblyName System.Drawing
}
$fontCollection = New-Object -TypeName System.Drawing.Text.PrivateFontCollection
$fontCollection.AddFontFile("$DownloadsFolder\CascadiaCode.zip")

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
$Installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families | Where-Object -FilterScript {$_.Name -eq "CaskaydiaCove NF"}

if (-not $Installed)
{
	# Installing fonts
	# https://docs.microsoft.com/en-us/windows/desktop/api/Shldisp/ne-shldisp-shellspecialfolderconstants
	# https://docs.microsoft.com/en-us/windows/win32/shell/folder-copyhere
	$ssfFONTS                  = 20

	# https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shfileopstructa
	$FOF_SILENT                = 4
	$FOF_NOCONFIRMATION        = 16
	$FOF_NOERRORUI             = 1024
	$FOF_NOCOPYSECURITYATTRIBS = 2048

	$CopyOptions = $FOF_SILENT + $FOF_NOCONFIRMATION + $FOF_NOERRORUI + $FOF_NOCOPYSECURITYATTRIBS

	$ShellApp = New-Object -ComObject Shell.Application
	$FontsFolder = $ShellApp.NameSpace($ssfFONTS)
	$null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ShellApp)

	$FontsFolder.CopyHere("$DownloadsFolder\CascadiaCode\CaskaydiaCoveNerdFont-Regular.ttf", $CopyOptions)

	Remove-Item -Path "$DownloadsFolder\CascadiaCode.zip", "$DownloadsFolder\CascadiaCode" -Recurse -Force

	$null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($FontsFolder)
	$null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ShellApp)
}
else
{
	Write-Warning -Message "Already Installed"
}
