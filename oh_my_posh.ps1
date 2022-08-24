# https://github.com/JanDeDobbeleer/oh-my-posh
winget install oh-my-posh --source winget --exact --accept-source-agreements

# Install Fira Code Nerd Font
# https://github.com/ryanoasis/nerd-fonts
# https://github.com/tonsky/FiraCode
$DownloadsFolder = Get-ItemPropertyValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name "{374DE290-123F-4565-9164-39C4925E467B}"

$Parameters = @{
	Uri            = "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
	UseBasicParsing = $true
}
$FiraCodeLatestRelease = ((Invoke-RestMethod @Parameters).assets | Where-Object -FilterScript {$_.name -eq "FiraCode.zip"}).browser_download_url

Write-Verbose -Message "Downloading Installing NuGet" -Verbose

$Parameters = @{
	Uri             = $FiraCodeLatestRelease
	OutFile         = "$DownloadsFolder\FiraCode.zip"
	UseBasicParsing = $true
	Verbose         = $true
}
Invoke-WebRequest @Parameters

$Parameters = @{
	Path            = "$DownloadsFolder\FiraCode.zip"
	DestinationPath = "$DownloadsFolder\FiraCode"
	Force           = $true
	Verbose         = $true
}
Expand-Archive @Parameters

Get-ChildItem -Path "$DownloadsFolder\FiraCode" -Recurse -Force | Unblock-File

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

$Fonts = Get-ChildItem -Path "$DownloadsFolder\FiraCode" | Where-Object -FilterScript {($_.Name -match "Fira") -and ($_.Name -match "Compatible")}
foreach ($Font in $Fonts)
{
	if (Test-Path -Path "$env:SystemRoot\Fonts\$($Font.Name)")
	{
		Remove-Item -Path "$env:SystemRoot\Fonts\$($Font.Name)" -Force
		Copy-Item -Path $Font.FullName -Destination $env:SystemRoot\Fonts -Force
	}
	else
	{
		(New-Object -ComObject Shell.Application).NameSpace($ssfFONTS).CopyHere($Font.FullName, $CopyOptions)
	}
}

Remove-Item -Path "$DownloadsFolder\FiraCode.zip", "$DownloadsFolder\FiraCode" -Recurse -Force

# Apply a PowerShell prompt theme
# https://ohmyposh.dev/docs/themes#m365princess
# %USERPROFILE%\Documents\WindowsPowerShell
$Theme = "M365Princess"
if (-not (Test-Path -Path $PROFILE))
{
	Set-Content -Path $PROFILE -Value "oh-my-posh init pwsh --config `"$env:POSH_THEMES_PATH\$($Theme).omp.json`" | Invoke-Expression" -Force
}
else
{
	Add-Content -Path $PROFILE -Value "`noh-my-posh init pwsh --config `"$env:POSH_THEMES_PATH\$($Theme).omp.json`" | Invoke-Expression" -Force
}

# https://www.powershellgallery.com/packages/terminal-icons
# https://github.com/devblackops/Terminal-Icons
Install-Module -Name Terminal-Icons -Force
Add-Content -Path $PROFILE -Value "`nImport-Module -Name Terminal-Icons" -Force

Invoke-Item -Path $PROFILE