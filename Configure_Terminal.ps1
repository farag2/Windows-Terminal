#Requires -RunAsAdministrator

# https://docs.microsoft.com/en-us/windows/terminal/
# https://github.com/microsoft/terminal/releases

Clear-Host

if ($psISE)
{
	exit
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Test-Path -Path "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json")
{
	$settings = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
}
else
{
	Start-Process -FilePath wt

	Write-Verbose -Message "Restart the PowerShell session and re-run the script" -Verbose

	exit
}

try
{
	$Terminal = Get-Content -Path $settings -Encoding UTF8 -Force | ConvertFrom-Json
}
catch [System.Exception]
{
	Write-Warning -Message "JSON is not valid!"

	Invoke-Item -Path "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState"

	exit
}

#region General
# Copy
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command -eq "copy"} | Where-Object -FilterScript {$_.keys -eq "ctrl+c"}))
{
	$closeTab = [PSCustomObject]@{
		"command" = "copy"
		"keys" = "ctrl+c"
	}
	$Terminal.actions += $closeTab
}

# Paste
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command -eq "paste"} | Where-Object -FilterScript {$_.keys -eq "ctrl+v"}))
{
	$closeTab = [PSCustomObject]@{
		"command" = "paste"
		"keys" = "ctrl+v"
	}
	$Terminal.actions += $closeTab
}

# Close tab
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command -eq "closeTab"} | Where-Object -FilterScript {$_.keys -eq "ctrl+w"}))
{
	$closeTab = [PSCustomObject]@{
		"command" = "closeTab"
		"keys" = "ctrl+w"
	}
	$Terminal.actions += $closeTab
}

# New tab
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command -eq "newTab"} | Where-Object -FilterScript {$_.keys -eq "ctrl+t"}))
{
	$newTab = [PSCustomObject]@{
		"command" = "newTab"
		"keys" = "ctrl+t"
	}
	$Terminal.actions += $newTab
}

# Find
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command -eq "find"} | Where-Object -FilterScript {$_.keys -eq "ctrl+f"}))
{
	$find = [PSCustomObject]@{
		"command" = "find"
		"keys" = "ctrl+f"
	}
	$Terminal.actions += $find
}

# Split pane
if (-not ($Terminal.actions | Where-Object -FilterScript {$_.command.action -eq "splitPane"} | Where-Object -FilterScript {$_.command.split -eq "auto"} | Where-Object -FilterScript {$_.command.splitMode -eq "duplicate"}))
{
	$split = [PSCustomObject]@{
		"action" = "splitPane"
		"split" = "auto"
		"splitMode" = "duplicate"
	}
	$splitPane = [PSCustomObject]@{
		"command" = $split
		"keys" = "ctrl+shift+d"
	}
	$Terminal.actions += $splitPane
}

# No confirmation when closing all tabs
if ($Terminal.confirmCloseAllTabs)
{
	$Terminal.confirmCloseAllTabs = $false
}
else
{
	$Terminal | Add-Member -Name confirmCloseAllTabs -MemberType NoteProperty -Value $false -Force
}

# Set default profile on PowerShell
if ($Terminal.defaultProfile)
{
	$Terminal.defaultProfile = "{61c54bbd-c2c6-5271-96e7-009a87ff44bf}"
}
else
{
	$Terminal | Add-Member -Name defaultProfile -MemberType NoteProperty -Value "{61c54bbd-c2c6-5271-96e7-009a87ff44bf}" -Force
}

# Show tabs in title bar
if ($Terminal.showTabsInTitlebar)
{
	$Terminal.showTabsInTitlebar = $false
}
else
{
	$Terminal | Add-Member -Name showTabsInTitlebar -MemberType NoteProperty -Value $false -Force
}

# Do not restore previous tabs and panes after relaunching
if ($Terminal.firstWindowPreference)
{
	$Terminal.firstWindowPreference = "defaultProfile"
}
else
{
	$Terminal | Add-Member -Name firstWindowPreference -MemberType NoteProperty -Value "defaultProfile" -Force
}
#endregion General

#region defaults
# Set Windows95.gif as a background image
# https://github.com/farag2/Windows_Terminal
if (-not (Test-Path -Path "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\RoamingState\Windows95.gif"))
{
	$Parameters = @{
		Uri             = "https://raw.githubusercontent.com/farag2/Windows_Terminal/main/Windows95.gif"
		OutFile         = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\RoamingState\Windows95.gif"
		UseBasicParsing = $true
		Verbose         = $true
	}
	Invoke-WebRequest @Parameters
}

if ($Terminal.profiles.defaults.backgroundImage)
{
	$Terminal.profiles.defaults.backgroundImage = "ms-appdata:///roaming/Windows95.gif"
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name backgroundImage -MemberType NoteProperty -Value "ms-appdata:///roaming/Windows95.gif" -Force
}

# Background image alignment
if ($Terminal.profiles.defaults.backgroundImageAlignment)
{
	$Terminal.profiles.defaults.backgroundImageAlignment = "bottomRight"
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name backgroundImageAlignment -MemberType NoteProperty -Value bottomRight -Force
}

# Background image opacity
$Value = 0.3
if ($Terminal.profiles.defaults.backgroundImageOpacity)
{
	$Terminal.profiles.defaults.backgroundImageOpacity = $Value
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name backgroundImageOpacity -MemberType NoteProperty -Value 0.3 -Force
}

# Background image stretch mode
if ($Terminal.profiles.defaults.backgroundImageStretchMode)
{
	$Terminal.profiles.defaults.backgroundImageStretchMode = "none"
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name backgroundImageStretchMode -MemberType NoteProperty -Value none -Force
}

# Starting directory
$DesktopFolder = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" -Name Desktop
if ($Terminal.profiles.defaults.startingDirectory)
{
	$Terminal.profiles.defaults.startingDirectory = $DesktopFolder
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name startingDirectory -MemberType NoteProperty -Value $DesktopFolder -Force
}

# Use acrylic
if ($Terminal.profiles.defaults.useAcrylic)
{
	$Terminal.profiles.defaults.useAcrylic = $true
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name useAcrylic -MemberType NoteProperty -Value $true -Force
}

# Acrylic opacity
if ($Terminal.useAcrylicInTabRow)
{
	$Terminal.useAcrylicInTabRow = $true
}
else
{
	$Terminal | Add-Member -Name useAcrylicInTabRow -MemberType NoteProperty -Value $true -Force
}

# Show acrylic in tab row
if ($Terminal.profiles.defaults.useAcrylic)
{
	$Terminal.profiles.defaults.useAcrylic = $true
}
else
{
	$Terminal.profiles.defaults | Add-Member -Name useAcrylic -MemberType NoteProperty -Value $true -Force
}


# Run profile as Administrator by default
if ($Terminal.profiles.defaults.elevate)
{
	$Terminal.profiles.defaults.elevate = $true
}
else
{
	$Terminal.profiles.defaults | Add-Member -MemberType NoteProperty -Name elevate -Value $true -Force
}

# Set "FiraCode NF" as a default font
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

if ((New-Object -TypeName System.Drawing.Text.InstalledFontCollection).Families.Name -contains "FiraCode NF")
{
	if ($Terminal.profiles.defaults.font.face)
	{
		$Terminal.profiles.defaults.font.face = "FiraCode Nerd Font Mono Retina"
	}
	else
	{
		$Terminal.profiles.defaults.font | Add-Member -Name face -MemberType NoteProperty -Value "FiraCode Nerd Font Mono Retina" -Force
	}
}

# Remove trailing white-space in rectangular selection
if ($Terminal.trimBlockSelection)
{
	$Terminal.trimBlockSelection = $true
}
else
{
	$Terminal | Add-Member -Name trimBlockSelection -MemberType NoteProperty -Value $true -Force
}

# Create new tabs in the most recently used window on this desktop. If there's not an existing window on this virtual desktop, then create a new terminal window
if ($Terminal.windowingBehavior)
{
	$Terminal.windowingBehavior = "useExisting"
}
else
{
	$Terminal | Add-Member -Name windowingBehavior -MemberType NoteProperty -Value "useExisting" -Force
}
#endregion defaults

#region Azure
# Hide Azure
if (($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{b453ae62-4e3d-5e58-b989-0a998ec441b8}"}).hidden)
{
	($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{b453ae62-4e3d-5e58-b989-0a998ec441b8}"}).hidden = $true
}
else
{
	$Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{b453ae62-4e3d-5e58-b989-0a998ec441b8}"} | Add-Member -MemberType NoteProperty -Name hidden -Value $true -Force
}
#endregion Azure

#region Powershell Core
if (Test-Path -Path "$env:ProgramFiles\PowerShell\7")
{
	# Set the PowerShell 7 tab name
	if (($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"}).name)
	{
		($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"}).name = "🏆 PowerShell 7"
	}
	else
	{
		$Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"} | Add-Member -MemberType NoteProperty -Name name -Value "🏆 PowerShell 7" -Force
	}

	# Run this profile as Administrator by default
	if (($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"}).elevate)
	{
		($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"}).elevate = $true
	}
	else
	{
		$Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{574e775e-4f2a-5b96-ac1e-a2962a402336}"} | Add-Member -MemberType NoteProperty -Name elevate -Value $true -Force
	}
}

if (Test-Path -Path "$env:ProgramFiles\PowerShell\7-preview")
{
	# Background image stretch mode
	if (($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{a3a2e83a-884a-5379-baa8-16f193a13b21}"}).name)
	{
		($Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{a3a2e83a-884a-5379-baa8-16f193a13b21}"}).name = "🐷 PowerShell 7 Preview"
	}
	else
	{
		$Terminal.profiles.list | Where-Object -FilterScript {$_.guid -eq "{a3a2e83a-884a-5379-baa8-16f193a13b21}"} | Add-Member -MemberType NoteProperty -Name name -Value "🐷 PowerShell 7 Preview" -Force
	}
}
#endregion Powershell Core

ConvertTo-Json -InputObject $Terminal -Depth 4 | Set-Content -Path $settings -Encoding UTF8 -Force
# Re-save in the UTF-8 without BOM encoding due to JSON must not has the BOM: https://datatracker.ietf.org/doc/html/rfc8259#section-8.1
Set-Content -Value (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false).GetBytes($(Get-Content -Path $settings -Raw)) -Encoding Byte -Path $settings -Force

# Set Windows Terminal as default terminal app to host the user interface for command-line applications
$TerminalVersion = (Get-AppxPackage -Name Microsoft.WindowsTerminal).Version
if ([System.Version]$TerminalVersion -ge [System.Version]"1.11")
{
	if (-not (Test-Path -Path "HKCU:\Console\%%Startup"))
	{
		New-Item -Path "HKCU:\Console\%%Startup" -Force
	}

	# Find the current GUID of Windows Terminal
	$PackageFullName = (Get-AppxPackage -Name Microsoft.WindowsTerminal).PackageFullName
		Get-ChildItem -Path "HKLM:\SOFTWARE\Classes\PackagedCom\Package\$PackageFullName\Class" | ForEach-Object -Process {
		if ((Get-ItemPropertyValue -Path $_.PSPath -Name ServerId) -eq 0)
		{
			New-ItemProperty -Path "HKCU:\Console\%%Startup" -Name DelegationConsole -PropertyType String -Value $_.PSChildName -Force
		}

		if ((Get-ItemPropertyValue -Path $_.PSPath -Name ServerId) -eq 1)
		{
			New-ItemProperty -Path "HKCU:\Console\%%Startup" -Name DelegationTerminal -PropertyType String -Value $_.PSChildName -Force
		}
	}
}
