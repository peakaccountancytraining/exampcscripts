# Configure-PeakStudent.ps1: configure student account.

Param([Switch]$AlwaysInstall)

# Version history
$Version = "1.03"
# v1.00 04/08/25 New: Original version.
# v1.01 20/08/25 New: Uninstall lots of apps
# v1.02 23/08/25 New: Various tweaks for main release of new exam PCs
# v1.03 02/09/25 Mod: Uninstalled OneDrive earlier in sequence
#				 New: Sets power settings

# Load Sapphire helper code.
Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# Packages to uninstall.
$Uninstalls = @(
	@{Name = "OneDrive"; Id = "microsoft.onedrive"; WinGet = $True}
	@{Name = "Camera"; Id = "Microsoft.WindowsCamera"}
	@{Name = "Co-Pilot"; Id = "Microsoft.Copilot"}
	@{Name = "Feedback Hub"; Id = "Microsoft.WindowsFeedbackHub"}
	@{Name = "Game Assist"; Id = "Microsoft.Edge.GameAssist"}
	@{Name = "Microsoft Bing"; Id = "Microsoft.BingSearch"}
	@{Name = "Microsoft Clipchamp"; Id = "Clipchamp.Clipchamp"}
	@{Name = "Microsoft To Do"; Id = "Microsoft.Todos"}
	@{Name = "News"; Id = "Microsoft.BingNews"}
	@{Name = "Outlook"; Id = "Microsoft.OutlookForWindows"}
	@{Name = "Paint"; Id = "Microsoft.Paint"}
	@{Name = "Power Automate"; Id = "Microsoft.PowerAutomateDesktop"}
	@{Name = "Solitaire & Casual Games"; Id = "Microsoft.MicrosoftSolitaireCollection"}
	@{Name = "Teams"; Id = "microsoft.teams"; WinGet = $True}
	@{Name = "Weather"; Id = "Microsoft.BingWeather"}
	@{Name = "XBox Live"; Id = "Microsoft.Xbox.TCUI"}
	@{Name = "Xbox"; Id = "9MV0B5HZVK9Z"; WinGet = $True}
	@{Name = "Microsoft Store"; Id = "Microsoft.WindowsStore"}
)

# ReportError: report (optional) error and exit.
Function ReportError($Message) {
	If ($Message) {Write-Warning $Message}
	Write-Host -Fore Red "Script aborted..."
	[Sapphire]::AnyKey()
	Exit
}

# ExpandVariables: expand variables in string.
Function ExpandVariables($Text) {
	ForEach ($EnvVar In $Variables) {$Text = $Text.Replace("%$EnvVar%", ([System.Environment]::GetEnvironmentVariable($EnvVar)))}
	Return $Text
}

# Starts here...
# If ([Sapphire]::FromDesktop()) {Write-Host ("`n" * 10)}
Write-Host "Configure-PEAKStudent v$Version running on $($Env:ComputerName)`n"
[Sapphire]::AnyKey("Use local.admin account if prompted to run as admin. Press any key to continue...", $False, [System.ConsoleColor]::Red)

# Uninstall packages and apps.
$Errors = $False; $RebootRequired = $False; $FirstWinGet = $True
ForEach ($Uninstall In $Uninstalls) {
	If ($Uninstall.WinGet) {
		If ($FirstWinGet) {winget list --accept-source-agreements | out-null; $FirstWinGet = $False}
		$Id = $Uninstall.Id
		$Result = winget list $Id | select-string " $([regex]::Escape($Id)) "
		If ($Result) {
			Write-Host "Uninstalling $($Uninstall.Name)"
			$Result = Uninstall-WinGetPackage -Id $Uninstall.Id
			If ($Result.UninstallerErrorCode -ne 0) {
				Write-Warning "Installer error code $($Result.UninstallerErrorCode)"
				$Errors = $True
			} ElseIf ($Result.RebootRequired) {
				$RebootRequired = $True
			}
		}
	} Else {
		$ProgressPreference = "SilentlyContinue"
		$Packages = Get-AppxPackage $Uninstall.Id
		If ($Packages) {
			Write-Host "Uninstalling $($Uninstall.Name)"
			$Packages | Remove-Appxpackage
		}
		$ProgressPreference = "Continue"
	}
}
If ($Errors) {Write-Warning "Script aborting"}
If ($RebootRequired -And ([Sapphire]::ReadKey("Reboot required. Reboot now? ", "Yellow") -eq "y")) {Restart-Computer}
If ($Errors) {ReportError}

# Configure screen saver to kick in after an hour.
$RegPath = "HKCU:\Control Panel\Desktop"
$RegistryValues = @(
	@{Path = $RegPath; Name = "SCRNSAVE.EXE"; Value = "$($Env:WinDir)\System32\Mystify.scr"}
	@{Path = $RegPath; Name = "ScreenSaveActive"; Value = "1"}
	@{Path = $RegPath; Name = "ScreenSaveTimeOut"; Value = "3600"}
	@{Path = $RegPath; Name = "ScreenSaverIsSecure"; Value = "1"}
)
If ([Sapphire]::ChangeRegistryList($RegistryValues)) {Write-Host "Screen saver configured"}

# Configure power settings.
$PowerCfgResult = powercfg /query
$MarkerPattern = "^\s+GUID Alias: (.*)$"; 
$AC = @{Name = "Value"; Pattern = "^\s+Current AC Power Setting Index: (.*)$"}
$DC = @{Name = "Value"; Pattern = "^\s+Current DC Power Setting Index: (.*)$"}
$ACPowerSettings = @(
	@{Name = "power screen off";	MarkerValue = "VIDEOIDLE"; 		Parameters = $AC;	Value = 120;	Text = "2 hours";	Setting = "monitor-timeout-ac"}
	@{Name = "power sleep"; 		MarkerValue = "STANDBYIDLE"; 	Parameters = $AC;	Value = 0; 		Text = "never"; 	Setting = "standby-timeout-ac"}
	@{Name = "power hibernate"; 	MarkerValue = "HIBERNATEIDLE"; 	Parameters = $AC;	Value = 0; 		Text = "never"; 	Setting = "hibernate-timeout-ac"}
)
$DCPowerSettings = @(
	@{Name = "battery screen off";	MarkerValue = "VIDEOIDLE"; 		Parameters = $DC;	Value = 120;	Text = "2 hours";	Setting = "monitor-timeout-dc"}
	@{Name = "battery sleep"; 		MarkerValue = "STANDBYIDLE"; 	Parameters = $DC;	Value = 0; 		Text = "never"; 	Setting = "standby-timeout-dc"}
	@{Name = "battery hibernate"; 	MarkerValue = "HIBERNATEIDLE"; 	Parameters = $DC;	Value = 0; 		Text = "never"; 	Setting = "hibernate-timeout-dc"}
)
$PowerSettings = $ACPowerSettings
If (Get-WmiObject -Class Win32_Battery) {$PowerSettings += $DCPowerSettings}
ForEach ($PowerSetting In $PowerSettings) {
	$Result = [Sapphire]::ScrapeArray($PowerCfgResult, $MarkerPattern, $PowerSetting.MarkerValue, $PowerSetting.Parameters)
	If ($Result.Value) {
		$Value = [int]$Result.Value
		If ($Value -ne ($PowerSetting.Value * 60)) {
			Write-Host "Changing $($PowerSetting.Name) to $($PowerSetting.Text)"
			$Cmd = "powercfg /change $($PowerSetting.Setting) $($PowerSetting.Value)"
			Invoke-Expression $Cmd
		}
	}
}

# Remove all desktop shortcuts.
$Items = Get-ChildItem ([Environment]::GetFolderPath('Desktop')) -File -Filter "*.lnk"
$Items = $Items | ? Name -ne "Atlas Cloud.lnk"
If ($Items) {Write-Host "Cleaning desktop"; $Items | Remove-Item}

# Add Atlas Cloud shortcut.
If ([Sapphire]::AddDesktopShortcut("Atlas Cloud", "Atlas Cloud website", "https://aat.psionline.com")) {Write-Host "Added Atlas Cloud shortcut"}
[Sapphire]::RefreshDesktop()

# Turn on file extensions.
If ([Sapphire]::ChangeRegistry("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0)) {Write-Host "Showing file name extensions in File Explorer"}

# Launch Excel.
Write-Host -Fore Green "Excel: Activate with officeXX account and then switch account to student"
Start-Process excel

# Finished.
[Sapphire]::AnyKey()