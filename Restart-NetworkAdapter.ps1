# Restart-NetworkAdapter.ps1: disables and enables network adapter.

Param(
	[Parameter(Mandatory)][String]$Name,		# name of adapter to Restart-Computer
	[Int]$Sleep = 0								# how long to sleep between disabling and enabling (seconds)
)

# Version history.
$Version = "v1.04"
# v1.00 01/01/24 New: Original version.
# v1.01 11/03/24 Mod: Can specify adapter name.
# 				 Mod: Renamed Restart-NetworkAdapter.
# v1.02 18/03/24 Fix: Status message referred to "Ethernet", not the name of the interface.
#				 Fix: Changed "interface" to "adapter"
# v1.03 05/04/24 Mod: The name of the adapter must be specified. No wildcards!
# v1.04 07/04/24 New: Restarts the adapter if duration is zero as opposed to disable/enable
#			     Mod: Duration defaults to 0

# Check if running as administrator
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Result = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
If (!$Result) {Write-Warning "Must be run as administrator"; Exit}

# Check adapter exists.
$Adapter = Get-NetAdapter $Name -EA SilentlyContinue
If (!$Adapter) {Write-Warning "Unable to find adapter $Name"; Exit}
If ($Adapter.Status -ne "Up") {Write-Warning "Adapter $Name isn't up"; Exit}

# Disable, sleep, enable...
$Name = $Adapter.Name
If ($Sleep -eq 0) {
	Write-Host "Restarting $Name adapter"
	Start-Sleep 1
	$Adapter | Restart-NetAdapter
} Else {
	Write-Host "Disabling $Name adapter before sleeping for $Sleep seconds"
	Start-Sleep 1
	$Adapter | Disable-NetAdapter -Confirm:$False
	Write-Host "Sleeping for $Sleep seconds"
	Start-Sleep $Sleep
	Write-Host "Enabling $Name adapter"
	$Adapter | Enable-NetAdapter
}
