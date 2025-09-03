# Get-NetworkInfo.ps1: get network information.

Param(
	[String]$Name = ".*",	# name of adapter to check (regex)
	[Switch]$Renew,			# renew IP address first
	[Switch]$Wait			# wait until adapter is up
)

# Version history:
$Version = "1.13"
# v1.00 03/04/23 New: Original version
# v1.01 04/04/23 New: Reports link speed
#				 Fix: Traps where no IP address assigned (virtual)
# v1.02 14/10/23 New: Includes DHCP server address
#				 New: Renews IP address first via switch
# v1.03 12/12/23 New: Includes subnet mask
# v1.04 05/04/24 Mod: Reports on all up adapters
# v1.05 06/04/24 Fix: Sleeps for 5 seconds after renewing to allow re-connection to Wi-Fi
# v1.06 07/04/24 New: Wait switch add to wait until (at least one) adapter is up
#				 New: Can specify name of adapter
# v1.07 08/04/24 New: WiFi switch added to return information about Wi-Fi
# v1.08 21/05/24 Mod: Uses wait switch when renewing
# v1.09 24/10/24 Mod: Adapters sorted alphabetically
# v1.10 27/03/25 Mod: Removed WiFi switch, always reports
#				 New: Pauses if run from desktop
# v1.11 28/03/25 Mod: Displays WiFi N if known
# v1.12 27/08/25 New: Includes MAC, WiFi band and AP SSID
# v1.12 31/08/25 Fix: Gets AP SSID on Windows 10 (different pattern)
#				 Fix: Band isn't known on Windows 10
# 				 Mod: Throws error if location permissions missing for getting Wi-Fi information
# v1.13 02/09/25 Mod: Name parameter is a regexp

# Load Sapphire helper.
Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# Variables.
$MarkerPattern = "^\s+Name\s+:\s+(.*)$"
$WiFiParameters = @(
	@{Name = "Radio Type"; 		Pattern = "^\s+Radio type\s+:\s+(.*)$"}
	@{Name = "Band"; 			Pattern = "^\s+Band\s+:\s+(.*)$"}
	@{Name = "Signal"; 			Pattern = "^\s+Signal\s+:\s+(.*)$"}
	@{Name = "SSID"; 			Pattern = "^\s+SSID\s+:\s+(.*)$"}
	@{Name = "AP SSID";			Pattern = "^\s+(?:AP )?BSSID\s+:\s+(.*)$"}
	@{Name = "Authentication";	Pattern = "^\s+Authentication\s+:\s+(.*)$"}
)
$RadioTypes = @{
	"802.11n" 	= "Wi-Fi 4"
	"802.11ac" 	= "Wi-Fi 5"
	"802.11ax" 	= "Wi-Fi 6"
}

# Check name is valid regex.
Try {$Result = "" -Match $Name} Catch {Write-Warning "Name isn't a valid regex"; Exit}

# Check name is valid regex.
Try {$Result = "" -Match $Name} Catch {Write-Warning "Name isn't a valid regex"; Exit}

# Renew DHCP IP address.
If ($Renew) {
	ipconfig /renew | Out-Null
	$Wait = $True
}

# Load JSON mapping of MAC to access point names.
$JsonPath = "$PSScriptRoot\" + [IO.Path]::GetFileNameWithoutExtension($Script:MyInvocation.MyCommand) + ".json"
If (Test-Path $JSONPath) {Try {$AccessPoints = [Sapphire]::LoadJson($JsonPath)} Catch {Write-Warning $_; Exit}}

# Loop until up.
$First = $True
While ($True) {
	
	# Get adapters.
	$Adapters = Get-NetAdapter | ? {($_.Name -Match $Name ) -And ($_.Status -eq "Up")} | Sort-Object Name
	
	# Loop for each adapter.
	ForEach ($NetAdapter In $Adapters) {
		
		# Skip any adapters without an IP address.
		$IPAddress = $NetAdapter | Get-NetIPAddress -AddressFamily IPv4 -EA Silent
		If (!$IPAddress) {Continue} # no IP address assigned so ignore
		
		# Get WiFi information for first adapter.
		If ($First) {
			$WlanInfo = netsh wlan show interfaces
			If (($WlanInfo -Match "location permission").Count) {Write-Warning "Location permission needed to read wlan interfaces: ms-settings:privacy-location"; Exit}
			$First = $False
		}

		# Build output.
		$IP = $IPAddress.IPAddress + "/" + $IPAddress.PrefixLength
		$DNS = ($NetAdapter | Get-DnsClientServerAddress -AddressFamily IPv4).ServerAddresses
		If ($DNS.Count -eq 0) {$DNS = ""}
		$GW = ($NetAdapter | Get-NetIPConfiguration).IPv4Defaultgateway.NextHop
		$DHCP = (Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "InterfaceIndex=$($NetAdapter.InterfaceIndex)").DHCPServer
		If (!$DHCP) {$DHCP = "Static"}
		$Speed = $NetAdapter.LinkSpeed
		$Output = [Ordered]@{Name = $NetAdapter.Name; IP = $IP; MAC = $NetAdapter.MacAddress.Replace("-",":"); DNS = $DNS; Gateway = $GW; DHCP = $DHCP; Speed = $Speed}
		$Info = [Sapphire]::ScrapeArray($WlanInfo, $MarkerPattern, $NetAdapter.Name, $WiFiParameters)
		If ($Info.Count) {
			$RadioType = $Info["Radio Type"]
			If ($RadioTypes.Contains($RadioType)) {$RadioType = $RadioTypes[$RadioType] + " (" + $RadioType + ")"}
			$Info["Radio Type"] = $RadioType
			$APSSID = $Info["AP SSID"].ToUpper()
			If ($AccessPoints -And $AccessPoints.AccessPoints.$APSSID) {$APSSID += " (" + $AccessPoints.AccessPoints.$APSSID + ")"}
			$Info["AP SSID"] = $APSSID
			If ($Info.Band) {$Info["Band"] = $Info["Band"].Replace(" ", "")}
			[Sapphire]::MergeHashTable($Output, $Info)
		}
		[PSCustomObject] $Output
		$Up = $True
	}
	If ($Up -Or !$Wait) {Break}
	Start-Sleep 5
}

# Report if no matching up adapters found.
If (!$Up) {Write-Host -Fore Cyan "No connected adapters found matching ""$Name"""}

# Pause if run from desktop.
If ([Sapphire]::FromDesktop()) {[Sapphire]::AnyKey()}