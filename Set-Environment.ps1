# Set-Environment.ps1: set environment variable permanently and in the current session.

Param(
	[Parameter(Mandatory)][String]$Name,		# name of environment variable to set
	[Parameter(Mandatory)][String]$Value,		# value to set
	[Switch]$Quiet								# quiet operation
)

# Version history
$Version = "1.00"
# v1.00 16/08/25 New: Original version
# v1.01 18/08/25 New: Quiet switch added

Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}
$EnvVar = Get-Item -Path "Env:$Name" -EA SilentlyContinue
If ($EnvVar) {
	$Message = [Sapphire]::IIf($EnvVar.Value -ne $Value, "Changing $Name", "$Name unchanged")
} Else {
	$Message = "Setting $Name"
}
If (!$Quiet) {Write-Host $Message}
setx $Name $Value | Out-Null
Set-Item -Path "Env:$Name" -Value $Value