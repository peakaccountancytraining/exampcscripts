# Edit-Environment.ps1: edit Windows environment variables.

# Version history
$Version = "1.00"
# v1.00 04/08/25 New: Original version

# Load Sapphire helper.
Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# Get existing environment variables.
$OldVariableSets = "Machine", "User" | % {[Environment]::GetEnvironmentVariables($_)}

# Shell the edit environment variables app.
Start-Process "rundll32.exe" -ArgumentList "sysdm.cpl,EditEnvironmentVariables" -Wait

# Get/refresh local environment variables.
. "$PSScriptRoot\Get-Environment.ps1"

# Remove any deleted variables.
$NewVariableSets = "Machine", "User" | % {[Environment]::GetEnvironmentVariables($_)}
ForEach ($OldVariableSet In $OldVariableSets) {
	ForEach ($OldKey In $OldVariableSet.Keys) {
		$Found = $False
		ForEach ($NewVariableSet In $NewVariableSets) {If ($NewVariableSet.Contains($OldKey)) {$Found = $True; Break}}
		If (!$Found) {Remove-Item "Env:$OldKey"}
	}
}
