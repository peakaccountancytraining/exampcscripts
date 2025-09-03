# Get-Environment.ps1: re-load the current session environment.

# Version history
$Version = "1.00"
# v1.00 16/08/25 New: Original version

# Load Sapphire helper.
Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# Query user and system environment variables.
$VariableSets = "Machine", "User" | % {[Environment]::GetEnvironmentVariables($_)}

# Set local environment, path is handled differently as it's the merge of user and system.
$Path = ""
ForEach ($VariableSet In $VariableSets) {
	ForEach ($Key In $VariableSet.Keys) {
		Set-Item -Path "Env:$Key" -Value $VariableSet[$Key]
		If ($Key -eq "Path") {$Path = [Sapphire]::Append($Path, ";", $VariableSet[$Key])}
	}
}
If ($Path -ne "") {$Env:Path = $Path}
