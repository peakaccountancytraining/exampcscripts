# Get-ExamPCScripts.ps1: downloads Peak exam PC scripts from github.

# Version history
$Version = "1.00"
# v1.00 03/09/25 New: Original version.

# Variables.
$DefaultGitRepository = "peakaccountancytraining/exampcscripts"

# Load Sapphire helper code.
Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# ReportError: report error and exit.
Function ReportError($Message) {
	Write-Warning $Message
	[Sapphire]::AnyKey()
	Exit
}

# ---- START HERE ----
Write-Host "Get-ExamPCScripts v$Version`n"

# Use Git repository environment variable if specified.
$Repository = $DefaultGitRepository
If ($Env:PeakExamScriptsRepository) {$Repository = $Env:PeakExamScriptsRepository}

# Download the file.
Write-Host "Downloading repository"
$TempFolder = [Sapphire]::GetTempFolder("Get-ExamPCScripts-")
$DownloadFile = "$TempFolder\Download.gzip"
$Uri = "https://github.com/$Repository/archive/refs/heads/main.tar.gz"
Try {$Response = Invoke-WebRequest -Uri $Uri -OutFile $DownloadFile -ErrorAction Stop} Catch {ReportError $_}

# Unarchive download.
Push-Location $TempFolder
$Cmd = "tar -xf ""$DownloadFile"""
Invoke-Expression $Cmd
Pop-Location

# Find the folder containing the files.
$Folders = Get-ChildItem $TempFolder -Directory
If (!$Folders) {ReportError "Unable to find scripts in download"}
If ($Folders.Count -gt 1) {ReportError "Too many folders in download"}
$SourcePath = $Folders[0].Fullname
Write-Host "Synchronising to current folder:"
& $PsScriptRoot\Sync-Folder $SourcePath $PSScriptRoot -CreateTarget -Indent 2

# Finished.
[Sapphire]::AnyKey()
