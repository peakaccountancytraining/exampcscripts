<#
.SYNOPSIS
    Synchronises a source folder with a target folder
.DESCRIPTION
    Wrapper for the robocopy.exe command with more user friendly options and reporting.
.PARAMETER Source
	Source folder path
.PARAMETER Target
	Target folder path
.PARAMETER ExcludeFolders
	Optional: List of folders to exclude from the sync
.PARAMETER ExcludeFiles
	Optional: List of files to exclude from the sync
.PARAMETER Gap
	Optional: slows down the copy by introducing a millisecond gap every 64k transferred. Forces the -Progress switch otherwise gap doesn't work (bug in robocopy)
.PARAMETER Retry
	Optional: retries the copy this many times before skipping the file. Default is to fail immediately
.PARAMETER Age
	Optional: limits the sync to files of this age, e.g. -Age 30 only copies files modified in last 30 days
.PARAMETER Indent
	Optional: indents the report output with indent spaces
.PARAMETER ExcludeBin
	Optional: excludes the Windows recycle bin
.PARAMETER ExcludeOlder
	Optional: Excludes older files
.PARAMETER ExcludeLonely
	Optional: Excludes lonely files (exist in source but not destination)
.PARAMETER ExcludeChanged
	Optional: exclude changed files (file size differs but same timestamp, often OneDrive sync)
.PARAMETER ExcludeExtra
	Optional: exclude extra files, i.e. files that exist in the target but not the source
.PARAMETER ExcludeOnline
	Optional: only copy offline files, i.e. those not downloaded from cloud. Disables mirror
.PARAMETER ExcludeShortcuts
	Optional: ignores OneDrive shortcuts
.PARAMETER OneDrive
	Optional: excludes the ".849C9593-D756-4E56-8D6E-42412F2A707B" file
.PARAMETER SingleThread
	Optional: just use one thread (mainly copying to USB). Show and Progress force single thread
.PARAMETER Preview
	Optional: preview the copy
.PARAMETER Show
	Optional: show files during copy
.PARAMETER CopyAll
	Optional: copy all file attributes
.PARAMETER NoAttributes
	Optional: don't copy file attributes like read-only, hidden
.PARAMETER Progress
	Optional: show percentage progress
.PARAMETER ReportAll
	Optional: report all counts, e.g. skipped, failed
.PARAMETER CreateTarget
	Optional: removes check that target exists and allows robocopy to create it
.PARAMETER SkipEmpty
	Optional: skip empty files (zero bytes)
.PARAMETER Restartable
	Optional: use robocopy in restartable mode
.PARAMETER ShowCommand
	Optional: show robocopy command
.PARAMETER FAT
	Optional: assume FAT file times (ignore 2 second differences)
.PARAMETER Mashup
	Optional: Throws errors instead of writing to console
.EXAMPLE
	Sync-Folder -Source "E:\Temp\Movies" -Target "S:\Movies"
.NOTES
    Author: Rob Nicholson
    Date:   31st Aug 2025

    v1.00 03/01/20 New: Original version
    v1.01 04/01/20 Mod: Removes trailing "\" from path
                   New: MakeFolder switch
                   New: SingleThread switch
    v1.02 19/04/20 New: Exclude folders parameter
                   Mod: Uses single quotes on exclude lists
                   Mod: MakeFolder commented out
                   New: Retry switch added
    v1.03 22/04/20 Fix: Allows E:\ type paths
                   Fix: Empty exclusions lists work
    v1.04 24/04/20 Fix: Unhides the target folder
    v1.05 03/05/20 Mod: Retry is a value
    v1.06 06/06/20 New: Reports robocopy exit codes
                   New: Show switch added
    v1.07 09/06/20 Fix: Final attrib can handle spaces in target folder
    v1.08 23/06/20 Mod: /np switch added
    v1.09 25/08/20 New: Displays a warning at end if preview used
    v1.10 18/10/20 New: ExcludeOlder switch
    v1.11 01/02/21 New: ExcludeOneDriveLock switch to exclude the ".849C9593-D756-4E56-8D6E-42412F2A707B" file
    v1.12 21/02/21 New: LogPrefix added
    v1.13 07/04/21 New: Security switch added
    v1.14 29/04/21 New: NoAttributes switched added
                   New: Progress switched added
    v1.15 05/05/21 New: Passes exit codes
    v1.16 23/06/21 New: Age parameter added
    v1.17 20/07/21 Mod: ExcludeOneDriveLock changed to just OneDrive
    v1.18 06/08/21 Mod: Added comments to options section
                   Mod: Display elapsed in minutes if < 60 mins
    v1.19 01/12/21 New: Can indent output
                   New: Displays number of bytes copied
    v1.20 20/12/21 Mod: If OneDrive parameter specified, then changed files are ignored. This is because OneDrive adds metadata to each new file cocking up file sizes!
    v1.21 14/01/22 Mod: Show switch inferred if Progress switch used
    v1.22 19/01/22 Mod: SingleThread inferred if Preview switch used (so any report is in directory order)
    v1.23 22/01/22 New: SkipZero switched added to suppress report of zero files copied etc.
    v1.24 09/02/22 Mod: SkipZeroReport switch changed to ReportAll
    v1.25 02/03/22 Fix: Grammar error in preview switch warning fixed
    v1.26 24/10/22 Mod: Security switch renamed CopyAll
                   New: CreateTarget switch added
    v1.27 26/10/22 New: SkipEmpty switch added
    v1.28 31/10/22 Mod: Re-organised the setting up of robocopy options so all together
    v1.29 15/12/22 New: Reports number of directories copied
    v1.30 18/01/23 New: IgnoreShortcuts switch added to skip shortcuts in OneDrive
    v1.31 13/02/23 Fix: Test-Path uses EA stop for more robust checking
    v1.32 25/03/23 New: -Restartable switch added
    v1.33 08/04/23 Mod: Time reported in seconds if <1 minute
    v1.34 05/06/23 Mod: ExcludeOlder switch also excludes extra
    v1.35 04/08/23 New: ExcludeChanged flag added
    v1.36 23/08/23 New: ShowCommand switch added
    v1.37 22/09/23 New: FAT switch added
    v1.38 28/09/23 New: ExcludeExtra flag added
                   Mod: LogPrefix parameter removed
                   New: All logs prefixed by "Sync-Folder"
                   New: Deletes very old log files in temp folder (90 days)
    v1.39 30/09/23 New: Sets $Global:FilesCopied variable with count of files copied
                   New: Sets $Global:LogFile variable with name of log file
    v1.40 27/10/23 Fix: If gap is specified, then progress is enabled due to bug in Robocopy . If /np is used, then /ipg is ignored!
    v1.41 06/11/23 Mod: Includes /XJ to exclude junctions
    v1.42 23/02/24 Fix: Added -Force to creating temporary file
    v1.43 25/02/24 New: Changed global variables to single global variable $SyncFolder and added skipped, failed & extra.
    v1.44 01/03/24 New: Added Bytes to $SyncFolder containing total bytes copied
    v1.45 04/04/24 New: Added -ExcludeOnline switch to only copy files downloaded locally from cloud. Disables mirror
                   Mod: -IgnoreShortcuts renamed to -ExcludeShortcuts
    v1.46 06/10/24 New: -ExcludeBin switch to exclude recycle bin
	v1.47 17/10/24 New: This header format added
			       Fix: Uses -LiteralPath on testing source and target folders exist
	v1.48 18/11/24 New: NothingCopied added to SyncFolder object
	v1.49 08/07/25 Mod: Changed code to throw errors in mash-up mode
	v1.50 11/07/25 Fix: Extra count works, was always one before
	v1.51 31/08/25 New: ExcludeLonely switch added
#>

Param(

	[Parameter(Mandatory = $True)][String]$Source,			# source path
	[Parameter(Mandatory = $True)][String]$Target,			# target path

	[String[]]$ExcludeFolders = @(),						# folders to exclude when copying
	[String[]]$ExcludeFiles = @(),							# files to exclude when copying

	[Int]$Gap = 0,											# delay between 64k blocks. forces display of copy otherwise doesn't make effect!
	[Int]$Retry = -1,										# retry the copy if failed
	[Int]$Age = -1,											# age of files to copy
	[Int]$Indent = 0,										# indent the output

	[SWitch]$ExcludeBin,									# exclude recycle bin
	[Switch]$ExcludeOlder,									# exclude older files
	[Switch]$ExcludeChanged,								# exclude changed files (file size differs but same timestamp, often OneDrive sync)
	[Switch]$ExcludeExtra,									# exclude extra files (exist in destination but not source)
	[Switch]$ExcludeLonely,									# exclude lonely files (exist in source but not destination)
	[Switch]$ExcludeOnline,									# only copy offline files, i.e. those not downloaded from cloud
	[Switch]$ExcludeShortcuts,								# ignores OneDrive shortcuts
	[Switch]$OneDrive,										# excludes the ".849C9593-D756-4E56-8D6E-42412F2A707B" file
	[Switch]$SingleThread,									# just use one thread (mainly copying to USB)
	[Switch]$Preview,										# preview the copy
	[Switch]$Show,											# show files during copy
	[Switch]$CopyAll,										# copy all attributes
	[Switch]$NoAttributes,									# don't copy attributes like read-only, hidden
	[Switch]$Progress,										# show percentage progress
	[Switch]$ReportAll,										# report all counts, e.g. skipped, failed
	[Switch]$CreateTarget,									# removes check that target exists and allows robocopy to create it
	[Switch]$SkipEmpty,										# skip empty files (zero bytes)
	[Switch]$Restartable,									# use robocopy in restartable mode
	[Switch]$ShowCommand,									# show robocopy command
	[Switch]$FAT,											# assume FAT file times (ignore 2 second differences)
	[Switch]$Mashup											# throw exception instead of reporting to console
	
)

# Variables.
$DrivesPattern = "^\w:\\$"		# match drive letter
$Spaces = " " * $Indent			# indent spaces
$LogPrefix = "Sync-Folder"		# prefix for temporary files
$LogAge = 30					# number of days to keep the log

# Used to return metadata about copy.
$Global:SyncFolder = @{
	Copied = 0; 
	Skipped = 0; 
	Failed = 0; 
	Extra = 0;
	Bytes = 0;
	NothingCopied = $False
}

# CreateTemporaryFile: create a temporary file.
Function CreateTemporaryFile($Prefix) {
	$TempFile = New-TemporaryFile
	$TempFileName = $TempFile.DirectoryName + "\$Prefix-" + $TempFile.Name
	If (Test-Path $TempFileName) {Remove-Item $TempFileName}
	Rename-Item $TempFile.FullName $TempFileName
	Return $TempFileName
}

# ReportError: report an error and exit.
Function ReportError($Message) {
	If ($Mashup) {Throw $Message.Trim()}
	Write-Warning $Message
	Exit
}

# Get temporary log file.
$Global:LogFile = CreateTemporaryFile $LogPrefix

# Remove trailing slash character which confuses robocopy
$Source = $Source.Trim(); $Target = $Target.Trim()
If ($Source.Substring($Source.Length -1) -eq "\" -and $Source -NotMatch $DrivesPattern) {$Source = $Source.Substring(0, $Source.Length - 1)}
If ($Target.Substring($Target.Length -1) -eq "\" -and $Target -NotMatch $DrivesPattern) {$Target = $Target.Substring(0, $Target.Length - 1)}

# Test folders exist and not instructed to create.
If (!(Test-Path -LiteralPath $Source -EA Stop)) {ReportError "$Spaces`Can't find ""$Source"" folder"}
If ((!$CreateTarget) -And (!(Test-Path -LiteralPath $Target -EA Stop))) {ReportError "$Spaces`Can't find ""$Target"" folder"}

# Turn on some switches based on other switches.
If ($Gap -ne 0) {$Progress = $True}															# bug in robocopy whereby /ipg doesn't work if /np is present
If ($Progress) {$Show = $True}																# show inferred if showing progress
If ($Show) {$SingleThread = $True}															# ensures the listing is in directory order otherwise all over the place

# Exclude OneDrive shortcuts. This is a little sketchy until Microsoft implement some better method of determining whether a folder is a OneDrive shortcut.
# This works by checking the icon on the folder is the the link icon!
If ($ExcludeShortcuts) {
	$Title = "Scanning for OneDrive shortcut folders"
	Write-Progress -Activity $Title
	ForEach ($Folder In Get-ChildItem -LiteralPath $Source -Recurse -Directory) {
		$DesktopIniPath = $Folder.FullName + "\desktop.ini"
		If (Test-Path $DesktopIniPath) {If ((Get-Content $DesktopIniPath) -Match "IconResource=.*OneDrive.exe,[4|7]") {$ExcludeFolders += $Folder.FullName}}
	}
	Write-Progress -Activity $Title -Completed
}

# Set Robocopy options.
$Opts = "/ndl /xj"																			# no directory listing and exclude all junctions/links
If ($Gap -ne 0) {$Opts += " /ipg:$Gap"} ElseIf (!$SingleThread) {$Opts += " /mt"}			# multithreading
If ($Preview) {$Opts += " /l"}																# don't copy, just preview
If ($Retry -eq -1) {$Opts += " /r:0"} Else {$Opts += " /r:$Retry"}							# retry locked files
If ($Show) {$Opts += " /tee /log:""$LogFile"""}												# show robocopy output during copy
If ($ExcludeOlder) {$Opts += " /xo"}														# only copy new/newer, not older
If ($ExcludeChanged) {$Opts += " /xc"}														# exclude changed files (file size different but timestamps equal)
If ($ExcludeLonely) {$Opts += " /xl"}														# exclude lonely files, i.e. new files in source
If ($ExcludeExtra) {$Opts += " /xx"}														# exclude extra files, i.e. new files in the target
If (!$Progress) {$Opts += " /np"}															# show progress percentage
If ($SkipEmpty) {$Opts += " /min:1"}														# skip empty zero byte files
If ($Restartable) {$Opts += " /z"}															# use restartable mode
If ($FAT) {$Opts += " /fft"}																# ignore two second time differences
$ExcludeFolders | % {$Opts += " /xd '" + $_ + "'"}											# exclude folders
If ($ExcludeBin) {$Opts += ' /xd ''$RECYCLE.BIN'''}											# exclude recycle bin
$ExcludeFiles | % {$Opts += " /xf '" + $_ + "'"}											# exclude files
If ($OneDrive) {$Opts += " /xf "".849C9593-D756-4E56-8D6E-42412F2A707B"" /xc"}				# exclude OneDrive file & modified (as OneDrive sync modifies files)
If ($CopyAll) {																				# file attributes
	$Opts += " /copyall"
} Else {
	$Opts += " /copy:DT"
	If (!$NoAttributes) {$Opts += "A"}
}
If ($ExcludeOnline) {$Opts += " /xa:o"}														# only copy files downloaded locally
If ($Age -eq -1) {																			# copying all files
	If ($ExcludeOnline) {$Opts += " /s"} Else {$Opts += " /mir"}							# if online, then don't copy empty folders otherwise mirror
} Else {
	$Opts += " /s /maxage:$Age"																# only copy over a certain age
}

# Execute robocopy.
$Cmd = "robocopy ""$Source"" ""$Target"" $Opts"
If ($ShowCommand) {Write-Host "$Spaces$Cmd"}
$Start = Get-Date
If ($Show) {
	Invoke-Expression $Cmd
} Else {
	Write-Host "$Spaces`Log file: $LogFile"
	Invoke-Expression $Cmd > $LogFile
}
$RobocopyError = $LASTEXITCODE

# Scrape the output log.
$Result = Get-Content $LogFile
$FilesPattern = "^\s+(Dirs|Files)\s+:\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$"
$Matches = $Result -Match $FilesPattern
If ($Matches.Count -ne 0) {
	# Scrape copied, skipped, failed and extra counts.
	ForEach ($Match In $Matches) {
		$Type = $Match -Replace $FilesPattern, '$1'
		If ($Type -eq "Dirs") {
			$DirsCopied = [int]($Match -Replace $FilesPattern, '$3')
		} Else {
			$FilesCopied = [int]($Match -Replace $FilesPattern, '$3')
			$FilesSkipped = [int]($Match -Replace $FilesPattern, '$4')
			$FilesFailed = [int]($Match -Replace $FilesPattern, '$6')
			$FilesExtra = [int]($Match -Replace $FilesPattern, '$7')
			$Global:SyncFolder.Copied = $FilesCopied
			$Global:SyncFolder.Skipped = $FilesSkipped
			$Global:SyncFolder.Failed = $FilesFailed
			$Global:SyncFolder.Extra = $FilesExtra
		}
	}
	# Scrape bytes copied and work out MB, GB etc.
	$BytesPattern = "^\s+Bytes\s+:\s+(?:\d+\.\d+\s+[bkmgt]|\d+)\s+(?:(\d+\.\d+)\s+([kmgt])|(\d+)).*$"
	$BytesMatch = $Result -Match $BytesPattern
	$Copied = ""
	If ($BytesMatch.Count -ne 0) {
		$Copied = $BytesMatch[0] -Replace $BytesPattern, '$3'
		If ($Copied -eq "") {
			$CopiedNumber = [decimal]($BytesMatch[0] -Replace $BytesPattern, '$1')
			$CopiedSuffix = ($BytesMatch[0] -Replace $BytesPattern, '$2').ToUpper()
			$Copied = " ({0:N1}{1}B)" -f $CopiedNumber, $CopiedSuffix
			ForEach ($Multiplier In "kmgt".ToCharArray()) {$CopiedNumber *= 1024; If ($CopiedSuffix -eq $Multiplier) {Break}}
			$Global:SyncFolder.Bytes = $CopiedNumber
		} ElseIf ($Copied -eq "0") {
			$Copied = ""
		} Else {
			$Global:SyncFolder.Bytes = $Copied
			$Copied = " (" + $Copied + "B)"
		}
	}
}

# Handle errors.
Switch ($RobocopyError) {
	0 					 {Write-Host "$Spaces`No errors and no files copied"; $Global:SyncFolder.NothingCopied = $True}
	{$PSItem -band 0x01} {Write-Host "$Spaces`One of more files were copied successfully"}
	{$PSItem -band 0x02} {Write-Host "$Spaces`Extra files or directories were detected"}
	{$PSItem -band 0x04} {ReportError "Mismatched files or directories were detected"}
	{$PSItem -band 0x08} {ReportError "Some files or directories could not be copied and the retry limit was exceeded"}
	{$PSItem -band 0x10} {ReportError "Robocopy did not copy any files (check log)"}
}

# Report file actions.
If (($FilesCopied -ne 0) -or $ReportAll) 	{Write-Host "$Spaces`Copied:  $("{0:N0}{1}" -f $FilesCopied, $Copied)"}
If (($FilesSkipped -ne 0) -or $ReportAll) 	{Write-Host "$Spaces`Skipped: $("{0:N0}" -f $FilesSkipped)"}
If (($FilesFailed -ne 0) -or $ReportAll) 	{Write-Host "$Spaces`Failed:  $("{0:N0}" -f $FilesFailed)"}
If (($FilesExtra -ne 0) -or $ReportAll) 	{Write-Host "$Spaces`Extra:   $("{0:N0}" -f $FilesExtra)"}
If (($DirsCopied -ne 0) -or $ReportAll) 	{Write-Host "$Spaces`Dirs:    $("{0:N0}" -f $DirsCopied)"}
$End = Get-Date
$Elapsed = (New-TimeSpan $Start $End).TotalSeconds
$Period = "seconds"
If ($Elapsed -ge 60) {$Elapsed /= 60; $Period = "minutes"; If ($Elapsed -ge 60) {$Elapsed /= 60; $Period = "hours"}}
Write-Host "$Spaces`Elapsed: $("{0:N}" -f $Elapsed) $Period"
If ($Show) {Write-Host "$Spaces`Log:     $LogFile"}
If ($Preview) {Write-Host; Write-Host "$Spaces`Preview switch used so NO changes made"  -Fore Cyan}

# Put attributes back
Attrib -r -h -s "$Target" | Out-Null

# Delete old log files.
$LogPath = Split-Path $LogFile
$OldDate = (Get-Date).AddDays(-$LogAge)
Get-ChildItem $LogPath -Filter "$LogPrefix*.tmp" | ? LastWriteTime -le $OldDate | Remove-Item
