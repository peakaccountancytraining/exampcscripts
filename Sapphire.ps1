# Sapphire Helper class: collection of useful functions used by other scripts.

# Load Sapphire helper code.
# Try {. $PSScriptRoot\Sapphire.ps1} Catch {Write-Warning "Unable to load Sapphire.ps1`n$_"; Exit}

# Version History:
# v1.00 12/08/21 New: Original version
# v1.52 16/01/25 Mod: Older versions moves to end
#			 	 Mod: Renamed MissingCSVHeadings to CheckCSVHeadings
# v1.53 04/07/25 Fix: CheckCSVHeadings always returned true!
#				 New: LoadJSON function added
# v1.54 07/07/25 Mod: CheckCSVHeadings throws exception
# v1.55 10/07/25 Mod: GetEnvironment can handle different environment variable name to PowerShell variable using comma format, e.g. PEAKOneDrive,Path
#				 New: UnicodePath function added to prefix path with "\\?\"
#				 Fix: ClearProgress pauses for a second on PowerShell v7 otherwise next progress now show
#				 New: IsVersion7 script variable defined
#			 	 New: IIf function added
# v1.56 14/07/25 Fix: FromDesktop handles PSv7
# v1.57 15/07/25 New: StripQuotes function added
# v1.58 23/07/25 Mod: AddCommandPath adds to local session and permanently via SETX. Also refactored to split and check
# v1.59 24/07/25 Mod: RunAdministrator doesn't need invocation parameter
# v1.60 30/07/25 Fix: CreateFolder and CopyNewer work with Unicode paths on PS7
# v1.61 31/07/25 New: Colour of the ReadKey prompt can be specified
#				 Mod: Renamed XXXRegistryValue to XXXRegistry
#				 New: ChangeRegistryList added
# v1.62 05/08/25 New: RefreshDesktop and AddDesktopShortcut added
# v1.63 05/08/25 New: AddCommandPath switch to remove missing and duplicate entries
#				 New: Added TrimTrailing
#				 Mod: GetFile throws an exception if it's unable to read the file
# v1.64 16/08/25 New: Left/Right with one parameter fetches one character
#				 New: Append function added
#				 New: LoadJSON can be called with no parameter in which case it uses the path of the script with .json extension
#			 	 Mod: AddCommandPath removed (see Add-CommandPath script)
#				 New: Added GetManufacturer
#				 Mod: SendEmail throws errors
#				 New: AnyKey can accept colour for output
#				 Mod: Install throws errors instead of printing
#				 New: Added Replace parameter to AddDesktopShortcut
#				 Mod: GetRegistry returns object so can be used with any data type
#				 Mod: AddDesktopShortcut parameters changed to allow arguments and icon location
#				 Mod: AddShortcut added
# 				 Mod: RunAdministrator converts scripts run from mapped drive into expanded version
#				 New: RunAdministrator can add -FromDesktop switch if needed
#				 Mod: CreateFolder and CopyNewer throw errors
#				 Fix: ScrapeArray returns ordered hash table
#				 Mod: ScrapeArray can handle multiple matches for the pattern in which case the value for that key is an array
# v1.65 03/09/25 Mod: GetEnvironment throws exception, not return message
#				 New: GetTempFolder added

# Script variables.
$Progress = @{}											# hash table of per level progress classes
$IsVersion7 = $PSVersionTable.PSVersion.Major -ge 7		# PS version 7 flag

# SapphireProgress: variables for the progress system.
Class SapphireProgress {

	# Member variables.
	[String]$Activity																	# Record of progress titles
	[DateTime]$StartTime = (Get-Date)													# Time progress started
	[System.Diagnostics.Stopwatch]$Timer = [System.Diagnostics.Stopwatch]::StartNew()	# Timers to prevent many calls to Write-Progress which can be slower than the code
	[Boolean]$First = $True																# Always display first call to ShowProgress (skip timer check)
	[Int]$Delay																			# Millisecond delay between updating ShowProgress. Default 500ms

	# Constructor.
	SapphireProgress($Activity, $Delay) {
		$this.Activity = $Activity
		$this.Delay = $Delay
	}

}

# SapphireShadowCopy: creates and deletes shadow copies.
Class SapphireShadowCopy {
	
	# Member variables.
	[Object]$ShadowCopy			# the shadowcopy object
	[String]$Junction			# temporary junction
	[String]$Path				# path within the junction
	
	# Constructor.
	SapphireShadowCopy($Path) {
		
		# Create shadow copy.
		$Drive = $Path.SubString(0, 2)
		$Volume = $Drive + "\\"
		$ID = (Invoke-CimMethod -MethodName Create -ClassName Win32_ShadowCopy -Arguments @{Volume = $Volume}).ShadowID
		$this.ShadowCopy = Get-CimInstance Win32_shadowcopy | ? ID -eq $ID

		# Create junction to shadow copy.
		$DevicePath = $this.ShadowCopy.DeviceObject
		$this.Junction = $Drive + "\" + [Guid]::NewGuid().ToString()
		cmd /c mklink /d $this.Junction $DevicePath
		
		# Set the path within the junction.
		$this.Path = $this.Junction + $Path.Substring($Drive.Length)
	}
	
	# Delete: delete shadow copy.
	Delete() {
		If ($this.Junction) {
			If (Test-Path $this.Junction) {(Get-Item $this.Junction).Delete()}
			$this.Junction = $Null
		}
		If ($this.ShadowCopy) {
			Remove-CimInstance $this.ShadowCopy
			$this.ShadowCopy = $Null
		}
	}
	
}	

# Main Sapphire helper class.
Class Sapphire {
	
	# Constants.
	Static [Decimal]$Version = 1.65
	
	# FromDesktop: Check if run from desktop. A rather sketchy solution. On Terminal, WT_SESSION is defined. On PS W10, the invocation is different from in session or elsewhere like File Explorer.
	Static [Boolean] FromDesktop() {
		If ($Script:FromDesktop) {Return $True}		# Desktop switch defined
		If ($Script:IsVersion7) {Return $False}		# From PowerShell v7 (never desktop)
		If ($Env:WT_SESSION) {Return $False}		# From Windows Terminal
		# Have to use registry hacks for Windows Server and Windows 10
		$RegVal = [Sapphire]::GetRegistry("HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\0\Command", "", "")						# Windows Server
		If ($RegVal -eq "") {$RegVal = [Sapphire]::GetRegistry("HKLM:\SOFTWARE\Classes\SystemFileAssociations\.ps1\Shell\0\Command", "", "")}	# Windows 10
		If ($RegVal -eq "") {Return $True}																										# unknown, assume desktop
		$RegCmd = $RegVal.Split(' ',3)[2].Replace('%1', $($Script:MyInvocation.MyCommand.Definition))
		$RunCmd = "`"$($Script:MyInvocation.Line)`""
		Return $RunCmd -eq $RegCmd
	}
	
	# AnyKey: press any key if run from desktop.
	Static AnyKey($Message, $Always, $Colour) {
		If ([Sapphire]::FromDesktop() -Or $Always) {
			If ($Message -eq "") {$Message = "`nPress any key to continue..."}
			Write-Host -NoNewLine -Fore $Colour $Message
			$Answer = [Console]::ReadKey()
			Write-Host
		}
	}
	Static AnyKey($Message, $Always) {[Sapphire]::AnyKey($Message, $Always, $Script:Host.UI.RawUI.ForegroundColor)}
	Static AnyKey($Message) {[Sapphire]::AnyKey($Message, $False)}
	Static AnyKey() {[Sapphire]::AnyKey("")}
	
	# ReadKey: Read key with prompt.
	Static [String] ReadKey($Prompt, $Colour) {
		[Console]::TreatControlCAsInput = $True
		Write-Host -Fore $Colour -NoNewLine $Prompt
		$Key = [Console]::ReadKey()
		[Console]::TreatControlCAsInput = $False
		If ($Key.Key -eq [ConsoleKey]::C -and ($Key.Modifiers -band [ConsoleModifiers]::Control)) {Exit}
		Write-Host
		Return $Key.Key
	}
	Static [String] ReadKey($Prompt) {Return [Sapphire]::ReadKey($Prompt, $Global:Host.UI.RawUI.ForegroundColor)}
	
	# ConfirmAction: confirm choice.
	Static ConfirmAction($Title, $Message) {
		$Options = @(
			[System.Management.Automation.Host.ChoiceDescription]::new("&Yes", "Yes, perform the action")
			[System.Management.Automation.Host.ChoiceDescription]::new("&No", "No, cancel the action")
		)
		If (!$Script:Confirm) {
			$Decision = $Script:Host.UI.PromptForChoice($Title, $Message, $Options, 1)
			If ($Decision -eq 1) {Exit}
		}
	}

	# IsAdministrator: check if running as administrator
	Static [Boolean] IsAdministrator($ShowError) {
		$User = [Security.Principal.WindowsIdentity]::GetCurrent()
		$Result = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
		If ($ShowError -And !$Result) {Write-Warning "Please run as administrator"}
		Return $Result
	}
	Static [Boolean] IsAdministrator() {Return [Sapphire]::IsAdministrator($True)}
	
	# RunAdministrator: run script as administrator. Returns true if escalated.
	Static [Bool] RunAdministrator($Wait, $Parameters, $FromDesktop) {
		
		# Return false if already administrator.
		If ([Sapphire]::IsAdministrator($False)) {Return $False}

		# Convert mapped Z:\xxx type paths to the expanded version.
		$File = $Script:MyInvocation.MyCommand.Definition
		$Pattern = "^([a-z]):(\\.*)$"
		If ($File -Match $Pattern) {
			$Letter = $File -Replace $Pattern, '$1'
			$Drive = Get-PSDrive -Name $Letter -PSProvider FileSystem -EA Silent
			If ($Drive -And ($Drive.Root -NotMatch $Pattern)) {$File = $Drive.DisplayRoot + ($File -Replace $Pattern, '$2')}
		}

		# Add -FromDesktop if requested. This is needed on W10 because the FromDesktop hack in here doesn't work when escalated to admin.
		If ($FromDesktop -And [Sapphire]::FromDesktop()) {$Parameters += " -FromDesktop"}
		
		# Create another PowerShell process as admin (runas).
		$PSI = New-Object System.Diagnostics.ProcessStartInfo
		$PSI.FileName = "powershell"
		$PSI.Arguments = "-File ""$File"" $Parameters"
		$PSI.Verb = "runas"
		$Process = $Null
		Try {$Process = [System.Diagnostics.Process]::Start($PSI)} Catch {Write-Warning $_; Return $True}
		If ($Wait) {$Process.WaitForExit()}
		
		# Has run as admin, parent child should exit.
		Return $True
	}
	Static [Bool] RunAdministrator($Wait, $Parameters) {Return [Sapphire]::RunAdministrator($Wait, $Parameters, $False)}
	Static [Bool] RunAdministrator($Wait) {Return [Sapphire]::RunAdministrator($Wait, "")}
	Static [Bool] RunAdministrator() {Return [Sapphire]::RunAdministrator($False)}
	
	# CreateFolder: Create a folder with error message. Returns true if folder created.
	Static [Bool] CreateFolder($Path) {
		If (Test-Path -LiteralPath $Path) {Return $False}
		[void](New-Item $Path -ItemType Directory)
		If (!(Test-Path -LiteralPath $Path)) {Throw "Unable to create ""$Path"" folder"}
		Return $True
	}
	
	# CopyNewer: Copy file if newer.
	Static [Bool] CopyNewer($SourcePath, $TargetFolder) {
		If (!(Test-Path -LiteralPath $SourcePath)) {Throw "Can't find ""$SourcePath"""}
		Try {[Sapphire]::CreateFolder($TargetFolder)} Catch {Throw $_}
		$TargetPath = $TargetFolder + "\" + $(Split-Path $SourcePath -Leaf)
		$SourceTimestamp = $Null
		$TargetTimestamp = $Null
		If (Test-Path -LiteralPath $SourcePath) {$SourceTimestamp = (Get-ItemProperty -LiteralPath $SourcePath -Name LastWriteTime).LastWriteTime}
		If (Test-Path -LiteralPath $TargetPath) {$TargetTimestamp = (Get-ItemProperty -LiteralPath $TargetPath -Name LastWriteTime).LastWriteTime}
		If (($SourceTimestamp -ne $Null) -and ($TargetTimestamp -ne $Null) -And ($SourceTimestamp -le $TargetTimestamp)) {Return $False}
		Copy-Item -LiteralPath $SourcePath $TargetFolder
		Return $True
	}
	
	# GetRegistry: get a value from the registry.
	Static [Object] GetRegistry($Path, $Name, $DefaultValue) {
		$Value = $DefaultValue
		$Key = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
		If ($Key) {$Value = $Key.GetValue($Name, $DefaultValue)}
		Return $Value
	}

	# GetRegistryDWord: get a value from the registry.
	Static [Int] GetRegistryDWord($Path, $Name, $Default) {
		$Value = $Null
		$Key = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
		If ($Key) {$Value = $Key.GetValue($name, $null)}
		If ($Value -eq $Null) {$Value = $Default}
		Return $Value
	}

	# # ChangeRegistrySZ: change a value in the registry.
	# Static [Boolean] ChangeRegistrySZ($Path, $Name, $Value) {
		# $OldValue = [Sapphire]::GetRegistry($Path, $Name, $Null)
		# [Sapphire]::SetRegistry($Path, $Name, $Value)
		# Return $Value -ne $OldValue
	# }
	
	# SetRegistry: set a value in the registry
	Static SetRegistry($Path, $Name, $Value) {
		If (!(Test-Path $Path)) {New-Item -Path $Path -Force}
		New-ItemProperty -Path $Path -Name $Name -Value $Value -Force
	}
	
	# ChangeRegistry: Change a registry value.
	Static [Boolean] ChangeRegistry($Path, $Name, $Value) {
		$OldValue = [Sapphire]::GetRegistry($Path, $Name, $Null)
		If ($Value -ne $OldValue) {[Sapphire]::SetRegistry($Path, $Name, $Value)}
		Return $Value -ne $OldValue
	}

	# ChangeRegistryList: Changes a list of registry values. Pass list of hash table items with Path, Name and Value properties.
	Static [Boolean] ChangeRegistryList($List) {
		$Changed = $False
		ForEach ($Item In $List) {If ([Sapphire]::ChangeRegistry($Item.Path, $Item.Name, $Item.Value)) {$Changed = $True}}
		Return $Changed
	}
	
	# Left: get characters from left of string.
	Static [String] Left($Text, $Length) {Return $Text.Substring(0, [Math]::Min($Length,$Text.length))}
	Static [String] Left($Text) {Return [Sapphire]::Left($Text, 1)}
	
	# Right: get characters from right of string.
	Static [String] Right($Text, $Length) {
		$Start = [Math]::Min($Text.Length - $Length, $Text.Length)
		$Start = [Math]::Max(0, $Start)
		Return $Text.SubString($Start ,[Math]::Min($Text.length, $Length))
	}
	Static [String] Right($Text) {Return [Sapphire]::Right($Text, 1)}

	# TrimTrailing: Trim a trailing character.
	Static [String] TrimTrailing($Text, $Char) {
		$Text = $Text.Trim()
		If ([Sapphire]::Right($Text, 1) -eq $Char) {$Text = [Sapphire]::Left($Text, $Text.Length - 1)}
		Return $Text
	}
	
	# Append: append text to string with separator.
	Static [String] Append($Text, $Seperator, $Append) {Return [Sapphire]::IIf($Text -eq "", "", $Text + $Seperator) + $Append}
	
	# IsNumeric: test for numeric.
	Static [Boolean] IsNumeric($Text) {Return $Text -Match "^[\d\.]+$"}
	
	# IsYes: test for starting with "Y".
	Static [Boolean] IsYes($Text) {Return $Text -Match "^.*y.*$"}
	
	# IsNo: test for starting with "N".
	Static [Boolean] IsNo ($Text) {Return $Text -Match "^.*n.*$"}
	
	# RemoveTrailingSlash: remove trailing slash character from paths.
	Static [String]RemoveTrailingSlash($Path) {
		While ([Sapphire]::Right($Path, 1) -eq "\") {$Path = [Sapphire]::Left($Path, $Path.Length - 1)}
		Return $Path
	}
		
	# Timestamp: get a now timestamp in format year-month-day-hour-min
	Static [String] Timestamp($IncludeSeconds) {
		$Now = Get-Date
		$Format = "{0:0000}-{1:00}-{2:00} {3:00}-{4:00}"
		If ($IncludeSeconds) {$Format += "-{5:00}"}
		Return $Format -f $Now.Year, $Now.Month, $Now.Day, $Now.Hour, $Now.Minute, $Now.Second
	}
	Static [String] Timestamp() {Return [Sapphire]::Timestamp($False)}
	
	# FriendlySize: get storage space in friendly amount.
	Static [String] FriendlySize($Bytes) {
		$Format = "{0:0}{1}"
		ForEach ($Suffix In "B", "KB", "MB", "GB", "TB", "PB") {
			If ($Bytes -lt 1KB) {Return $Format -f $Bytes, $Suffix}
			$Bytes /= 1KB
			$Format = "{0:N1}{1}"
		}
		Return "Error"
	}
	
	# FriendlyTime: get millisecond time in friendly amount.
	Static [String] FriendlyTime($Value) {
		$Intervals = @(
			@{Divider = 1000;	Limit = 60; 				Suffix = "second"}
			@{Divider = 60;		Limit = 60; 				Suffix = "minute"}
			@{Divider = 60; 	Limit = 24; 				Suffix = "hour"}
			@{Divider = 24; 	Limit = [int]::MaxValue; 	Suffix = "day"}
		)
		ForEach ($Interval In $Intervals) {
			$Value /= $Interval.Divider
			If ($Value -lt $Interval.Limit) {
				$Value = [Decimal]("{0:N1}" -f $Value)
				Return [Sapphire]::Plural("{0:N1} $($Interval.Suffix){1}", $Value)
			}
		}
		Return $Null # should never occur
	}
	
	# UploadFtp Upload file to FTP
	Static UploadFtp($Site, $Username, $Password, $LocalPath, $RemoteFolder, $RemoteName) {
	
		# Create the folder tree.
		If ($RemoteFolder -ne "") {
			$Folders = $RemoteFolder.Split("/")
			$Path = ""
			ForEach ($Folder In $Folders) {
				If ($Path -ne "") {$Path += "/"}
				$Path += $Folder
				$Uri = "ftp://" + $Site + "/" + $Path
				$FTPRequest = [System.Net.FtpWebRequest]::Create($Uri)
				$FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::MakeDirectory
				$FTPRequest.Credentials = new-object System.Net.NetworkCredential($Username, $Password)
				$FTPRequest.UseBinary = $True
				Try {$Response = $FTPRequest.GetResponse()} Catch {}		# ignore errors on creating (already exists)
			}
		}

		# Build the full remote path.
		$Uri = "ftp://" + $Site + "/" 
		If ($RemoteFolder -ne "") {$Uri += $RemoteFolder + "/"}
		$Uri += $RemoteName
		# Create FTP request object
		$FTPRequest = [System.Net.FtpWebRequest]::Create($Uri)
		$FTPRequest = [System.Net.FtpWebRequest]$FTPRequest
		$FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
		$FTPRequest.Credentials = new-object System.Net.NetworkCredential($Username, $Password)
		$FTPRequest.UseBinary = $true
		$FTPRequest.UsePassive = $true
		# Read the file for upload
		$FileContent = [System.IO.File]::ReadAllBytes($LocalPath)
		$FTPRequest.ContentLength = $FileContent.Length
		# Get Stream Request by bytes
		$Run = $FTPRequest.GetRequestStream()
		$Run.Write($FileContent, 0, $FileContent.Length)
		# Cleanup
		$Run.Close()
		$Run.Dispose()
	}
	
	# StartProgress: start progress bar. In fast loops, the call to ShowProgress (Write-Progress) can take more time than the actual code so by default, it only updates every 500ms. 
	#				 $Delay can be used to change this. 0 = update every time ShowProgress is called
	Static StartProgress($Activity, $Level, $Delay) {
		$Script:Progress[$Level] = [SapphireProgress]::New($Activity, $Delay)
		$ProgressPreference = "continue" # fixes a bug where progress bars stop working
		Write-Progress -Activity $Activity -Id $Level
	}
	Static StartProgress($Activity) {[Sapphire]::StartProgress($Activity, 1, 500)}

	# ShowProgress: display progress bar. StartProgress must be called first.
	Static ShowProgress($PercentComplete, $Status, $Level, $Elapsed) {
		# Only Write-Progress every 500ms or first time called or if Always flag set. Otherwise many calls actually slows the program down.
		$Progress = $Script:Progress[$Level]
		If (!$Progress) {Return} # forget to call StartProgress
		If (($Progress.Delay -ne 0) -And (!$Progress.First)) {
			If ($Progress.Timer.Elapsed.TotalMilliseconds -le $Progress.Delay) {Return}
			$Progress.Timer.Reset(); $Progress.Timer.Start()
		}
		$Progress.First = $False
		# Display activity without progress bar
		$Activity = $Progress.Activity
		If ($PercentComplete -eq 0) {
			Write-Progress -Activity $Activity -Status $Status -Id $Level -Parent ($Level - 1)
		} Else {
			# Calculate time to finish.
			If (!$Elapsed) {$Elapsed = New-TimeSpan -Start $Progress.StartTime -End (Get-Date)}
			$Finish = ($Elapsed.TotalSeconds * 100) / $PercentComplete
			$SecondsRemaining = $Finish - $Elapsed.TotalSeconds
			If ($SecondsRemaining -gt [int32]::MaxValue) {$SecondsRemaining = 0} # Write-Progress can't handle massive seconds remaining!
			If ($SecondsRemaining -lt 1) {
				Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -Id $Level -ParentId ($Level - 1)
			} Else {
				Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -SecondsRemaining $SecondsRemaining -Id $Level -ParentId ($Level - 1)
			}
		}
	}
	Static ShowProgress($PercentComplete, $Status) {[Sapphire]::ShowProgress($PercentComplete, $Status, 1, $Null)}
	Static ShowProgress($PercentComplete) {[Sapphire]::ShowProgress($PercentComplete, "Processing", 1, $Null)}
	
	# ClearProgress: Clear a progress bar.
	Static ClearProgress($Level) {
		$Activity = $Script:Progress[$Level]
		If ($Activity) {
			Write-Progress -Activity $Activity -Completed -Id $Level
			If ($Script:IsVersion7) {Start-Sleep 1}
		}
	}
	Static ClearProgress() {[Sapphire]::ClearProgress(1)}
	
	# WaitForFile: Wait until a file appears.
	Static [Boolean] WaitForFile($Path, $Timeout) {
		$StartTime = Get-Date
		While ($True) {
			Start-Sleep 1
			If (Test-Path $Path) {Break}
			$Elapsed = New-TimeSpan -Start $StartTime -End (Get-Date)
			If ($Elapsed.TotalSeconds -ge $Timeout) {Return $False}
		}
		Return $True
	}
	
	# WaitForFileAndDelete: wait until a file exists and then delete.
	Static WaitForFileAndDelete($Path, $Timeout) {If ([Sapphire]::WaitForFile($Path, $Timeout)) {Remove-Item $Path}}
	
	# StopJobs: delete old jobs left due to break from script. Matches some watermark text in the script block.
	Static StopJobs($Watermark) {Get-Job | ? {($_.Command.Contains($Watermark)) -And ($_.State -eq "Running")} | Stop-Job}

	# GetFile: Get first byte of a file - forces cloud provider to download the file.
	Static GetFile($Path) {
		If (Test-Path $Path) {
			Try {
				$File = [System.IO.File]::OpenRead($Path)
				$File.ReadByte() | Out-Null
				$File.Close()
			} Catch {
				Throw "Unable to read ""$Path"""
			}
		}
	}

	# GetFiles: get cloud files, i.e. make local copy. Doesn't work with Dropbox.
	Static GetFiles($Path, $Wait) {
		
		# Variables.
		$Offline 	= 0x001000	# file is available locally
		$Pinned 	= 0x080000	# file is always available locally
		$Unpinned 	= 0x100000	# file is not always available locally (download again on demand?)

		# Check folder exists.
		If (!(Test-Path $Path)) {Throw "Unable to find ""$Path"""}

		# Get list of offline files.
		[Sapphire]::StartProgress("Getting list of offline files in ""$Path""")
		[System.Collections.ArrayList]$Files = @(Get-ChildItem -LiteralPath $Path -File -Recurse -Attributes Offline)
		$TotalLength = ($Files | Measure-Object -Property Length -Sum).Sum
		[Sapphire]::ClearProgress()

		# Change attributes to pin/download.
		[Sapphire]::StartProgress([Sapphire]::Plural("Flagging {0:N0} file{1} to download from ""$(Split-Path $Path -Leaf)""", $Files.Count))
		$Index = 0
		ForEach ($OfflineFile In $Files) {
			[Sapphire]::ShowProgress($Index / $Files.Count * 100)
			$OfflineFile.Attributes = ($OfflineFile.Attributes -band (-bnot $Unpinned)) -bor $Pinned
			$Index += 1
		}
		[Sapphire]::ClearProgress()

		# Wait until files downloaded. Progress by file size.
		If ($Wait) {
			[Sapphire]::StartProgress([Sapphire]::Plural("Waiting for {0:N0} file{1}", $Files.Count) + "/" + [Sapphire]::FriendlySize($TotalLength) + " to download from ""$(Split-Path $Path -Leaf)""")
			$NumFiles = $Files.Count
			$Downloaded = 0
			While ($Files.Count) {
				[Sapphire]::ShowProgress($Downloaded / $TotalLength * 100, "Downloaded: " + [Sapphire]::FriendlySize($Downloaded))
				$RemoveList = @(); $Sleep = 1
				ForEach ($File In $Files) {
					Try {
						$Attributes = (Get-Item -LiteralPath $File.FullName -EA Stop).Attributes
						If (($Attributes -band $Offline) -eq 0) {
							$Downloaded += $File.Length
							$RemoveList += $File
							$Sleep = 0
						}
					} Catch {}
				}
				ForEach ($File In $RemoveList) {$Files.Remove($File)}
				Start-Sleep $Sleep
			}
			[Sapphire]::ClearProgress()
		}
	}
	Static GetFiles($Path) {[Sapphire]::GetFiles($Path, $True)}

	# GetDropboxFiles: Get first byte of all files in folder - forces cloud system to get files. Uses PowerShell jobs to force multiple downloads at once.
	#   Path: path to the folder
	#   Retry: how long to pause between retries for cloud storage. 0=don't loop, report error
	#   NumberOfJobs: number of jobs to run in parallel
	#   FilesPerJob: number of files to open (causes download) per job
	#   Returns Null or an error message string

	Static [String] GetDropboxFiles($Path, $Retry, $NumberOfJobs, $FilesPerJob) {

		# Test path exists.
		If (!(Test-Path $Path)) {Return "Path not found"}

		# Get list of files and calculate total length.
		[Sapphire]::StartProgress("Getting list of files from ""$Path""")
		$Files = Get-ChildItem -LiteralPath $Path -File -Recurse:$True -Attributes o
		$TotalBytes = ($Files | Measure-Object -Property Length -Sum).Sum

		# Script block to get files. Returns hash with two fields: RetryCount and Exception.
		$ScriptBlock = {
			# Sapphire:GetFiles
			Param($Files, $Retry)
			$CloudErrors = "The cloud operation is invalid", "The cloud operation was not completed before the time-out period expired"
			$TimeoutAfter = 10 # minutes
			$Result = @{}; $Result.RetryCount = 0
			ForEach ($File In $Files) {
				$StartTime = Get-Date
				While ($True) {
					Try {
						$Handle = [System.IO.File]::OpenRead($File.FullName)		# this doesn't trigger cloud download
						$Handle.ReadByte() | Out-Null								# but this does!
						$Handle.Close()
						Break
					} Catch {
						$ErrorMessage = $_.Exception.Message
						$CloudError = $False; $CloudErrors | % {If ($ErrorMessage.Contains($_)) {$CloudError = $True; Break}}
						$Timeout = ((New-Timespan $StartTime (Get-Date)).TotalMinutes) -ge $TimeoutAfter
						If ($Timeout) {$ErrorMessage = "TIMEOUT: $ErrorMessage"}
						If ((!$CloudError) -Or ($Retry -eq 0) -Or $Timeout) {$Result.Exception = $ErrorMessage; Return $Result}
						Start-Sleep $Retry
						$Result.RetryCount += 1
					} Finally {
						If ($Handle) {$Handle.Close()}
					}
				}
			}
			Return $Result
		}
		
		# Stop old Sapphire jobs. Must match the heading in the script block above.
		[Sapphire]::StopJobs("# Sapphire:GetFiles")
		
		# The main job loop.
		[Sapphire]::StartProgress(("Getting {0:N0} files/{1} from ""{2}""" -f $Files.Count, [Sapphire]::FriendlySize($TotalBytes), $Path))
		$Index = 0; $RetryCount = 0; $BytesCount = 0; $Processed = 0
		[System.Collections.ArrayList]$Jobs = @()
		Do {
			
			# If getting close to end, balance out files per job.
			$NumLeftPerJob = [int](($Files.Count - $Index) / $NumberOfJobs)
			$MaxFiles = $FilesPerJob
			If ($MaxFiles -gt $NumLeftPerJob) {$MaxFiles = $NumLeftPerJob}
			If ($MaxFiles -eq 0) {$MaxFiles = 1}
			
			# Add jobs until limit reached or out of files to get.
			While (($Jobs.Count -lt $NumberOfJobs) -And ($Index -lt $Files.Count)) {
				$Job = @{}; $Job.Items = @(); $Job.Length = 0
				While (($Job.Items.Count -lt $MaxFiles) -And ($Index -lt $Files.Count)) {
					$File = $Files[$Index]
					$Job.Items += $File; $Job.Length += $File.Length
					$Index += 1
				}
				If ($Job.Items.Count -ne 0) {
					$Job.PsJob = Start-Job $ScriptBlock -ArgumentList $Job.Items, $Retry
					$Jobs += $Job
				}
			}
			
			# Update progress.
			$Complete = ($Processed / $Files.Count) * 100
			$Message = "Processed: {0:N0}/{1}, Retries: {2:N0}, Retry: {3} seconds, Jobs: {4}, Files per job: {5}"
			[Sapphire]::ShowProgress($Complete, ($Message -f $Processed, [Sapphire]::FriendlySize($BytesCount), $RetryCount, $Retry, $Jobs.Count, $FilesPerJob))
			
			# Remove jobs if they've finished. If one job throws an exception, stop all the other jobs. Jobs to be removed have to be done outside the loop. Can't remove object during iteration!
			$RemoveJobs = @()
			ForEach ($Job In $Jobs) {
				If ($Job.PsJob.State -eq "Completed") {
					$Result = Receive-Job $Job.PsJob
					If ($Result.Exception) {
						$Jobs | ? {$_.PsJob.State -eq "Running"} | % {$_.PsJob | Stop-Job}
						[Sapphire]::ClearProgress()
						Return $Result.Exception
					}
					$Processed += $Job.Items.Count		# count number of files processed
					$BytesCount += $Job.Length			# count number of file lengths processed
					$RetryCount += $Result.RetryCount	# count retries 
					$RemoveJobs += $Job
				}
			}
			$RemoveJobs | % {$Jobs.Remove($_)}
			
			# Keep looping until processed all bytes.
			
		} Until ($BytesCount -eq $TotalBytes)
		
		# All done, no result.
		[Sapphire]::ClearProgress()
		Return $Null
	}
	Static [String] GetDropboxFiles($Path) {Return [Sapphire]::GetDropbox($Path, 0, 4, 50)}

	# CheckM365Tenant: Check M365 tenant is defined. MA = modern authentication. Returns $True if okay.
	Static [Boolean] CheckM365Tenant() {

		# Check the variables have been defined.
		$VariableNames = @("M365Tenant", "M365Username", "M365PnPClientId")
		$List = ""
		ForEach ($VariableName in $VariableNames) {
			$Variable = Get-Variable $VariableName -EA SilentlyContinue
			If (!$Variable) {If ($List -ne "") {$List += ", "}; $List += $VariableName}
		}
		If ($List -ne "") {
			Write-Warning "M365 tenant variables(s) missing: $List. Run Select-[Tenant] script"
		} Else {
			# Set the Url of the admin portal.
			$Pattern = "(https:\/\/.*)(\.sharepoint\.com)"
			$Global:M365AdminUrl = "$($Global:M365Tenant -replace $Pattern, '$1')-admin$($Global:M365Tenant -replace $Pattern, '$2')"
		}
		
		# Return true if all okay.
		Return $List -eq ""
	}
	
	# GetM365Password: Gets the M365 password (being phased out).
	Static [Boolean] GetM365Password() {
		If (!$Global:M365Password) {
			$Global:M365Password = Read-Host "Microsoft 365 password for $($Global:M365Username)" -AsSecureString
			If ($Global:M365Password.Length -eq 0) {
				$Global:M365Password = $Null
				Return $False
			}
		}
		Return $True
	}
	
	# GetExeType: get exe type, e.g. 32-bit
	Static [String] GetExeType([String]$Path) {
		$exeType = 'Unknown'
		$Stream = $Null
		Try {
			$Stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
			$Bytes = New-Object Byte[](4)
			If ($Stream.Length -ge 64 -and $Stream.Read($bytes, 0, 2) -eq 2 -and $bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A) {
				$exeType = '16-bit'
				If ($Stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) -eq 0x3C -and $Stream.Read($bytes, 0, 4) -eq 4) {
					If (-not [System.BitConverter]::IsLittleEndian) {[Array]::Reverse($bytes, 0, 4)}
					$peHeaderOffset = [System.BitConverter]::ToUInt32($bytes, 0)
					If ($Stream.Length -ge $peHeaderOffset + 6 -and 
							$Stream.Seek($peHeaderOffset, [System.IO.SeekOrigin]::Begin) -eq $peHeaderOffset -and 
							$Stream.Read($bytes, 0, 4) -eq 4 -and 
							$bytes[0] -eq 0x50 -and $bytes[1] -eq 0x45 -and 
							$bytes[2] -eq 0 -and $bytes[3] -eq 0) {
							$exeType = 'Unknown'
						if ($Stream.Read($bytes, 0, 2) -eq 2) {
							if (-not [System.BitConverter]::IsLittleEndian) {[Array]::Reverse($bytes, 0, 2)}
							$MachineType = [System.BitConverter]::ToUInt16($bytes, 0)
							Switch ($MachineType) {
								0x014C { $exeType = '32-bit' }
								0x0200 { $exeType = '64-bit' }
								0x8664 { $exeType = '64-bit' }
							}
						}
					}
				}
			}
		} 
		Catch {}
		Finally {If ($Stream -ne $Null) {$Stream.Dispose()}}
		Return $exeType
	}
	
	# GetEnvironment: Gets environment variables. Returns true if successful.
	Static GetEnvironment([String[]]$Variables) {
		$Message = ""
		ForEach ($Variable In $Variables) {
			# If comma variant, then PowerShell variable name is 2nd param, e.g. PEAKOneDrive,Path"
			$Split = $Variable.Split(",")
			If ($Split.Count -eq 2) {
				$EnvName = $Split[0].Trim()
				$VarName = $Split[1].Trim()
			} Else {
				$EnvName = $Variable
				$VarName = $Variable
			}
			$Item = Get-Item -Path "Env:$EnvName"
			If (!$Item) {
				If ($Message -ne "") {$Message += "; "}
				$Message += $EnvName
			} Else {
				Set-Variable $VarName -Value $Item.Value -Scope Script
			}
		}
		If ($Message -ne "") {Throw "Missing environment variable(s): $Message"}
	}
	
	# GetRelativeUrl: gets the relative part of a Url.
	Static [String] GetRelativeUrl([String]$Url) {Return $Url -Replace "^https:\/\/.*?(\/.*$)$",'$1'}
	
	# CheckCSVHeadings: check CSV headings present.Returns true if all found.
	Static CheckCSVHeadings($CSV, $Headings) {
		$Missing = ""
		$CSV_Headings = ($CSV | Get-Member -MemberType NoteProperty).Name
		ForEach ($Heading In $Headings) {
			If (!($CSV_Headings.Contains($Heading))) {
				If ($Missing -ne "") {$Missing += ", "}
				$Missing += $Heading
			}
		}
		If ($Missing -ne "") {Throw "Missing headings in CSV: $Missing"}
	}
	
	# CreateTemporaryFile: create a temporary file.
	Static [String] CreateTemporaryFile($Prefix) {
		$TempFile = New-TemporaryFile
		$TempFileName = $TempFile.DirectoryName + "\$Prefix-" + $TempFile.Name
		Rename-Item $TempFile.FullName $TempFileName
		Return $TempFileName
	}
	
	# SendEmail: send email via SMTP. Returns "" if sent okay.
	Static SendEmail($SMTP, $Subject, $ToAddress, $FromAddress, $Body) {
		$Message = New-Object Net.Mail.MailMessage
		$Message.Subject = $Subject
		$Message.From = New-Object Net.Mail.MailAddress -ArgumentList $FromAddress
		Try {$Message.To.Add($ToAddress)} Catch {Throw "Bad email address ""$ToAddress"""}
		$Message.IsBodyHtml = $True
		$Message.Body = $Body
		$SMTPClient = New-Object Net.Mail.SmtpClient($SMTP.Server, $SMTP.Port)
		$SMTPClient.EnableSsl = $True
		$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTP.Username, $SMTP.Password)
		Try {$SMTPClient.Send($Message)} Catch {Throw "Problem sending email to $ToAddress`n$Error"}
	}
	
	# ScrapeArray: Scrape array for values using regexp patterns. Used to parse non-PowerShell command output. Scraping starts when item matches $MarkerPattern and stops when it next matches. 
	#			   Items that match $Parameters (an array) are added to the returned hash table. Example of $Parameters entry: @{Name = "Radio Type"; 		Pattern = "^\s+Radio type\s+:\s+(.*)$"}

	Static [System.Collections.Specialized.OrderedDictionary] ScrapeArray($Array, $MarkerPattern, $MarkerValue, $Parameters) {

		# Scrape the array line by line.
		$Output = @{}; $Scraping = $False
		ForEach ($Line In $Array) {
			If ($Line -Match $MarkerPattern) {
				If ($Scraping) {Break} # found next marker so stop scraping.
				If (($Line -Replace $MarkerPattern, '$1') -eq $MarkerValue) {$Scraping = $True}
			} ElseIf ($Scraping) {
				ForEach ($Parameter In $Parameters) {
					$Key = $Parameter.Name
					If ($Line -Match $Parameter.Pattern) {
						$Value = $Line -Replace $Parameter.Pattern, '$1'
						# If key already exists, the value is turned into an array.
						If ($Output.Contains($Key)) {
							If ($Output[$Key] -Is [Array]) {$Output[$Key] = $Output[$Key] + $Value} Else {$Output[$Key] = @($Output[$Key], $Value)}
						} Else {
							$Output[$Key] = $Value
						}
					}
				}
			}
		}

		# Return results in order parameters specified, not order found in array.
		$Results = [Ordered]@{}
		ForEach ($Parameter In $Parameters) {
			If ($Output.Contains($Parameter.Name)) {
				$Results[$Parameter.Name] = $Output[$Parameter.Name]
			}
		}
		Return $Results

	}

	# MergeHashTable: merge two hash tables.
	Static MergeHashTable($Hash1, $Hash2) {ForEach ($Key In $Hash2.Keys) {$Hash1.$Key = $Hash2.$Key}}
	
	# Install: Unzip and run setup.exe.
	Static Install($ZipFile, $Setup) {

		# Check files exist.
		If (!(Test-Path $ZipFile)) {Throw "Unable to find ""$ZipFile"""}
		$Path = [System.IO.Path]::GetDirectoryName($ZipFile)
		$Exe = "$Path\7za.exe"
		If (!(Test-Path $Exe)) {Throw "Unable to find ""$Exe"""}
		
		# Unzip into temp folder.
		$ZipName = [System.IO.Path]::GetFileName($ZipFile)
		[Sapphire]::StartProgress("Unzipping $ZipName")
		$TempFolder = $Env:Temp
		$Cmd = "& ""$Exe"" x -aoa ""$ZipFile"" -o$TempFolder"
		Invoke-Expression $Cmd | Out-Null
		[Sapphire]::ClearProgress()

		# Start the installer.
		$TempPath = $TempFolder + "\" + ([System.IO.Path]::GetFileNameWithoutExtension($ZipFile))
		$Installer = "$TempPath\$Setup"
		Try {
			Start-Process "$TempPath\$Setup" -Wait
		} Catch {
			Throw "Unable to run ""$Installer"""
		} Finally {
			If (Test-Path $TempPath) {Remove-Item $TempPath -Recurse}
		}
	}
	
	# Plural: return plural text. {0} is the count, {1} is the plural suffix ("s" or "") and P1-P3 are optional {2}-{4} formatting parameters.
	Static [String] Plural($Text, [int]$Count, $P1, $P2, $P3) {
		$Suffix = ""; If ($Count -ne 1) {$Suffix = "s"}
		Return $Text -f $Count, $Suffix, $P1, $P2, $P3
	}
	Static [String] Plural($Text, $Count) {Return [Sapphire]::Plural($Text, $Count, $Null, $Null, $Null)}
	Static [String] Plural($Text, $Count, $P1) {Return [Sapphire]::Plural($Text, $Count, $P1, $Null, $Null)}
	Static [String] Plural($Text, $Count, $P1, $P2) {Return [Sapphire]::Plural($Text, $Count, $P1, $P2, $Null)}
	
	# RemoveEmptyFolders: removes empty folders from a path.
	Static RemoveEmptyFolders($Path) {
		
		# Keep looping whilst there is an error.
		$RemoveError = $False
		Do {
			
			# Get hash of folders at the start.
			$Folders = @{};	$Removed = @()
			ForEach ($Folder In Get-ChildItem -LiteralPath $Path -Recurse -Directory -Force) {$Folders[$Folder.Fullname] = $Folder}

			# Start the repeated move. Loop around until none deleted. This is because removing an empty folder in a child folder might then mean the parent folder is empty.
			$AllEmptyFolders = @{}
			$RemoveError = $False
			Do {

				# Build list of empty folders.
				$EmptyFolders = @()
				$Index = 0
				ForEach ($Folder In $Folders.Values) {
					$Count = (Get-ChildItem -LiteralPath $Folder.FullName -Force).Count
					If ($Count -eq 0) {
						$EmptyFolders += $Folder
						If (!($AllEmptyFolders.Contains($Folder.Fullname))) {$AllEmptyFolders[$Folder.Fullname] = $Folder}
					}
					$Index += 1
				}

				# Remove empty folders in this pass.
				$Removed = @()
				ForEach ($Folder In $EmptyFolders) {
					Try {
						$Folder | Remove-Item -Force -EA Stop | Out-Null
						$Removed += $Folder.Fullname
					} Catch {
						$RemoveError = $True
					}
				}
				ForEach ($Folder In $Removed) {$Folders.Remove($Folder)}
				
				# Loop again if something was removed and not preview.
				
			} Until (($Removed.Count -eq 0) -Or $RemoveError)
			
		} Until (!$RemoveError)
	}
	
	# GetRegExGroups: get array of RegEx group values. Adds dummy [0] entry so same convention as {N} used in pattern. Returns $Null if pattern doesn't match.
	Static [String[]] GetRegExGroups($Text, $Pattern, $NumGroups) { 
		$Result = $Null
		If ($Text -Match $Pattern) {
			$Result = @("Dummy")
			For ($i = 1; $i -le $NumGroups; $i++) {$Result += $Text -Replace $Pattern, "`$$i"}
		}
		Return $Result
	}
	
	# ZeroMilliseconds: zero milliseconds on a date.
	Static [DateTime] ZeroMilliseconds($Date) {Return Get-Date -Year $Date.Year -Month $Date.Month -Day $Date.Day -Hour $Date.Hour -Minute $Date.Minute -Second $Date.Second -Millisecond 0}
	
	# Uncache: uncache global variables after a certain period. Default is 2 hours.
	Static Uncache([String[]]$Names, $Hours, $NoCache) {
		ForEach ($Name In $Names) {
			$ExpireVariable = Get-Variable "$Name`_Expire" -Scope Global -EA SilentlyContinue
			If ($ExpireVariable) {If (($ExpireVariable.Value -ge (Get-Date)) -And !$NoCache) {Return}}
			Set-Variable -Name "$Name`_Expire" -Value (Get-Date).AddHours($Hours) -Scope Global
			Remove-Variable -Name $Name -Scope Global -EA SilentlyContinue
		}
	}
	Static Uncache([String[]]$Names) {[Sapphire]::Uncache($Names, 2, $False)}
	
	# MangleUsername: Mangle a username as a variable name.
	Static [String] MangleUsername($Username) {Return $Username.Replace(".","")}
	
	# GetCredential: get cached credentials.
	Static [System.Management.Automation.PSCredential] GetCredential($Username) {
		$Name = [Sapphire]::MangleUsername($Username)
		[Sapphire]::Uncache($Name)
		$Variable = Get-Variable -Name $Name -Scope Global -EA SilentlyContinue
		If ($Variable) {
			$Cred = $Variable.Value
		} Else {
			Try {
				$Cred = Get-Credential $Username -EA Stop
			} Catch {
				Return $Null
			}
			Set-Variable -Name $Name -Value $Cred -Scope Global
		}
		Return $Cred
	}
	
	# DeleteCredential: delete cached credentials.
	Static DeleteCredential($Username) {
		$Name = [Sapphire]::MangleUsername($Username)
		Remove-Variable -Name $Name -Scope Global -EA Stop
	}
	
	# WriteLog: write entry to console and event log. $Script:LogMessage defines the event log source.
	Static WriteLog($Message, $Colour, $EntryType) {
		Write-Host -Fore $Colour $Message
		If ($Script:LogSource) {
			If (![System.Diagnostics.EventLog]::SourceExists($Script:LogSource)) {New-EventLog -LogName Application -Source $Script:LogSource}
			Write-EventLog -LogName Application -Source $Script:LogSource -EventId 1000 -EntryType $EntryType -Message $Message -Category 0
		}
	}
	Static WriteLog($Message, $Colour) {[Sapphire]::WriteLog($Message, $Colour, "Information")}
	Static WriteLog($Message) {[Sapphire]::WriteLog($Message, $Global:Host.UI.RawUI.ForegroundColor, "Information")}
	
	# LoadJson: Load a JSON file. If not path supplied, uses same path as script with .json extension.
	Static [PSCustomObject] LoadJson($Path) {
		If (!(Test-Path $Path)) {Throw "Unable to find ""$Path"" file"}
		Try {$Config = Get-Content -Raw $Path | ConvertFrom-Json} Catch {Throw "Invalid JSON ""$Path"""}
		Return $Config
	}
	Static [PSCustomObject] LoadJson() {Return [Sapphire]::LoadJson("$PSScriptRoot\" + [IO.Path]::GetFileNameWithoutExtension($Script:MyInvocation.MyCommand) + ".json")}
	
	# UnicodePath: prefixes local drives with Unicode prefix.
	Static [String] UnicodePath($Path) {
		If ($Path -Match "^[a-zA-Z]:\\") {$Path = "\\?\" + $Path}
		Return $Path
	}
	
	# IIf: returns one of two parts.
	Static [String] IIf($Condition, $TruePart, $FalsePart) {If ($Condition) {Return $TruePart} Else {Return $FalsePart}}
	
	# StripQuotes: strip leading and trailing quotes from a string/path.
	Static [String] StripQuotes([String]$Text) {
		$Text = $Text.Trim()
		If ([Sapphire]::Left($Text, 1) -eq """") {$Text = $Text.Substring(1)}
		If ([Sapphire]::Right($Text, 1) -eq """") {$Text = [Sapphire]::Left($Text, $Text.Length - 1)}
		Return $Text
	}
	
	# RefreshDesktop: Refresh desktop icons.
	Static RefreshDesktop() {
		$Code = `
			"[System.Runtime.InteropServices.DllImport(""Shell32.dll"")]`n" + `
			"private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);`n" + `
			"public static void RefreshDesktop() {SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);}"
		Add-Type -MemberDefinition $Code -Namespace WinAPI -Name Explorer
		Invoke-Expression "[WinAPI.Explorer]::RefreshDesktop()"
	}
	
	# AddShortcut: add a shortcut link in a folder. Returns shortcut if added.
	Static [Object] AddShortcut($Folder, $Name, $Description, $Path, $Params, $IconLocation, $Replace) {
		$LinkPath = "$Folder\$Name.lnk"
		If ((Test-Path $LinkPath) -And !$Replace) {Return $Null}
		If (Test-Path $LinkPath) {Remove-Item $LinkPath | Out-Null}
		$WShell = New-Object -ComObject WScript.Shell
		$Shortcut = $WShell.CreateShortcut($LinkPath)
		$Shortcut.TargetPath = $Path
		If ($Params) {$Shortcut.Arguments = $Params}
		$Shortcut.Description = $Description
		If ($IconLocation) {
			$Shortcut.IconLocation = $IconLocation
		} ElseIf ($Path -Match "^http[s]?:.*$") {
			$Shortcut.IconLocation = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe, 0"
		}
		$Shortcut.Save()
		# ---- DEBUG ----
		If ($Params -And ($Shortcut.Arguments -ne $Params)) {Write-Host -Fore Red "Arguments blank for $Name shortcut!"}
		Return $Shortcut
	}
	Static [Object] AddShortcut($Folder, $Name, $Description, $Path) {Return [Sapphire]::AddShortcut($Folder, $Name, $Description, $Path, $Null, $Null, $False)}

	# AddDesktopShortcut: add a shortcut to desktop.
	Static [Object] AddDesktopShortcut($Name, $Description, $Path, $Params, $IconLocation, $Replace) {
		Return [Sapphire]::AddShortcut([Environment]::GetFolderPath('Desktop'), $Name, $Description, $Path, $Params, $IconLocation, $Replace)
	}
	Static [Object]AddDesktopShortcut($Name, $Description, $Path) {Return [Sapphire]::AddDesktopShortcut($Name, $Description, $Path, $Null, $Null, $False)}
	
	# GetManufacturer: get the computer manufacturer.
	Static [String] GetManufacturer() {
		$Info = Get-ComputerInfo
		$Result = "Unknown manufacturer ($($Info.BiosManufacturer))"
		$Manufacturers = @("Dell", "Lenovo","VMware", "Microsoft", "HP")
		ForEach ($Manufacturer In $Manufacturers) {If ($Info.BiosManufacturer -like "*$Manufacturer*") {$Result = $manufacturer; Break}}
		Return $Result
	}
	
	# GetVariables: get hash table of variable names.
	Static [Hashtable] GetVariables() {
		$Variables = @{}; Get-Variable | Select -ExpandProperty Name | Sort | % {$Variables[$_] = $_}
		Return $Variables
	}
		
	# RemoveVariables: removes variables added since getting a hash table.
	Static RemoveVariables($OriginalVariables, $ExceptVariables) {
		$RemoveQueue = New-Object System.Collections.Queue
		ForEach ($VariableName in Get-Variable -Scope Script | Select -ExpandProperty Name) {
			If ((!($OriginalVariables.Contains($VariableName))) -And (!$ExceptVariables.Contains($VariableName))) {$RemoveQueue.Enqueue($VariableName)}
		}
		While ($RemoveQueue.Count) {Remove-Variable $RemoveQueue.Dequeue() -Scope Script}
		While ($RemoveQueue.Count) {Remove-Variable $RemoveQueue.Dequeue() -Scope Script}
	}
	
	# GetTempFolder: get a temporary folder in temp area
	Static [String] GetTempFolder($Prefix) {
		$TempFolder = [System.IO.Path]::GetTempFileName()
		Remove-Item $TempFolder -Force
		$TempFolder = (Split-Path $TempFolder) + "\$Prefix" + (Split-Path $TempFolder -Leaf)
		New-Item -ItemType Directory -Path $TempFolder -Force
		Return $TempFolder
	}
	Static [String] GetTempFolder() {Return [Sapphire]::GetTempFolder("")}

}

# Older version notes
# v1.01 01/09/21 New: $M365AzureMA checked added to CheckM365Tenant
# v1.02 07/01/22 Mod: GetFiles recurse switch added
# v1.03 01/02/22 Fix: RightString function works in all situations
# v1.04 12/02/22 New: LeftString function added
# v1.05 02/03/22 New: CheckEnvironment function added
# v1.06 14/03/22 Mod: CheckM365Tenant recoded to report all missing variables
# v1.07 28/03/22 Mod: Renamed LeftString and RightString as Left and Right
# v1.08 15/04/22 Mod: GetM365Credential renamed GetM365Password
# v1.09 20/04/22 Mod: GetM365Password doesn't return credentials
# v1.10 25/04/22 New: GetRegistry function added
# v1.11 05/08/22 New: $Global:M365AdminUrl set check CheckM365Tenant
# v1.12 11/11/22 Fix: $ProgressPreference = "continue" added to force progress bars
# v1.13 11/11/22 New: CheckM365Tenant checks "M365SPOMA" and "M365EXOMA"
# v1.14 05/12/22 New: Added GetRelativeUrl function
# v1.15 13/03/23 New: GetRegistrySZ depreciated
# v1.16 07/04/23 Mod: CheckEnvironment returns false if failed
# v1.17 20/07/23 Mod: Removed MFA variables from CheckM365Tenant
# v1.18 02/08/23 New: FriendlySize function added
# v1.19 04/08/23 New: MissingCSVHeadings function added
# v1.20 09/08/23 New: CreateTemporaryFile function added
# v1.21 06/10/23 New: ConfirmAction function added
# v1.22 13/10/23 New: ShowProgress can pass a timespan if known more accurately than average per percent
# v1.23 20/11/23 New: RemoveTrailingSlash function added
#				 Mod: GetSoftware removed
# v1.24 11/12/23 New: SendEmail added
#				 Mod: Removed ScriptsFolder - can use "$PSScriptRoot
# v1.25 16/01/24 Mod: Progress system optimised to run faster/only update every 500ms
#				 Fix: Progress bar displayed even when percent complete = 0
# v1.26 22/02/24 New: Recoded progress system to use a class. Always flag added to StartProgress
# v1.27 24/02/24 Mod: Refactored some progress code
# v1.28 01/03/24 Mod: IsAdministrator defaults to showing message
# v1.29 07/03/24 Mod: GetFiles can show progress
# v1.30 31/03/24 Mod: Renamed FriendlyStorage as FriendlySize
# v1.31 08/04/24 New: Added ScrapeArray and MergeHashTable functions
# v1.32 10/04/24 Mod: Changed AnyKey so message can be specified
# v1.33 12/04/24 New: Install function added
#				 New: Plural function added
# v1.34 22/04/24 Mod: GetFiles can handle progress string with {0} for count and {1} for friendly size
#				 Mod: GetFiles can retry if required
#				 Fix: Doesn't crash is time remaining exceeds [in32] maximum. Doesn't show time remaining.
# v1.35 24/04/24 Mod: GetFiles massively reworked to use multiple jobs
#				 New: FriendlyTime added
#				 New: StopJobs added
#				 New: RemoveEmptyFolders added
#				 New: GetRegExGroups added
#				 New: ZeroMilliseconds added
# v1.36 15/05/24 Mod: Plural can accept up to additional parameters
# v1.37 01/06/24 New: Added CreateShadowCopy and DeleteShadowcopy functions
# v1.38 15/08/24 Mod: StartProgress always replaced by a millisecond delay
#				 Fix: Plural count parameter defined as [int] to force conversion
# v1.39 18/08/24 New: Cache function added to expire cached global variables
# v1.40 25/08/24 Mod: Renamed CheckEnvironment as GetEnvironment and also sets the script variables.
# v1.41 30/08/24 Mod: GetCredential and DeleteCredential added
# v1.42 01/09/24 Mod: Old logging functions removed
# v1.43 10/09/24 New: WriteLog to event log added
# v1.44 05/10/24 New: WriteLog can define event type, e.g. warning. Task Category set to none (0)
# v1.45 06/10/24 Fix: FromDesktop works on Windows 11 with Terminal
# v1.46 19/10/24 New: Added NoCache to Uncache
# v1.47 15/11.24 New: Added Wait to GetFiles
# v1.48 19/11/24 New: Added IsYes and IsNo functions
# v1.49 29/11/24 New: Added M365PnPClientId to CheckM365Tenant
# v1.50 01/12/24 New: ReadKey function added
# 				 Mod: AnyKey uses [Console]::ReadKey
# v1.51 11/12/24 New: Breaking change to implement SapphireShadowCopy class/PS v7 compatibility
