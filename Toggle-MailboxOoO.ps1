<#

	.SYNOPSIS
	Toggles the OoO messages for mailboxes to reset the Sender list
	
	.DESCRIPTION
	Used as a scheduled task to reset the OoO sender list to prevent the list from filling up and to allow daily OoO messages to be 
	sent to every sender

	.PARAMETER CallingLogPath
	When calling from another script, this allows you to specify which log file to write log output from this script

	.PARAMETER CallingLogLevel
	String to specify the level of data to log. (Error, Warning, Info, Debug, Verbose)

	.PARAMETER CallingLogType
	Switch to allow this script to write its log data to various locations (Screen, ScreenFile, File, FileWEL, WEL, ScreenWEL, All)

	.INPUTS
	None

	.OUTPUTS
	None

	.NOTES
	Author: Zabolyx
	Version: 0.1.0
	Template Version: 1.5
	Date: 12/12/2023

	Needs to be ran as Administrator if running the first time to create the Windows Event Log source.

	Since the senders list that is used to store who all has emailed a mailbox (for knowing who to send a response to) has finite
	storage, it will eventually fill up and start sending OoO messages to every sender on every email if they are not in the sender
	list. This can be prevented by toggling the OoO state to clear the sender list. This prevents the listed issue above as well as 
	allow senders to get an OoO once a day while the OoO is enabled.

	.LINK
	Zabolyx's Github page
	https://github.com/zabolyx

	.LINK
	Template script's repository
	https://github.com/zabolyx/Powershell-Template

	.LINK
	Template script's online Wiki
	https://github.com/zabolyx/Powershell-Template/Wiki

	.ROLE
	Domain User
	External Script

#>


#*=================================================================================================
#*
#*							Script:		Toggle-MailboxOoO.ps1
#*							Author:		Zabolyx
#*							Version:	0.1.0 (Template ver. 0.1.5)
#*							Date: 		12/12/2023
#*
#*							Changelog at the bottom of the script
#*
#*	Used as a scheduled task to reset the OoO sender list to prevent the list from filling up and 
#*	to allow daily OoO messages to be sent to every sender
#*
#*=================================================================================================



#*=================================================================================================
#*	Parameters
#*=================================================================================================

#region [Parameters]


[CmdletBinding(
	#specifies the level of impact the script will have for use with the -Confirm parameter
	ConfirmImpact = "Medium", #default is medium but you can set None, Low, Medium, High

	#allows the use of the -WhatIf parameter when running
	SupportsShouldProcess = $True

	#set the default parameter set to use for the script
#	DefaultParameterSetName = "",

	#URL for the online help page that you want displayed if the user uses Get-Help -Online
#	HelpURI = "https://",

	#allows for the paging options when outputting data (First, Skip, IncludeTotalCount)
#	SupportsPaging = $True

	)
]


Param (
	
	#region #!MAIN----------------------------------------------------------


	#endregion #!MAIN-------------------------------------------------------



	#region #!TEMPLATE-------------------------------------------------------

	#region #?=====LOGGING FUNCTION============================
	
		#The following parameters are used when an external script calles this
		
		#when calling from another script this will allow the logging function to write to the calling scripts log
		[String]$CallingLogPath,
		
		#allows logging of debug messages into the calling scripts logs
		[ValidateSet("Error", "Warning", "Info", "Debug", "Verbose")]	
		[String]$CallingLogLevel,
		
		#allow various logging locations from the calling script (Windows Event Logs should be handled on the calling script's side)
		[ValidateSet("Screen", "ScreenFile", "File", "FileWEL", "WEL", "ScreenWEL", "All")]
		[String]$CallingLogType,

		#enable a transcription to run along with the script
		[switch]$Transcript,

		#enable dumping variables, and the stack with verbose calls as needed for validation and testing
		[switch]$DumpVariables

		#the following switches are used to force the script to log verbose or debug data to the log file and console
		#this replaces the default fucntions built into powershell with a bunch more control

		#force the script to log verbose data
#		[switch]$EnhancedVerbose, #replace the default functionality for CmdletBinding

		#force the script to log debug data and pause before desctructive commands
#		[switch]$EnhancedDebug #replace the default functionality for CmdletBinding


	#endregion #?=====LOGGING FUNCTION============================

	#endregion #!TEMPLATE----------------------------------------------------

)


#endregion


#*=================================================================================================
#*	Settings
#*=================================================================================================

#region [Settings]


#region #!MAIN----------------------------------------------------------

#hashtable to list 
$arrMailboxGUIDs2Proceess = @(

    [PSCustomObject]@{

        EmailAddress = "AccountsPayable@cmh.edu"
        ExchangeGUID = "61f9ac8a-0bb3-4094-b004-228f22151f4f"

    }

)

#endregion #!MAIN-------------------------------------------------------


#region #!TEMPLATE-------------------------------------------------------

#region #?=====LOGGING SETTINGS============================

#name for the log file to use, if no log file name specified it will create one based on the name of the script 
$strLogFile = ""

#path for the log file to use, if no path is specified it will use the scripts current path. This can be a direct path or a relative path (C:\logs or ..\logs)
$strLogPath = ""

#set the logging type (Screen, ScreenFile, File, FileWEL, WEL, ScreenWEL, All)
$strLoggingType = "All"

#set the logging level (Error, Warning, Info, Debug, Verbose)
$strLoggingLevel = "Info"

#set the option to write the logs to the event log ($True,$False)
$bolLogWinEvent = $True

#setting for use with autocoloring of console messages
$arrAutoColorSettings = @(

    [PSCustomObject]@{
		Filter = "WARNING:*"
		LabelSplit = 8
		LabelColor = "Yellow"
		MessageColor = "Yellow"
	}
    [PSCustomObject]@{
		Filter = "ERROR:*"
		LabelSplit = 6
		LabelColor = "Red"
		MessageColor = "Red"
	}
	[PSCustomObject]@{
		Filter = "INFO:*"
		LabelSplit = 5
		LabelColor = "Cyan"
		MessageColor = "Cyan"
	}
	[PSCustomObject]@{
		Filter = "SUCCESS:*"
		LabelSplit = 8
		LabelColor = "Green"
		MessageColor = "Green"
	}
	[PSCustomObject]@{
		Filter = "DONE:*"
		LabelSplit = 5
		LabelColor = "Green"
		MessageColor = "Green"
	}
	[PSCustomObject]@{
		Filter = "DEBUG:*"
		LabelSplit = 6
		LabelColor = "Magenta"
		MessageColor = "Gray"
	}
	[PSCustomObject]@{
		Filter = "*"
		LabelSplit = 0
	}
)

#settings for use with the Windows Event Logging messages
$arrWELSettings = @(

    [PSCustomObject]@{
		Filter = "WARNING:*"
		EventType = "Warning"
		EventId = 64012
	}
    [PSCustomObject]@{
		Filter = "ERROR:*"
		EventType = "Error"
		EventId = 64013
	}
    [PSCustomObject]@{
		Filter = "INFO:*"
		EventType = "Information"
		EventId = 64011
	}
    [PSCustomObject]@{
		Filter = "SUCCESS:*"
		EventType = "SuccessAudit"
		EventId = 64014
	}
    [PSCustomObject]@{
		Filter = "DONE:*"
		EventType = "SuccessAudit"
		EventId = 64014
	}
    [PSCustomObject]@{
		Filter = "DEBUG:*"
		EventType = "Information"
		EventId = 64011
	}
    [PSCustomObject]@{
		Filter = "*"
		EventType = "Information"
		EventId = 64011
	}
)

#endregion #?=====LOGGING SETTINGS============================


#endregion #!TEMPLATE----------------------------------------------------


#endregion


#*=================================================================================================
#*	Functions
#*=================================================================================================


#region [Functions]


#region #!TEMPLATE-------------------------------------------------------


#region #?=====COMMON FUNCTIONS=============================


Function fncGetTimeStamp {

	<#
	.SYNOPSIS
		Retrieves the current timestamp formatted as "[MM/dd/yy HH:mm:ss]".

	.DESCRIPTION
		This function retrieves the current timestamp and formats it as "[MM/dd/yy HH:mm:ss]".

	.EXAMPLE 
		PS> fncGetTimeStamp
		[10/25/23 14:30:45]

		Get the current timestamp.

	.EXAMPLE 
		PS> $logEntry = "[$(fncGetTimeStamp)] This is a log entry."
		PS> Write-Host $logEntry
		[10/25/23 14:30:45] This is a log entry.

		Use the timestamp in a log entry.

	#>

	#get the timestamp
	$datTimeStamp = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)

	#return the timestamp
	Return $datTimeStamp

} #get the time stamp for the log file


#endregion #?=====COMMON FUNCTIONS============================


#region #?=====LOGGING FUNCTIONS============================

Function fncLogThis {

	<#

	.SYNOPSIS
		All in one logging function set that should handle console, file, and Windows Event Log logging needs.

	.DESCRIPTION
		All in one logging function set that should handle console, file, and Windows Event Log logging needs. Along with some 
		enhanced logging functions for debug and verbose that can show the current change to be made as well as dumping all 
		variables in use for easier troubleshooting. By passing various parameters in the in script calls to this function you 
		can specify the location and types of logging to do. This also has functionality at the overall script level to allow 
		a calling script to have the logs from the current script to log to the calling scripts log files and to the screen if 
		requested.

	.PARAMETER LogData
		Log message to record in the logs

	.PARAMETER Color
		Set the color for the console logged data, overriding the preconfigured colors

	.PARAMETER AutoColor
		Allow the predetermined color scheme in the script settings to be use based on the start of the log data
		Accepted labels are INFO:, WARNING:, DEBUG:, INFO:, ERROR:, DONE:, SUCCESS:

	.PARAMETER SameLine
		Allows writing to the same line (defaults to the next line)

	.PARAMETER SkipLine
		Used to skip lines (add blank lines to console) specify the number of blank lineks to add to the console

	.PARAMETER WELBlock
		Used to start a block for the Event Log write allowing for a complete log block to be written as a single items in 
		the Windows Event Log

	.PARAMETER WELDelete
		#Used to delete a block for the Windows Event Log if no longer needed or required

	.PARAMETER WELWrite
		Used to write a the current WEL block to the Windows Event Log

	.PARAMETER Variables
		An object used to pass variables for verbose logging.

	.PARAMETER Function
		If specified, passes the function data from the call stack.

	.PARAMETER LastError
		The current error message to write to a file or log.	
	
	.EXAMPLE
		PS> fncLogThis -LogData "Message to log"

		Simply log a message to all set logging sources (defaults basic messages to white on screen)

	.EXAMPLE
		PS> fncLogThis -LogData "Message to log" -Color Cyan

		Write a log message with the foreground color of cyan

	.EXAMPLE
		PS> fncLogThis -LogData "Total: " -Sameline
		PS> fncLogThis -LogData "100" -Color Red

		Write a 2 part log message on the same line first in white and then finish in red

	.EXAMPLE
		PS> fncLogThis -SkipeLine 3

		Skip 3 lines on the console for readability (doesn't write to file or Event Log)

	.EXAMPLE
		PS> fncLogThis -LogData "Message to log" -SkipLine 3

		Skip 3 lines before writing the current message (doesn't write the skipped lines to file or Event Log)

	.EXAMPLE
		PS> fncLogThis -LogData "ERROR: User account not found in AD"

		Color logs based on their label
		Accepted labels are INFO:, WARNING:, DEBUG:, INFO:, ERROR:, DONE:, SUCCESS:
		If you specify split coloring in the settings area for the functions then the colors listed in the function are followed

	.EXAMPLE
		PS> fncLogThis -LogData "DEBUG: Value is set to $Value"

		Write debug messages for troubleshooting your code
		Messages with the DEBUG label are only written if the debug option is enabled in the functions settings

	.EXAMPLE
		PS> fncLogThis "INFO: Starting the process" -WELBlock

		Write to the standard logging and add to or start a Windows Event Log block

	.EXAMPLE
		PS> fncLogThis "INFO: Testing block deletion and how it handles it" -WELDelete

		Write to the logs and delete the current running Windows Event Log block

	.EXAMPLE
		PS> fncLogThis "ERROR: I'm sorry I can't do that, Dave" -WELWrite

		Write to the logs and write the entire Windows Event Log block file to the Windows Event Log and delete the block file

	.EXAMPLE
		PS> fncLogThis "ERROR: I'm sorry I can't do that, Dave" -LastError $_

		Write to the logs and write the entire Windows Event Log block file to the Windows Event Log and delete the block file 
		as well as write the system's last error thrown to the log file.

	.EXAMPLE
		PS> fncLogThis -LogData "DEBUG: Add-MailboxPermission -Identity $($strMailbox) -User $($strAlias) -AccessRights FullAccess"
		DEBUG: Add-MailboxPermission -Identity TestMailbox -User HanSolo -AccessRights FullAccess
		Would you like to continue with this action? ('Y' or 'N', default is 'Y' ):

		Prompts the user with a command to run and verifies with the user if they want to run the command. Uses the EnhancedDebug 
		switch of the script

	.EXAMPLE

		PS> fncLogThis -LogData "VERBOSE: Testing the verbose logging" -Function -Variables (Get-Variable)

		Pulls the current function, call stack, and variables in use to write to log, which can aid in troubleshooting. Uses the 
		EnhancedDebug switch of the script

	#>

	Param(
		
		#the message to log
		[string]$LogData,
		#color to use
		[string]$Color,
#		#set to auto color
#		[switch]$AutoColor,
		#allows writing to the same line (defaults to the next line)
		[switch]$SameLine,
		#used to skip lines (add blank lines to console)
		[int]$SkipLine = 0,
		#used to start a block for the Event Log write
		[switch]$WELBlock,
		#used to delete a block for the Event Log write
		[switch]$WELDelete,
		#used to write a block to the Even Log
		[switch]$WELWrite,
		#object used to allow pulling in the vairables to log when doing verbose logging
		[object]$Variables,
		#used to pass the function data from the call stack
		[switch]$Function,
		#used to pass the current error to write to file
		[string]$LastError,
		#force the pause to confirm function of the EnhancedDebug with either exit or continue running
		[ValidateSet("Block", "Skip")]
		[string]$ConfirmAction

	)

	#set the logging level to verbose or debug is explictly passed
	If ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) {$strLoggingLevel = "DEBUG"}
	If ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {$strLoggingLevel = "VERBOSE"}	

	#reset the log level comparison flags
	$bolWriteLogData = $False
	$bolWriteCallingLogData = $False

	#verify the logging data will need to be written
	Switch -Regex ($LogData) {
		"^ERROR:" {
			If ($strLoggingLevel -in "ERROR", "WARNING", "INFO", "DEBUG", "VERBOSE") {$bolWriteLogData = $True}
			If ($CallingLogLevel -in "ERROR", "WARNING", "INFO", "DEBUG", "VERBOSE") {$bolWriteCallingLogData = $True}
		}
		"^WARNING:" {
			If ($strLoggingLevel -in "WARNING", "INFO", "DEBUG", "VERBOSE") {$bolWriteLogData = $True}
			If ($CallingLogLevel -in "WARNING", "INFO", "DEBUG", "VERBOSE") {$bolWriteCallingLogData = $True}
		}
		"^INFO:" {
			If ($strLoggingLevel -in "INFO", "DEBUG", "VERBOSE") {$bolWriteLogData = $True}
			If ($CallingLogLevel -in "INFO", "DEBUG", "VERBOSE") {$bolWriteCallingLogData = $True}
		}
		"^DEBUG:" {
			If ($strLoggingLevel -in "DEBUG", "VERBOSE") {$bolWriteLogData = $True}
			If ($CallingLogLevel -in "DEBUG", "VERBOSE") {$bolWriteCallingLogData = $True}
		}
		"^VERBOSE:" {
			If ($strLoggingLevel -in "VERBOSE") {$bolWriteLogData = $True}
			If ($CallingLogLevel -in "VERBOSE") {$bolWriteCallingLogData = $True}
		}
		Default {
			$bolWriteLogData = $True
			$bolWriteCallingLogData = $True
		}
	} #evaluate if the logs need written to file/screen

	#kick from function if we are not supposed to log the data
	If (!($bolWriteLogData) -and !($bolWriteCallingLogData)) {Return}

	#removing the repeat calls to the timestamp function will improve speed and allow the calling script logging to have the same time
	$datTimeStamp = $(fncGetTimeStamp)

	#check if the auto color selection is set
	If ([string]::IsNullOrEmpty($Color)) {

		#run trough the auto color settings for the colors to use
		ForEach ($elmAutoColorSet in $arrAutoColorSettings) {

			#check if the filter matches and assign the variables
			If ($LogData -like "$($elmAutoColorSet.Filter)") {

				$intLabelSplit = $elmAutoColorSet.LabelSplit
				$strLabelColor = $elmAutoColorSet.LabelColor
				$strMessageColor = $elmAutoColorSet.MessageColor

				#on first match kick from the loop
				Break				

			} #assign the variables with the required settings

		} #pull the needed auto color settings

	} #set the color of the text for console based on message

	#check if the color is null and supply a default
	If ([string]::IsNullOrEmpty($Color)) {$Color = "White"}

	#check if the Windows Event Logging selection is set
	If (($strLoggingType -like "*WEL*") -or ($CallingLogType -like "*WEL*")) {

		#run trough the auto color settings for the colors to use
		ForEach ($elmWELSetting in $arrWELSettings) {

			#check if the filter matches and assign the variables
			If ($LogData -like $($elmWELSetting.Filter)) {

				$strEventType = $elmWELSetting.EventType
				$strEventId = $elmWELSetting.EventId

				#on first match kick from the loop
				Break

			} #assign the variables with the required settings

		} #pull the needed auto color settings



	} #set the needed eventid and log type for the incoming data

	#add blank lines on the console
	If (($SkipLine -ne 0) -and (($strLoggingType -or $CallingLogType -like "*Screen*") -or ($strLoggingType -or $CallingLogType -eq "All"))) {
		
		#add the count of blank lines
		For ($intSkipLineCounter = 0; $intSkipLineCounter -lt $SkipLine; $intSkipLineCounter++) {
			
			#write the data to the screen
			Write-Host " " 
			
		} #count through the lines to write to the console
		
		#check if there is no message to log and exit function
		If ($LogData -eq "") {

			#exit function
			Return
			
		} #exit function if there is no data to write
		
	} #check if the new line needs forced and move to the next line	

	#log to screen or file
	If (($strLoggingType -like "*Screen*") -or ($CallingLogType -like "*Screen*") -or ($strLoggingType -eq "All") -or ($CallingLogType -eq "All")) {
		
		#pass the data to the function to write the logs to console
		fncLogThisScreen -LogData $LogData -Color $Color -LabelSplit $intLabelSplit -LabelColor $strLabelColor -MessageColor $strMessageColor
			
	} #log the data to the screen

	#check if we need to log to the file as well
	If (($strLoggingType -like "*File*") -or ($strLoggingType -eq "All") -and ($bolWriteLogData)) {

		#if the error is passed write the error
		If (!([string]::IsNullOrEmpty($LastError))) {

			#write the data to the log file with the function
			fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "$datTimeStamp : $LogData" -LastError $LastError

		}
		Else {

			#write the data to the log file with the function
			fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "$datTimeStamp : $LogData"
			
		}

	} #log the data to the log file
	
	#check if we need to log to the calling file as well
	If ((($CallingLogType -like "*File*") -or ($CallingLogType -eq "All")) -and ($bolWriteCallingLogData) -and (!([string]::IsNullOrEmpty($CallingLogPath)))) {
		
		#write the data to the log file with the function
		fncLogThisFileWrite -LogFileLocation $CallingLogPath -LogData "$datTimeStamp : $LogData"

	} #log the data to the log file	

	#check if the Windows Event block file needs written
	If ((($strLoggingType -like "*WEL*") -or ($CallingLogType -like "*WEL*")) -and $WELBlock) {
	
		
		#if the error is passed write the error
		If (!([string]::IsNullOrEmpty($LastError))) {

			#pass the log data to the fncLogThisEventBlockWriteFile function
			fncLogThisEventBlockWriteFile -LogData $LogData -LastError $LastError

		}
		Else {

			#pass the log data to the fncLogThisEventBlockWriteFile function
			fncLogThisEventBlockWriteFile -LogData $LogData
			
		}

#		#pass the log data to the fncLogThisEventBlockWriteFile function
#		fncLogThisEventBlockWriteFile -LogData $LogData -LastError $LastError
		
	} #write the events to the Windows Event Log block file if needed
	
	#check if the WEL block file need written to the Windows Event
	If ((($strLoggingType -like "*WEL*") -or ($CallingLogType -like "*WEL*")) -and $WELWrite -and $WELBlock) {
		
		#if the error is passed write the error
		If (!([string]::IsNullOrEmpty($LastError))) {

			#write the WEL block file to the Windows Event Log
			fncLogThisEventWriteBlock2WEL -LogData $LogData -EventID $strEventId -EventType $strEventType -LastError $LastError

		}
		Else {

			#write the WEL block file to the Windows Event Log
			fncLogThisEventWriteBlock2WEL -LogData $LogData -EventID $strEventId -EventType $strEventType
			
		}		

		
	} #write the events to the Windows Event Log from the WEL block file

	#check if the Windows Event needs written
	If ((($strLoggingType -like "*WEL*") -or ($CallingLogType -like "*WEL*")) -and $WELWrite) {
	
		#write the lone data to the Windows Event Log
		fncLogThisEventWrite -LogData $LogData -EventID $strEventId -EventType $strEventType
	
	} #write the events to the Windows Event Log if needed

	#check if the WEL block file needs deleted
	If ((($strLoggingType -like "*WEL*") -or ($CallingLogType -like "*WEL*")) -and $WELDelete) {
	
		#remove the WEL block file
		fncLogThisEventClearBlockFile
	
	} #write the events to the Windows Event Log if needed

	#check if the data to log is verbose and then process the variables and call stack
#	If (($LogData -match '^VERBOSE:') -and $EnhancedVerbose) { 
	If (($LogData -match '^VERBOSE:') -and ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)) { 

		#verify the function parameter has data and pull the data for logging the verbose data
		If ($Function -and $DumpVariables) {

			If($MyInvocation -ne $null){

				Try {
					$strScriptLine = $MyInvocation.ScriptLineNumber
				}
				Catch {
					$strScriptLine = "N/A"
				}

				Try {
					$strFunctionName = $MyInvocation.MyCommand.Name
				}
				Catch {
					$strFunctionName = "N/A"
				}

				Try {
					$strStackTrace = (Get-PSCallStack | Select-Object -Skip 1 | Select-Object -Last 5 | ForEach-Object { "    $_.Command" }) -join "`n"
				}
				Catch {
					$strStackTrace = "N/A"
				}

			}

			#pass the data to the verbose logger to write the data
			fncWriteVerboseFunction -ScriptName $strScriptName -FunctionName $strFunctionName -ScriptLine $strScriptLine -StackTrace $strStackTrace

		} #process the function call stack

		#pull in the current variables if passed
		If ($Variables -and $DumpVariables) {

			#pass the data to the verbose logger to write the data
			fncWriteVerboseVariables -Variables $Variables

		} #process the passed variable data

	} #process the needed data for the verbose logging

	#check if we need to process the pause when running debug data
#	If (($LogData -match '^DEBUG:') -and $EnhancedDebug -and ($ConfirmAction -eq "Exit")) {
	If (($LogData -match '^DEBUG:') -and ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) -and ($ConfirmAction -eq "Block")) {

		#prompt the user to confirm they want to continue
		$strContinue = Read-Host "Would you like to continue with this action? ('Y' or 'N', default is 'Y' ) No exits the script."

		#check if the user needs to stop the script
		If ($strContinue -eq "N" -or $strContinue -eq "n") {

			#write to screen that we are closing the script
			fncLogThisScreen -LogData "User choose to cancel the script" -Color Magenta

			#write the data to the log file with the function
			fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "$datTimeStamp : User choose to cancel the script"

			#kill the script
			Exit

		} #close the script if the user wants to not run it

	} #process the needed debug pause for running debug data
#	ElseIf (($LogData -match '^DEBUG:') -and $EnhancedDebug -and ($ConfirmAction -eq "Continue")) {
	ElseIf (($LogData -match '^DEBUG:') -and ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) -and ($ConfirmAction -eq "Skip")) {

		#prompt the user to confirm they want to continue
		$strContinue = Read-Host "Would you like to continue with this action? ('Y' or 'N', default is 'Y' ) No will skip the command."

		#check if the user needs to stop the script
		If ($strContinue -eq "N" -or $strContinue -eq "n") {

			#write to screen that we are closing the script
			fncLogThisScreen -LogData "User choose to skip the command" -Color Magenta

			#write the data to the log file with the function
			fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "$datTimeStamp : User choose to skip the command"

		} #close the script if the user wants to not run it

		Return "Skip"

	} #process the needed debug pause for running debug data

} #function to simplify and clean up the logging needs of the script


Function fncWriteVerboseFunction {
	
	Param (

		#pass the script name
		[string]$ScriptName,
		#pass the function name
		[string]$FunctionName,
		#pass the script line
		[string]$ScriptLine,
		#pass the stack trace
		[string]$TracedStack

	)
	
	#check if there is a file to write the verbose log data to
	If ($strLogFilePath -ne $Null) {

		#write the function data to the log file
		fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "Function: $FunctionName"
		fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "StackTrace: $TracedStack"

	} #write the data to the log file

	#write the data to the console
	fncLogThisScreen -LogData "ScriptName: $ScriptName" -LabelSplit 11 -LabelColor Magenta -MessageColor Yellow
	fncLogThisScreen -LogData "Function: $FunctionName" -LabelSplit 9 -LabelColor Magenta -MessageColor Yellow
	fncLogThisScreen -LogData "Line: $ScriptLine" -LabelSplit 5 -LabelColor Magenta -MessageColor Yellow
	fncLogThisScreen -LogData "StackTrace: $TracedStack" -LabelSplit 11 -LabelColor Magenta -MessageColor Yellow

} #process writing the function data for verbose to file and screen


Function fncWriteVerboseVariables {

	Param (

		#pass the function call stack data to write to the console and file
		[Object]$Variables

	)

	$Variables = $($Variables | Where-Object {$_.Name -like "$sScriptName_*"})
	
	#process the list of variables
	ForEach ($elmVariable in $Variables) {

		#check if there is a file to write the verbose log data to
		If ($strLogFilePath -ne $Null) {

			#write the function data to the log file
			fncLogThisFileWrite -LogFileLocation $strLogFilePath -LogData "$($elmVariable.Name): $($elmVariable.Value)"

		} #write the data to the log file

		#write the data to the console
		fncLogThisScreen -LogData "$($elmVariable.Name): $($elmVariable.Value)" -LabelSplit ("$($elmVariable.Name): $($elmVariable.Value)".IndexOf(":") + 1) -LabelColor Green -MessageColor Yellow
		
	} #process the list of variables to write

} #writes the variables currently in use by the script to file


Function fncLogThisScreen {

	Param (

		#the log data to write to the console
		[string]$LogData,
		#the split location of the log data if needing to dual color console messages
		[int]$LabelSplit,
		#the label color to use
		[string]$LabelColor,
		#the message color to use
		[string]$MessageColor,
		#passes a specified color if needed
		[string]$Color,
		#switch to force it to stay on the same line
		[switch]$SameLine

	)

	#split and color the console output if needed
	If ($LabelSplit -ge 1) {
		
		#split the message into two parts
		$strLabel = $LogData.Substring(0,$LabelSplit)
		$strMesssage = $LogData.Substring($LabelSplit)
		
		#write the label in the label color
		Write-Host "$strLabel" -ForegroundColor $LabelColor -NoNewline
		
		#write the data to the screen
		Write-Host "$strMesssage" -ForegroundColor $MessageColor -NoNewline
		
	} #split message display both parts
	Else {
		
			#write the data to the screen
			Write-Host "$LogData" -ForegroundColor $Color -NoNewline
		
	} #nonsplit message displays
	
	#check if no same line is specified
	If (!$SameLine) {
		
		#write blank line and carriage return
		Write-Host " "
		
	} #line feed if there was no request for same line

} #logs the messages to the console


Function fncLogThisFileWrite {

	Param(
		
		#log file path
		[string]$LogFileLocation,
		#the message to log
		[string]$LogData,
		#error from failure to write to the log file
		[string]$LastError

	)

	#set retry counter
	$intRetryCounter = 0

	#try writing to the file until it is successful
	:LogFileRetryMessage While ($intRetryCounter -le 5) {

		#try to write to the file and delay for 1 second if it fails
		Try {

			#write the data to the log file
			Add-Content -Path $LogFileLocation "$LogData" -ErrorAction SilentlyContinue

			#if this succeeds break from the current loop
			Break LogFileRetryMessage

		}
		Catch {

			#pause for the file to become available
			Start-Sleep -Seconds 1

			#increment the rety counter
			$intRetryCounter++

		} #try to write to the file until it succeeds
		
	} #while the file write fails continue to try again

	#if the error is passed write the error
	If (!([string]::IsNullOrEmpty($LastError))) {

		#set retry counter
		$intRetryCounter = 0	

		#try writing to the file until it is successful
		:LogFileRetryError While ($intRetryCounter -le 5) {

			#try to write to the file and delay for 1 second if it fails
			Try {

				#write the data to the log file
				Add-Content -Path $LogFileLocation "$LastError" -ErrorAction SilentlyContinue

				#if this succeeds break from the current loop
				Break LogFileRetryError

			}
			Catch {

				#pause for the file to become available
				Start-Sleep -Seconds 1

				#increment the rety counter
				$intRetryCounter++

			} #try to write to the file until it succeeds
			
		} #while the file write fails continue to try again

	} #write the last error is passed

} #write the log data to the log file


Function fncLogThisEventBlockWriteFile {

	Param(
			
		#the message to log
		[string]$LogData,
		#error from failure to write to the log file
		[string]$LastError

	)

	#check if the block file exists and if it doesn't create it
	If (![bool](Test-Path $strWELBlockFile)) {
		
		#get the account running the script and the computer
		$strRunningUser = $env:USERNAME
		$strRunningComp = $env:COMPUTERNAME
		
		#write the user's name and computer to the file
		fncLogThisFileWrite -LogFileLocation $strWELBlockFile -LogData "[$strRunningUser]"	
		fncLogThisFileWrite -LogFileLocation $strWELBlockFile -LogData "[$strRunningComp]"	
	
	} #create the file with the user's name and computer

	#write the data to the block file
	fncLogThisFileWrite -LogFileLocation $strWELBlockFile -LogData $LogData


	If (!([string]::IsNullOrEmpty($LastError))) {
		
		#write the error to the block file
		fncLogThisFileWrite -LogFileLocation $strWELBlockFile -LogData $LastError
		
	}

} #write the log data to the WEL block file


Function fncLogThisEventClearBlockFile {

	#check if the block file exists and delete it
	IF ([bool](Test-Path $strWELBlockFile)) {Remove-Item $strWELBlockFile}

	
} #delete the WEL block file


Function fncLogThisEventWriteBlock2WEL {

	Param(
			
		#the message to log
		[string]$LogData,
		#the event ID to use
		[String]$EventId,
		#the event type to use
		[String]$EventType,
		#error from failure to write to the log file
		[string]$LastError

	)

	#if the error is passed write the error
	If (!([string]::IsNullOrEmpty($LastError))) {

		#pass the log data to the fncLogThisEventBlockWriteFile function
		fncLogThisEventBlockWriteFile -LogData $LogData -LastError $LastError

	}
	Else {

		#pass the log data to the fncLogThisEventBlockWriteFile function
		fncLogThisEventBlockWriteFile -LogData $LogData
		
	}	

	#get the data from the block file
	$arrWELMessage = Get-Content -Path $strWELBlockFile -Raw

	#attempt to write the data to the log
	Try {

		#write the data to the Windows Event Log
		Write-EventLog -LogName "Application" -Source "Powershell-$($strScriptName)" -EventID $EventId -EntryType $EventType -Category 0 -Message "$arrWELMessage"

	}
	Catch{}
	
	#call the function to delete the WEL block file
	fncLogThisEventClearBlockFile

} #write the block file to the WEL and delete the block file


Function fncLogThisEventWrite {

	Param(
			
		#the message to log
		[string]$LogData,
		#the event ID to use
		[String]$EventId,
		#the event type to use
		[String]$EventType

	)

	#attempt to write the data to the log
	Try {

		#write the data to the Windows Event Log
		Write-EventLog -LogName "Application" -Source "Powershell-$($strScriptName)" -EventID $EventId -EntryType $EventType -Category 0 -Message "$LogData"

	}
	Catch{}	

} #writes the event log messages to the Windows Event Log

#endregion #?=====LOGGING FUNCTIONS============================


#endregion #!TEMPLATE----------------------------------------------------


#region #!MAIN----------------------------------------------------------


#endregion #!MAIN-------------------------------------------------------


#endregion


#*=================================================================================================
#*	Initialization
#*=================================================================================================

#region [Initialization]


#region #!TEMPLATE-------------------------------------------------------

<#

Script usable ariables created by the initialization code

strScriptRoot - the folder of the current running script
strScriptFile - the filename.ext of the current running script
strScriptName - the name of the currently running script

#>

#region #?=====SCRIPT INITIALIZATION========================

#if WhatIf is passed, disable until needed in the script
#this will keep cmdlets from triggering on the WhatIf property outside of where this script needs to trigger
If (($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {$global:WhatIfPreference = $False}

#write the start of script message if WhatIf is not passed
If (!($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {fncLogThis "INFO:Starting script $strScriptName" -WelWrite}

#get the current path of this script for calling other scripts (snagged from https://stackoverflow.com/questions/6816450/call-powershell-script-ps1-from-another-ps1-script-inside-powershell-ise)
$strScriptRoot = $null
If(!$strScriptRoot){$strScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent}
#$X = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, "filename"))
#& $X -Alternative

#get the script name from the currently running path
$strScriptFile = Split-Path $MyInvocation.InvocationName -Leaf

#try testing the script name
Try {

	#check if the returned info is a powershell file and get the command if it is not
	If (($strScriptFile.Substring($strLogFile.length - 4, 4)) -ne ".ps1") {

		#get the MyCommand portion if the InvocationName is not a powershell script name
		$strScriptFile = $MyInvocation.MyCommand

	} #get the MyCommand name if the InvocationName did not return the powershell script name

}
Catch{

	#get the MyCommand portion if the InvocationName is not a powershell script
	$strScriptFile = $MyInvocation.MyCommand

} #test and get the command name

#get script file name root
$strScriptName = $strScriptFile -Replace ".ps1",""



#endregion #?=====SCRIPT INITIALIZATION========================


#region #?=====LOGGING INITIALIZATION=======================

#if the log file is needed create the folder directory and file if missing
If (($strLoggingType -like "*File*") -or ($strLoggingType -eq "All")) {

	#check if the log file name is not specified, create from the script name
	If ($strLogFile -eq "") {$strLogFile = "$strScriptName.log"}

	#check if a path is specified and what type it is to configure the complete log file path
	If ($strLogPath -eq "") {
		
		#find a logical location for the logs to be stored
		If (![bool]($strLogPath = fncFindSupportFolderOrFile -Name "Logs" -IsFolder)) {

			#if no folder is found hunt for an existing log file
			If (![bool]($strLogFilePath = fncFindSupportFolderOrFile -Name "$strLogFile" -IsFile)) {

				#set the log file path to the script's directory
				$strLogPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strLogPath))
				$strLogFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strLogPath, $strLogFile))

			}
			
		} 
		Else {
				
			#set the log file path of the found directory
			$strLogFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strLogPath, $strLogFile))

		}

	}
	Else {
		
		#check if the path is a direct path or a relative path
		If ($strLogPath -like "*:*") {
			
			#set the log file path based on a direct path
			$strLogFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strLogPath, $strLogFile))
					
		}
		Else{
			
			#set the log file path based on a relative path
			$strTempPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strLogPath))
			$strLogFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strTempPath, $strLogFile))
			
		} #create the log file path based on the data given
		
	} #building the log file path based off of the passed data

	#check if the log file folder exists and create it if it doesn't
	If (![bool](Test-Path -Path $strLogPath)) {

		#create the folder path
		New-Item -ItemType Directory -Force -Path $strLogPath

	} #check and create log file folder if needed

	#check if the log file exists and create the file if needed
	If (![bool](Test-Path -Path $strLogFilePath)) {

		#create the log file
		New-Item -Force -Path $strLogFilePath

	} #check and create the log file if needed

} #find the log file folder and file and create path and file if needed

#check if the Windows Event needs to be used and create the log source if needed and set the WEL block file location
If ($bolLogWinEvent) {

	#create the event log source name
	$strEventLogName = "Powershell-$($strScriptName)"

	#check if the event log source doesn't exist
#	If (![bool](Get-EventLog -LogName "Application" -Source "Powershell-$($strScriptName)" -ErrorAction SilentlyContinue)) {
	If (![bool](Get-EventLog -LogName "Application" -Source $strEventLogName -ErrorAction SilentlyContinue)) {

		#check if the script is running as admin and try to create the 
		If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))	{

			#attempt to create the log source (if it errors it's likely already been created or the script runner doesn't have permission)
			Try {

				#create the log source
#				New-EventLog -LogName "Application" -Source "Powershell-$($strScriptName)" #-ErrorAction SilentlyContinue
				New-EventLog -LogName "Application" -Source $strEventLogName #-ErrorAction SilentlyContinue

			}
			Catch{

				#log the failure to create the Windows Event Log source
				If (!($PSCmdlet.MyInvocation.BoundParameters["WhatIf"].IsPresent)) {fncLogThis -LogData "WARNING: Unable to create the Windows Event Log source"}

				#force the script to skip using the WEL functions
				$bolLogWinEvent = $False

			}
		
		} #verify that the script is running as admin
		Else {

				#log the failure to create the Windows Event Log source
				If (!($PSCmdlet.MyInvocation.BoundParameters["WhatIf"].IsPresent)) {fncLogThis -LogData "WARNING: Cannot create Windows Event Log source because the sript is not running as admin"}

		}

	} #check if the event log source exists and create it if it doesn't

	#set the block file name and location for creating the Windows Event Log blobs
	$strWELBlockFile = "$strScriptName.wel"

	#set the WEL block file path based on a relative path
	$strWELBlockFile = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strWELBlockFile))

}


#endregion #?=====LOGGING INITIALIZATION=======================



#region #?=====TRANSCRIPT INITIALIZATION=====================

	#if the transcripts are to be used
	If ($Transcript) {

		#check if the log file name is not specified, create from the script name
		$strTranscriptFile = $("$strScriptName-$(fncGetTimeStamp).tran") -Replace '\[|\]','' -Replace ':',"" -Replace '/',''

		#find a logical location for the logs to be stored
		If (![bool]($strTranscriptPath = fncFindSupportFolderOrFile -Name "Transcripts" -IsFolder)) {

			#set the log file path to the script's directory
			$strTranscriptPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strTranscriptPath))
			$strTranscriptFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strTranscriptPath, $strTranscriptFile))

			
		} 
		Else {
				
			#set the log file path of the found directory
			$strTranscriptFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strTranscriptPath, $strTranscriptFile))

		}

		#check if the log file folder exists and create it if it doesn't
		If (![bool](Test-Path -Path $strTranscriptPath)) {

			#create the folder path
			New-Item -ItemType Directory -Force -Path $strTranscriptPath

		} #check and create log file folder if needed

		#check if the log file exists and create the file if needed
		If (![bool](Test-Path -Path $strTranscriptFilePath)) {

			#create the log file
			New-Item -Force -Path $strTranscriptFilePath

		} #check and create the log file if needed		

		#start the transcript at the location given
		Start-Transcript -Path $strTranscriptFilePath

	}



#endregion #?=====TRANSCRIPT INITIALIZATION====================

#endregion #!TEMPLATE----------------------------------------------------


#region #!MAIN----------------------------------------------------------



#endregion #!MAIN-------------------------------------------------------


#endregion


#*=================================================================================================
#*	Main Code
#*=================================================================================================


#region [Main]

#if WhatIf was passed, re-enable for use in the scriptin the script
#this will keep cmdlets from triggering on the WhatIf property outside of where this script needs to trigger
If (($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {$global:WhatIfPreference = $True}

#call script to connect to Exchange Online
If (!($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {

	$strScriptPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, "..\Support\Connect-zExO.ps1"))
	& $strScriptPath -CredSet ExO -CallingLogPath $sLogFilePath -CallingLogBoth -NoRemote

}

#process the emails in the configured email list
ForEach ($pcoMailbox in $arrMailboxGUIDs2Proceess) {

	#process the changes if the 
	If ($PSCmdlet.ShouldProcess("$($pcoMailbox.EmailAddress)", "Disable OoO state")) {

		#write to the log
		fncLogThis "INFO:Disabling OoO for mailbox $($pcoMailbox.EmailAddress)" -WelBlock

		#attempt to disable the OoO message
		Try {

			#disable the OoO settings for the mailbox
			Set-MailboxAutoReplyConfiguration accountspayable@cmh.edu -AutoReplyState Disable

			#write to the log
			fncLogThis "SUCCESS:Disabled OoO for mailbox $($pcoMailbox.EmailAddress)" -WelBlock

		}
		Catch {

			#write to the log
			fncLogThis "WARNING:Failed to disable OoO for mailbox $($pcoMailbox.EmailAddress)" -WelBlock -LastError $_

		}	

	} #process disabling OoO

	If ($PSCmdlet.ShouldProcess("$($pcoMailbox.EmailAddress)", "Enable OoO state")) {
		
		#write to the log
		fncLogThis "INFO:Enabling OoO for mailbox $($pcoMailbox.EmailAddress)" -WelBlock

		#attempt to enable the OoO message
		Try {

			#disable the OoO settings for the mailbox
			Set-MailboxAutoReplyConfiguration accountspayable@cmh.edu -AutoReplyState Enabled

			#write to the log
			fncLogThis "SUCCESS:Enabled OoO for mailbox $($pcoMailbox.EmailAddress)" -WelWrite

		}
		Catch {

			#write to the log
			fncLogThis "ERROR:Failed to enable OoO for mailbox $($pcoMailbox.EmailAddress)" -WelWrite -LastError $_

		}

	} #process enabling OoO

} #process each email in the list


#write script completion
If (!($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {fncLogThis "DONE:Script $strScriptName has completed" -WelWrite}


#endregion

#*=================================================================================================
#*	Further Documentation
#*=================================================================================================

#region [Documentation]


#=================================================================================================
#	Change Log - TEMPLATE
#=================================================================================================

#	Ver 0.1.0: 00/00/0000 INT - Base code
#	Ver 0.1.3: 07/04/2020 ADD - Logging back to the calling script's log file
#	                      ADD - Windows Event Logging (for logging service scraping) using idividual log messages or 
#								saving them up in a block and writing as a single message
#	                      CHG - Reduced the calls to fncGetTimeStamp to speed up fncLogThis
#	                      ADD - Option to not log to screen if the script is logging to the calling script (file and 
#								Windows Event Logging still occurs)
#	Ver 0.1.5: 12/12/2023 ADD - Added function fncChunkArray to allow splitting of large arrays for easier processing
#	                      CHG - Improved the speed of the WEL logging creation
#	                      ADD - Function level documentation
#	                      FIX - Now will create the log file/path if it doesn't exist
#	                      ADD - Gets the script file name as well as script root name for use in the script
#						  ADD - Broke up the fncLogThis function into helper functions to reduce code processing
#	                      CHG - Improved the WEL block file creation
#	                      ADD - Function for hunting for existing logging and support file locations
#	                      CHG - Calling script parameters now allow for multiple levels of logging
#	                      CHG - Calling script parameters now allow for multiple types/locations of logging
#	                      CHG - You can now write Event Logs separate from the block file when it exists
#	                      ADD - fncChunkArray now supports sorting the array by specified object property/element
#	                      ADD - fncChunkArray now supports chunking alphanumerically
#	                      ADD - fncChunkArray now supports sorting, chunking alphanumerically, and by specified object property/element
#	                      ADD - fncLogThis function now has a template specific VERBOSE and DEBUG functionality with enhancements over common
#								CmdletBinding functionality
#	                      ADD - Cmdlet settings - backup of current cmdlet properties before running destructive cmdlets
#	                      ADD - Using Better Comments extention for color coding sections dividers
#	                      ADD - Better region formatting for template management when working on a script
#	                      ADD - Simple nestable menu system that allows for read-host or cursor selection
#	                      ADD - Basic job broker functions for creating jobs, allowing for asynchronous job retrieval, and removal
#	                      ADD - Transcript creation and logging in script
#	                      CHG - Added a ConfirmAction option to fncLogThis for the Debug to pause for action approval
#	                      ADD - Option to allow for the debug question to skip the command asked about or exit the script if N is selected
#	                      ADD - WhatIf functionality accounts for possible cmdlet interaction and works around this



#=================================================================================================
#	Change Log - 
#=================================================================================================

#	Ver 0.1.0: 12/12/2023 ADD - Intial base code



#endregion
