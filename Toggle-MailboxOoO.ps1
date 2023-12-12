<#

	.SYNOPSIS

	.DESCRIPTION

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
	Version:
	Template Version: 2.0
	Date: 

	Needs to be ran as Administrator if running the first time to create the Windows Event Log source.

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
#*							Date: 		00/00/0000
#*
#*							Changelog at the bottom of the script
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


#region #?=====LAST RAN SETTINGS===========================

#name for the tim file that stores the last run time, if no name is specified it will use the script name for the file name
$strTimeFile = ""

#endregion #?=====LAST RAN SETTINGS===========================


#region #?=====SCRIPT RESUME SETTINGS======================


#name of the file to store the index data
$strIndexFile = ".ndx"

#name of the file to store the items to process
$strProcessFile = ".csv"

#headers for the index file if desired (other than the index since it is first)
$strIndexDataHeaders = ""


#endregion #?=====SCRIPT RESUME SETTINGS======================


#region #?=====BACKUP CMDLET SETTINGS======================


#name of the folder to store the transaction logs in
$strTransactionLogPath = "TRANSACTION LOGS"

#name of the file to write the transaction logs to
$strTransactionLogFile = ""



#endregion #?=====BACKUP CMDLET SETTINGS======================

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


#region #?=====LAST RAN FUNCTIONS===========================

Function fncLastRan {

	<#

	.SYNOPSIS

	Function to create and pull when needed the last time the script was ran.
	
	.DESCRIPTION

	Function to create and pull when needed the last time the script was ran. This allows for each run of
	a script to pull changes since the last time the script was ran.

	.NOTES

	Below code example of getting the last date the script was ran as saved in the *.tim file
	This would pull all users created in AD since the last time the script was ran
	Running the function without using the EnhancedDebug switch overwrites the *.tim file with the current date/time

	Get-ADUser -Filter {WhenCreated -ge $datLastRan} -Properties DisplayName, Manager, WhenCreated, Description

	The above code would pull all the AD users created since the last time the parent script was ran because
	of the date/time saved in the *.tim file.	

	.PARAMETER EnhancedDebug,
	Allow debugging to set a default time to simulate
	
	.PARAMETER Months,
	Pass a month amount to subtract from the current date when using the debug function of this function
	
	.PARAMETER Days, 
	Pass a day amount to subtract from the current date when using the debug function of this function
	
	.PARAMETER Hours,
	Pass an hour amount to subtract from the current time when using the debug function of this function
	
	.PARAMETER Minutes
	Pass an minute amount to subtract from the current time when using the debug function of this function

	.EXAMPLE
	$GetLastRan = (fncLastRan)
	This will return the last time/date in the *.tim file to the variable $GetLastRan and overwrite the current 
	date/time to the *.tim file

	The below examples are used to pull the date for debugging purposes (so your *.tim file doesn't get modified)
	Let's say the date for the examples is June 1, 2020 10:00 (24HR)

	.EXAMPLE
	PS> fncLastRan -Simulate

	Returns the days date at midnight (06/01/2020 00:00)

	.EXAMPLE
	PS> fncLastRan -Simulate -Months 2

	You can specify an amount of time to subtract from the returned date with one of the following
	Returns (04/01/2020 10:00)

	.EXAMPLE
	PS> fncLastRan -Simulate -Days 2

	Returns (05/30/2020 10:00)

	.EXAMPLE
	PS> fncLastRan -Simulate -Hours 2

	Returns (06/01/2020 08:00)

	.EXAMPLE
	PS> fncLastRan -Simulate -Minutes 15

	Returns (06/01/2020 09:45)

	#>

	Param (
		
		#allow debugging to set a default time to simulate
		[switch]$Simulate,
		
		#pass a month amount to subtract from the current date
		[int]$Months,
		
		#pass a day amount to subtract from the current date
		[int]$Days, 
		
		#pass an hour amount to subtract from the current time
		[int]$Hours,
		
		#pass an minute amount to subtract from the current time
		[int]$Minutes
		
	)

	#set the validation flag to false
	$bolValidInput = $False

	#validate that one of the time parameters have been passed if the simulate switch is on
	If ($Simulate) {
		
		If ($Months -match '^(?:[1-9]\d*|10)$') {$bolValidInput = $True}
		ElseIf ($Days -match '^(?:[1-9]\d*|10)$') {$bolValidInput = $True}
		ElseIf ($Hours -match '^(?:[1-9]\d*|10)$') {$bolValidInput = $True}
		ElseIf($Minutes -match '^(?:[1-9]\d*|10)$') {$bolValidInput = $True}

		#check if the input was not valid and return the error
		If (!$bolValidInput) {

			#since the needed parameters are not met, error and exit the function
			#write to the log
			fncLogThis "ERROR: fncLastRan - Need to specify a time when using the Simulate switch"

			Return "Error"
		
		}

	}

	
	#check if debugging is enabled
	If ($Simulate) {
		
		#set the last ran variable based on the passed info Months
		If ($Months -ne "") {
			
			#subtract the months from now
			$datLastRan = (Get-Date).AddMonths(-($Months))
			
		} #subtract the number of specified days from today for last ran
		
		#set the last ran variable based on the passed info Days
		If ($Days -ne "") {
			
			#subtract the days from now
			$datLastRan = (Get-Date).AddDays(-($Days))
			
		} #subtract the number of specified days from today for last ran
		
		#set the last ran variable based on the passed info Hours
		If ($Hours -ne "") {
			
			#subtract the hours from now
			$datLastRan = (Get-Date).AddHours(-($Hours))
			
		} #subtract the number of specified hours from today for last ran
		
		#set the last ran variable based on the passed info Minutes
		If ($Minutes -ne "") {
			
			#subtract the minutes from now
			$datLastRan = (Get-Date).AddMinutes(-($Minutes))
			
		} #subtract the number of specified hours from today for last ran

		#set the last ran variable based on the passed info Hours
		If (($Months -eq "") -and ($Days -eq "") -and ($Hours -eq "") -and ($Minutes -eq "")) {

			#subtract the days from today
			$datLastRan = ((Get-Date).Date)
			
		} #subtract the number of specified hours from today for last ran
		
	} #run a debug scenario
	#otherwise check and set the last run file
	Else {
	
		#check if there is a date/time file for the last run
		If (Test-Path $strTimeFile) {
			
			#pull the time and date from the file
			[DateTime]$datLastRan = Get-Content -Path $strTimeFile
		
		} #if there is a time file get the last run time from it
		Else{
		
			#get current date
			$datLastRan = ((Get-Date).Date)
		
		} #if no time fild found default to midnight of today
		
		#delete the old time file
		If (Test-Path -Path $strTimeFile) {Remove-Item -Path $strTimeFile}
		
		#write the current time to the time file
		(Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt') | Add-Content -Path $strTimeFile
	
	} #get and set the last time
	
	Return $datLastRan

} #funtion to get the last time the script was ran or create an new time if it had never been ran

#endregion #?=====LAST RAN FUNCTIONS===========================


#region #?=====RESUME SCRIPT FUNCTIONS======================

Function fncScriptResume { #function to help a failed run to resume on a large process

	<#

	.SYNOPSIS

	Function to keep track of the current scripts progress on a large task to allow recovery
	and resume from where it left off in case of cancelling or failure of the script.

	.DESCRIPTION

	Function to keep track of the current scripts progress on a large task to allow recovery
	and resume from where it left off in case of cancelling or failure of the script.

	This is accomplished by an overall data file of the objects that are being tracked and an
	file being used as an index of the processed items to be used to track and resume the script
	as needed.

	.NOTES
	Example of the use of the code in a real application use

	Write the file which the allows the script to pull the next item to resume the processing of.
	Here I'm getting mailboxes with the pertinent information needed for the rest of the script
	so including that in this file near the beginning of the processing is ideal.

	Get-Mailbox -ResultSize 300 -RecipientTypeDetails UserMailbox | Select-Object Name, DisplayName, WindowsEmailAddress, Identity, ArchiveState, RecipientTypeDetails, ArchiveName, GUID, RecoverableItemsQuota, ProhibitSendReceiveQuota, LitigationHoldEnabled, LitigationHoldDate, AutoExpandingArchiveEnabled | Export-Csv -Path "$sMailboxesFile"

	Get the index value if it is needed (last item recorded as been processed)
	$fltSkipValue = nc -NDXpath $strIndexFile -Request

	Grab the mailboxes to work with
	$aMailboxes = Import-Csv $sMailboxesFile
	
	Parse through the mailboxes to grab the statistics from
	ForEach ($eMailbox in $aMailboxes | Select-Object -Skip $fltSkipValue) {
	
		do stuff
		
		Write the index and search name to the index file
		fncScriptResume -NDXpath $strIndexFile -SetIndex -Index $fltIndex -IndexData "$eMailbox.Name,$eMailbox.WindowsEmailAddress,$eMailbox.Identity,$eMailbox.ArchiveState,$eMailbox.RecipientTypeDetails,$eMailbox.ArchiveName"

	}

	.PARAMETER NDXpath,
	Pass the index file/path to use

	.PARAMETER Request,
	Request the index of the last record recorded in the 

	.PARAMETER SetIndex,
	Set the current index

	.PARAMETER Index,
	Passes the current index number to save which should be tracked by the script likely using incrementation of a counter

	.PARAMETER IndexData
	Passes additional data to be saved in the index file that you might want to track from the processes being applied

	.EXAMPLE
	PS> $value = fncScriptResume -NDXpath ".\index.ndx" -Request

	Get the last line of the index file to start the script process at

	.EXAMPLE
	PS> fncScriptResume -NDXpath ".\index.ndx" -SetIndex -Index $Counter -IndexData "$UserName, SAMAccountName, $ObjectGUID" 

	Write data to the index file

	#>


	Param (

		#pass the index file/path
		[string]$NDXpath,

		#request the index
		[switch]$Request,

		#set the current index
		[switch]$SetIndex,

		#passes the current index number to save
		[string]$Index,

		#passes additional data to be saved in the index file
		[string]$IndexData

	)

	#verify that not more than one switch is passed
	If ($Request -and $SetIndex) {

		#test if the index number is passed if so then default to getting index
		If (($Index -ne $Null) -or ($Index -ne "")) {

			#since the index exists nullify the request and set the new index
			$Request = $Null

		} #if the index is passed then nullify the 
		ElseIf (($Index -eq $Null) -or ($Index -eq "")) {

			#sind the index has not been passed nulify the request to set it
			$SetIndex = $Null

		} #test if the index number is passed if so then default to getting index

	} #validate only one switch is active and default to the one that makes sense

	#verify if the index file is passed as it is always needed
	If ($NDXpath -eq $Null -or $NDXpath -eq "") {

		#write error to log and kill script as this is a fatal error
		fncLogThis -LogData "ERROR: FATAL - no index file passed to nc function" -NextLine
		Exit

	} #verify if the index file is passed and 

	#check if this is a request and pull the index number
	If ($Request) {

		#clear all input values
		$Request = $Null
		$SetIndex = $Null
		$Index = $Null 

		#set the default skip value
		$fltSkipValue = 0

		#check if there is already a partial output file
		If ([bool](Test-Path "$NDXpath")){

			#get the skip value
			$arrIndexFileRead = (Import-Csv "$NDXpath")
			$fltSkipValue = ([double]($($arrIndexFileRead[$arrIndexFileRead.Count-1]).Index))
			If (($fltSkipValue -eq $Null) -or ($fltSkipValue -eq "")){$fltSkipValue = 0}
			$arrIndexFileRead = $Null

			#pass the skup value back to the 
			Return $fltSkipValue

		} #get the number of items to skip if needed
		Else { #if no file then return a skip value of 0

			#return a value of zero for the skip value
			$fltSkipValue = 0
			Return $fltSkipValue 

		} #if no file then return a skip value of 0

	} #getting the index number

	#check if the index file exists and if not create the index file
	If (![bool](Test-Path $NDXpath)){

		#if the file doesn't exist create the file with the headers needed
		Add-Content -Path $NDXpath "Index,$strIndexDataHeaders"

	} #if the index file doesn't exist create the index file	

	#check if the update the index file switch is on
	If ($SetIndex) {

		#write the data in the index file
		Add-Content -Path $NDXpath "$Index, $($IndexData)"

	} #if the index file is being updated

	#clear all input values
	$Request = $Null
	$SetIndex = $Null
	$Index = $Null

} #function to help a failed run to resume on a large process

#endregion #?=====RESUME SCRIPT FUNCTIONS======================


#region #?=====ARRAY CHUNKER FUNCTIONS======================

Function fncChunkArray{

	<#
	.SYNOPSIS
		Splits an array or object into smaller chunks for processing based on specified criteria.

	.DESCRIPTION
		This function splits a given array or object into smaller chunks based on various criteria such as the number of chunks, chunk size, alphanumeric characters, or date fields. It returns an array of chunked data.

	.PARAMETER Array
		The array or object to split into smaller chunks for processing (mandatory).

	.PARAMETER SortElement
		The element by which to sort the input array before splitting it.

	.PARAMETER ByCount
		The number of chunks to create when splitting the array.

	.PARAMETER BySize
		The size of each chunk when splitting the array.

	.PARAMETER ByAlphaNumerical
		Split the array based on alphanumeric characters.

	.PARAMETER ByDate
		Split the array based on a date field.

	.PARAMETER ChunkBy
		The element to use for chunking when splitting by date or alphanumeric criteria.

	.PARAMETER Year
		Chunk the data by year when splitting by date.

	.PARAMETER Month
		Chunk the data by month when splitting by date.

	.EXAMPLE
		PS> $data = 1..9
		PS> fncChunkArray -Array $data -ByCount 3
		
		Split an array into three equal-sized chunks.

	.EXAMPLE
		PS> $data = 1..10
		PS> fncChunkArray -Array $data -BySize 5
		
		Split an array into chunks of five elements each.

	.EXAMPLE
		PS> $data = "apple", "banana", "cherry", "date", "fig", "grape"
		PS> fncChunkArray -Array $data -ByAlphaNumerical
		
		Split an array into chunks based on alphanumeric characters.

	.EXAMPLE
		PS> $data = @(Get-Date "2023-01-01", Get-Date "2024-02-01", Get-Date "2023-02-01")
		PS> fncChunkArray -Array $data -ByDate -ChunkBy "Year"
		
		Split an array into chunks based on a date field (year).

	.NOTES
		Based on code from Doug Finke on StackOverflow
		https://stackoverflow.com/questions/13888253/powershell-break-a-long-array-into-a-array-of-array-with-length-of-n-in-one-line

	#>

	Param (

		#array or object to split into smaller arrays for processing
		[Parameter(Mandatory = $True)]				
		$Array,

		#array or object to split into smaller arrays for processing
		[Parameter(ParameterSetName = 'ByCount')]
		[Parameter(ParameterSetName = 'BySize')]
		[Parameter(ParameterSetName = 'ByAlphaNumerical')]	
		[String]$SortElement,	

		#tell the function how to split the array
		#number of chunks
		[Parameter(Mandatory = $True, ParameterSetName = 'ByCount')]
		[Int]$ByCount,

		#size of chunk
		[Parameter(Mandatory = $True, ParameterSetName = 'BySize')]		
		[Int]$BySize,

		#chunk my alphanumerical
		[Parameter(Mandatory = $True, ParameterSetName = 'ByAlphaNumerical')]		
		[Switch]$ByAlphaNumerical,

		#chunk by date field
		[Parameter(Mandatory = $True, ParameterSetName = 'ByDate')]
		[Switch]$ByDate,	

		#element to chunk by
		[Parameter(Mandatory = $True, ParameterSetName = 'ByDate')]
		[Parameter(ParameterSetName = 'ByAlphaNumerical')]
		[String]$ChunkBy,

		#chunk date by year
		[Parameter(ParameterSetName = 'ByDate')]
		[Switch]$Year,

		#chunk date by year
		[Parameter(ParameterSetName = 'ByDate')]
		[Switch]$Month

	)

	#set the chunk start to 0
	$dblChunkStart = 0

	#set the chunked array array to empty
	$arrChunkedArray = @()

	#see if the SortElement is passed then sort the array
	If ($SortElement -ne "") {

		#sort the input array by the default element
		$Array = $Array | Sort-Object -Property $SortElement

	} #sort by passed element

	#check if needing to break up into a fixed number of chunks
	If ($ByCount -gt 0 ) {

		#calculate the size of each chunk from the array
		$dblChunkSize = [Math]::Floor($Array.Count / $ByCount)

		#calculate the array's leftovers for the last chunk
		$dblChunkRemainder = $Array.Count % $ByCount

		#count through the chucks to create and create them
		For ($dblChunkCounter = 1; $dblChunkCounter -le $ByCount; $dblChunkCounter++) {

			#get the current end of chunk
			$dblChunkEnd = $dblChunkCounter * $dblChunkSize -1

			#create the array chunk with the data provided
			$arrChunkedArray += @(,$Array[$dblChunkStart..$dblChunkEnd])

			#increment the start of the next chunk based on the processed chunk data
			$dblChunkStart = $dblChunkEnd + 1

		} #create the chunks per the information given

		#if there is any remainder then add the last chunk with that
		If ($dblChunkRemainder -gt 0) {

			#write a chunk with the leftover data
			$arrChunkedArray += @(,$Array[$dblChunkStart..($dblChunkEnd + $dblChunkRemainder)])

		} #write the last chunk with the remainder of data if needed

	} #check if needing to break up the array by a fixed number of chunks
	#check if needing to break up into even set size chunks
	ElseIf ($BySize -gt 0 ) {

		#calculate the number of chunks
		$iChunkCount = [Math]::Floor($Array.Count / $BySize)

		#calculate the array's leftovers for the last chunk
		$dblChunkRemainder = $Array.Count % $BySize

		#count through the chucks to create and create them
		For ($dblChunkCounter = 1; $dblChunkCounter -le $iChunkCount; $dblChunkCounter++) {

			#get the current end of chunk
			$dblChunkEnd = $dblChunkCounter * $BySize -1

			#create the array chunk with the data provided
			$arrChunkedArray += @(,$Array[$dblChunkStart..$dblChunkEnd])

			#increment the start of the next chunk based on the processed chunk data
			$dblChunkStart = $dblChunkEnd + 1

		} #create the chunks per the information given

		#if there is any remainder then add the last chunk with that
		If ($dblChunkRemainder -gt 0) {

			#write a chunk with the leftover data
			$arrChunkedArray += @(,$Array[$dblChunkStart..($dblChunkEnd + $dblChunkRemainder)])

		} #write the last chunk with the remainder of data if needed

	} #check if needing to break up the array by a fixed number of chunks
	#check if needing to break up the data alphanurmerically
	ElseIf ($ByAlphaNumerical) {

		#set the array used for alphanumeric chunking
		[char[]]$chrAlphaNumeric = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

		#get the default property to use for extracting the objects
		$strDefaultProperty = ($Array | Get-Member | Where-Object {$_.MemberType -eq "Property"} | Select-Object -First 1 Name).Name

		#process through the AlphaNumeric array to chunk the data into
		ForEach ($elmCharacter in $chrAlphaNumeric) {

			#use the chunk element if passed
			If ($ChunkBy -ne "") {

				#create the array chunk with the data provided
				$arrChunkedArray += @(,($Array | Where-Object {$_.($ChunkBy) -like "$($elmCharacter)*"}))

			}
			Else {
			
				#create the array chunk with the data provided
				$arrChunkedArray += @(,($Array | Where-Object {$_.($strDefaultProperty) -like "$($elmCharacter)*"}))
			
			} #filter the array by the current character

		} #process each of the characters in the AlphaNumeric string

	} #check if needing to break up the data alphanumerically
	#check if needing to break up the data by a date field
	ElseIf ($ByDate) {

		#check if the year, month, and chunkby criteria are included
		If (((!$Year) -and (!$Month)) -and ($ChunkBy -eq "")) {

			#kick an error message back to the calling code
			Return "ERROR"

		} #throw an error and kick from the function

		#check if the year is to be used
		If ($Year -and (!$Month)) {

			#get available years by grouping the array and pulling the values
			$arrAvailableSets = ($Array | Group-Object -Property {$_.($ChunkBy).ToString("yyyy")} | Sort-Object Name)

		} #check if the year is to be used
		##check if the month is to be used
		ElseIf ($Month -and (!$Year)) {

			#get available months by grouping the array and pulling the values
			$arrAvailableSets = ($Array | Group-Object -Property {$_.($ChunkBy).ToString("MM")} | Sort-Object Name)	

		} #check if the month is to be used
		#check if the year and month is to be used
		ElseIf ($Year -and $Month) {

			#get available years by grouping the array and pulling the values
			$arrAvailableSets = ($Array | Group-Object -Property {$_.($ChunkBy).ToString("yyyy-MM")} | Sort-Object Name)		

		} #check if the year and month is to be used

		#loop through the available sets to chunk the data by
		ForEach ($elmAvailableSet in $($arrAvailableSets.Name)) {

#			#create the array chunk with the data provided
#			$arrChunkedArray += @(,($Array | Where-Object {$_.($ChunkBy) -like "*$($elmAvailableSet)*"}))

			#create the array chunk with the data provided
			$arrChunkedArray += @(,($arrAvailableSets | Where-Object {$_.Name -eq $elmAvailableSet}).Group)

		} #process each of the characters in the AlphaNumeric string

	} #check if needing to break up the data by a date field

	#return the new chunked array
	Return $arrChunkedArray
	
} #function to break up large arrays/object sets into more managable chunks for faster processing

#endregion #?=====ARRAY CHUNKER FUNCTIONS======================


#region #?=====SUPPORT FOLDER FILE HUNTER FUNCTIONS=========

Function fncFindSupportFolderOrFile {

	<#

	.SYNOPSIS
		Searches scriptpath, parent, sibling, and children folders for needed support folder.

	.DESCRIPTION
		Searches scriptpath, parent, sibling, and children folders for needed support folders or files. This allows the end 
		user more flexability in folder structure and file storage.	Such as finding the location to store log files 
		in a folder called LOGS or a hunting for a script in a neighboring folder.

	.PARAMETER Name
		The name of the folder or file to find (mandatory).

	.PARAMETER IsFolder
		Specifies that the search is for a folder (mandatory if searching for a folder).

	.PARAMETER IsFile
		Specifies that the search is for a file (mandatory if searching for a file).

	.EXAMPLE
		PS> fncFindSupportFolderOrFile -Name "Data" -IsFolder

		Search for a folder named "Data" in the script's location.

	.EXAMPLE
		PS> fncFindSupportFolderOrFile -Name "Config.xml" -IsFile

		Search for a file named "Config.xml" in the script's location.


	#>

	Param (
		
		#pass the name of the file to find
		[Parameter(Mandatory = $True, ParameterSetName = 'Folder')]
		[Parameter(Mandatory = $True, ParameterSetName = 'File')]
		[string]$Name,
		
		#switch to specify if a folder
		[Parameter(Mandatory = $True, ParameterSetName = 'Folder')]
		[switch]$IsFolder,

		#switch to specify if is is a file
		[Parameter(Mandatory = $True, ParameterSetName = 'File')]
		[switch]$IsFile

	)

	#build the type of test-path cmdlet to use
	If ($IsFile) {$strRunCommand = 'Test-Path -Path $($strFilePath) -PathType Leaf -ErrorAction SilentlyContinue'}
	Else {$strRunCommand = 'Test-Path -Path $($strFilePath) -PathType Container -ErrorAction SilentlyContinue'}

	#checking the script path for a matching folder or file
	#generate the path to check
	$strFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, ".\$Name"))

	#check the scriptroot path
#	If (Test-Path $strFilePath -ErrorAction SilentlyContinue) {
	If (Invoke-Expression $strRunCommand) {

		#return the full file path name
		Return $strFilePath

	} #check the location and kick back the path if the file is found

	#checking the script's parent path for a matching folder or file
	#generate the path to check
	$strFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, "..\$Name"))

	#check the script parent path
#	If (Test-Path $strFilePath -ErrorAction SilentlyContinue) {
	If (Invoke-Expression $strRunCommand) {

		#return the full file path name
		Return $strFilePath

	} #check the location and kick back the path if the file is found

	#checking the script's sibling folders for a matching folder or file
	#loop through children folders
	ForEach ($elmFolder in (Get-ChildItem $([System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, ".\"))) -Directory -Depth 1 -ErrorAction SilentlyContinue)) {

		#generate the path to check
		$strFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($($elmFolder.FullName), ".\$Name"))

		#check the script parent path
#		If (Test-Path $strFilePath -ErrorAction SilentlyContinue) {
		If (Invoke-Expression $strRunCommand) {

			#return the full file path name
			Return $strFilePath

		} #check each location and kick back the 
	
	} #look in each of the children folders	

	#checking the script's children paths for a matching folder or file
	#loop through sibling folders
	ForEach ($elmFolder in (Get-ChildItem $([System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, "..\"))) -Directory -Depth 1 -ErrorAction SilentlyContinue)) {

		#generate the path to check
		$strFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($($elmFolder.FullName), ".\$Name"))


		#check the script parent path
#		If (Test-Path $strFilePath -ErrorAction SilentlyContinue) {
		If (Invoke-Expression $strRunCommand) {

			#return the full file path name
			Return $strFilePath

		} #check each location and kick back the 
	
	} #look in each of the sibling folders

	#if nothing was found pass back $false as the answer so that the rest of the script can deal with it
	Return $False

} #used to find files or folders used by the script in the near by folders of the current folder path


#endregion #?=====SUPPORT FOLDER FILE HUNTER FUNCTIONS=========


#region #?=====BACKUP CMDLET SETTINGS FUNCTIONS=============

Function fncBackupCmdletSettingsPull {

	<#

	.SYNOPSIS
		Pulls data from a specified PowerShell cmdlet and exports it to a CSV file while also running a command.

	.DESCRIPTION
		This function pulls data from a specified PowerShell cmdlet and exports it to a CSV file. Additionally, it runs a provided command after exporting the data.

	.PARAMETER Command
		The command to process (mandatory).

	.PARAMETER LogFilePath
		The path to the logfile where the exported data will be stored (mandatory).

	.EXAMPLE
		PS> fncBackupCmdletSettingsPull -Command "Get-Service -Name ServiceName" -LogFilePath "C:\Logs\ServiceData.csv"

		Pull data from the "Get-Service" cmdlet and export it to a CSV file while running the command "Restart-Service -Name ServiceName."

	#>

    Param (

        #pass the command to process
        [string]$Command,

        #pass the logfile path
        [string]$LogFilePath

    )


    #get the identity info to use on the generated command
    If ($(($Command.Split(" ")[1]).Substring(0,1)) -eq "-"){

        #gather the first property and dat passed in it (for cases of -Identity or like with Get-MSOLUser -UserPrincipalName chump@chumpland.com)
        $strCommandIdentity = "$($Command.Split(" ")[1]) $($Command.Split(" ")[2])"

    }
    Else {

        $strCommandIdentity = $($Command.Split(" ")[1])

    }

    #pull the root of the command to run
    $strCommandRoot = $($Command.Split(" ")[0]).Split("-")[1]    

    #generate new command to export data from
    $strExportCommand = "Get-$strCommandRoot $strCommandIdentity"

    #get current date time to add to the exported data log
#    [string]$strLogDate = "{0:yyyy-MM-dd} {0:HH:mm:ss}" -f (Get-Date)
	[string]$strLogDate = fncGetTimeStamp

    #check if the command supports the -property parameter
    $objCMDletInfo = Get-Command -Name $strExportCommand.Split(" ")[0]
    $bolHasPropertyParam = $objCMDletInfo.Parameters.ContainsKey("Properties")

    #run the get variation of the command if it supports the -properties parameter and adjust if not
    If ($bolHasPropertyParam) {

        #try to pull the data with the -property 
        $objCommandOutput = (Invoke-Expression "$strExportCommand -Properties * -ErrorAction Stop" | Select-Object -Property *)

    }
    Else {

        #try to pull the data without the -property 
        $objCommandOutput = (Invoke-Expression "$strExportCommand -ErrorAction Stop" | Select-Object -Property *)

    }

    #create an object to hold the data for the csv
    $objOutputCsvData = @()

    #counter to track the line to write
    $intLineCounter = 0

    #process each item in the output if it's an object
    $objCommandOutput | ForEach-Object {

        #create an object to hold the cmdlet data chunks
        $objOutputData = New-Object -TypeName PSCustomObject

        #on first item in the list add the date and commands
        If ($intLineCounter -eq 0) {

            #create an object to write to csv
            $objOutputData = [pscustomobject]@{
                Date = $strLogDate
                Command = $Command
                Backup = $strExportCommand
            }

        }

        #write the properties of the backed up data to the output object
        ForEach ($elmProperty in ($objCommandOutput| Get-Member -MemberType *Property).Name) {

            #extract the property value to evaluate
            $strPropertyValue = $objCommandOutput[$intLineCounter]  | Select-Object -ExpandProperty $elmProperty -ErrorAction SilentlyContinue

            #convert arrays or complex objects to string representation
            If ($strPropertyValue -is [System.Array] -or $propertyValue -is [psobject]) {

                #pull the property value from the property and covert to json
                $strPropertyValue = $strPropertyValue | ConvertTo-Json -Compress

            }
            Else {

                #if not an array or object cast to a string
                [string]$strPropertyValue | Out-Null

            }

            #add the value to the object to write to file
            $objOutputData | Add-Member -MemberType NoteProperty -Name $elmProperty -Value $strPropertyValue

        }

        #gather the data into the object for writing to csv
        $objOutputCsvData += $objOutputData

        #counter to track the object
        $intLineCounter++

    }

    #call the function to write the data to the transaction log
    fncBackupCmdletWriteTransactionLog -LogFile $LogFilePath -LogData $objOutputCsvData

    #try running the command
    Try {

        #run the passed command after the data was written
        Invoke-Expression $Command -ErrorVariable CommandError -ErrorVariable Brokey | Out-Null

    }
    Catch {

        #pass the error back to the calling function to aggregate
        Return $Brokey

    }

}


Function fncBackupCmdletWriteTransactionLog {

    Param (

		#pass the logfile location
		[string]$LogFilepath,

		#pass the data to write to the transaction log
		[PSCustomObject]$LogData

    )

    #convert the hastable to csv compatable data
#    $sCSVData = $LogData | ConvertTo-Csv -NoTypeInformation -Delimiter ","

    #set the failure counter to 0
    $intFailureCounter = 0

    #write the command passed to the log file
    While ($True) {

        Try {

            #write data to the log
#			[PSCustomObject]$hOutputCSVData | Export-CSV -Path $strLogFile -NoTypeInformation -Append
            $LogData | Export-CSV -Path $LogFilepath -NoTypeInformation -Append
#           $sCSVData | Out-File -FilePath $LogFilepath -Append

            #if write is successful break from while loop
            Break

        }
        Catch {

            Write-Host "WARNING: Cannot write to the log file, trying again" -ForegroundColor Yellow

            Start-Sleep -Seconds 5

            #increment the failure counter to see if the data is unwritable
            $intFailureCounter++

            If ($intFailureCounter -ge 5) {

                #inform the user there was an unrecoverable error
                Write-Host "ERROR: Cannot write to the log file, trying again" -ForegroundColor Red

                #prompt the user to confirm they want to continue
                $strContinue = Read-Host "Would you like to continue with this action? ('Y' or 'N', default is 'Y' )"

                #check if the user needs to stop the script
                If ($strContinue -eq "N" -or $strContinue -eq "n") {

                    #write to screen that we are closing the script
                    fncLogThis -LogData "User choose to cancel the script" -Color Magenta

                    #kill the script
                    Exit

                } #close the script if the user wants to not run it

            }

        }

    }

}


Function fBackupCmdletSettings() {

	[cmdletbinding()]
    Param(

        #input object from the pipeline
        [Parameter(
            Position=0,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True
        )]
        $PipelineInput,

        #pass the command that needs the settings backed up on
        [Parameter(Mandatory)]
        [string]$Command,

		#pass the transaction log file name to use
        [string]$TransactionLogPrefix


        
    )

    Begin {

        #set an empty array to collect errors that happen
        $arrErrCollector = @()

        #pull the root of the command to run
        $strCommandRoot = $($Command.Split(" ")[0]).Split("-")[1]    

		#if there is a transaction log prefix specified then add the prefix
		If ($TransactionLogPrefix -ne "") {

			#set logging filename
			$strTransactionLogFileName = "$TransactionLogPrefix - $strCommandRoot.csv"

		} #add the prefix to the name
		Else {

			#set logging filename
			$strTransactionLogFileName = "$strCommandRoot.csv"

		} #creat the file name from the command root

        #set the log file path based on the transaction log path
        $strTransactionLogFileName = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strTransactionLogPath, $strTransactionLogFileName))

    }

    #process the pipeline
    Process {

        #check if there was pipeline data and process it
        If ($null -ne $PipelineInput) {

            #process the pipeline data
            ForEach ($emlInput in $PipelineInput) {

                #swap values in command
                $Command = $Command.replace('$($_.','$($emlInput.')

                #call the function to pull the cmdlet settings and collect the errors passed back
                $arrErrCollector += fncBackupCmdletSettingsPull -Command $Command -LogFilePath $strTransactionLogFileName

            }

        }
        Else {

            #call the function to pull the cmdlet settings and collect the errors passed back
            $arrErrCollector = fncBackupCmdletSettingsPull -Command $Command -LogFilePath $strTransactionLogFileName

        }

        Return $arrErrCollector

    }

}


#endregion #?=====BACKUP CMDLET SETTINGS FUNCTIONS=============


#region #?=====MENU MANAGER FUNCTIONS=======================

Function fncMenuManagerShowMain {

	<#
	.SYNOPSIS
		Displays a customizable menu in the console and allows the user to select options.

	.DESCRIPTION
		This function displays a menu in the console, where you can specify the title, menu items, description, and various formatting options. It provides the user with the ability to navigate and select menu options using arrow keys or typing the corresponding number. Additionally, you can enable "Back" and "Quit" options.

	.PARAMETER Title
		The title for the menu.

	.PARAMETER MenuItems
		An ArrayList of menu items, where each item should be an object with properties "Title," "ScriptBlock," and optionally "ToolTip."

	.PARAMETER Description
		A description to be displayed below the title.

	.PARAMETER MenuWidth
		The width of the menu in characters. Defaults to 80.

	.PARAMETER BoxStyle
		Specifies the style of the menu box. Options are "Single," "Double," "Round," "Bold," or custom characters.

	.PARAMETER BoxCharacter
		Specifies the character to use for custom box styles.

	.PARAMETER CursorSelect
		Enables cursor-based menu item selection.

	.PARAMETER ToolTips
		Enables the display of tooltips for menu items.

	.PARAMETER Quit
		Enables the "Quit" option (press 'Q' to quit).

	.PARAMETER Back
		Enables the "Back" option (press 'B' to go back).

	.EXAMPLE
		PS> $MenuItems = @(
			@{ Title = "Option 1"; ScriptBlock = { Write-Host "You selected Option 1" } },
			@{ Title = "Option 2"; ScriptBlock = { Write-Host "You selected Option 2" } }
		)
		
		PS> fncMenuManagerShowMain -Title "Main Menu" -MenuItems $MenuItems -Description "Select an option:"
		
		Display a menu with title, options, and descriptions.

	.EXAMPLE
		PS> $MenuItems = @(
			@{ Title = "Option 1"; ScriptBlock = { Write-Host "You selected Option 1" } },
			@{ Title = "Option 2"; ScriptBlock = { Write-Host "You selected Option 2" } }
		)
		
		PS> fncMenuManagerShowMain -Title "Main Menu" -MenuItems $MenuItems -Description "Select an option:" -Back -Quit
		
		Display a menu with title, options, and descriptions and allow "B" for back and "Q" for quitting the menu.

	.EXAMPLE
		PS> $MenuItems = @(
			@{ Title = "Option 1"; ScriptBlock = { Write-Host "You selected Option 1" }; ToolTip = "Select option 1" },
			@{ Title = "Option 2"; ScriptBlock = { Write-Host "You selected Option 2" }; ToolTip = "Select option 1" }
		)
		
		PS> fncMenuManagerShowMain -Title "Main Menu" -MenuItems $MenuItems -ToolTips
		
		Display a menu with title, options, and tool tips for the individual menu items.

	#>

	Param (

		#title for the menu
		[string]$Title,

		#options to use
		[Parameter (Mandatory=$True)]
		[System.Collections.ArrayList]$MenuItems,

		#offer menu description
		[string]$Description,

		#specify a width
		[int]$MenuWidth = 80,

		#specify if box design
		[string]$BoxStyle,

		#specify a character for box style 0
		[string]$BoxCharacter,

		#specify if the menu should allow for cursor selection
		[switch]$CursorSelect,

		#enable the tool tips to be displayed
		[switch]$ToolTips,

		#enable the quit option
		[switch]$Quit,

		#enable the back option
		[switch]$Back

	)

	#set the default selection to less than the menu supports
	$intSelection = -1

	#default the selector position to zero
	$intSelectCursor = 0

	#clear the number selection to prevent double selections
	$strTypedSelection = $null	
	
	#if the box is needed get the characters needed
	If (!([string]::IsNullOrEmpty($BoxStyle))) {$chrBox = fncMenuManagerBoxStyle -Style $BoxStyle -Character $BoxCharacter}

	:MenBuilder While (($intSelection -lt 1) -or ($intSelection -ge $MenuItems.Count)) {

		#create an array to hold the menu for displaying all at once
		$objMenuWindow = @()

		#top row of box
		If (!([string]::IsNullOrEmpty($BoxStyle))) {

			#write the row data to the menu window object
			$objMenuWindow += New-Object -TypeName PSObject -Property  @{
				Row = "$($chrBox.TopLeftCorner)" + "$($chrBox.HorizontalLine)" * $($MenuWidth - 2) + "$($chrBox.TopRightCorner)"
			}
		}
		
		#write the title
		#add space to end of title if odd number of characters
		If ($($Title.Length % 2) -eq 1) {$Title = $Title + " "} 

		#title rows
		If(!([string]::IsNullOrEmpty($Title))) {

			#if the box style is needed create in the box style
			If (!([string]::IsNullOrEmpty($BoxStyle))) {

				#calculate spaces for menu line
				$intSpaces = ($MenuWidth - 2 - $($Title.Length))/2

				#write the title line data to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.VerticalLine)" + ' '*$intSpaces + "$Title" + ' '*$intSpaces + "$($chrBox.VerticalLine)"
				}

				#write the bottom of the title box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.LeftSplit)" + "$($chrBox.HorizontalSplit)" * $($MenuWidth - 2) + "$($chrBox.RightSplit)"
				}	

			}
			Else {

				#calculate spaces for menu
				$intSpaces = ($MenuWidth - $($Title.Length))/2

				#write the bottom of the title to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = " "*$intSpaces + "$Title"
				}				
			}
		}
		
		#description rows
		If(!([string]::IsNullOrEmpty($Description))) {

			#add space to end of description if odd number of characters
			If ($($Description.Length % 2) -eq 1) {$Description = $Description + " "}		

			#if the box style is needed create in the box style
			If (!([string]::IsNullOrEmpty($BoxStyle))) {

				#calculate spaces for menu line
				$intSpaces = ($MenuWidth - 2 - $($Description.Length))/2

				#write the description line data to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.VerticalLine)" + ' '*$intSpaces + "$Description" + ' '*$intSpaces + "$($chrBox.VerticalLine)"
				}

				#write the bottom of the description box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.LeftSplit)" + "$($chrBox.HorizontalSplit)" * $($MenuWidth - 2) + "$($chrBox.RightSplit)"
				}	

			}
			Else {

				#calculate spaces for menu
				$intSpaces = ($MenuWidth - $($Description.Length))/2

				#write the bottom of the description box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = " "*$intSpaces + "$Description"
				}				
			}
		}

		#blank row
		If (!([string]::IsNullOrEmpty($BoxStyle))) {

			#write a blank row before the menu options to the menu window object
			$objMenuWindow += New-Object -TypeName PSObject -Property  @{
				Row = "$($chrBox.VerticalLine)" + " " * $($MenuWidth - 2) + "$($chrBox.VerticalLine)"
			}

		}

		#write the menu items by looping through the items passed
		For ($intMenuItemNumber = 0; $intMenuItemNumber -lt $MenuItems.Count; $intMenuItemNumber++) {

			#create the cursor selector if needed
			If (($CursorSelect) -and ($intSelectCursor -ne $intMenuItemNumber + 1)) {$strOptionSelector = "   "}
			ElseIf (($CursorSelect) -and ($intSelectCursor -eq $intMenuItemNumber + 1)) {$strOptionSelector = "-> "}

			#convert menu item number to string to fix padding if needed
			[string]$strMenuOptionNumber = $intMenuItemNumber + 1

			#see if we need to pad the option number
			If (($strMenuOptionNumber.Length % 2) -eq 1) {$strMenuOptionNumber = " $strMenuOptionNumber"}

			#create the menu item string
			$strMenuItem = "$strOptionSelector" + "$strMenuOptionNumber). $(($MenuItems[$intMenuItemNumber]).Title)"

			#if the box style is needed create the menu options
			If (!([string]::IsNullOrEmpty($BoxStyle))) {

				#calculate the spaces to use after the menu item to the right side of the box
				$intSpaces = $MenuWidth - 2 - $($strMenuItem.Length)				

				#write the bottom of the title box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.VerticalLine)" + "$strMenuItem" + " "*$intSpaces + "$($chrBox.VerticalLine)"
				}	

			}
			Else {

				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$strMenuItem"
				}				
				
			}

		}

		#blank row
		If (!([string]::IsNullOrEmpty($BoxStyle))) {

			#write a blank row before the menu options to the menu window object
			$objMenuWindow += New-Object -TypeName PSObject -Property  @{
				Row = "$($chrBox.VerticalLine)" + " " * $($MenuWidth - 2) + "$($chrBox.VerticalLine)"
			}

		}

		#Menu option tool tip rows
		If($ToolTips -and ($intSelectCursor -ne 0)) {

			#if the box style is needed create in the box style
			If (!([string]::IsNullOrEmpty($BoxStyle))) {

				#calculate spaces for tooltip line
				$intSpaces = ($MenuWidth - 4 - $($($MenuItems[$intSelectCursor-1].ToolTip).Length))

				#write the bottom of the description box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.LeftSplit)" + "$($chrBox.HorizontalSplit)" * $($MenuWidth - 2) + "$($chrBox.RightSplit)"
				}				
				#write the description line data to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = "$($chrBox.VerticalLine)" + "  $($MenuItems[$intSelectCursor-1].ToolTip)" + ' '*$intSpaces + "$($chrBox.VerticalLine)"
				}

			}
			Else {

				#rite a blank row before tool tip
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = " "
				}	

				#write the bottom of the description box to the menu window object
				$objMenuWindow += New-Object -TypeName PSObject -Property  @{
					Row = " $($MenuItems[$intSelectCursor-1].ToolTip)"
				}				
			}
		}

		#if the box is needed create the bottom of the box
		If (!([string]::IsNullOrEmpty($BoxStyle))) {

			#write the row data to the menu window object
			$objMenuWindow += New-Object -TypeName PSObject -Property  @{
				Row = "$($chrBox.BotLeftCorner)" + "$($chrBox.HorizontalLine)" * $($MenuWidth - 2) + "$($chrBox.BotRightCorner)"
			}
		}

		#build input prompt
		$strInputPrompt = "Enter selection (1-$($MenuItems.Count)"
		If ($Back) {$strInputPrompt = $strInputPrompt + ", `'B`' for Back"}
		If ($Quit) {$strInputPrompt = $strInputPrompt + ", `'Q`' for Quit"}
		$strInputPrompt = $strInputPrompt + ")"

		#refreshe the screen
		Clear-Host

		#write the menu object
		ForEach ($elmMenuWindow in $objMenuWindow) {Write-Host $($elmMenuWindow.Row)}
		
		#if the cursor selector is allowed and capture keypresses
		If (($CursorSelect)) {

			#prompt for input
			Write-Host " " ; Write-Host $strInputPrompt": " -NoNewline

			#write current input if available
			If (!([string]::IsNullOrEmpty($strTypedSelection))) {Write-Host "$strTypedSelection" -NoNewline}

			#check captured keys for usable ones
			:KeyTrap Do {

				If ([Console]::KeyAvailable) {
					
					#trap the pressed key
					$objKeyInfo = [Console]::ReadKey($True)

					#check if the key is an allowed direction
					If ($objKeyInfo.Key -eq "UpArrow") {

						#increment the selector location
						If ($intSelectCursor -gt 1) {$intSelectCursor--} Else {$intSelectCursor = $($MenuItems.Count)}

						#clear the number selection to prevent double selections
						$strTypedSelection = $null

						#exit the loop to force a redraw of the menu
						Break KeyTrap
					
					}
					#check if the key is an allowed direction
					ElseIf ($objKeyInfo.Key -eq "DownArrow") {

						#increment the selector location
						If ($intSelectCursor -lt $($MenuItems.Count)) {$intSelectCursor++} Else {$intSelectCursor = 1}

						#clear the number selection to prevent double selections
						$strTypedSelection = $null
						
						#exit the loop to force a redraw of the menu
						Break KeyTrap
					
					}
					ElseIf ($($objKeyInfo.KeyChar) -match '\d') {

						#clear the cursor selector
						$intSelectCursor = 0
						
						#write the key pressed to the selection value
						$strTypedSelection = $strTypedSelection + [string]$($objKeyInfo.KeyChar)

						#write the key to the screen
						Write-Host "$($objKeyInfo.KeyChar)" -NoNewLine

					}
					ElseIf ($Quit -and ($($objKeyInfo.KeyChar) -eq "q")) {

						#exit the current menu to the previous if there is one
						Return $strReturn = "Quit"

					}
					ElseIf ($Back -and ($($objKeyInfo.KeyChar) -eq "b")) {

						#exit the current menu to the previous if there is one
						Return $strReturn = "Back"

					}
					ElseIf (($($objKeyInfo.Key) -eq "Backspace") -and ($strTypedSelection.Length -gt 0)) {

						#remove last character typed in the selector
						$strTypedSelection = $strTypedSelection.Substring(0,$($strTypedSelection.Length -1))

						#force a redraw of the menu
						Break KeyTrap

					}
					ElseIf ($($objKeyInfo.Key) -eq "Delete") {

						#remove last character typed in the selector
						$strTypedSelection = $Null

						#force a redraw of the menu
						Break KeyTrap

					}
					ElseIf ($objKeyInfo.Key -eq "Enter") {

						#check if there is a value to use
						If (!([string]::IsNullOrEmpty($strTypedSelection))) {

							#force the string to an int for evaluation
							$intTypedSelection = [Int]$strTypedSelection

							#check if the selection is within the available selections
							If (($intTypedSelection -gt 0) -and ($intTypedSelection -le $($MenuItems.Count))) {

								#clear the number selection to prevent double selections
								$strTypedSelection = $null

								#run the returned menu item's scriptblock
								$objRanCommand = & $(($MenuItems[$intTypedSelection - 1]).ScriptBlock)

								#if back or quit was returned drop from key trap
								If ($objRanCommand -eq "Quit") {Return "Quit"}
#								ElseIf ($objRanCommand -eq "Back") {$strTypedSelection = ""}
								Else {Break KeyTrap}

							}
						}
						ElseIf (!([string]::IsNullOrEmpty($intSelectCursor)) -and ($intSelectCursor -ne 0)) {

							#check if the selection is within the available selections
							If (($intSelectCursor -gt 0) -and ($intSelectCursor -le $($MenuItems.Count))) {

								#clear the number selection to prevent double selections
								$strTypedSelection = $null

								#run the returned menu item's scriptblock
								$objRanCommand = & $(($MenuItems[$intSelectCursor - 1]).ScriptBlock)

								#if back or quit was returned drop from key trap
								If ($objRanCommand -eq "Quit") {Return "Quit"}
#								ElseIf ($objRanCommand -eq "Back") {$strTypedSelection = ""}
								Else {Break KeyTrap}

							}

						}
						Else {

							#clear the number selection to prevent double selections
							$strTypedSelection = $null

							#exit the keytrap
							Break KeyTrap

						}
					
					}

				} 


			} While ($True)

		}
		Else {

			#prompt for input
			Write-Host " " ; $intSelection = Read-Host $strInputPrompt

			#if a valid selection is returned, process the selection
			If (1..$($MenuItems.Count) -contains $intSelection) {

				#run the returned menu item's scriptblock
				$objRanCommand = & $(($MenuItems[$strSelection-1]).ScriptBlock)

				#if the retuned value is quit process as needed
				If ($objRanCommand -eq "Quit") {Break MenuBuilder}

				#if the retuned value is quit process as needed
				If ($objRanCommand -eq "Back") {$intSelection = ""}

			} 
			ElseIf (($intSelection -eq "q") -and ($Quit)) {

				#exit the current menu to the previous if there is one
				Return $strReturn = "Quit"

			}
			ElseIf (($intSelection -eq "b") -and ($Back)) {

				#exit the current menu to the previous if there is one
				Return $strReturn = "Back"

			}
			Else {

				#neutralize any items outside of the allowed selections
				$intSelection = -1

			}

		}

	}

}


Function fncMenuManagerBoxStyle {

	Param (

		#allow an integer to specify the style
		[Parameter(Mandatory=$True)]
		[string]$Style,

		#custom character to use selected
		[string]$Character = " "

	)


	#create the empty object to put the design in
	$arrMenuBorder = @()

	#check which style is needed
	Switch ($Style) {

		{"1", "Box" -eq $_} {
			$arrMenuBorder = New-Object -Type PSCustomObject -Property @{
				HorizontalLine = [char]0x2588;
				VerticalLine = [char]0x2588;
				TopLeftCorner = [char]0x2588;
				TopRightCorner = [char]0x2588;
				BotLeftCorner = [char]0x2588;
				BotRightCorner = [char]0x2588;
				TopSplit = [char]0x2588;
				BotSplit = [char]0x2588;
				HorizontalSplit = [char]0x2588;
				VerticalSplit = [char]0x2588;
				LeftSplit = [char]0x2588;
				RightSplit = [char]0x2588;
				Junction = [char]0x2588
			}
		}
		{"2", "Line" -eq $_} {
			$arrMenuBorder = New-Object -Type PSCustomObject -Property @{
				HorizontalLine = [char]0x2500;
				VerticalLine = [char]0x2502;
				TopLeftCorner = [char]0x250C;
				TopRightCorner = [char]0x2510;
				BotLeftCorner = [char]0x2514;
				BotRightCorner = [char]0x2518;
				TopSplit = [char]0x252C;
				BotSplit = [char]0x2534;
				HorizontalSplit = [char]0x2500;
				VerticalSplit = [char]0x2502;
				LeftSplit = [char]0x251C;
				RightSplit = [char]0x2524;
				Junction = [char]0x253C
			}
		}
		{"3", "DoubleLine" -eq $_} {
			$arrMenuBorder = New-Object -Type PSCustomObject -Property @{
				HorizontalLine = [char]0x2550;
				VerticalLine = [char]0x2551;
				TopLeftCorner = [char]0x2554;
				TopRightCorner = [char]0x2557;
				BotLeftCorner = [char]0x255A
				BotRightCorner = [char]0x255D;
				TopSplit = [char]0x2566;
				BotSplit = [char]0x2569;
				HorizontalSplit = [char]0x2550;
				VerticalSplit = [char]0x2551;
				LeftSplit = [char]0x2560;
				RightSplit = [char]0x2563;
				Junction = [char]0x256C
			}
		}
		{"4", "MixedLine" -eq $_} {
			$arrMenuBorder = New-Object -Type PSCustomObject -Property @{
				HorizontalLine = [char]0x2550;
				VerticalLine = [char]0x2551;
				TopLeftCorner = [char]0x2554;
				TopRightCorner = [char]0x2557;
				BotLeftCorner = [char]0x255A
				BotRightCorner = [char]0x255D;
				TopSplit = [char]0x2564;
				BotSplit = [char]0x2567;
				HorizontalSplit = [char]0x2500;
				VerticalSplit = [char]0x2502;
				LeftSplit = [char]0x255F;
				RightSplit = [char]0x2562;
				Junction = [char]0x253C
			}
		}
		{"0","Character" -eq $_} {
			$arrMenuBorder = New-Object -Type PSCustomObject -Property @{
				HorizontalLine = $Character;
				VerticalLine = $Character;
				TopLeftCorner = $Character;
				TopRightCorner = $Character;
				BotLeftCorner = $Character;
				BotRightCorner = $Character;
				TopSplit = $Character;
				BotSplit = $Character;
				HorizontalSplit = $Character;
				VerticalSplit = $Character;
				LeftSplit = $Character;
				RightSplit = $Character;
				Junction = $Character
			}
		}
 
	}

	#return the selected menu box design
	Return $arrMenuBorder

}


#endregion #?=====MENU MANAGER=================================


#region #?=====JOBS MANAGER FUNCTIONS=======================

Function fncJobManagerCreateJob {

	<#

	.SYNOPSIS
		Creates and starts a background job using a specified script block.

	.DESCRIPTION
		This function creates and starts a background job using a specified script block. You can provide additional parameters such as job name, session (for remote sessions), maximum concurrent jobs, job prefix, and more.

	.PARAMETER ScriptBlock
		The script block to run as a background job.

	.PARAMETER JobPrefix
		A prefix to prepend to the job name (optional).

	.PARAMETER JobName
		A custom name for the job (optional).

	.PARAMETER Session
		The session in which to run the job (for remote sessions, optional).

	.PARAMETER MaxJobs
		The maximum number of concurrent jobs allowed (default is 10).

	.PARAMETER NoNewWindow
		Prevents the job from running in a new window (default is $False).

	.EXAMPLE
		PS> $ScriptBlock = {
			Write-Host "Background job is running."
			Start-Sleep -Seconds 5
			Write-Host "Background job completed."
		}
		
		PS> fncJobManagerCreateJob -ScriptBlock $ScriptBlock -JobName "MyJob" -MaxJobs 5
		
		Create and start a background job to run a script block.

	#>

	Param (

		[Parameter(Mandatory=$True)]
		[scriptblock]$ScriptBlock,

		[Parameter(Mandatory=$False)]
		[string]$JobPrefix,

		[Parameter(Mandatory=$False)]
		[string]$JobName,

		[Parameter(Mandatory=$False)]
		[string]$Session,

		[Parameter(Mandatory=$False)]
		[int]$MaxJobs = 10, 

		[Parameter(Mandatory=$False)]
		[switch]$NoNewWindow = $False

	)

	While (((($arrJobs = Get-Job) | Where-Object {$_.State -eq "Running"} | Measure-Object).Count) -gt $MaxJobs) {

		#write to the log
		fncLogThis "WARNING: Waiting on jobs to finish - Job Limit: $($MaxJobs)"

		#start wait timer
		Start-Sleep -Seconds 5

	} #seeing if we need to wait for jobs to complete before firing off more jobs

    #determine if we are in a remote session
    $bolRemoteSession = $false
    if ($null -ne $PSSenderInfo) {$bolRemoteSession = $true}

	#build the job name to use with script name, prefix, and job name
	$strFullJobName = ""
	If (!([string]::IsNullOrEmpty($JobPrefix))) {$strFullJobName = "$sScriptName-$JobPrefix-$JobName"} Else {$strFullJobName = "$sScriptName-$JobName"}

	#write to the log
	fncLogThis "INFO: Starting job $strFullJobName"

	#clear the variable used to return the job
	$objJob = $Null

	#run invoke-command for a remote session and start-job for local
	If ($bolRemoteSession) {

		#try to create the job
		Try {

			#create the job to process in the background
			$objJob = Invoke-Command -ScriptBlock Start-Job -ScriptBlock $ScriptBlock -Name $strFullJobName -Session $Session -NoNewWindow:$NoNewWindow -AsJob

			#write to the log
			fncLogThis "SUCCESS: Job $strFullJobName created"

			#return the job information for tracking
			Return $objJob		
		
		}
		Catch {

			#write to the log
			fncLogThis "WARNING: Failed to create job $strFullJobName" -WELBlock -LastError "$_"

		} #try to create the job

	} #run invoke-command for the remote session
	Else {

		#try to create the job
		Try {

			#create the job to process in the background
			$objJob = Start-Job -ScriptBlock $ScriptBlock -Name $strFullJobName #-NoNewWindow:$NoNewWindow

			#write to the log
			fncLogThis "SUCCESS: Job $strFullJobName created"

			#return the job information for tracking
			Return $objJob		
		
		}
		Catch {

			#write to the log
			fncLogThis "WARNING: Failed to create job $strFullJobName" -WELBlock -LastError "$_"

		} #try to create the job

	} #run start-job for the local session

}

Function fncJobManagerRetrieveJob {

	<#
	.SYNOPSIS
		Retrieve completed background jobs based on specified criteria.

	.DESCRIPTION
		This function retrieves completed background jobs based on specified criteria, such as job prefix, session, return count, and more. It can also remove jobs as they are retrieved and optionally wait for jobs to finish.

	.PARAMETER JobPrefix
		The prefix of the jobs to search for (optional).

	.PARAMETER Session
		The session to search for running jobs (optional).

	.PARAMETER ReturnCount
		The number of completed jobs to retrieve. Use 'All' to retrieve all completed jobs.

	.PARAMETER Remove
		Remove jobs as they are retrieved (default is $False).

	.PARAMETER Wait
		Wait for all jobs to finish before retrieving them (default is $False).

	.EXAMPLE
		PS> $CompletedJobs = fncJobManagerRetrieveJob -JobPrefix "MyJobPrefix" -ReturnCount "All"
		PS> $CompletedJobs | ForEach-Object {
			Write-Host "Job Name: $($_.JobName)"
			Write-Host "Command: $($_.Command)"
			Write-Host "Job Data: $($_.JobData)"
			Write-Host "---"
		}
		
		Retrieve and display information for all completed jobs with a specific prefix.

	.EXAMPLE
		PS> $CompletedJobs = fncJobManagerRetrieveJob -JobPrefix "MyJobPrefix" -ReturnCount 1 -Remove
		PS> $CompletedJobs | ForEach-Object {
			Write-Host "Job Name: $($_.JobName)"
			Write-Host "Command: $($_.Command)"
			Write-Host "Job Data: $($_.JobData)"
			Write-Host "---"
		}
		
		Retrieve and display information for the first completed job and removes that job after getting the data	

	.EXAMPLE
		PS> $CompletedJobs = fncJobManagerRetrieveJob -JobPrefix "MyJobPrefix" -ReturnCount "All" -Wait -Remove
		PS> $CompletedJobs | ForEach-Object {
			Write-Host "Job Name: $($_.JobName)"
			Write-Host "Command: $($_.Command)"
			Write-Host "Job Data: $($_.JobData)"
			Write-Host "---"
		}
		
		Retrieve and display information for the all of the completed jobs, but waits until all running jobs 
		are completed before returning data and removing them.	

	#>

    param(
		#specify the prefix of the jobs to search for
		[Parameter(Mandatory=$False)]
		[string]$JobPrefix,
		
		#specify the session if needed for running an invoke-command job
		[Parameter(Mandatory=$False)]
		[string]$Session,

		#
		[Parameter(Mandatory=$True)]
		[ValidateScript({
			If($_ -eq "All" -or $_ -match "^\d+$") { $True }
			Else { Throw "ReturnCount must be a numerical value or 'All'" }
		})]
		[string]$ReturnCount,

		#specify if this function should remove the jobs as it pulls the data
		[switch]$Remove,

		#have the function wait until all jobs are done
		[switch]$Wait

    )

	#write to the log
	fncLogThis "INFO: Checking finished jobs"	

	#build the job name to use with script name and prefix to search for
	$strFullJobName = ""
	If (!([string]::IsNullOrEmpty($JobPrefix))) {
		
		#calculate the name to use with the prefix
		$strFullJobName = "$sScriptName-$JobPrefix-*"

	} 
	Else {
		
		#calculate the name to use without the prefix
		$strFullJobName = "$sScriptName-*"

	}

	#if needing to wait set the wait value to wait for
	If (($Wait) -and ($ReturnCount -eq "all")) {$intJobsCount = (Get-Job | Measure-Object).Count}
	ElseIf (($Wait) -and ($ReturnCount -ne "all")) {$intJobsCount = $ReturnCount}
	Else {$intJobsCount = -1}

	#if needing to wait alert the user
	If ($Wait) {

		#write to the log
		fncLogThis "WARNING: Waiting on jobs to finish"

	}

	#pull jobs and wait if needed
	Do {

		If ($ReturnCount -eq "All") {

			#check for jobs to retrieve 
			$arrJobs = Get-Job -IncludeChildJob | Where-Object {($_.Name -like "$strFullJobName") -and ($_.State -eq "Completed")}
		
		}
		Else {

			#check for jobs to retrieve by count
			$arrJobs = Get-Job -IncludeChildJob | Where-Object {($_.Name -like "$strFullJobName") -and ($_.State -eq "Completed")} | Select-Object -First $ReturnCount

		}

	} While (($arrJobs.Count) -lt $intJobsCount)

	#create an empty array to store the returned data
	$objJobData = @()

	#process each job and dump into the object to send back to the calling code
	ForEach ($elmJob in $arrJobs) {

		$objJobData += New-Object -TypeName PSObject -Property @{
			JobName = "$($elmJob.Name)";
			JobObject = "$elmJob";
			Command = "$($elmJob.Command)";
			JobData = $(Receive-Job $elmJob)
		}

		#if the jobs should be removed
		If ($Remove) {fncJobManagerRemoveJob -Jobs $elmJob}

	}

	#return the item(s) back to the calling script
	Return $objJobData
	
}

Function fncJobManagerRemoveJob {

	<#
	.SYNOPSIS
		Remove background jobs based on specified criteria, such as specific jobs or all completed jobs.

	.DESCRIPTION
		This function removes background jobs based on specified criteria. You can remove specific jobs or all completed jobs. Additionally, you can choose to wait for all jobs to finish before removal.

	.PARAMETER Jobs
		The job(s) to remove. Use this parameter to specify specific jobs (mandatory when not using -AllJobs).

	.PARAMETER AllJobs
		Remove all completed jobs (mandatory when not specifying specific jobs).

	.PARAMETER Wait
		Wait for all jobs to finish before removing them (default is $False).

	.EXAMPLE
		PS> $JobsToRemove = Get-Job | Where-Object { $_.Name -like "MyJob*" }
		PS> fncJobManagerRemoveJob -Jobs $JobsToRemove
		
		Remove specific jobs by name.

	.EXAMPLE
		PS> fncJobManagerRemoveJob -AllJobs -Wait
		
		Wait for all jobs to complete and remove them.

	#>

	Param(

		#job(s) to remove
		[Parameter(Mandatory, ParameterSetName = "Some")]
		$Jobs,

		#or you can remove all completed jobs
		[Parameter(Mandatory, ParameterSetName = "All")]
		[switch]$AllJobs,

		#have the function wait for all jobs to finish before removing them
		[Parameter(ParameterSetName = "All")]
		[switch]$Wait
	)

	#if needing to wait set the wait value to wait for
	If (($Wait) -and ($AllJobs)) {$intJobsCount = (Get-Job | Measure-Object).Count}
	ElseIf (($Wait) -and !($AllJobs)) {$intJobsCount = 1}
	Else {$intJobsCount = -1}

	#if waiting write a message for the user
	If ($Wait) {

		#write to the log
		fncLogThis "WARNING: Waiting on jobs to finish before removal"

	}

	Do {

		#if all jobs are needed then set the $Jobs value with the completed jobs
		If ($AllJobs) {$Jobs = Get-Job | Where-Object {($_.State -eq "Completed")}}

	} While (($Jobs.Count) -lt $intJobsCount)


	#process the jobs to remove
	ForEach ($elmJob in $Jobs) {

		#remove the current job
		Remove-Job -Job $elmJob

		#write to the log
		fncLogThis "SUCCESS: Job $($elmJob.Name) deleted"

	}



}


#endregion #?=====JOBS MANAGER=================================


#endregion #!TEMPLATE----------------------------------------------------


#region #!MAIN----------------------------------------------------------


#endregion #!MAIN-------------------------------------------------------


#endregion


#*=================================================================================================
#*	Initialization
#*=================================================================================================

#region [Initialization]

#if WhatIf is passed, disable until needed in the script
#this will keep cmdlets from triggering on the WhatIf property outside of where this script needs to trigger
If (($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {$global:WhatIfPreference = $False}

#write the start of script message if WhatIf is not passed
If (!($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf'))) {fncLogThis "INFO:Starting script $strScriptName" -WelWrite}


#region #!TEMPLATE-------------------------------------------------------

<#

Script usable ariables created by the initialization code

strScriptRoot - the folder of the current running script
strScriptFile - the filename.ext of the current running script
strScriptName - the name of the currently running script

#>

#region #?=====SCRIPT INITIALIZATION========================

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


#region #?=====LAST RAN INITIALIZATION======================

#check if the log file name is not specified
If ($strTimeFile -eq "") {

	#for loging with the script name
	$strTimeFile = Split-Path $MyInvocation.InvocationName -Leaf
	$strTimeFile = $strTimeFile + ".tim"
	$strTimeFile = $strTimeFile -Replace ".ps1",""
	
	#set the log file path based on the current script's location
	$strTimeFile = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strTimeFile))	
		
} #set the log file to be named after the script name

#set the log file path based on the current script's location
$strTimeFile = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strScriptRoot, $strTimeFile))	

#endregion #?=====LAST RAN INITIALIZATION======================


#region #?=====BACKUP CMDLET SETTINGS INITIALIZATION========

#find the transaction logs folder path and set to the current script's path if no folder is found
If (![bool]($strTransactionLogPath = fncFindSupportFolderOrFile -Name $strTransactionLogPath -IsFolder)) {$strTransactionLogPath = $strScriptRoot}

#check if the log file name is not specified
If ($strTransactionLogFile -eq "") {

	#for loging with the script name
	$strTransactionLogFile = Split-Path $MyInvocation.InvocationName -Leaf
	$strTransactionLogFile = $strTransactionLogFile -Replace ".ps1",""
		
} #set the log file to be named after the script name

#set the log file path based on the current script's location
$strTransactionLogFile = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($strTransactionLogPath, $strTransactionLogFile))	

#endregion #?=====BACKUP CMDLET SETTINGS INITIALIZATION========


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
			fncLogThis "WARNING:Failed to disable OoO for mailbox $($pcoMailbox.EmailAddress)" -WelBlock

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
			fncLogThis "ERROR:Failed to enable OoO for mailbox $($pcoMailbox.EmailAddress)" -WelWrite

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
#	Ver 0.1.5: 00/00/0000 ADD - Added function fncChunkArray to allow splitting of large arrays for easier processing
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
#	                      CHG - Added a ConfirmAction option to fncLogThis for the EnhancedDebug to pause for action approval
#	                      ADD - Option to allow for the debug question to skip the command asked about or exit the script if N is selected
#	                      ADD - 





#=================================================================================================
#	Change Log - 
#=================================================================================================

#	Ver 0.1.0: 12/11/2023 ADD - Intial base code



#endregion
