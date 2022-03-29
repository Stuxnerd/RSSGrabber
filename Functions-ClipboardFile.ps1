<#
.SYNOPSIS
	Helps to copy easily any links from RSS readers, as it is sufficient to copy it to the clipboard and it will be stored in a file.
	Audio signals (beep) help to recognize the status.
.DESCRIPTION
	TODO
	Assumptions:
	* The user copies several links to the clipboard.
	* Only Strings are copied.
	This script includes these functions:
	* Confirm-SoundDuplicateEntry
	* Confirm-SoundError
	* Confirm-SoundException
	* Confirm-SoundSuccess
	* Get-Clipboard
	* Add-EntryToFile
	* Write-EndlessClipboardToFile
.LINK
	TODO: http://psbus.sourceforge.net
.NOTES
	VERSION: 0.1.0 - 2022-03-30

	AUTHOR: @Stuxnerd
		If you want to support me: bitcoin:19sbTycBKvRdyHhEyJy5QbGn6Ua68mWVwC

	LICENSE: This script is licensed under GNU General Public License version 3.0 (GPLv3).
		Find more information at http://www.gnu.org/licenses/gpl.html

	TODO: These tasks have to be implemented in the following versions:
	* Exception handling
#>

###########################################
#INTEGRATION OF EXTERNAL FUNCTION PACKAGES#
###########################################

#is based on https://github.com/Stuxnerd/PsBuS
. ../PsBuS/Functions-Logging.ps1


##################################
#GLOBAL VARIABLES - RETURN VALUES#
##################################
#the global variables are used to save the return values of functions; they are just for usage of the funtions

#not used in this script


####################################
#GLOBAL VARIABLES - VARIABLE VALUES#
####################################
#these values are used during execution, but are independent from a single fuction invocation

#not used in this script


#############################################
#VARIABLES FOR THE SETTING - CONSTANT VALUES#
#############################################
#these values define the configuration of the script; the might be overwritten by an external script which is using included functions

#Link to the file, where the links are stored to
[String]$FileWithLinks = "./articles.txt"

#Frequency in msec how often the clipboard is checked for new links
[int]$SleepTime = 100

#term to end the endless loop of Write-EndlessClipboardToFile
[String]$ExitTerm = "exitRSS"


#####################
#FUNCTION DEFINITION#
#####################

<#
.SYNOPSIS
	The function makes a specific sound - to recognize a duplicate and proceed
.DESCRIPTION
	The function makes a specific sound which indicates a duplicate enrty to the user.
#>
Function Confirm-SoundDuplicateEntry() {
	[System.Console]::Beep(1300,100)
	[System.Console]::Beep(1300,100)
	[System.Console]::Beep(1300,100)
}

<#
.SYNOPSIS
	The function makes a specific sound - to recognize an error and do NOT proceed
.DESCRIPTION
	The function makes a specific sound which indicates an exception to the user.
#>
Function Confirm-SoundError() {
	[system.media.systemsounds]::Exclamation.play()
}

<#
.SYNOPSIS
	The function makes a specific sound - to recognize an exception and do NOT proceed
.DESCRIPTION
	The function makes a specific sound which indicates an exception to the user.
#>
Function Confirm-SoundException() {
	[System.Console]::Beep(1500,100)
	[System.Console]::Beep(1300,100)
	[System.Console]::Beep(1100,100)
	[System.Console]::Beep(900,100)
}

<#
.SYNOPSIS
	The function makes a specific sound - to recognize success and proceed
.DESCRIPTION
	The function makes a specific sound which indicates an exception to the user.
#>
Function Confirm-SoundSuccess() {
	[System.Console]::Beep(500,300)
}

<#
.SYNOPSIS
	ToDo
.DESCRIPTION
	This function will write the value of the clipoboard to the referenced String $ClipBoardOutput
	Assumptions:
	* The user copies several links to the clipboard.
	* Only Strings are copied.
.EXAMPLE
	ToDo
#>
Function Get-Clipboard() {
	Param(
		#REFERENCE to a [String] to save the clipboard content
		[Parameter(Mandatory = $True, Position = 1)]
		[REF]$ClipBoardOutput		
	)
	try {
		#Create a ClipBoardObject
		Add-Type -AssemblyName System.Windows.Forms
		$ClipBoard = New-Object System.Windows.Forms.TextBox
		#links are always only one line
		$ClipBoard.Multiline = $false
		$ClipBoard.Paste()
		#save value in reference variable
		$ClipBoardOutput.Value = $ClipBoard.Text
        Trace-LogMessage -Level 9 -MessageType Info -Message "Clipboard: ($ClipBoardOutput.Value)"
	}
	catch [Exception] {
		Trace-LogMessage -Indent 1 -Level 1 -MessageType Exception -Message "Exception at using the content of the clipboard"
		#accustic message to the user
		Confirm-SoundException
		$ClipBoardOutput.Value = $null
	}
}

<#
.SYNOPSIS
	TODO
.PARAM
	TODO: Add Parameter for $ExitTerm and $FileWithLinks and $SleepTime
.DESCRIPTION
	This function will endlessly check the clipboiard and wrtie the content (if a link) into the file at $FileWithLinks
	The endless loop can be stopped by copying the term strored in $ExitTerm ("exitRSS").
.EXAMPLE
	TODO
#>
Function Add-EntryToFile() {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$Entry,
		[Parameter(Mandatory=$false, Position=1)]
		[String]
		$FileWithLinks = "./articles.txt"
	)
	#to check if the storage was successfull, the items in the file are counted
	$SizeOld = (Get-Content -Path $FileWithLinks).Length

	#Append to file
	Out-File -FilePath $FileWithLinks -InputObject "$Entry" -Encoding utf8 -Append
	Trace-LogMessage -Indent 5 -Level 5 -MessageType Info -Message "Link $Entry was successfully added"

	#check if the link was really added
	$SizeNew = (Get-Content -Path $FileWithLinks).Length
	if (($SizeOld + 1) -eq $SizeNew) {
		#inform user about the success
		Confirm-SoundSuccess
	} else {
		#Notify user abour error
		Trace-LogMessage -Indent 1 -Level 1 -MessageType Error -Message "Error adding $Entry to $FileWithLinks"
		Confirm-SoundError
	}
}

<#
.SYNOPSIS
	TODO
.PARAM
	TODO: Add Parameter for $ExitTerm and $FileWithLinks and $SleepTime
.DESCRIPTION
	This function will endlessly check the clipboiard and wrtie the content (if a link) into the file at $FileWithLinks
	The endless loop can be stopped by copying the term strored in $ExitTerm ("exitRSS").
.EXAMPLE
	TODO
#>
Function Write-EndlessClipboardToFile() {
	try {
		#To avoid duplicates a list of existing entries is managed
		[System.Collections.ArrayList]$ArticleList = @()

		#The last entries are separatly stored (if the same link is already twice copied, it will not raise a notification
		[String]$LastLinkAdded = ""
		[String]$LastNonLinkContent = ""

		Trace-LogMessage -Indent 3 -Level 3 -MessageType Info -Message "Read old content of file"
		#All existing entries in the file are read to ensure they are included to duplicate check
		Get-Content -Path $FileWithLinks | ForEach-Object { $ArticleList.Add($_) | Out-Null }
		$ArticleList | ForEach-Object { Trace-LogMessage -Indent 5 -Level 5 -MessageType Info -Message ("Link "+$_)}
		[int]$NumberOfOldContent = $ArticleList.Count
		Trace-LogMessage -Indent 3 -Level 3 -MessageType Info -Message "Read $NumberOfOldContent existing entries."
		[int]$NumberOfNewContent = 0

		#an endless loop is used to check the clipboard while a user is copying links
		[Boolean]$GoOn = $true
		while ($GoOn) {
			#get any content from clipboard
			[String]$global:ClipBoardContent = ""
			#the function will store it into the referenced variable
			Get-Clipboard -ClipBoardOutput ([REF]$ClipBoardContent)
# [String]$ClipBoardContent = $global:VARIABLE_ClipBoardOutput #the function will store it into a global variable
			#TODO: What if the content is a picture?
			#null is catched below

			#a special term is is used to exit the endless loop "exitRSS"
			if ($ClipBoardContent -eq $ExitTerm) {
				Trace-LogMessage -Indent 8 -Level 8 -MessageType Info -Message "Program was end by user ($ExitTerm)"
				return
			}

			#check if the clipboard did contain a link and not null
			if ($null -ne $ClipBoardContent -and $ClipBoardContent.StartsWith("http")) {
				#check if the link is already contained
				if($ArticleList.Contains($ClipBoardContent)) {
					#the user is only notified if it is not the last link (he may just copied the link again)
					if ($LastLinkAdded -ne $ClipBoardContent) {
						#not it is the last item
						[String]$LastLinkAdded = $ClipBoardContent
						Trace-LogMessage -Indent 3 -Level 3 -MessageType Warning -Message "$ClipBoardContent is already in the list"
						Confirm-SoundDuplicateEntry
					}
				} else {
					Add-EntryToFile -File $FileWithLinks -Entry $ClipBoardContent
					Out-File -FilePath $global:CONSTANT_LogFilePathError -InputObject $LogMessage -Append
					#save this entry to the list of added content
					$ArticleList.Add($ClipBoardContent) | Out-Null
					#ensure there is no notification, if this link is copied a second time next
					[String]$LastLinkAdded = $ClipBoardContent
					$NumberOfNewContent++
				}
			} else {
				if ($LastNonLinkContent -ne $ClipBoardContent) {
					#if the same item was not a link, the use will not be notified twice
					[String]$LastNonLinkContent = $ClipBoardContent
					Trace-LogMessage -Indent 3 -Level 3 -MessageType Warning -Message "$ClipBoardContent is not a link"
					Confirm-SoundDuplicateEntry
				}
			}
			#wait to not block the processor
			Start-Sleep -Milliseconds $SleepTime
		}
	}
	catch [Exception] {
		Trace-LogMessage -Indent 1 -Level 1 -MessageType Exception -Message "Exception at writing the content of the clipboard into a file"
		Trace-LogMessage -Indent 1 -Level 1 -MessageType Exception -Message $Error
		#accustic message to the user
		Confirm-SoundException
	}
	Trace-LogMessage -Indent 3 -Level 3 -MessageType Info -Message "Added $NumberOfNewContent new entries."
}