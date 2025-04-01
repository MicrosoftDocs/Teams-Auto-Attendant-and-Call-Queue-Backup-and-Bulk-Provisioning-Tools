# Version: 1.0.1
# Date: 2025.04.01



###########################
#  AudioFileExport
###########################
function AudioFileExport
{
	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[string] $fileName,
		[Parameter(Mandatory=$true, Position=1)]
		[string] $fileID,
		[Parameter(Mandatory=$true, Position=2)]
		[string] $CQID
	)

	$currentDIR = (Get-Location).Path
	$audioFilesDIR = $currentDIR + "\AudioFiles\"

	if ( ! (Test-Path -Path $audioFilesDIR) )
	{
		$null = New-Item -Path $currentDIR -Name "AudioFiles" -ItemType Directory
	}
	
	$currentDIR = $audioFilesDIR
	$callQueueDIR = $currentDIR + $CQID

	if ( ! (Test-Path -Path $callQueueDIR) )
	{
		$null = New-Item -Path $currentDIR -Name $CQID -ItemType Directory
	}
	
	$currentDIR = $callQueueDIR + "\" + $fileID + "_" + $fileName

	Write-Host "`t`tDownloading filename: " $fileName
	$content = (Export-CsOnlineAudioFile -ApplicationId "HuntGroup" -Identity $fileID)
	[System.IO.File]::WriteAllBytes($currentDIR, $content)	
		
    return
}

#
# PowerShell Streams
#
#Stream #	Description			Write Cmdlet		Variable				Default
#1			Success stream		Write-Output
#2			Error stream		Write-Error			$ErrorActionPreference	Continue
#3			Warning stream		Write-Warning		$WarningPreference		Continue
#4			Verbose stream		Write-Verbose		$VerbosePrefernce		SilentlyContinue
#5			Debug stream		Write-Debug			$DebugPreference		SilentlyContinue
#6			Information stream	Write-Information	$InformationPreference	SilentlyContinue
#
#Preference Variable Options
# Use name or value
#
#Name				Value
#Break				6
#Suspend			5
#Ignore				4
#Inquire			3
#Continue			2
#Stop				1
#SilentlyContinue	0
#

#
# Confirm running in PowerShell v5.x
#
if ( $PSVersionTable.PSVersion.Major -ne 5 )
{
	Write-Error "This script is only supported in PowerShell v5.x"
	exit
}

#
# Process arguments
#
$args = @()
$arguments = (Get-PSCallStack).Arguments
$arguments = $arguments -replace '[{}]', ''
$arguments = $arguments[0].ToLower()
$arguments = $arguments -split ", "
$args += $arguments

if ( $args -ne "" )
{
	for ( $i = 0; $i -lt $args.length; $i++ )
	{
		switch ( $args[$i] )
		{
			"-aacount"   			{ $AACount = $args[$i+1]
									  $i++
									}
			"-cqcount"   			{ $CQCount = $args[$i+1]
									  $i++
									}
			"-download"				{ $Download = $true }
			"-excelfile" 			{ $ExcelFilename = $args[$i+1]
									  $i++
									}
			"-help"   				{ $Help = $true }
			"-noautoattendants"		{ $NoAutoAttendants = $true }
			"-nocallqueues"			{ $NoCallQueues = $true }
			"-nophonenumbers"		{ $NoPhoneNumbers = $true }
			"-noresourceaccounts"	{ $NoResourceAccounts = $true }
			"-noteamschannels"		{ $NoTeamsChannels = $true }
			"-nousers"				{ $NoUsers = $true }
			"-noopen"               { $NoOpen = $true}
			"-verbose"   			{ $Verbose = $true }
			"-view"      			{ $View = $true }	  
			Default      			{ $ArgError = $true
									  $arg = $args[$i]
									  Write-Warning  "Unknown argument passed: $arg" 
									}
		}
	}
	
	if ( $Download )
	{
		$NoAutoAttendants = $false
		$NoCallQueues = $false
		$NoPhoneNumbers = $false
		$NoResourceAccounts = $false
		$NoTeamsChannels = $false
		$NoUsers = $false
		$Verbose = $true
	}
}

if ( $ArgError )
{
	Write-Host "An unknown argument was encountered. Processing has been halted." -f Red
	Write-Host ""
}


if ( ( $Help ) -or ( $ArgError ) )
{
	Write-Host "The following options are avaialble:"
	Write-Host "`t-AACount <n> - the number of Auto Attendants in tenant, only needed if greater than 100"
	Write-Host "`t-CQCount <n> - the number of Call Queues in tenant, only needed if greater than 100"
	Write-Host "`t-Download - download all Call Queue configuration, including audio file. WARNING - may take a long time"
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkCQs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)"	
	Write-Host "`t-NoAutoAttendants - do not download existing auto attendant information"
	Write-Host "`t-NoCallQueues - do not download existing call queue information"
	Write-Host "`t-NoPhoneNumbers - do not download Voice Apps phone numbers"
	Write-Host "`t-NoResourceAccounts - do not download existing resource account information"
	Write-Host "`t-NoTeamsChannels - do not download existing Teams channels information"
	Write-Host "`t-NoUsers - do not download existing EV enabled user information"	
	Write-Host "`t-NoOpen - do not open spreadsheet when finished"
	Write-Host "`t-Verbose - provides extra messaging during the process"
	Write-Host "`t-View - watch the spreadsheet as the script modifies it"
	exit
}

Write-Host "Starting BulkCQsPreparation."
# Write-Host "Cleaning up from any previous runs."

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Check that minimum verion of required modules are installed - install if not
#
# MicrosoftTeams
#
Write-Host "Checking for MicrosoftTeams module 6.7.0 or later."
$Version = ( (get-installedmodule -Name MicrosoftTeams -MinimumVersion "6.7.0").Version 2> $null )
if ( ( $Version.Major -ge 6 ) -and ( $Version.minor -ge 7 ) )
{
   Write-Host "Connecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion 6.7.0
   
   try
   { 
      Get-CsTenant -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}
else
{
   Write-Host "Module MicrosoftTeams does not exist - installing."
   Install-Module -Name MicrosoftTeams -MinimumVersion 6.7.0 -Force -AllowClobber

   Write-Host "Connecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion 6.7.0
   
   try
   { 
      Get-CsTenant -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}

#
# ImportExcel
#
Write-Host "Checking for ImportExcel module 7.8.0 or later."
$Version = ( (get-installedmodule -Name ImportExcel -MinimumVersion "7.8.0").Version 2> $null )
if ( ( $Version.Major -ge 7 ) -and ( $Version.minor -ge 8 ) )
{
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion 7.8.0
}
else
{
   Write-Host "Module ImportExcel - installing."
   Install-Module -Name ImportExcel -MinimumVersion 7.8.0 -Force -AllowClobber
   
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion 7.8.0
}

#
# setup filename
#
if ( $ExcelFilename -eq $null )
{
   $ExcelFilename = "BulkCQs.xlsm"
}
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename

Write-Host "Accessing the $ExcelFilename worksheet (this may take some time, please be patient)."

#
# check if supplied filename exists
#
if ( !( Test-Path -Path $ExcelFullPathFilename ) )
{
	Write-Error "ERROR: $ExcelFilename does not exist."
	exit
}

$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open($ExcelFullPathFilename)

if ( $View )
{
	$ExcelObj.visible = $true
}

#
# Resource Accounts
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("ResourceAccountsAll")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoResourceAccounts )
{
	Write-Host "Getting list of Resource Accounts."

	$ResourceAccounts = @(get-csonlineapplicationinstance | Sort-Object ApplicationId, DisplayName)

	$j = 2
	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		#
		# Make sure resource account is not Deleted
		#
		$ResourceAccountUserDetails = (get-csonlineuser -identity $ResourceAccounts.ObjectId[$i])
		
		if ( $ResourceAccountUserDetails.SoftDeletionTimeStamp.length -eq 0 )
		{
			if ( $ResourceAccountUserDetails.UsageLocation.length -eq 0 )
			{
				$ResourceAccountUserDetails.UsageLocation = "US"
			}

			if ( $ResourceAccounts.ApplicationId[$i] -eq "ce933385-9390-45d1-9512-c8d228074e07" )
			{
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}) [RA-AA] Resource Account : {1,-50}" -f ($i + 1), $ResourceAccounts.DisplayName[$i] )
				}

				# $ExcelWorkSheet.Cells.Item($j,1) = ("[RA-AA] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" + $ResourceAccounts.PhoneNumber[$i] + "~" + $ResourceAccountUserDetails.UsageLocation + "~" + $ResourceAccountPriority )
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-AA] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" + $ResourceAccounts.PhoneNumber[$i] + "~" + $ResourceAccountUserDetails.UsageLocation )

			}
			else
			{
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}) [RA-CQ] Resource Account : {1,-50}" -f ($i + 1), $ResourceAccounts.DisplayName[$i] )
				}

				# request will generate a "Correlation id for this request" message when the RA is not assigned to anything, also generates error so redirecting that to null
				$ResourceAccountPriority = ( (get-csonlineapplicationinstanceassociation -identity $ResourceAccounts.ObjectId[$i]).CallPriority 2> $null  )
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-CQ] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" +$ResourceAccounts.PhoneNumber[$i] + "~" + $ResourceAccountUserDetails.UsageLocation  + "~" + $ResourceAccountPriority )
			}
			$j++
		}
		else
		{
			if ( $Verbose )
			{
				Write-Host "`tResource Account Not Added (Soft Deleted): " $ResourceAccounts.DisplayName[$i]
			}
		}
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + $j + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Resource Accounts skipped."	
}

#
# Auto Attendants
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("AutoAttendants")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoAutoAttendants )
{
	Write-Host "Getting list of Auto Attendants."

	if ( $AACount -gt 0 )
	{
		$loops = [int] [Math]::Truncate($AACount / 100) + 1
	}
	else
	{
		$loops = 1
	}

	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100
		
		if ( $Verbose )
		{
			Write-Host "`tRetrieving list of auto attendants $($j+1) to $($j+100)"
		}
	
		$AutoAttendants = @(Get-CsAutoAttendant -Skip $j -First 100)

		for ($k=0; $k -lt  $AutoAttendants.length; $k++)
		{
			if ( $Verbose )
			{
				Write-Host ( "`t({0,4}) Auto Attendant : {1,-50}" -f ($k + $j + 1), $AutoAttendants.Name[$k] )
			}

			$Row = $k + $j + 2
			$ExcelWorkSheet.Cells.Item($Row,1) = ($AutoAttendants.Name[$k] + "~" + $AutoAttendants.Identity[$k])
		}
	}

	#
	# Blank out remaining rows
	#
	$Row += 1
	$Range = "A" + ($Row) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Auto Attendants skipped."	
}

#
# Call Queues
#
if ( ! $NoCallQueues )
{
	Write-Host "Getting list of Call Queues."

	if ( $CQCount -gt 0 )
	{
		$loops = [int] [Math]::Truncate($CQCount / 100) + 1
	}
	else
	{
		$loops = 1
	}

	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100

		if ( $Verbose )
		{
			Write-Host "`tRetrieving list of call queues $($j+1) to $($j+100)"
		}

		$CallQueues = @(Get-CsCallQueue -Skip $j -First 100 3> $null )

		if ( $Download )
		{
			$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")
			
			#
			# Blank out any existing data if first time through
			#
			if ( $i -eq 0 )
			{
				$Range = "A3:ZZ2002"
				$ExcelWorkSheet.Range($Range).value = ""
			}
			
			for ($k=0; $k -lt  $CallQueues.length; $k++)
			{
				$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")

				if ( $View )
				{
					$ExcelWorkSheet.activate()
				}

				$RowOffset = $k + $j + 3
				$ColumnOffset = 1
				
				$CQ = (Get-CsCallQueue -Identity $CallQueues.Identity[$k] 3> $null)
			
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Name
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Identity

				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}) Downloading Call Queue : {1,-50}" -f ($k + $j + 1), $CQ.Name )
				}
				
				switch ( $CQ.RoutingMethod )
				{
					"Attendant"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Attendant routing"}
					"Serial" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Serial routing"}
					"RoundRobin" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Round robin (Default)"}
					"LongestIdle" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Longest idle"}
					default  		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.RoutingMethod }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AllowOptOut
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ConferenceMode
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.PresenceBasedRouting
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AgentAlertTime
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.LanguageId
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicResourceId

				if ( $CQ.WelcomeMusicResourceId.length -gt 0 )
				{
					AudioFileExport $CQ.WelcomeMusicFileName $CQ.WelcomeMusicResourceId $CQ.Identity
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeTextToSpeechPrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.UseDefaultMusicOnHold
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldResourceId

				if ( $CQ.MusicOnHoldResourceId.length -gt 0 )
				{
					AudioFileExport $CQ.MusicOnHoldFileName $CQ.MusicOnHoldResourceId $CQ.Identity
				}

				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ServiceLevelThresholdResponseTimeInSecond

				# Overflow
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowThreshold
				
				switch ( $CQ.OverflowAction	)
				{
					"DisconnectWithBusy" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Disconnect (Default)" }
					"Forward" 				{ 	switch ( $CQ.OverflowActionTarget.Type )
												{
													"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
													"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
													"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
													"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
												}												
											} 
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowAction }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowActionCallPriority
				
				# Overflow Shared Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailTextToSpeechPrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailTranscription
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailSystemPromptSuppression

				if ( $CQ.OverflowSharedVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowSharedVoicemailAudioFilePromptFileName $CQ.OverflowSharedVoicemailAudioFilePrompt $CQ.Identity
				}
				
				# Overflow Disconnect
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectTextToSpeechPrompt
				
				if ( $CQ.OverflowDisconnectAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowDisconnectAudioFilePromptFileName $CQ.OverflowDisconnectAudioFilePrompt $CQ.Identity
				}
				
				# Overflow Redirect - Person
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonTextToSpeechPrompt
				
				if ( $CQ.OverflowRedirectPersonAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowRedirectPersonAudioFilePromptFileName $CQ.OverflowRedirectPersonAudioFilePrompt $CQ.Identity
				}
				
				# Overflow Redirect - Voice App
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppTextToSpeechPrompt
				
				if ( $CQ.OverflowRedirectVoiceAppAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName $CQ.OverflowRedirectVoiceAppAudioFilePrompt $CQ.Identity
				}
				
				# Overflow Redirect - Phone Number
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberTextToSpeechPrompt
				
				if ( $CQ.OverflowRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName $CQ.OverflowRedirectPhoneNumberAudioFilePrompt $CQ.Identity
				}
				
				# Overflow Redirect - Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailTextToSpeechPrompt

				if ( $CQ.OverflowRedirectVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.OverflowRedirectVoicemailAudioFilePromptFileName $CQ.OverflowRedirectVoicemailAudioFilePrompt $CQ.Identity
				}

				
				# Timeout
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeOutThreshold
				
				switch ( $CQ.TimeoutAction	)
				{
					"Disconnect" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Disconnect (Default)" }
					"Forward" 		{ 	switch ( $CQ.TimeoutActionTarget.Type )
										{
											"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
											"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
											"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
											"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
										}												
									} 
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutAction }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutActionCallPriority
				
				# Timeout Shared Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailTextToSpeechPrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailTranscription
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailSystemPromptSuppression
				
				if ( $CQ.TimeoutSharedVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutSharedVoicemailAudioFilePromptFileName $CQ.TimeoutSharedVoicemailAudioFilePrompt $CQ.Identity
				}
				
				# Timeout Disconnect
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectTextToSpeechPrompt
				
				if ( $CQ.TimeoutDisconnectAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutDisconnectAudioFilePromptFileName $CQ.TimeoutDisconnectAudioFilePrompt $CQ.Identity
				}
				
				# Timeout Redirect - Person
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonTextToSpeechPrompt
				
				if ( $CQ.TimeoutRedirectPersonAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutRedirectPersonAudioFilePromptFileName $CQ.TimeoutRedirectPersonAudioFilePrompt $CQ.Identity
				}
				
				# Timeout Redirect - Voice App
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppTextToSpeechPrompt
				
				if ( $CQ.TimeoutRedirectVoiceAppAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName $CQ.TimeoutRedirectVoiceAppAudioFilePrompt $CQ.Identity
				}
				
				# Timeout Redirect - Phone Number
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt
				
				if ( $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt $CQ.Identity
				}
				
				# Timeout Redirect - Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailTextToSpeechPrompt

				if ( $CQ.TimeoutRedirectVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName $CQ.TimeoutRedirectVoicemailAudioFilePrompt $CQ.Identity
				}
				
				
				# NoAgent
				switch ( $CQ.NoAgentAction	)
				{
					"Queue" 				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Queue call (Default)" }
					"Disconnect" 			{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Disconnect" }
					"Forward" 				{ 	switch ( $CQ.NoAgentActionTarget.Type )
												{
													"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
													"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
													"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
													"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
												}
											}
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutAction }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentActionCallPriority
				
				# No Agent Shared Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailTextToSpeechPrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailTranscription
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailSystemPromptSuppression
				
				if ( $CQ.NoAgentSharedVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentSharedVoicemailAudioFilePromptFileName $CQ.NoAgentSharedVoicemailAudioFilePrompt $CQ.Identity
				}
				
				switch ( $CQ.NoAgentApplyTo	)
				{
					"NewCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "New Calls" }
					"AllCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "All calls (Default)" }
				}

				# No Agent Disconnect
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectTextToSpeechPrompt

				if ( $CQ.NoAgentDisconnectAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentDisconnectAudioFilePromptFileName $CQ.NoAgentDisconnectAudioFilePrompt $CQ.Identity
				}
				
				# No Agent Redirect - Person
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonTextToSpeechPrompt
				
				if ( $CQ.NoAgentRedirectPersonAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentRedirectPersonAudioFilePromptFileName $CQ.NoAgentRedirectPersonAudioFilePrompt $CQ.Identity
				}
				
				# No Agent Redirect - Voice App
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppTextToSpeechPrompt

				if ( $CQ.NoAgentRedirectVoiceAppAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName $CQ.NoAgentRedirectVoiceAppAudioFilePrompt $CQ.Identity
				}
				
				# No Agent Redirect - Phone Number
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt
				
				if ( $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt $CQ.Identity
				}
				
				# No Agent Redirect - Voicemail
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailTextToSpeechPrompt

				if ( $CQ.NoAgentRedirectVoicemailAudioFilePrompt.length -gt 0 )
				{
					AudioFileExport $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName $CQ.NoAgentRedirectVoicemailAudioFilePrompt $CQ.Identity
				}


				# Callback
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.IsCallbackEnabled
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackRequestDtmf
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WaitTimeBeforeOfferingCallbackInSecond
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NumberOfCallsInQueueBeforeOfferingCallback
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallToAgentRatioThresholdBeforeOfferingCallback
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptResourceId
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptFileName
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferTextToSpeechPrompt
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackEmailNotificationTarget.Id

				if ( $CQ.CallbackOfferAudioFilePromptResourceId.length -gt 0 )
				{
					AudioFileExport $CQ.CallbackOfferAudioFilePromptFileName $CQ.CallbackOfferAudioFilePromptResourceId $CQ.Identity
				}

				
				# Agents
				if ( $CQ.Agents.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.Agents.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.Agents[$l].ObjectId
					}
				}
				
				# Channel
				$ColumnOffset += 20
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ChannelId
				

				# Authorized Users
				if ( $CQ.AuthorizedUsers.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.AuthorizedUsers.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.AuthorizedUsers[$l].Guid
					}
				}
				
				# Hidden Authorized Users
				$ColumnOffset += 15
				if ( $CQ.HideAuthorizedUsers.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.HideAuthorizedUsers.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.HideAuthorizedUsers[$l].Guid
					}
				}


				# Numbers in lists
				$ColumnOffset += 15
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.DistributionLists.length
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ApplicationInstances.length
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OboResourceAccounts.length
				
				# Spare Numbers in Lists
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 04"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 05"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 06"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 07"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 08"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 09"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 10"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 11"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 12"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 13"
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Spare 14"
				
				# Distributions Lists
				if ( $CQ.DistributionLists.length -gt 0 )
				{
					for ($l=0; $l -lt 4; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.DistributionLists[$l].Guid
					}
				}
				
				# Resource Accounts
				$ColumnOffset += $CQ.DistributionLists.length
				
				if ( $CQ.ApplicationInstances.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.ApplicationInstances.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset,$ColumnOffset + $l) = $CQ.ApplicationInstances[$l]
					}
				}
				
				# Obo Resource Accounts
				$ColumnOffset += $CQ.ApplicationInstances.length

				if ( $CQ.OboResourceAccounts.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.OboResourceAccounts.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset,$ColumnOffset + $l) = $CQ.OboResourceAccounts[$l].ObjectId
					}
				}
				
				
				$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
		
				if ( $View )
				{
					$ExcelWorkSheet.activate()
				}
				
				$AssignedResourceAccounts = ( (Get-CsCallQueue -Identity $CallQueues.Identity[$k]).ApplicationInstances 3> $null )
				$ExcelWorkSheet.Cells.Item($RowOffset - 1,1) = ($CallQueues.Name[$k] + "~" + $CallQueues.Identity[$k] + "~" + ($AssignedResourceAccounts -join ","))
			}
			
			#
			# Blank out remaining rows
			#
			$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
		
			if ( $View )
			{
				$ExcelWorkSheet.activate()
			}
			
			$Range = "A" + ($RowOffset) + ":A2001"
			$ExcelWorkSheet.Range($Range).value = ""
			
			$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")

			if ( $View )
			{
				$ExcelWorkSheet.activate()
			}

			$RowOffset += 1
			$Range = "A" + ($RowOffset) + ":ZZ2002"
			$ExcelWorkSheet.Range($Range).value = ""
		} # Download

		if ( ! $Download )
		{
			
			$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
		
			if ( $View )
			{
				$ExcelWorkSheet.activate()
			}

			for ($k=0; $k -lt  $CallQueues.length; $k++)
			{
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}) Call Queue : {1,-50}" -f ($k + $j + 1), $CallQueues.Name[$k] )
				}

				$RowOffset = $k + $j + 2
				$AssignedResourceAccounts = ( (Get-CsCallQueue -Identity $CallQueues.Identity[$k]).ApplicationInstances 3> $null )
				$ExcelWorkSheet.Cells.Item($RowOffset,1) = ($CallQueues.Name[$k] + "~" + $CallQueues.Identity[$k] + "~" + ($AssignedResourceAccounts -join ","))
			}
		}

		#
		# Blank out remaining rows
		#
		$RowOffset += 1
		$Range = "A" + ($RowOffset) + ":A2001"
		$ExcelWorkSheet.Range($Range).value = ""
	}
}
else
{
	Write-Host "Downloading Call Queues skipped."		
}

#
# Phone Numbers
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("PhoneNumbers")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoPhoneNumbers )
{
	Write-Host "Getting list of existing voice application phone numbers"

	$PhoneNumbers = @(get-csphonenumberassignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned")

	for ($i=0; $i -lt  $PhoneNumbers.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}) Phone Number : {1,-50}" -f ($i + 1), $PhoneNumbers.TelephoneNumber[$i] )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($PhoneNumbers.TelephoneNumber[$i] + "~" + $PhoneNumbers.NumberType[$i] + "~" + $PhoneNumbers.IsoSubdivision[$i] + "~" + $PhoneNumbers.IsoCountryCode[$i])
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Voice Applications phone numbers skipped."	
}

#
# Team Channels
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("TeamsChannels")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoTeamsChannels )
{
	Write-Host "Geting list of existing Teams and Channels."

	$Teams = @(get-team | Sort-Object DisplayName)

	$row = 1
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsChannels = @(get-teamchannel -groupId $Teams.GroupId[$i] | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName)

		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}) Processing Team : {1,-50}" -f ($i + 1), $Teams.DisplayName[$i] )

			if ( $TeamsChannels.length -eq 1 )
			{
				Write-Host "`t`tChannel: " $TeamsChannels.DisplayName
			}
			else
			{
				Write-Host "`t`tChannel: " $TeamsChannels.DisplayName[$j]
			}
		}

		for ($j=0; $j -lt $TeamsChannels.length; $j++)
		{
			$row += 1
			if ( $TeamsChannels.length -eq 1 )
			{
				# Write-Host ([string]$row + " : " + $Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id + "~" + $TeamsChannels.DisplayName)
				$ExcelWorkSheet.Cells.Item($row,1) = ($Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id + "~" + $TeamsChannels.DisplayName)

			}
			else
			{
				# Write-Host ([string]$row + " : " + $Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id[$j] + "~" + $TeamsChannels.DisplayName[$j])
				$ExcelWorkSheet.Cells.Item($row,1) = ($Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id[$j] + "~" + $TeamsChannels.DisplayName[$j])
			}
		}
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($row + 1) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Teams and Channels skipped."	
}	

#
# Users
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Users")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoUsers )
{
	Write-Host "Getting list of enterprise voice enabled users."

	$Users = @(get-csonlineuser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias)

	for ( $i = 0; $i -lt $Users.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}) User : {1,-50}" -f ($i + 1), $Users.UserPrincipalName[$i] )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($Users.UserPrincipalName[$i] + "~" + $Users.Identity[$i])
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Users skipped."	
}

#
# Save and close the Excel file
#
if ( $View )
{
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Config-CallQueue")
   $ExcelWorkSheet.activate()
}

$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)
$ExcelObj.Quit()

if ( Test-Path -Path ".\`$null" )
{
	Remove-Item -Path ".\`$null" | Out-Null
}


if ( ! $NoOpen )
{
	Write-Host "Preparation complete.  Opening $ExcelFilename.  "
	Write-Host "Please complete the configuration, save and exit the spreadsheet and then run the BulkCQsConfig script."
	Write-Host -NoNewLine "Press any key to continue..."
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

	Invoke-item "$ExcelFullPathFilename"
}
else
{
	Write-Host "Preparation complete."
}

