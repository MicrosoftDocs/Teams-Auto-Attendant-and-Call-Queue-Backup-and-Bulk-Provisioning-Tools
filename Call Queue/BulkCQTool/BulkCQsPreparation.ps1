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
		
	$content = (Export-CsOnlineAudioFile -ApplicationId "HuntGroup" -Identity $fileID)
	[System.IO.File]::WriteAllBytes($currentDIR, $content)	
		
		
    #$content = [System.IO.File]::ReadAllBytes($currentDIR) 
    #$audioFileID = (Import-CsOnlineAudioFile -ApplicationID HuntGroup -FileName $fileName -Content $content).ID

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
			Default      			{ Write-Warning "Unknown argument passed: $args[$i]" }
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
	}
}

if ( $Help )
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
Write-Host "Cleaning up from any previous runs."

if ( Test-Path -Path ".\List-Resource-Accounts.csv" )
{
   Remove-Item -Path ".\List-Resource-Accounts.csv" | Out-Null
}

if ( Test-Path -Path ".\List-AAs.csv" )
{
   Remove-Item -Path ".\List-AAs.csv" | Out-Null
}

if ( Test-Path -Path ".\List-CQs.csv" )
{
   Remove-Item -Path ".\List-CQs.csv" | Out-Null
}

if ( Test-Path -Path ".\List-Teams.csv" )
{
   Remove-Item -Path ".\List-Teams.csv" | Out-Null
}

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Check that required modules are installed - install if not
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
      # Get-CsTenant -ErrorAction Stop 2>&1> $null
      Get-CsTenant -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      # Get-CsTenant -ErrorAction Stop 2>&1> $null
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
	# get-csonlineapplicationinstance | Sort-Object ApplicationId, DisplayName | Export-csv -Path .\List-Resource-Accounts.csv
	# $ResourceAccounts = @(Import-csv -Path .\List-Resource-Accounts.csv)

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
			if ( $Verbose )
			{
				Write-Host "`tResource Account : " $ResourceAccounts.DisplayName[$i]
			}

			if ( $ResourceAccountUserDetails.UsageLocation.length -eq 0 )
			{
				$ResourceAccountUserDetails.UsageLocation = "US"
			}

			# request will generate a "Correlation id for this request" message when the RA is not assigned to anything, also generates error so redirecting that to null
			$ResourceAccountPriority = ( (get-csonlineapplicationinstanceassociation -identity $ResourceAccounts.ObjectId[$i]).CallPriority 2> $null  )
			
			if ( $ResourceAccounts.ApplicationId[$i] -eq "ce933385-9390-45d1-9512-c8d228074e07" )
			{
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-AA] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" + $ResourceAccounts.PhoneNumber[$i] + "~" + $ResourceAccountUserDetails.UsageLocation + "~" + $ResourceAccountPriority )
			}
			else
			{
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

if ( Test-Path -Path ".\List-Resource-Accounts.csv" )
{
   Remove-Item -Path ".\List-Resource-Accounts.csv" | Out-Null
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
		Get-CsAutoAttendant -Skip $j 3> $null | Export-csv -Path .\List-AA-$j.csv
	}

	$command = "Get-Content "
	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100
		$command += "List-AA-$j.csv, "
	}
	$command = $command.Substring(0, $command.length -2)
	$command += " | Out-File .\List-AAs.csv"

	Invoke-Expression $command


	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100
		Remove-Item -Path ".\List-AA-$j.csv" | Out-Null
	}

	$AutoAttendants = @(Import-csv -Path .\List-AAs.csv)

	for ($i=0; $i -lt  $AutoAttendants.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host "`tAuto Attendant : " $AutoAttendants.Name[$i]
		}

		#$AssignedResourceAccounts = (Get-CsAutoAttendant -Identity $AutoAttendants.Identity[$i]).ApplicationInstances
		#$ExcelWorkSheet.Cells.Item($i + 2,1) = ($AutoAttendants.Name[$i] + "~" + $AutoAttendants.Identity[$i] + "~" + ($AssignedResourceAccounts -join ","))
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($AutoAttendants.Name[$i] + "~" + $AutoAttendants.Identity[$i])
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Auto Attendants skipped."	
}

if ( Test-Path -Path ".\List-AAs.csv" )
{
   Remove-Item -Path ".\List-AAs.csv" | Out-Null
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
		Get-CsCallQueue -Skip $j 3> $null | Export-csv -Path .\List-CQ-$j.csv
	}

	$command = "Get-Content "
	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100
		$command += "List-CQ-$j.csv, "
	}
	$command = $command.Substring(0, $command.length -2)
	$command += " | Out-File .\List-CQs.csv"

	Invoke-Expression $command

	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100
		Remove-Item ".\List-CQ-$j.csv" | Out-Null
	}

	$CallQueues = @(Import-csv -Path .\List-CQs.csv)

	if ( $Download )
	{
		$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")
		
		if ( $View )
		{
			$ExcelWorkSheet.activate()
		}
		
		#
		# Blank out any existing data
		#
		$Range = "A3:ZZ2002"
		$ExcelWorkSheet.Range($Range).value = ""
		
		$RowOffset = 3
		
		for ($i=0; $i -lt  $CallQueues.length; $i++)
		{
			$ColumnOffset = 1
			
			$CQ = (Get-CsCallQueue -Identity $CallQueues.Identity[$i] 3> $null)

			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.Name							# A
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.Identity						# B

			if ( $Verbose )
			{
				Write-Host "`tDownloading Call Queue : " $CQ.Name
			}
			
			switch ( $CQ.RoutingMethod )																	# C
			{
				"Attendant"		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Attendant routing"}
				"Serial" 		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Serial routing"}
				"RoundRobin" 	{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Round robin (Default)"}
				"LongestIdle" 	{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Longest idle"}
				default  		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.RoutingMethod }
			}
			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.AllowOptOut						# D
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.ConferenceMode					# E
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.PresenceBasedRouting				# F
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.AgentAlertTime					# G
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.LanguageId						# H
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicFileName				# I --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicResourceId 			# **** LETTERS ARE WRONG FROM HERE DOWN

			if ( $CQ.WelcomeMusicResourceId.length -gt 0 )
			{
				AudioFileExport $CQ.WelcomeMusicFileName $CQ.WelcomeMusicResourceId $CQ.Identity
			}
			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.WelcomeTextToSpeechPrompt		# J
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.UseDefaultMusicOnHold			# K
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldFileName				# L --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldResourceId

			if ( $CQ.MusicOnHoldResourceId.length -gt 0 )
			{
				AudioFileExport $CQ.MusicOnHoldFileName $CQ.MusicOnHoldResourceId $CQ.Identity
			}

			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.ServiceLevelThresholdResponseTimeInSecond	# M

			# Overflow
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowThreshold										# N
			
			switch ( $CQ.OverflowAction	)																		# O 
			{
				"DisconnectWithBusy" 	{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Disconnect (Default)" }
				"Forward" 				{ 	switch ( $CQ.OverflowActionTarget.Type )
											{
												"User"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
												"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
												"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
												"Phone"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
											}												
										} 
				"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
				"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
				default  				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowAction }
			}
			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowActionTarget.Id									# P
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowActionCallPriority								# Q
			
			# Overflow Shared Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailTextToSpeechPrompt				# R
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePrompt					# S --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePromptFileName			# T
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailTranscription				# U
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailSystemPromptSuppression		# V

			if ( $CQ.OverflowSharedVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowSharedVoicemailAudioFilePromptFileName $CQ.OverflowSharedVoicemailAudioFilePrompt $CQ.Identity
			}
			
			# Overflow Disconnect
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePrompt						# W --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePromptFileName				# X
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectTextToSpeechPrompt						# Y
			
			if ( $CQ.OverflowDisconnectAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowDisconnectAudioFilePromptFileName $CQ.OverflowDisconnectAudioFilePrompt $CQ.Identity
			}
			
			# Overflow Redirect - Person
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePrompt					# Z --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePromptFileName			# AA
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonTextToSpeechPrompt				# AB
			
			if ( $CQ.OverflowRedirectPersonAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowRedirectPersonAudioFilePromptFileName $CQ.OverflowRedirectPersonAudioFilePrompt $CQ.Identity
			}
			
			# Overflow Redirect - Voice App
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePrompt				# AC --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName		# AD
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppTextToSpeechPrompt				# AE
			
			if ( $CQ.OverflowRedirectVoiceAppAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName $CQ.OverflowRedirectVoiceAppAudioFilePrompt $CQ.Identity
			}
			
			# Overflow Redirect - Phone Number
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePrompt				# AF --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName		# AG
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberTextToSpeechPrompt			# AH
			
			if ( $CQ.OverflowRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName $CQ.OverflowRedirectPhoneNumberAudioFilePrompt $CQ.Identity
			}
			
			# Overflow Redirect - Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePrompt				# AI --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePromptFileName		# AJ
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailTextToSpeechPrompt			# AK

			if ( $CQ.OverflowRedirectVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.OverflowRedirectVoicemailAudioFilePromptFileName $CQ.OverflowRedirectVoicemailAudioFilePrompt $CQ.Identity
			}

			
			# Timeout
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeOutThreshold										# AL
			
			switch ( $CQ.TimeoutAction	)																				# AM 
			{
				"Disconnect" 	{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Disconnect (Default)" }
				"Forward" 		{ 	switch ( $CQ.TimeoutActionTarget.Type )
									{
										"User"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
										"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
										"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
										"Phone"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
									}												
								} 
				"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
				"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
				default  				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutAction }
			}
			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutActionTarget.Id									# AN
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutActionCallPriority								# AO
			
			# Timeout Shared Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailTextToSpeechPrompt				# AP
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePrompt					# AQ --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePromptFileName			# AR
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailTranscription				# AS
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailSystemPromptSuppression	# AT
			
			if ( $CQ.TimeoutSharedVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutSharedVoicemailAudioFilePromptFileName $CQ.TimeoutSharedVoicemailTextToSpeechPrompt $CQ.Identity
			}
			
			# Timeout Disconnect
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePrompt						# AU --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePromptFileName				# AV
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectTextToSpeechPrompt					# AW
			
			if ( $CQ.TimeoutDisconnectAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutDisconnectAudioFilePromptFileName $CQ.TimeoutDisconnectAudioFilePrompt $CQ.Identity
			}
			
			# Timeout Redirect - Person
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePrompt					# AX --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePromptFileName			# AY
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonTextToSpeechPrompt				# AZ
			
			if ( $CQ.TimeoutRedirectPersonAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutRedirectPersonAudioFilePromptFileName $CQ.TimeoutRedirectPersonAudioFilePrompt $CQ.Identity
			}
			
			# Timeout Redirect - Voice App
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePrompt					# BA --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName			# BB
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppTextToSpeechPrompt				# BC
			
			if ( $CQ.TimeoutRedirectVoiceAppAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName $CQ.TimeoutRedirectVoiceAppAudioFilePrompt $CQ.Identity
			}
			
			# Timeout Redirect - Phone Number
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt				# BD --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName		# BE
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt			# BF
			
			if ( $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt $CQ.Identity
			}
			
			# Timeout Redirect - Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePrompt				# BG --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName		# BH
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailTextToSpeechPrompt				# BI

			if ( $CQ.TimeoutRedirectVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName $CQ.TimeoutRedirectVoicemailAudioFilePrompt $CQ.Identity
			}
			
			
			# NoAgent
			switch ( $CQ.NoAgentAction	)																				# BJ 
			{
				"Queue" 				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Queue call (Default)" }
				"Disconnect" 			{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Disconnect" }
				"Forward" 				{ 	switch ( $CQ.NoAgentActionTarget.Type )
											{
												"User"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Person in organization" }  
												"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
												"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voice app" }  
												"Phone"						{ $ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - External phone number" }
											}
										}
				"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail Personal" }
				"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Redirect - Voicemail (shared)" }
				default  				{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.TimeoutAction }
			}
			
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentActionTarget.Id									# BK
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentActionCallPriority								# BL
			
			# No Agent Shared Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailTextToSpeechPrompt				# BM
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePrompt					# BN --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePromptFileName			# BO
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailTranscription				# BP
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailSystemPromptSuppression	# BQ
			
			if ( $CQ.NoAgentSharedVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentSharedVoicemailAudioFilePromptFileName $CQ.NoAgentSharedVoicemailAudioFilePrompt $CQ.Identity
			}
			
			switch ( $CQ.NoAgentApplyTo	)																				# BR
			{
				"NewCalls" 	{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "New Calls" }
				"AllCalls" 	{ 	$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "All calls (Default)" }
			}

			# No Agent Disconnect
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePrompt						# BS --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePromptFileName				# BT
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectTextToSpeechPrompt					# BU

			if ( $CQ.NoAgentDisconnectAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentDisconnectAudioFilePromptFileName $CQ.NoAgentDisconnectAudioFilePrompt $CQ.Identity
			}
			
			# No Agent Redirect - Person
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePrompt					# BV --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePromptFileName			# BW
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonTextToSpeechPrompt				# BX
			
			if ( $CQ.NoAgentRedirectPersonAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentRedirectPersonAudioFilePromptFileName $CQ.NoAgentRedirectPersonAudioFilePrompt $CQ.Identity
			}
			
			# No Agent Redirect - Voice App
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePrompt					# BY --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName			# BZ
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppTextToSpeechPrompt				# CA

			if ( $CQ.NoAgentRedirectVoiceAppAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName $CQ.NoAgentRedirectVoiceAppAudioFilePrompt $CQ.Identity
			}
			
			# No Agent Redirect - Phone Number
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt				# CB --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName		# CC
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt			# CD
			
			if ( $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt $CQ.Identity
			}
			
			# No Agent Redirect - Voicemail
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePrompt				# CE --- need to download
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName		# CF
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailTextToSpeechPrompt				# CG

			if ( $CQ.NoAgentRedirectVoicemailAudioFilePrompt.length -gt 0 )
			{
				AudioFileExport $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName $CQ.NoAgentRedirectVoicemailAudioFilePrompt $CQ.Identity
			}


			# Callback
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.IsCallbackEnabled										# CH
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallbackRequestDtmf									# CI
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.WaitTimeBeforeOfferingCallbackInSecond					# CJ
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.NumberOfCallsInQueueBeforeOfferingCallback				# CK
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallToAgentRatioThresholdBeforeOfferingCallback		# CL
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptResourceId					# CM
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptFileName					# CN
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallbackOfferTextToSpeechPrompt						# CO
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.CallbackEmailNotificationTarget.Id						# CP

			if ( $CQ.CallbackOfferAudioFilePromptResourceId.length -gt 0 )
			{
				AudioFileExport $CQ.CallbackOfferAudioFilePromptFileName $CQ.CallbackOfferAudioFilePromptResourceId $CQ.Identity
			}

			
			# Agents
			if ( $CQ.Agents.length -gt 0 )
			{
				for ($j=0; $j -lt $CQ.Agents.length; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset + $j) = $CQ.Agents[$j].ObjectId						# CQ - DJ (20) / 95-114
				}
			}
			
			# Channel
			$ColumnOffset += 20
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.ChannelId												# DK
			

			# Authorized Users
			if ( $CQ.AuthorizedUsers.length -gt 0 )
			{
				for ($j=0; $j -lt $CQ.AuthorizedUsers.length; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset + $j) = $CQ.AuthorizedUsers[$j].Guid					# DL - DZ (15) / 116-130
				}
			}
			
			# Hidden Authorized Users
			$ColumnOffset += 15
			if ( $CQ.HideAuthorizedUsers.length -gt 0 )
			{
				for ($j=0; $j -lt $CQ.HideAuthorizedUsers.length; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset + $j) = $CQ.HideAuthorizedUsers[$j].Guid				# EA - EO (15) / 131-145
				}
			}


			# Numbers in lists
			$ColumnOffset += 15
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.DistributionLists.length								# EP
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.ApplicationInstances.length							# EQ
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = $CQ.OboResourceAccounts.length							# ER
			
			# Spare Numbers in Lists
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 04"														# ES
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 05"														# ET
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 06"														# EU
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 07"														# EV
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 08"														# EW
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 09"														# EX
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 10"														# EY
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 11"														# EZ
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 12"														# FA
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 13"														# FB
			$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset++) = "Spare 14"														# FC
			
			# Distributions Lists
			if ( $CQ.DistributionLists.length -gt 0 )
			{
				for ($j=0; $j -lt 4; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset, $ColumnOffset + $j) = $CQ.DistributionLists[$j].Guid				# FD+ / 160+
				}
			}
			
			# Resource Accounts
			$ColumnOffset += $CQ.DistributionLists.length
			
			if ( $CQ.ApplicationInstances.length -gt 0 )
			{
				for ($j=0; $j -lt $CQ.ApplicationInstances.length; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset,$ColumnOffset + $j) = $CQ.ApplicationInstances[$j]		# FD+ / 160+
				}
			}
			
			# Obo Resource Accounts
			$ColumnOffset += $CQ.ApplicationInstances.length

			if ( $CQ.OboResourceAccounts.length -gt 0 )
			{
				for ($j=0; $j -lt $CQ.OboResourceAccounts.length; $j++)
				{
					$ExcelWorkSheet.Cells.Item($i + $RowOffset,$ColumnOffset + $j) = $CQ.OboResourceAccounts[$j].ObjectId	# FD+ / 160+
				}
			}
		}
		
		#
		# Blank out remaining rows
		#
		$Range = "A" + ($i + $RowOffset) + ":ZZ2002"
		$ExcelWorkSheet.Range($Range).value = ""
	}
	
	$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
	
	if ( $View )
	{
		$ExcelWorkSheet.activate()
	}

	for ($i=0; $i -lt  $CallQueues.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host "`tCall Queue : " $CallQueues.Name[$i]
		}

		$AssignedResourceAccounts = ( (Get-CsCallQueue -Identity $CallQueues.Identity[$i]).ApplicationInstances 3> $null )
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($CallQueues.Name[$i] + "~" + $CallQueues.Identity[$i] + "~" + ($AssignedResourceAccounts -join ","))
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + $RowOffset) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading Call Queues skipped."		
}

if ( Test-Path -Path ".\List-CQs.csv" )
{
	Remove-Item -Path ".\List-CQs.csv" | Out-Null
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
	#get-csphonenumberassignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned" | Export-csv -Path .\List-PhoneNumbers-VA.csv
	#$PhoneNumbers = @(Import-csv -Path .\List-PhoneNumbers-VA.csv)

	$PhoneNumbers = @(get-csphonenumberassignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned")

	for ($i=0; $i -lt  $PhoneNumbers.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host "`tPhone Number : " $PhoneNumbers.TelephoneNumber[$i]
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

if ( Test-Path -Path ".\List-PhoneNumbers-VA.csv" )
{
   Remove-Item -Path ".\List-PhoneNumbers-VA.csv" | Out-Null
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
	# get-team | Sort-Object DisplayName | Export-csv -Path .\List-Teams.csv
	# $Teams = @(Import-csv -Path .\List-Teams.csv)

	$Teams = @(get-team | Sort-Object DisplayName)

	$row = 1
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		# get-teamchannel -groupId $Teams.GroupId[$i] | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName | Export-csv -Path .\List-Team-Channel.$i.csv
		# $TeamsChannels = @(Import-csv -Path .\List-Team-Channel.$i.csv)

		$TeamsChannels = @(get-teamchannel -groupId $Teams.GroupId[$i] | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName)

		if ( $Verbose )
		{
			Write-Host "`tProcessing Team: " $Teams.DisplayName[$i]

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

		if ( Test-Path -Path ".\List-Team-Channel.$i.csv" )
		{
			Remove-Item -Path ".\List-Team-Channel.$i.csv" | Out-Null
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

if ( Test-Path -Path ".\List-Teams.csv" )
{
   Remove-Item -Path ".\List-Teams.csv" | Out-Null
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
	# get-csonlineuser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias | Export-csv -Path .\List-EV-Users.csv
	# $Users = @(Import-csv -Path .\List-EV-Users.csv)

	$Users = @(get-csonlineuser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias)

	for ( $i = 0; $i -lt $Users.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host "`tUser : " $Users.UserPrincipalName[$i]
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

if ( Test-Path -Path ".\List-EV-Users.csv" )
{
   Remove-Item -Path ".\List-EV-Users.csv" | Out-Null
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

