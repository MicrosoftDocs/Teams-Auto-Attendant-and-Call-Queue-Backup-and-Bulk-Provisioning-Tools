# Version: 1.0.5
# Date: 2025.04.28

# Changelog: https://github.com/MicrosoftDocs/Teams-Auto-Attendant-and-Call-Queue-Backup-and-Bulk-Provisioning-Tools/blob/main/Call%20Queue/CHANGELOG.md

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


###########################
#  AudioFileExport
###########################
function AudioFileExport
{
	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[string] $fileType,
		[Parameter(Mandatory=$true, Position=1)]
		[string] $fileName,
		[Parameter(Mandatory=$true, Position=2)]
		[string] $fileID,
		[Parameter(Mandatory=$true, Position=3)]
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

	Write-Host "`t`t`tDownloading $fileType filename: " $fileName
	$content = (Export-CsOnlineAudioFile -ApplicationId "HuntGroup" -Identity $fileID 2> $null)
	if ( $content.length -ne 0 )
	{
		[System.IO.File]::WriteAllBytes($currentDIR, $content)
		return $true
	}
	else
	{
		return $false
	}
}

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

#
# setting default 
#
$AACount = [int]0
$CQCount = [int]0

if ( $args -ne "" )
{
	for ( $i = 0; $i -lt $args.length; $i++ )
	{
		switch ( $args[$i] )
		{
			"-aacount"   				{ 	if ( ( $args[$i+1] -eq $null ) -or ( !($args[$i+1] -match "^[\d\.]+$" ) ) )
											{
												$NoAutoAttendants = $true
												$AACount = [int]0
												$i++
											}
											else
											{
												$AACount = [int]$args[$i+1]
												$i++
											}
										}
			"-cqcount"   				{ 	if ( ( $args[$i+1] -eq $null ) -or ( !($args[$i+1] -match "^[\d\.]+$" ) ) )
											{
												$NoCallQueues = $true
												$CQCount = [int]0
												$i++
											}
											else
											{
												$CQCount = [int]$args[$i+1]
												$i++
											}
										}
			"-download"					{ 	$Download = $true }
			"-excelfile" 				{ 	$ExcelFilename = $args[$i+1]
											$i++
										}
			"-help"   					{ 	$Help = $true }
			"-noautoattendants"			{ 	$NoAutoAttendants = $true }
			"-nocallqueues"				{ 	$NoCallQueues = $true }
			"-nocr4cqtemplates"			{	$NoCR4CQTemplates = $true }
			"-nophonenumbers"			{ 	$NoPhoneNumbers = $true }
			"-noresourceaccounts"		{ 	$NoResourceAccounts = $true }
			"-noteamschannels"			{ 	$NoTeamsChannels = $true }
			"-noteamsschedulegroups" 	{ 	$NoTeamsScheduleGroups = $true }
			"-nousers"					{ 	$NoUsers = $true }
			"-noopen"               	{ 	$NoOpen = $true}
			"-verbose"   				{ 	$Verbose = $true }
			"-view"      				{ 	$View = $true }	  
			Default      				{ 	$ArgError = $true
											$arg = $args[$i]
										}
		}
	}
	
	if ( $Download )
	{
#		$NoAutoAttendants = $false
		$NoCallQueues = $false
#		$NoPhoneNumbers = $false
#		$NoResourceAccounts = $false
#		$NoTeamsChannels = $false
#		$NoTeamsScheduleGroups = $false
#		$NoUsers = $false
#		$NoCR4CQTemplates = $false
		$Verbose = $true
	}
	
	if ( (! $NoAutoAttendants) -and $AACount -eq 0 )
	{
		$AACount = 100
	}
	
	if ( (! $NoCallQueues) -and $CQCount -eq 0 )
	{
		$CQCount = 100
	}
}

if ( $ArgError )
{
	Write-Host "An unknown argument was encountered: $arg" -f Red
}

if ( $AACount -gt 0 -and $NoAutoAttendants )
{
	Write-Host "Conflicting parameters. -NoAutoAttendants specified but -AACount supplied. Processing has been halted." -f Red
	$ArgError = $true
}

if ( $CQCount -gt 0 -and $NoCallQueues )
{
	Write-Host "Conflicting parameters. -NoCallQueues specified but -CQCount supplied. Processing has been halted." -f Red
	$ArgError = $true
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
	Write-Host "`t-NoCR4CQTemplates - do not download existing compliance recording templates"
	Write-Host "`t-NoPhoneNumbers - do not download Voice Apps phone numbers"
	Write-Host "`t-NoResourceAccounts - do not download existing resource account information"
	Write-Host "`t-NoTeamsChannels - do not download existing Teams channels information"
	Write-Host "`t-NoTeamsScheduleGroups - do not download existing Teams schedule group information (removes requirement for Graph access)"
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
# Module Min Supported Versions
#
$MicrosoftTeamsMinVersion = [version]"7.0.0"
$MicrosoftGraphMinVersion = [version]"2.24.0"
$ImportExcelMinVersion = [version]"7.8.0"

#
# Set range variables
#
$Range_ResourceAccounts = "A2:A2001"
$Range_AutoAttendants = "A2:A2001"
$Range_CallQueues = "A2:A2001"
$Range_CallQueueDownload = "A3:ZZ2002"
$Range_PhoneNumbers = "A2:A2001"
$Range_TeamsChannels = "A2:A2001"
$Range_TeamsScheduleGroups = "A2:A2001"
$Range_Users = "A2:A2001"
$Range_CR4CQ = "A2:A51"

#
# Check that minimum verion of required modules are installed - install if not
#
# MicrosoftTeams
#
Write-Host "Checking for MicrosoftTeams module $MicrosoftTeamsMinVersion or later."
$Version = ( (Get-InstalledModule -Name MicrosoftTeams -MinimumVersion "$MicrosoftTeamsMinVersion").Version 2> $null )

if ( $Version -match "-preview" )
{
	$Version = $Version.Replace("-preview", "")
}

$Version = [version]$Version

Write-Host "`tMicrosoftTeams Version: $Version"
if ( $Version -ge $MicrosoftTeamsMinVersion )
{
   Write-Host "`tConnecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion

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
      Write-Error "`tNot signed into Microsoft Teams!" 
      exit
   }
   Write-Host "`tConnected to Microsoft Teams."
}
else
{
   Write-Host "`tThe MicrosoftTeams module is not installed or does not meet the minimum requirements - installing."
   Install-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion -Force -AllowClobber

   Write-Host "`tConnecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion
   
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
      Write-Error "`tNot signed into Microsoft Teams!" 
      exit
   }
   Write-Host "`tConnected to Microsoft Teams."
}

if ( ! $NoTeamsScheduleGroups )
{
	Write-Host "Checking for Microsoft.Graph module $MicrosoftGraphMinVersion or later."
	$Version = ( (Get-InstalledModule -Name Microsoft.Graph -MinimumVersion "$MicrosoftGraphMinVersion").Version 2> $null)
	
	if ( $Version -match "-preview" )
	{
		$Version = $Version.Replace("-preview", "")
	}

	$Version = [version]$Version

	Write-Host "`tMicrosoft.Graph Version: $Version"
	if ( $Version -ge $MicrosoftGraphMinVersion )
	{
		Write-Host "`tConnecting to Microsoft Graph."
   
		Connect-MgGraph -Scopes "Schedule.Read.All" -NoWelcome | Out-Null
		
		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Connect-MgGraph -Scopes "Schedule.Read.All" -NoWelcome | Out-Null
		}
		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Write-Error "`tNot signed into Microsoft Graph!" 
			exit
		}
		Write-Host "`tConnected to Microsoft Graph."
	}
	else
	{
		Write-Host "`tThe Microsoft.Graph module is not installed or does not meet the minimum requirements - installing."
		Install-Module -Name Microsoft.Graph -MinimumVersion $MicrosoftGraphMinVersion -Force -AllowClobber

		Connect-MgGraph -Scopes "Schedue.Read.All" -NoWelcome | Out-Null

		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
		}
		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Write-Error "`tNot signed into Microsoft Graph!" 
			exit
		}
		Write-Host "`tConnected to Microsoft Graph."
	}
}

#
# ImportExcel
#
Write-Host "Checking for ImportExcel module $ImportExcelMinVersion or later."
$Version = ( (Get-InstalledModule -Name ImportExcel -MinimumVersion "$ImportExcelMinVersion").Version 2> $null )

if ( $Version -match "-preview" )
{
	$Version = $Version.Replace("-preview", "")
}

$Version = [version]$Version

Write-Host "`tImportExcel Version: $Version"
if ( $Version -ge $ImportExcelMinVersion )
{
   Write-Host "`tImporting ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
}
else
{
   Write-Host "`tThe ImportExcel module is not installed or does not meet the minimum requirements - installing."
   Install-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion -Force -AllowClobber
   
   Write-Host "`tImporting ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
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

#
# Check if file is already open
#
$ExcelFileOpen =(Get-CimInstance Win32_Process -Filter "CommandLine like '%$ExcelFileName%'").ProcessId

if ( $ExcelFileOpen.length -eq 0 )
{
	$ExcelObj = New-Object -comobject Excel.Application
	$ExcelWorkBook = $ExcelObj.Workbooks.Open($ExcelFullPathFilename)
}
else
{
	Write-Host "The $ExcelFileName appears to be open already. Please close the file and try again." -f Red
	exit
}

#
# Turn off AutoSave & Calculation
#
$AutoSave = $ExcelWorkBook.AutoSaveOn
$AutoCalc = $ExcelWorkBook.Parent.Calculation
$ExcelWorkBook.AutoSaveOn = $false
$ExcelWorkBook.Parent.Calculation = -4135

if ( $View )
{
	$ExcelObj.Visible = $true
}

#
# Resource Accounts
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("ResourceAccountsAll")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoResourceAccounts )
{
	#
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_ResourceAccounts).Value = ""

	Write-Host "Getting list of Resource Accounts."

	$ResourceAccounts = @(Get-CsOnlineApplicationInstance | Sort-Object ApplicationId, DisplayName)

	$j = 2
	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		#
		# Make sure resource account is not Deleted
		#
		$ResourceAccountUserDetails = (Get-CsOnlineUser -Identity $ResourceAccounts[$i].ObjectId)
		
		if ( $ResourceAccountUserDetails.SoftDeletionTimeStamp.length -eq 0 )
		{
			if ( $ResourceAccountUserDetails.UsageLocation.length -eq 0 )
			{
				$ResourceAccountUserDetails.UsageLocation = "US"
			}

			if ( $ResourceAccounts[$i].ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07" )
			{
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}/{1,4}) [RA-AA] Resource Account : {2,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].DisplayName )
				}

				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-AA] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" + $ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails.UsageLocation )

			}
			else
			{
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}/{1,4}) [RA-CQ] Resource Account : {2,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].DisplayName )
				}

				# request will generate a "Correlation id for this request" message when the RA is not assigned to anything, also generates error so redirecting that to null
				$ResourceAccountPriority = ( (Get-CsOnlineApplicationInstanceAssociation -Identity $ResourceAccounts[$i].ObjectId).CallPriority 2> $null  )
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-CQ] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" +$ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails.UsageLocation  + "~" + $ResourceAccountPriority )
			}
			$j++
		}
		else
		{
			if ( $Verbose )
			{
				Write-Host "`tResource Account Not Added (Soft Deleted): " $ResourceAccounts[$i].DisplayName
			}
		}
	}
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
   $ExcelWorkSheet.Activate()
}

if ( ( ! $NoAutoAttendants ) -and ( $AACount -ne 0 ) )
{
	#
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_AutoAttendants).Value = ""

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
		
		if ( $AACount -le 100 )
		{
			$First = $AACount
		}
		elseif ( ( $AACount - $j ) -le 100 )
		{
			$First =  $AACount - $j
		}
		else
		{
			$First = 100
		}
		
		if ( $Verbose )
		{
			if ( $First -eq 100 )
			{
				Write-Host "`tRetrieving list of auto attendants $($j+1) to $($j+100) of $AACount"
			}
			else
			{
				Write-Host "`tRetrieving list of auto attendants $($j+1) to $($j+$First) of $AACount"
			}
		}
		
		$AutoAttendants += @(Get-CsAutoAttendant -Skip $j -First $First)
	}

	$RowOffset = 1
	for ($i=0; $i -lt $AutoAttendants.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Auto Attendant : {2,-50}" -f ($i + 1), $AACount, $AutoAttendants[$i].Name )
		}

		$RowOffset += 1
		$ExcelWorkSheet.Cells.Item($RowOffset, 1) = ($AutoAttendants[$i].Name + "~" + $AutoAttendants[$i].Identity)
	}
}
else
{
	Write-Host "Downloading Auto Attendants skipped."	
}

#
# Call Queues
#
if ( ( ! $NoCallQueues ) -and ( $CQCount -ne 0 ) )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
		
	if ( $View )
	{
		$ExcelWorkSheet.Activate()
	}
		
	$ExcelWorkSheet.Range($Range_CallQueues).Value = ""

	if ( $Download )
	{
		$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")

		if ( $View )
		{
			$ExcelWorkSheet.Activate()
		}
			
		$ExcelWorkSheet.Range($Range_CallQueueDownload).Value = ""
	}

	if ( $CQCount -gt 0 )
	{
		$loops = [int] [Math]::Truncate($CQCount / 100) + 1
	}
	else
	{
		$loops = 1
	}

	Write-Host "Getting list of Call Queues."
	
	$CallQueues = @()
	for ( $i = 0; $i -lt $loops; $i++ )
	{
		$j = $i * 100

		if ( $CQCount -le 100 )
		{
			$First = $CQCount
		}
		elseif ( ( $CQCount - $j ) -le 100 )
		{
			$First =  $CQCount - $j
		}
		else
		{
			$First = 100
		}
		
		if ( $Verbose )
		{
			if ( $First -eq 100 )
			{
				Write-Host "`tRetrieving list of call queues $($j+1) to $($j+100) of $CQCount"
			}
			else
			{
				Write-Host "`tRetrieving list of call queues $($j+1) to $($j+$First) of $CQCount"
			}
		}

		$CallQueues += @(Get-CsCallQueue -Skip $j -First $First 3> $null )
	}
	
	if ( $Download )
	{
		$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")

		if ( $View )
		{
			$ExcelWorkSheet.Activate()
		}

		Write-Host "Processing list of Call Queues for Download."		

		$RowOffset = 2
		for ( $i = 0; $i -lt $CallQueues.length; $i++ )
		{
				$RowOffset += 1
				$ColumnOffset = 1
				
				$CQ = $CallQueues[$i]
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Identity
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Name
				
				if ( $Verbose )
				{
					Write-Host ( "`t({0,4}/{1,4}) Processing Call Queue : {2,-50}" -f ($i + 1), $CQCount, $CQ.Name )
				}
			
				switch ( $CQ.RoutingMethod )
				{
					"Attendant"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Attendant routing"}
					"Serial" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Serial routing"}
					"RoundRobin" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Round robin (Default)"}
					"LongestIdle" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Longest idle"}
					default  		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.RoutingMethod }
				}

				if ( $CQ.AllowOptOut.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AllowOptOut
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.ConferenceMode.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ConferenceMode
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.PresenceBasedRouting.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.PresenceBasedRouting
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.AgentAlertTime.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AgentAlertTime
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.LanguageId.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.LanguageId
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( ( $CQ.WelcomeMusicFileName.length -ne 0 ) -and ( $CQ.WelcomeMusicResourceId.length -ne 0 ) )
				{
					if ( AudioFileExport "WelcomeMusic" $CQ.WelcomeMusicFileName $CQ.WelcomeMusicResourceId $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicFileName
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicResourceId
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.WelcomeMusicFileName
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.WelcomeMusicResourceId
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.WelcomeTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
				
				if ( $CQ.UseDefaultMusicOnHold.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.UseDefaultMusicOnHold
				}
				else
				{
					$ColumnOffset++
				}

				if ( ( $CQ.MusicOnHoldFileName.length -ne 0 ) -and ( $CQ.MusicOnHoldResourceId.length -ne 0 ) )
				{
					if ( AudioFileExport "MusicOnHold" $CQ.MusicOnHoldFileName $CQ.MusicOnHoldResourceId $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldFileName
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldResourceId
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.MusicOnHoldFileName
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.MusicOnHoldResourceId
					}
				}
				else
				{
					 $ColumnOffset += 2
				}

				if ( $CQ.ServiceLevelThresholdResponseTimeInSecond.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ServiceLevelThresholdResponseTimeInSecond
				}
				else
				{
					 $ColumnOffset++
				}

				#
				# Overflow
				#
				if ( $CQ.OverflowThreshold.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowThreshold
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.OverflowAction.length -ne 0 )
				{
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
				}
				else
				{
					$ColumnOffset++
				}

				if ( $CQ.OverflowActionTarget.Id.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowActionTarget.Id
				}
				else
				{
					 $ColumnOffset++
				}

				if ( $CQ.OverflowActionCallPriority.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowActionCallPriority
				}
				else
				{
					 $ColumnOffset++
				}
				
				# Overflow Shared Voicemail
				if ( $CQ.OverflowSharedVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
				
				if ( ( $CQ.OverflowSharedVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.OverflowSharedVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "OverflowSharedVoicemail" $CQ.OverflowSharedVoicemailAudioFilePromptFileName $CQ.OverflowSharedVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowSharedVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.EnableOverflowSharedVoicemailTranscription.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailTranscription
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.EnableOverflowSharedVoicemailSystemPromptSuppression.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableOverflowSharedVoicemailSystemPromptSuppression
				}
				else
				{
					$ColumnOffset++
				}


				# Overflow Disconnect
				if ( ( $CQ.OverflowDisconnectAudioFilePrompt.length -ne 0 ) -and ($CQ.OverflowDisconnectAudioFilePromptFileName.length -ne 0) )
				{
					if ( AudioFileExport "OverflowDisconnect" $CQ.OverflowDisconnectAudioFilePromptFileName $CQ.OverflowDisconnectAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowDisconnectAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.OverflowDisconnectTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
				
				
				# Overflow Redirect - Person
				if ( ( $CQ.OverflowRedirectPersonAudioFilePrompt.length -ne 0 ) -and ( $CQ.OverflowRedirectPersonAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "OverflowRedirectPerson" $CQ.OverflowRedirectPersonAudioFilePromptFileName $CQ.OverflowRedirectPersonAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectPersonAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.OverflowRedirectPersonTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
				

				# Overflow Redirect - Voice App
				if ( ( $CQ.OverflowRedirectVoiceAppAudioFilePrompt.length -ne 0 ) -and ( $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "OverflowRedirectVoiceApp" $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName $CQ.OverflowRedirectVoiceAppAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.OverflowRedirectVoiceAppTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}

				
				# Overflow Redirect - Phone Number
				if ( ( $CQ.OverflowRedirectPhoneNumberAudioFilePrompt.length -ne 0 ) -and ( $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "OverflowRedirectPhoneNumber" $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName $CQ.OverflowRedirectPhoneNumberAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.OverflowRedirectPhoneNumberTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
								
				
				# Overflow Redirect - Voicemail
				if ( ( $CQ.OverflowRedirectVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.OverflowRedirectVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "OverflowRedirectVoicemail" $CQ.OverflowRedirectVoicemailAudioFilePromptFileName $CQ.OverflowRedirectVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.OverflowRedirectVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.OverflowRedirectVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailTextToSpeechPrompt
				}
				else
				{
					 $ColumnOffset++
				}
				

				
				#
				# Timeout
				#
				if ( $CQ.TimeOutThreshold.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeOutThreshold
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.TimeoutAction.length -ne 0 )
				{
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
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.TimeoutActionTarget.Id.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutActionTarget.Id
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.TimeoutActionCallPriority.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutActionCallPriority
				}
				else
				{
					$ColumnOffset++
				}
				
				# Timeout Shared Voicemail
				if ( $CQ.TimeoutSharedVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( ( $CQ.TimeoutSharedVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutSharedVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutSharedVoicemail" $CQ.TimeoutSharedVoicemailAudioFilePromptFileName $CQ.TimeoutSharedVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutSharedVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.EnableTimeoutSharedVoicemailTranscription.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailTranscription
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.EnableTimeoutSharedVoicemailSystemPromptSuppression.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableTimeoutSharedVoicemailSystemPromptSuppression
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# Timeout Disconnect
				if ( ( $CQ.TimeoutDisconnectAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutDisconnectAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutDisconnect" $CQ.TimeoutDisconnectAudioFilePromptFileName $CQ.TimeoutDisconnectAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutDisconnectAudioFilePromptFileName
					}						
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.TimeoutDisconnectTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# Timeout Redirect - Person
				if ( ( $CQ.TimeoutRedirectPersonAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutRedirectPersonAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutRedirectPerson" $CQ.TimeoutRedirectPersonAudioFilePromptFileName $CQ.TimeoutRedirectPersonAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectPersonAudioFilePromptFileName
					}						
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.TimeoutRedirectPersonTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# Timeout Redirect - Voice App
				if ( ( $CQ.TimeoutRedirectVoiceAppAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutRedirectVoiceApp" $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName $CQ.TimeoutRedirectVoiceAppAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.TimeoutRedirectVoiceAppTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# Timeout Redirect - Phone Number
				if ( ( $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutRedirectPhoneNumber" $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# Timeout Redirect - Voicemail
				if ( ( $CQ.TimeoutRedirectVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "TimeoutRedirectVoicemail" $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName $CQ.TimeoutRedirectVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.TimeoutRedirectVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}

				
				
				#
				# NoAgent
				#
				if ( $CQ.NoAgentAction.length -ne 0 )
				{
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
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.NoAgentActionTarget.Id.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentActionTarget.Id
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.NoAgentActionCallPriority.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentActionCallPriority
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# No Agent Shared Voicemail
				if ( $CQ.NoAgentSharedVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( ( $CQ.NoAgentSharedVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentSharedVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentSharedVoicemail" $CQ.NoAgentSharedVoicemailAudioFilePromptFileName $CQ.NoAgentSharedVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentSharedVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentSharedVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.EnableNoAgentSharedVoicemailTranscription.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailTranscription
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.EnableNoAgentSharedVoicemailSystemPromptSuppression.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.EnableNoAgentSharedVoicemailSystemPromptSuppression
				}
				else
				{
					$ColumnOffset++
				}
								
				switch ( $CQ.NoAgentApplyTo	)
				{
					"NewCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "New Calls" }
					"AllCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "All calls (Default)" }
				}

				# No Agent Disconnect
				if ( ( $CQ.NoAgentDisconnectAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentDisconnectAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentDisconnect" $CQ.NoAgentDisconnectAudioFilePromptFileName $CQ.NoAgentDisconnectAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentDisconnectAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentDisconnectAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.NoAgentDisconnectTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}

				
				# No Agent Redirect - Person
				if ( ( $CQ.NoAgentRedirectPersonAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentRedirectPersonAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentRedirectPerson" $CQ.NoAgentRedirectPersonAudioFilePromptFileName $CQ.NoAgentRedirectPersonAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectPersonAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectPersonAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.NoAgentRedirectPersonTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# No Agent Redirect - Voice App
				if ( ( $CQ.NoAgentRedirectVoiceAppAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentRedirectVoiceApp" $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName $CQ.NoAgentRedirectVoiceAppAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectVoiceAppAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.NoAgentRedirectVoiceAppTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}

				
				# No Agent Redirect - Phone Number
				if ( ( $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentRedirectPhoneNumber" $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				
				# No Agent Redirect - Voicemail
				if ( ( $CQ.NoAgentRedirectVoicemailAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "NoAgentRedirectVoicemail" $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName $CQ.NoAgentRedirectVoicemailAudioFilePrompt $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectVoicemailAudioFilePrompt
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}

				if ( $CQ.NoAgentRedirectVoicemailTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}



				#
				# Callback
				#
				if ( $CQ.IsCallbackEnabled.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.IsCallbackEnabled
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.CallbackRequestDtmf.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackRequestDtmf
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.WaitTimeBeforeOfferingCallbackInSecond.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WaitTimeBeforeOfferingCallbackInSecond
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.NumberOfCallsInQueueBeforeOfferingCallback.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NumberOfCallsInQueueBeforeOfferingCallback
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.CallToAgentRatioThresholdBeforeOfferingCallback.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallToAgentRatioThresholdBeforeOfferingCallback
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( ( $CQ.CallbackOfferAudioFilePromptResourceId.length -ne 0 ) -and ( $CQ.CallbackOfferAudioFilePromptFileName.length -ne 0 ) )
				{
					if ( AudioFileExport "CallbackOffer" $CQ.CallbackOfferAudioFilePromptFileName $CQ.CallbackOfferAudioFilePromptResourceId $CQ.Identity )
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptResourceId
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptFileName
					}
					else
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.CallbackOfferAudioFilePromptResourceId
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.CallbackOfferAudioFilePromptFileName
					}
				}
				else
				{
					$ColumnOffset += 2
				}
				
				if ( $CQ.CallbackOfferTextToSpeechPrompt.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferTextToSpeechPrompt
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.CallbackEmailNotificationTarget.Id.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackEmailNotificationTarget.Id
				}
				else
				{
					$ColumnOffset++
				}

				
				#
				# Channel
				#
				if ( $CQ.ChannelId.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ChannelId
				}
				else
				{
					$ColumnOffset++
				}


				#
				# Schedule Group 
				#
				if ( $CQ.ShiftsTeamId.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ShiftsTeamId
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.ShiftsSchedulingGroupId.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ShiftsSchedulingGroupId
				}
				else
				{
					$ColumnOffset++
				}


				#
				# Compliance Recording For Call Queues
				#
				if ( $CQ.ComplianceRecordingForCallQueueTemplateId.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ComplianceRecordingForCallQueueTemplateId
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.CustomAudioFileAnnouncementForCR.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CustomAudioFileAnnouncementForCR
				}
				else
				{
					$ColumnOffset++
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "FILENAME01"
				
				if ( $CQ.TextAnnouncementForCR.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TextAnnouncementForCR
				}
				else
				{
					$ColumnOffset++
				}
				
				if ( $CQ.CustomAudioFileAnnouncementForCRFailure.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CustomAudioFileAnnouncementForCRFailure 
				}
				else
				{
					$ColumnOffset++
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "FILENAME02"
				
				if ( $CQ.TextAnnouncementForCRFailure.length -ne 0 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TextAnnouncementForCRFailure
				}
				else
				{
					$ColumnOffset++
				}
				
				
				#
				# Authorized Users (AuthorizedUsers is array)
				#
				if ( $CQ.AuthorizedUsers.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.AuthorizedUsers.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.AuthorizedUsers[$l].Guid
					}
				}
				$ColumnOffset += 15
				
				
				#
				# Hidden Authorized Users (HideAuthorizedUsers is array)
				#
				if ( $CQ.HideAuthorizedUsers.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.HideAuthorizedUsers.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.HideAuthorizedUsers[$l].Guid
					}
				}
				$ColumnOffset += 15
				

				# Numbers in lists
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.DistributionLists.length
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ApplicationInstances.length
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OboResourceAccounts.length
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Agents.length
				
				# Spare Numbers in Lists
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
				
				# Distributions Lists (DistributionLists is array)
				if ( $CQ.DistributionLists.length -gt 0 )
				{
					for ($l=0; $l -lt 4; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.DistributionLists[$l].Guid
					}
				}
				
				# Resource Accounts (ApplicationInstances is array)
				$ColumnOffset += $CQ.DistributionLists.length
				
				if ( $CQ.ApplicationInstances.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.ApplicationInstances.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset,$ColumnOffset + $l) = $CQ.ApplicationInstances[$l]
					}
				}
				
				# Obo Resource Accounts (OboResourceAccounts is array)
				$ColumnOffset += $CQ.ApplicationInstances.length

				if ( $CQ.OboResourceAccounts.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.OboResourceAccounts.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset,$ColumnOffset + $l) = $CQ.OboResourceAccounts[$l].ObjectId
					}
				}
				
				
				# Agents (Agents is array)
				$ColumnOffset += $CQ.OboResourceAccounts.length
				
				if ( $CQ.Agents.length -gt 0 )
				{
					for ($l=0; $l -lt $CQ.Agents.length; $l++)
					{
						$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset + $l) = $CQ.Agents[$l].ObjectId
					}
				}
		} # CQ Loop
	} # Download

	$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")
		
	if ( $View )
	{
		$ExcelWorkSheet.Activate()
	}

	Write-Host "Processing list of Call Queues."

	$RowOffset = 1
	for ($k=0; $k -lt  $CallQueues.length; $k++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Call Queue : {2,-50}" -f ($k + 1), $CQCount, $CallQueues[$k].Name )
		}

		$RowOffset += 1
		$AssignedResourceAccounts = ( (Get-CsCallQueue -Identity $CallQueues[$k].Identity).ApplicationInstances 3> $null )
		$ExcelWorkSheet.Cells.Item($RowOffset,1) = ($CallQueues[$k].Name + "~" + $CallQueues[$k].Identity + "~" + ($AssignedResourceAccounts -join ","))
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
   $ExcelWorkSheet.Activate()
}

if ( ! $NoPhoneNumbers )
{
	#
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_PhoneNumbers).Value = ""

	Write-Host "Getting list of existing voice application phone numbers"

	$PhoneNumbers = @(Get-CsPhoneNumberAssignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned")

	for ($i=0; $i -lt  $PhoneNumbers.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Phone Number : {2,-50}" -f ($i + 1), $PhoneNumbers.length, $PhoneNumbers[$i].TelephoneNumber )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($PhoneNumbers[$i].TelephoneNumber + "~" + $PhoneNumbers[$i].NumberType + "~" + $PhoneNumbers[$i].IsoSubdivision + "~" + $PhoneNumbers[$i].IsoCountryCode)
	}
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
   $ExcelWorkSheet.Activate()
}

if ( ! $NoTeamsChannels )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_TeamsChannels).Value = ""

	Write-Host "Geting list of existing Teams and Channels."

	$Teams = @(Get-Team | Sort-Object DisplayName)

	$RowOffset = 1
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsChannels = @(Get-TeamChannel -GroupId $Teams[$i].GroupId | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName)
		
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Processing Team : {2,-50}" -f ($i + 1), $Teams.length, $Teams[$i].DisplayName )
		}

		for ($j=0; $j -lt $TeamsChannels.length; $j++)
		{
			if ( $Verbose )
			{
				Write-Host "`t`t`tChannel: " $TeamsChannels[$j].DisplayName
			}
			
			$RowOffset += 1
#			$ExcelWorkSheet.Cells.Item($RowOffset, 1) = ($Teams[$i].GroupId + "~" + $Teams[$i].DisplayName + "~" + $TeamsChannels[$j].Id + "~" + $TeamsChannels[$j].DisplayName)


			$ExcelWorkSheet.Cells.Item($RowOffset, 1) = ($Teams[$i].DisplayName + "~[" +$Teams[$i].DisplayName + "] - " + $TeamsChannels[$j].DisplayName + "~" + $Teams[$i].GroupId + "~" + $TeamsChannels[$j].Id + "~" + $TeamsChannels[$j].DisplayName)


		}
	}
}
else
{
	Write-Host "Downloading Teams and Channels skipped."	
}	

#
# Team Schedule Groups
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("TeamsScheduleGroups")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoTeamsScheduleGroups )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_TeamsScheduleGroups).Value = ""

	Write-Host "Geting list of existing Teams Schedule Groups."
	
	if ( $NoTeamsChannels )
	{
		$Teams = @(Get-Team | Sort-Object DisplayName)
	}

	$RowOffset = 1
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsScheduleGroups = @(Get-MgTeamScheduleSchedulingGroup -TeamId $Teams[$i].GroupId 2> $null | Sort-Object DisplayName)
		
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Processing Team : {2,-50}" -f ($i + 1), $Teams.length, $Teams[$i].DisplayName )
		}
		
		for ($j=0; $j -lt $TeamsScheduleGroups.length; $j++)
		{
			if ( $Verbose )
			{
				Write-Host "`t`t`tScheduling Group: " $TeamsScheduleGroups[$j].DisplayName
			}
							
			$RowOffset += 1
			$ExcelWorkSheet.Cells.Item($RowOffset, 1) = ($Teams[$i].DisplayName + "~[" +$Teams[$i].DisplayName + "] - " + $TeamsScheduleGroups[$j].DisplayName + "~" + $Teams[$i].GroupId + "~" + $TeamsScheduleGroups[$j].Id + "~" + $TeamsScheduleGroups[$j].DisplayName)
		}
	}
}
else
{
	Write-Host "Downloading Teams and Schedule Groups skipped."	
}	

#
# Users
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Users")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoUsers )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_Users).Value = ""

	Write-Host "Getting list of enterprise voice enabled users."

	$Users = @(Get-CsOnlineUser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias)

	for ( $i = 0; $i -lt $Users.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) User : {2,-50}" -f ($i + 1), $Users.length, $Users[$i].UserPrincipalName )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ( (($Users[$i].UserPrincipalName -split "@")[0]) + "~" + $Users[$i].UserPrincipalName + "~" + $Users[$i].Identity)
	}
}
else
{
	Write-Host "Downloading Users skipped."	
}

#
# CR4CQ Templates
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CR4CQ")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoCR4CQTemplates )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_CR4CQ).Value = ""

	Write-Host "Getting list of compliance recording for call queue templates."

	$CR4CQTemplates = @(Get-CsComplianceRecordingForCallQueueTemplate)

	for ( $i = 0; $i -lt $CR4CQTemplates.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) CR4CQ Template : {2,-50}" -f ($i + 1), $CR4CQTemplates.length, $CR4CQTemplates[$i].Name )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($CR4CQTemplates[$i].Name + "~" + $CR4CQTemplates[$i].Id + "~" + $CR4CQTemplates[$i].Description + "~" + $CR4CQTemplates[$i].BotId + "~" + $CR4CQTemplates[$i].RequiredBeforeCall + "~" + $CR4CQTemplates[$i].RequiredDuringCall + "~" + $CR4CQTemplates[$i].ConcurrentInvitationCount + "~" + $CR4CQTemplates[$i].PairedApplication)
	}
}
else
{
	Write-Host "Downloading CR4CQ Templates skipped."	
}

#
# Save and close the Excel file
#
if ( $View )
{
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Config-CallQueue")
   $ExcelWorkSheet.Activate()
}

#
# Restore AutoSave
#
$ExcelWorkBook.AutoSaveOn = $AutoSave
$ExcelWorkBook.Parent.Calculation = $AutoCalc
$ExcelWorkBook.Save()
$ExcelWorkBook.Close($true)
$ExcelObj.Quit()

if ( Test-Path -Path ".\`$null" )
{
	Remove-Item -Path ".\`$null" | Out-Null
}


if ( ! $NoOpen )
{
	Write-Host "Preparation complete.  Opening $ExcelFilename.  "
	Write-Host "Please complete the configuration, save and exit the spreadsheet and then run the BulkCQsProvisioning script."
	Write-Host -NoNewLine "Press any key to continue..."
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

	Invoke-item "$ExcelFullPathFilename"
}
else
{
	Write-Host "Preparation complete."
}

