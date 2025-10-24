# Version: 1.0.7
# Date: 2025.10.24

#
# Changelog: https://github.com/MicrosoftDocs/Teams-Auto-Attendant-and-Call-Queue-Backup-and-Bulk-Provisioning-Tools/blob/main/Call%20Queue/CHANGELOG.md

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
		[string] $CQID,
		[Parameter(Mandatory=$true, Position=4)]
		[string] $CQName
		
	)

	$currentDIR = "\\?\" + (Get-Location).Path
	$audioFilesDIR = $currentDIR + "\AudioFiles\"

	if ( ! (Test-Path -Path $audioFilesDIR) )
	{
		$null = New-Item -Path $currentDIR -Name "AudioFiles" -ItemType Directory -Force
	}
	
	# Setup Directory For Call Queue
	$callQueueDIR = $audioFilesDIR + $CQID + "_" + $CQName + "\"
	
	# Does Call Queue directory exist - make it if not
	if ( ! (Test-Path -Path $callQueueDIR) )
	{
		$DIRName = $CQID + "_" + $CQName
		$null = New-Item -Path $audioFilesDIR -Name $DIRName -ItemType Directory -Force
	}
	
	$AudioExportPathDIR = $callQueueDIR + $fileID + "_" + $fileName

	Write-Host "`t`t`tDownloading ""$fileType"" : ($fileName)"
	$content = (Export-CsOnlineAudioFile -ApplicationId "HuntGroup" -Identity $fileID 2> $null)
	if ( $content.length -ne 0 )
	{
		[System.IO.File]::WriteAllBytes($AudioExportPathDIR, $content)
		$global:AudioFileExportPath = $AudioExportPathDIR.SubString(4)
		return $true
	}
	else
	{
		return $false
	}
}


###########################
#  ExcelScroll
###########################
function ExcelScroll
{
	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[int]$currentRowOffset,
		[Parameter(Mandatory=$true, Position=1)]
		[int]$scrollWindow
	)

	if ( $scrollWindow -eq 0 )
	{
		$ExcelObj.ActiveWindow.ScrollRow = $currentRowOffset
		$ExcelObj.ActiveWindow.ScrollColumn = 1
		return
	}
	
	$modCurrentRowOffset = ( $currentRowOffset % $scrollWindow )
	if ( $modCurrentRowOffset -eq 0 )
	{
		if ( ( $currentRowOffset - 5) -lt 1 )
		{
			$ExcelObj.ActiveWindow.ScrollRow = $currentRowOffset
		}
		else
		{
			$ExcelObj.ActiveWindow.ScrollRow = ( $currentRowOffset - 5 )
		}
	}
	return
}


###########################
#  TargetIDLookup
###########################
function TargetIDLookup
{
	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[AllowEmptyString()]
		[string] $targetID,
		[Parameter(Mandatory=$true, Position=1)]
		[AllowEmptyString()]
		[string] $targetIDType
	)
	
	switch ( $targetIDType )
	{
		"User"						{	$index = [Array]::IndexOf($Users.Identity, $targetID)
										
										if ( $index -ge 0 )
										{
											return $Users[$index].DisplayName 
										}
										else
										{
											return ""
										}
									}
		"ApplicationEndpoint"		{	$index = [Array]::IndexOf($ResourceAccounts.ObjectId, $targetID)
										
										if ( $index -ge 0 )
										{
											return $ResourceAccounts[$index].DisplayName
										}
										else
										{
											return ""
										}
									}
		"ConfigurationEndpoint"		{ 	$index = [Array]::IndexOf($AutoAttendants.Identity, $targetID)
										
										if ( $index -ge 0 )
										{
												return $AutoAttendants[$index].Name
										}
										else
										{
											$index = [Array]::IndexOf($CallQueues.Identity, $targetID)
											
											if ( $index -ge 0 )
											{
												return $CallQueues[$index].Name
											}
										}
										return ""
									}
		"SharedVoicemail"			{ 	$index = [Array]::IndexOf($Teams.GroupId, $targetID)
										
										if ( $index -ge 0 )
										{
											return $Teams[$index].DisplayName
										}
										return ""
									}
		"Voicemail"					{	$index = [Array]::IndexOf($Users.Identity, $targetID)
										
										if ( $index -ge 0 )
										{
											return $Users[$index].DisplayName 
										}
										return ""
									}


		"CR4CQ"						{	$index = [Array]::IndexOf($CR4CQTemplates.Id, $targetID)

										if ( $index -ge 0 )
										{
											return $CR4CQTemplates[$index].Name 
										}
										return ""
									}
		"SCH"						{	$index = [Array]::IndexOf($SCHTemplates.Id, $targetID)
										if ( $index -ge 0 )
										{
											return $SCHTemplates[$index].Name 
										}
										return ""
									}
		"Channel"					{	$index = [Array]::IndexOf($TeamsChannels.Id, $targetID)

										if ( $index -ge 0 )
										{
											return $TeamsChannels[$index].DisplayName
										}
										return ""
									}
		Default						{	return ""	}
	}
	return ""
}

#
# Confirm running in PowerShell v5.x
#
if ( $PSVersionTable.PSVersion.Major -ne 5 )
{
	Write-Error "This script is only supported in PowerShell v5.x" -f Red
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
			"-noschtemplates"			{	$NoSCHTemplates = $true }
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
		$NoAutoAttendants = $false
		$NoCallQueues = $false
		$NoPhoneNumbers = $false
		$NoResourceAccounts = $false
		$NoTeamsChannels = $false
		$NoTeamsScheduleGroups = $false
		$NoUsers = $false
		$NoCR4CQTemplates = $false
		$NoSCHTempaltes = $false

		$Verbose = $true
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

if ( $NoAutoAttendants -and $NoCallQueues -and $NoPhoneNumbers -and $NoResourceAccounts -and $NoTeamsChannels -and $NoTeamsScheduleGroups -and $NoUsers -and $NoCR4CQTemplates -and $NoSCHTemplates )
{
	Write-Host "All options disabled. Nothing to do. Processing has been halted." -f Red
	$ArgError = $true
}

if ( ( $Help ) -or ( $ArgError ) )
{
	Write-Host "The following options are avaialble:"
	Write-Host "`t-AACount <n> - provide to limit processing to the first AACount Auto Attendants otherwise all in tenant."
	Write-Host "`t-CQCount <n> - provide to limit processing to the first CQCount Call Queues otherwise all in tenant."
	Write-Host "`t-Download - download all Call Queue configuration, including audio file. WARNING - may take a long time"
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkCQs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)"	
	Write-Host "`t-NoAutoAttendants - do not download existing auto attendant information"
	Write-Host "`t-NoCallQueues - do not download existing call queue information"
	Write-Host "`t-NoCR4CQTemplates - do not download existing compliance recording templates"
	Write-Host "`t-NoSCHTemplates - do not download existing shared call queue history templates" 
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

$StartTime = (Get-Date -Format "HH:mm:ss")
Write-Host "$StartTime - Starting BulkCQsPreparation." -f Green

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Module Min Supported Versions
#
$MicrosoftTeamsMinVersion = [version]"7.1.0"
$MicrosoftGraphMinVersion = [version]"2.24.0"
$ImportExcelMinVersion = [version]"7.8.0"

#
# BulkCQs.xlms Required Version
#
$ExcelSpreadsheetRequiredVersion = "1.0.6"
$ExcelSpreadsheetVersionSheet = "Data"
$ExcelSpreadsheetVersionRowRef = 2
$ExcelSpreadsheetVersionColRef = 50

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
$Range_SCH = "A2:A51"

#
# Set types, declare global variables and assign defaults
#
$global:AudioFileExportPath = $null
[int]$RowOffset = 0
[int]$ScrollWindow = 20

#
# Check that minimum verion of required modules are installed - install if not
#
# MicrosoftTeams
#
Write-Host "Checking for MicrosoftTeams module $MicrosoftTeamsMinVersion or later." -f Green
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
      Write-Error "`tNot signed into Microsoft Teams!" -f Red
      exit
   }
   Write-Host "`tConnected to Microsoft Teams." -f Green
}
else
{
   Write-Host "`tThe MicrosoftTeams module is not installed or does not meet the minimum requirements - installing." -f Yellow
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
      Write-Error "`tNot signed into Microsoft Teams!" -f Red
      exit
   }
   Write-Host "`tConnected to Microsoft Teams." -f Green
}

if ( ! $NoTeamsScheduleGroups )
{
	Write-Host "Checking for Microsoft.Graph module $MicrosoftGraphMinVersion or later." -f Green
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
			Write-Error "`tNot signed into Microsoft Graph!" -f Red
			exit
		}
		Write-Host "`tConnected to Microsoft Graph." -f Green
	}
	else
	{
		Write-Host "`tThe Microsoft.Graph module is not installed or does not meet the minimum requirements - installing." -f Yellow
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
			Write-Error "`tNot signed into Microsoft Graph!" -f Red
			exit
		}
		Write-Host "`tConnected to Microsoft Graph." -f Green
	}
}

#
# ImportExcel
#
Write-Host "Checking for ImportExcel module $ImportExcelMinVersion or later." -f Green
$Version = ( (Get-InstalledModule -Name ImportExcel -MinimumVersion "$ImportExcelMinVersion").Version 2> $null )

if ( $Version -match "-preview" )
{
	$Version = $Version.Replace("-preview", "")
}

$Version = [version]$Version

Write-Host "`tImportExcel Version: $Version"
if ( $Version -ge $ImportExcelMinVersion )
{
   Write-Host "`tImporting ImportExcel." -f Green
   Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
}
else
{
   Write-Host "`tThe ImportExcel module is not installed or does not meet the minimum requirements - installing." -f Yellow
   Install-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion -Force -AllowClobber
   
   Write-Host "`tImporting ImportExcel." -f Green
   Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
}

#
# Setup filename
#
if ( $ExcelFilename -eq $null )
{
   $ExcelFilename = "BulkCQs.xlsm"
}
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename

Write-Host "Opening the $ExcelFilename worksheet (this may take some time, please be patient)." -f Green

#
# Check if supplied filename exists
#
if ( !( Test-Path -Path $ExcelFullPathFilename ) )
{
	Write-Error "`tERROR: $ExcelFilename does not exist." -f red
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
	Write-Host "`tThe $ExcelFileName appears to be open already. Please close the file and try again." -f Red
	exit
}

#
# Checking Spreadsheet Version
#
Write-Host "`tChecking that $ExcelFilename is the correct version."
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($ExcelSpreadsheetVersionSheet)
$ExcelSpreadsheetVersion = $ExcelWorkSheet.Cells.Item($ExcelSpreadsheetVersionRowRef, $ExcelSpreadsheetVersionColRef).Text

if ( $ExcelSpreadsheetVersion -ne $ExcelSpreadsheetRequiredVersion )
{
	Write-Host "`tThe $ExcelFileName version ($ExcelSpreadsheetVersion) does not match the required version ($ExcelSpreadsheetRequiredVersion)."  -f Red
	$ExcelWorkBook.Close($true)
	$ExcelObj.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelObj) | Out-Null
	exit
}


#
# Configure workbook
#
if ( $Verbose )
{
	Write-Host "`tConfiguring workbook"
	$AutoSave = $ExcelWorkBook.AutoSaveOn
	if ( $AutoSave )
	{
		Write-Host "`t`tAuto Save is currently enabled."
		
		$ExcelWorkBook.AutoSaveOn = $false
		Write-Host "`t`t`tAuto Save has been disabled"
	}
	else
	{
		Write-Host "`t`tAuto Save is currently disabled."
	}
	
	$AutoCalc = $ExcelWorkBook.Parent.Calculation
	if ( $AutoCalc -eq -4105 )
	{
		Write-Host "`t`tAutomatic Calculation is currently enabled."
		
		$ExcelWorkBook.Parent.Calculation = -4135
		Write-Host "`t`t`tAutomatic Calculation has been disabled" 
	}
	else
	{
		Write-Host "`t`tAutomatic Calculation is currently disabled. " $AutoCalc
	}
}

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

	Write-Host "Retrieving list of Resource Accounts." -f Green

	$ResourceAccounts = @(Get-CsOnlineApplicationInstance | Where {$_.ApplicationId -ne ""} | Sort-Object ApplicationId, DisplayName)
	
	if ( $Verbose )
	{	
		Write-Host ( "`tTotal number of Resource Accounts: {0,4}" -f $ResourceAccounts.length )
	}

	$ResourceAccountUserDetails = @()

	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		$PercentProcessed = ( ( $i / $ResourceAccounts.length ) * 100 )
		Write-Host -NoNewLine ( "`tRetrieving list of Resource Account User Details (this takes some time, please be patient [{0:F1}%])`r" -f $PercentProcessed )
		$ResourceAccountUserDetails += (Get-CsOnlineUser -Identity $ResourceAccounts[$i].ObjectId)
	}

	Write-Host "`nProcessing list of Resource Accounts." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		#
		# Make sure resource account is not Deleted
		#
		if ( $ResourceAccountUserDetails[$i].SoftDeletionTimeStamp.length -eq 0 )
		{
			if ( $ResourceAccountUserDetails[$i].UsageLocation.length -eq 0 )
			{
				$ResourceAccountUserDetails[$i].UsageLocation = "US"
			}

			switch ( $ResourceAccounts[$i].ApplicationId)
			{
				"ce933385-9390-45d1-9512-c8d228074e07"	{	# Auto Attendant
															if ( $Verbose )
															{
																Write-Host ( "`t({0,4}/{1,4}) [RA-AA] Resource Account ({2,-36}): {3,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].ObjectId, $ResourceAccounts[$i].DisplayName )
															}

															$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ("[RA-AA] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" + $ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails[$i].UsageLocation )
														}
				"7d563055-b6cc-4d7a-93c1-38aa762b34ae"	{	# Mainline Attendant
															if ( $Verbose )
															{
																Write-Host ( "`t({0,4}/{1,4}) [RA-MA] Resource Account ({2,-36}): {3,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].ObjectId, $ResourceAccounts[$i].DisplayName )
															}

															$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ("[RA-MA] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" + $ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails[$i].UsageLocation )
														}
				"11cd3e2e-fccb-42ad-ad00-878b93575e07"	{	# Call Queue
															if ( $Verbose )
															{
																Write-Host ( "`t({0,4}/{1,4}) [RA-CQ] Resource Account ({2,-36}): {3,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].ObjectId, $ResourceAccounts[$i].DisplayName)
															}
															
															# request will generate a "Correlation id for this request" message when the RA is not assigned to anything, also generates error so redirecting that to null
															$ResourceAccountPriority = ((Get-CsOnlineApplicationInstanceAssociation -identity $ResourceAccounts[$i].ObjectId).CallPriority 2> $null )
															$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ("[RA-CQ] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" +$ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails[$i].UsageLocation  + "~" + $ResourceAccountPriority )
														}
				default									{	# Other
															if ( $Verbose )
															{
																Write-Host ( "`t({0,4}/{1,4}) [OTHER] Resource Account ({2,-36}): {3,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].ObjectId, $ResourceAccounts[$i].DisplayName )
															}

															$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ("[OTHER] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" + $ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails[$i].UsageLocation )
														}
			}
		}
		else
		{
			if ( $Verbose )
			{
				Write-Host "`tResource Account Not Added (Soft Deleted): " $ResourceAccounts.DisplayName[$i]
			}
		}
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Resource Accounts skipped."	-f Yellow
}

#
# Auto Attendants
#
if ( ! $NoAutoAttendants )
{
	Write-Host "Retrieving list of Auto Attendants." -f Green

	$AutoAttendants = @()

	if ( $AACount -gt 0 )
	{
		if ( ( $AACount % 100 ) -eq 0 )
		{
			$loops = [int] [Math]::Truncate($AACount / 100)
		}
		else
		{
			$loops = [int] [Math]::Truncate($AACount / 100) + 1
		}
		
		for ( $i = 0; $i -lt $loops; $i++ )
		{
			$j = $i * 100

			if ( $AACount -le 100 )
			{
				$First = $AACount
			}
			else
			{
				$Remaining = $AACount - $j
				if ( $Remaining -gt 100 )
				{
					$First = 100
				}
				else
				{
					$First = $Remaining
				}
			}

			if ( $Verbose )
			{
				if ( $First -eq 100 )
				{
					Write-Host ( "`tRetrieving list of Auto Attendants {0,4} to {1,4} of {2,4}" -f ($j+1), ($j+100), $AACount )
				}
				else
				{
					Write-Host ( "`tRetrieving list of Auto Attendants {0,4} to {1,4} of {2,4}" -f ($j+1), ($j+$First), $AACount )
				}
			}

			$AutoAttendants += @(Get-CsAutoAttendant -Skip $j -First $First 3> $null)
		}
	}
	else
	{
		# Download all auto attendants
		$skipCount = 0
		
		do
		{
			if ( $Verbose )
			{
				Write-Host ( "`tRetrieving list of Auto Attendants {0,4} to {1,4}" -f ($skipCount+1), ($skipCount - ($skipCount % 100) + 100) )
			}
			$AA = @(Get-CsAutoAttendant -Skip $skipCount -First 100 3> $null)
			$AutoAttendants += $AA
			$skipCount += $AA.length
		}
		until ( $AA.length -lt 100 )
	}
	
	if ( $Verbose )
	{
		Write-Host ( "`tTotal number of Auto Attendants: {0,4}" -f $AutoAttendants.length )
	}

	Write-Host "Processing list of Auto Attendants." -f Green
	
	$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("AutoAttendants")

	if ( $View )
	{
		$ExcelWorkSheet.Activate()
	}

	#
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_AutoAttendants).Value = ""

	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ($i=0; $i -lt $AutoAttendants.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Auto Attendant ({2,36}) : {3,-50}" -f ($i + 1), $AutoAttendants.length, $AutoAttendants[$i].Identity,$AutoAttendants[$i].Name )
		}

		$AssignedResourceAccounts = $AutoAttendants[$i].ApplicationInstances
		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($AutoAttendants[$i].Name + "~" + $AutoAttendants[$i].Identity + "~" + ($AssignedResourceAccounts -join ","))
		
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Auto Attendants skipped." -f Yellow
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

	Write-Host "Retrieving list of existing voice application phone numbers" -f Green

	$PhoneNumbers = @(Get-CsPhoneNumberAssignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned" | Sort-Object TelephoneNumber)
	
	if ( $Verbose )
	{
		Write-Host "`tTotal number of voice application phone numbers: " $PhoneNumbers.length
	}

	Write-Host "Processing list of voice application phone numbers." -f Green
	
	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ($i=0; $i -lt  $PhoneNumbers.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Phone Number : {2,-50}" -f ($i + 1), $PhoneNumbers.length, $PhoneNumbers[$i].TelephoneNumber )
		}

		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($PhoneNumbers[$i].TelephoneNumber + "~" + $PhoneNumbers[$i].NumberType + "~" + $PhoneNumbers[$i].IsoSubdivision + "~" + $PhoneNumbers[$i].IsoCountryCode)
		
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Voice Applications phone numbers skipped."	-f Yellow
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

	Write-Host "Retrieving list of existing Teams and Channels." -f Green

	$Teams = @(Get-Team | Sort-Object DisplayName)
	
	if ( $Verbose )
	{
		Write-Host "`tTotal number of Teams: " $Teams.length
	}

	Write-Host "Processing list of Teams and Channels." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0
	
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsChannels = @(Get-TeamChannel -GroupId $Teams[$i].GroupId | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName)
		
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Team : {2,-50}" -f ($i + 1), $Teams.length, $Teams[$i].DisplayName )
			
			if ( $TeamsChannels.length -eq 1 )
			{
				Write-Host "`t`tChannel: " $TeamsChannels.DisplayName
			}
			else
			{
				Write-Host "`t`tChannel: " $TeamsChannels[0].DisplayName ### was $j which at this stage could be anything from usage above
			}
		}

		for ($j=0; $j -lt $TeamsChannels.length; $j++)
		{
			$ExcelWorkSheet.Cells.Item($RowOffset++, 1) = ($Teams[$i].DisplayName + "~[" +$Teams[$i].DisplayName + "] - " + $TeamsChannels[$j].DisplayName + "~" + $Teams[$i].GroupId + "~" + $TeamsChannels[$j].Id + "~" + $TeamsChannels[$j].DisplayName)
		}
	}
}
else
{
	Write-Host "Downloading Teams and Channels skipped." -f Yellow
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

	Write-Host "Retrieving list of existing Teams Schedule Groups." -f Green
	
	if ( $NoTeamsChannels )
	{
		$Teams = @(Get-Team | Sort-Object DisplayName)
	}
	
	if ( $Verbose )
	{
		Write-Host "`tTotal number of Teams: " $Teams.length
	}

	Write-Host "Processing list of Teams and Schedule Groups." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsScheduleGroups = @(Get-MgTeamScheduleSchedulingGroup -TeamId $Teams[$i].GroupId 2> $null | Sort-Object DisplayName)
		
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Team : {2,-50}" -f ($i + 1), $Teams.length, $Teams[$i].DisplayName )
		}
		
		if ( $TeamsScheduleGroups.length -eq 0 )
		{
			if ( $Verbose )
			{
				Write-Host "`t`t`tNo Schedules Groups for this Team " -f Yellow
			}
		}
		
		for ($j=0; $j -lt $TeamsScheduleGroups.length; $j++)
		{
			if ( $Verbose )
			{
				Write-Host "`t`t`tScheduling Group: " $TeamsScheduleGroups[$j].DisplayName
			}
							
			$ExcelWorkSheet.Cells.Item($RowOffset++, 1) = ($Teams[$i].DisplayName + "~[" +$Teams[$i].DisplayName + "] - " + $TeamsScheduleGroups[$j].DisplayName + "~" + $Teams[$i].GroupId + "~" + $TeamsScheduleGroups[$j].Id + "~" + $TeamsScheduleGroups[$j].DisplayName)
		}
	}
}
else
{
	Write-Host "Downloading Teams and Schedule Groups skipped." -f Yellow
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
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_Users).Value = ""

	Write-Host "Retrieving list of enterprise voice enabled users." -f Green

	$Users = @(Get-CsOnlineUser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias)

	if ( $Verbose )
	{
		Write-Host "`tTotal number of enterprise voice enabled users: " $Users.length
	}

	Write-Host "Processing list of enterprise voice enabled users." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0
	
	for ( $i = 0; $i -lt $Users.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) User : {2,-50}" -f ($i + 1), $Users.length, $Users[$i].UserPrincipalName )
		}

		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($Users[$i].UserPrincipalName + "~" + $Users[$i].Identity)

		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading EV enabled users skipped." -f Yellow
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

	Write-Host "Retrieving list of Compliance Recording for Call Queue templates." -f Green

	$CR4CQTemplates = @(Get-CsComplianceRecordingForCallQueueTemplate)
	
	if ( $Verbose )
	{
		Write-Host ( "`tTotal number of Compliance Recording for Call Queue templates: {0,4}" -f $CR4CQTemplates.length )
	}

	Write-Host "Processing list of Compliance Recording for Call Queue templates." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ( $i = 0; $i -lt $CR4CQTemplates.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Template : {2,-50}" -f ($i + 1), $CR4CQTemplates.length, $CR4CQTemplates[$i].Name )
		}

		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($CR4CQTemplates[$i].Name + "~" + $CR4CQTemplates[$i].Id + "~" + $CR4CQTemplates[$i].Description + "~" + $CR4CQTemplates[$i].BotId + "~" + $CR4CQTemplates[$i].RequiredBeforeCall + "~" + $CR4CQTemplates[$i].RequiredDuringCall + "~" + $CR4CQTemplates[$i].ConcurrentInvitationCount + "~" + $CR4CQTemplates[$i].PairedApplication)
	
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Compliance Recording for Call Queue templates skipped." -f Yellow
}

#
# Shared Call History (SCH) Templates
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("SCH")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoSCHTempaltes )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_SCH).Value = ""

	Write-Host "Retrieving list of Shared Call Queue History templates ." -f Green

	$SCHTemplates = @(Get-CsSharedCallQueueHistoryTemplate)
	
	if ( $Verbose )
	{
		Write-Host ( "`tTotal number of Shared Call Queue History templates: {0,4}" -f $SCHTemplates.length )
	}

	Write-Host "Processing list of Shared Call Queue History templates." -f Green

	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ( $i = 0; $i -lt $SCHTemplates.length; $i++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Template : {2,-50}" -f ($i + 1), $SCHTemplates.length, $SCHTemplates[$i].Name )
		}

		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($SCHTemplates[$i].Name + "~" + $SCHTemplates[$i].Id + "~" + $SCHTemplates[$i].Description + "~" + $SCHTemplates[$i].IncomingMissedCalls + "~" + $SCHTemplates[$i].AnsweredAndOutboundCalls)
	
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Shared Call Queue History templates skipped." -f Yellow
}


#
# Call Queues
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueues")

if ( $View )
{
	$ExcelWorkSheet.Activate()
}

if ( ! $NoCallQueues )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_CallQueues).Value = ""

	Write-Host "Retrieving list of Call Queues." -f Green

	$CallQueues = @()

	if ( $CQCount -gt 0 )
	{
		$loops = [int] [Math]::Truncate($CQCount / 100) + 1

		for ( $i = 0; $i -lt $loops; $i++ )
		{
			$j = $i * 100
			
			if ( $CQCount -le 100 )
			{
				$First = $CQCount
			}
			else
			{
				$Remaining = $CQCount - $j
				if ( $Remaining -gt 100 )
				{
					$First = 100
				}
				else
				{
					$First = $Remaining
				}
			}

			if ( $Verbose )
			{
				if ( $First -eq 100 )
				{
					Write-Host ( "`tRetrieving list of Call Queues {0,4} to {1,4} of {2,4}" -f ($j+1), ($j+100), $CQCount )
				}
				else
				{
					Write-Host ( "`tRetrieving list of Call Queues {0,4} to {1,4} of {2,4}" -f ($j+1), ($j+$First), $CQCount )
				}
			}

			$CallQueues = @(Get-CsCallQueue -Skip $j -First $First 3> $null )
		}
	}
	else
	{
		# Download all call queues
		$skipCount = 0
		$CallQueues = @()
		
		do
		{
			if ( $Verbose )
			{
				Write-Host ( "`tRetrieving list of Call Queues {0,4} to {1,4}" -f ($skipCount+1), ($skipCount - ($skipCount % 100) + 100) )
			}
			$CQ = @(Get-CsCallQueue -Skip $skipCount -First 100 3> $null)
			$CallQueues += $CQ
			$skipCount += $CQ.length
		}
		until ( $CQ.length -lt 100 )
	}
	
	if ( $Verbose )
	{
		Write-Host "`tTotal number of Call Queues: " $CallQueues.length
	}

	Write-Host "Processing list of Call Queues." -f Green
	
	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ($k=0; $k -lt  $CallQueues.length; $k++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Call Queue ({2,-36}) : {3,-50}" -f ($k + 1), $CallQueues.length, $CallQueues.Identity[$k], $CallQueues.Name[$k] )
		}

		$AssignedResourceAccounts = ( (Get-CsCallQueue -Identity $CallQueues.Identity[$k]).ApplicationInstances 3> $null )
		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($CallQueues.Name[$k] + "~" + $CallQueues.Identity[$k] + "~" + ($AssignedResourceAccounts -join ","))
		
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0

	if ( $Download )
	{
		$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("CallQueuesDownload")

		if ( $View )
		{
			$ExcelWorkSheet.Activate()
		}
			
		$ExcelWorkSheet.Range($Range_CallQueueDownload).Value = ""

		Write-Host "Processing list of Call Queues for Download." -f Green

		$RowOffset = 2
		ExcelScroll 2 0
		
		for ( $i = 0; $i -lt $CallQueues.length; $i++ )
		{
			$RowOffset += 1
			$ColumnOffset = 2
			
			$CQ = $CallQueues[$i]
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Identity
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.Name
			
			if ( $Verbose )
			{
				Write-Host ( "`t({0,4}/{1,4}) Call Queue ({2,-36}) : {3,-50}" -f ($i + 1), $CallQueues.length, $CQ.Identity, $CQ.Name )
			}
		
			switch ( $CQ.RoutingMethod )
			{
				"Attendant"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Attendant"}
				"Serial" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Serial"}
				"RoundRobin" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "RoundRobin"}
				"LongestIdle" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "LongestIdle"}
				default  		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.RoutingMethod }
			}

			if ( $CQ.AllowOptOut.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AllowOptOut
			}
			else
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "FALSE"
			}
			
			if ( $CQ.ConferenceMode.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.ConferenceMode
			}
			else
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "FALSE"
			}
			
			if ( $CQ.PresenceBasedRouting.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.PresenceBasedRouting
			}
			else
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "FALSE"
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
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "en-US"
			}
			
			if ( ( $CQ.WelcomeMusicFileName.length -ne 0 ) -and ( $CQ.WelcomeMusicResourceId.length -ne 0 ) )
			{
				if ( AudioFileExport "WelcomeMusic" $CQ.WelcomeMusicFileName $CQ.WelcomeMusicResourceId $CQ.Identity $CQ.Name)
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.WelcomeMusicResourceId
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.WelcomeMusicFileName) | Out-Null
				}
				else
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.WelcomeMusicResourceId
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.WelcomeMusicFileName
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

			# f2e3feed-ab0b-439e-b60a-61eb26df0c53 - default music on hold file id
			if ( ( $CQ.MusicOnHoldResourceId -ne "f2e3feed-ab0b-439e-b60a-61eb26df0c53" ) -and ( $CQ.MusicOnHoldFileName.length -ne 0 ) -and ( $CQ.MusicOnHoldResourceId.length -ne 0 ) )
			{
				if ( AudioFileExport "MusicOnHold" $CQ.MusicOnHoldFileName $CQ.MusicOnHoldResourceId $CQ.Identity $CQ.Name)
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.MusicOnHoldResourceId
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.MusicOnHoldFileName) | Out-Null
				}
				else
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.MusicOnHoldResourceId
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $CQ.MusicOnHoldFileName
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
					"DisconnectWithBusy" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "DisconnectWithBusy"
												$ColumnOffset++
											}
					"Forward" 				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Forward"
												switch ( $CQ.OverflowActionTarget.Type )
												{
													"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" }  
													"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndPoint" }  
													"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndPoint" }  
													"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Phone" }
												}
												$targetIDType = $CQ.OverflowActionTarget.Type
											} 
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Voicemail"
												$targetIDType = "Voicemail"
												$ColumnOffset++
											}
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
												$targetIDType = "SharedVoicemail"
												$ColumnOffset++
											}
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowAction
												$ColumnOffset++
											}
				}
			}
			else
			{
				$ColumnOffset += 2
			}

			if ( $CQ.OverflowActionTarget.Id.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.OverflowActionTarget.Id $targetIDType )
			}
			else
			{
				 $ColumnOffset += 2
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
				if ( AudioFileExport "OverflowSharedVoicemail" $CQ.OverflowSharedVoicemailAudioFilePromptFileName $CQ.OverflowSharedVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowSharedVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowSharedVoicemailAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "OverflowDisconnect" $CQ.OverflowDisconnectAudioFilePromptFileName $CQ.OverflowDisconnectAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowDisconnectAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowDisconnectAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "OverflowRedirectPerson" $CQ.OverflowRedirectPersonAudioFilePromptFileName $CQ.OverflowRedirectPersonAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPersonAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowRedirectPersonAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "OverflowRedirectVoiceApp" $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName $CQ.OverflowRedirectVoiceAppAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoiceAppAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowRedirectVoiceAppAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "OverflowRedirectPhoneNumber" $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName $CQ.OverflowRedirectPhoneNumberAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectPhoneNumberAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowRedirectPhoneNumberAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "OverflowRedirectVoicemail" $CQ.OverflowRedirectVoicemailAudioFilePromptFileName $CQ.OverflowRedirectVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.OverflowRedirectVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.OverflowRedirectVoicemailAudioFilePromptFileName) | Out-Null
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
					"Disconnect" 			{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Disconnect" 
												$ColumnOffset++
											}
					"Forward" 				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Forward" 
												switch ( $CQ.TimeoutActionTarget.Type )
												{
													"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" }  
													"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndPoint" }  
													"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndPoint" }  
													"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Phone" }
												}
												$targetIDType = $CQ.TimeoutActionTarget.Type
											} 
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Voicemail"
												$targetIDType = "Voicemail"
												$ColumnOffset++
											}
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
												$targetIDType = "SharedVoicemail"
												$ColumnOffset++
											}
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutAction
												$ColumnOffset++
											}
				}
			}
			else
			{
				$ColumnOffset += 2
			}
			
			if ( $CQ.TimeoutActionTarget.Id.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.TimeoutActionTarget.Id $targetIDType )
			}
			else
			{
				$ColumnOffset += 2
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
				if ( AudioFileExport "TimeoutSharedVoicemail" $CQ.TimeoutSharedVoicemailAudioFilePromptFileName $CQ.TimeoutSharedVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutSharedVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutSharedVoicemailAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "TimeoutDisconnect" $CQ.TimeoutDisconnectAudioFilePromptFileName $CQ.TimeoutDisconnectAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutDisconnectAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutDisconnectAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "TimeoutRedirectPerson" $CQ.TimeoutRedirectPersonAudioFilePromptFileName $CQ.TimeoutRedirectPersonAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPersonAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutRedirectPersonAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "TimeoutRedirectVoiceApp" $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName $CQ.TimeoutRedirectVoiceAppAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoiceAppAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutRedirectVoiceAppAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "TimeoutRedirectPhoneNumber" $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectPhoneNumberAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutRedirectPhoneNumberAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "TimeoutRedirectVoicemail" $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName $CQ.TimeoutRedirectVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.TimeoutRedirectVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.TimeoutRedirectVoicemailAudioFilePromptFileName) | Out-Null
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
					"Queue" 				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Queue"
												$ColumnOffset++}
					"Disconnect" 			{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Disconnect"
												$ColumnOffset++
											}
					"Forward" 				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Forward"
												switch ( $CQ.NoAgentActionTarget.Type )
												{
													"User"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" }  
													"ConfigurationEndPoint" 	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndPoint" }  
													"ApplicationEndPoint" 		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndPoint" }  
													"Phone"						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Phone" }
												}
												$targetIDType = $CQ.NoAgentActionTarget.Type
											}
					"Voicemail"				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Voicemail"
												$targetIDType = "Voicemail"
												$ColumnOffset++
											}
					"SharedVoicemail" 		{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
												$targetIDType = "SharedVoicemail"
												$ColumnOffset++
											}
					default  				{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentAction
												$ColumnOffset++
											}
				}
			}
			else
			{
				$ColumnOffset += 2
			}
			
			if ( $CQ.NoAgentActionTarget.Id.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentActionTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.NoAgentActionTarget.Id $targetIDType )
			}
			else
			{
				$ColumnOffset += 2
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
				if ( AudioFileExport "NoAgentSharedVoicemail" $CQ.NoAgentSharedVoicemailAudioFilePromptFileName $CQ.NoAgentSharedVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentSharedVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentSharedVoicemailAudioFilePromptFileName) | Out-Null
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
				"NewCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "NewCalls" }
				"AllCalls" 	{ 	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AllCalls" }
			}

			# No Agent Disconnect
			if ( ( $CQ.NoAgentDisconnectAudioFilePrompt.length -ne 0 ) -and ( $CQ.NoAgentDisconnectAudioFilePromptFileName.length -ne 0 ) )
			{
				if ( AudioFileExport "NoAgentDisconnect" $CQ.NoAgentDisconnectAudioFilePromptFileName $CQ.NoAgentDisconnectAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentDisconnectAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentDisconnectAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "NoAgentRedirectPerson" $CQ.NoAgentRedirectPersonAudioFilePromptFileName $CQ.NoAgentRedirectPersonAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPersonAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentRedirectPersonAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "NoAgentRedirectVoiceApp" $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName $CQ.NoAgentRedirectVoiceAppAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoiceAppAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentRedirectVoiceAppAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "NoAgentRedirectPhoneNumber" $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectPhoneNumberAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentRedirectPhoneNumberAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "NoAgentRedirectVoicemail" $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName $CQ.NoAgentRedirectVoicemailAudioFilePrompt $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.NoAgentRedirectVoicemailAudioFilePrompt
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.NoAgentRedirectVoicemailAudioFilePromptFileName) | Out-Null
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
				if ( AudioFileExport "CallbackOffer" $CQ.CallbackOfferAudioFilePromptFileName $CQ.CallbackOfferAudioFilePromptResourceId $CQ.Identity $CQ.Name )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.CallbackOfferAudioFilePromptResourceId
					
					$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
					$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $CQ.CallbackOfferAudioFilePromptFileName) | Out-Null
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
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.ChannelId "Channel" )
			}
			else
			{
				$ColumnOffset += 2
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
				
				# $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.ComplianceRecordingForCallQueueTemplateId "CR4CQ" )
				$ColumnOffset++
			}
			else
			{
				$ColumnOffset += 2
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
			# Shared Call Queue History
			#
			if ( $CQ.SharedCallQueueHistoryTemplateId.length -ne 0 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.SharedCallQueueHistoryTemplateId
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.SharedCallQueueHistoryTemplateId "SCH" )
			}
			else
			{
				$ColumnOffset += 2
			}
			
			#
			# Authorized Users (AuthorizedUsers is array)
			#
			for ($l=0; $l -lt $CQ.AuthorizedUsers.length; $l++)
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.AuthorizedUsers[$l].Guid
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.AuthorizedUsers[$l].Guid "User")
			}
			$ColumnOffset += ( 30 - ($l * 2) ) # Max 15 auth users x 2 fields per user
			
			#
			# Hidden Authorized Users (HideAuthorizedUsers is array)
			#
			for ($l=0; $l -lt $CQ.HideAuthorizedUsers.length; $l++)
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $CQ.HideAuthorizedUsers[$l].Guid
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ( TargetIDLookup $CQ.AuthorizedUsers[$l].Guid "User")
			}
			$ColumnOffset += ( 30 - ($l * 2) ) # Max 15 auth users x 2 fields per user
			

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
			ExcelScroll $RowOffset $ScrollWindow
		} # CQ Loop
		ExcelScroll 2 0
	} # Download
}
else
{
	Write-Host "Downloading Call Queues skipped." -f Yellow
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
# Restore workbook configuration
#
Write-Host "Saving the $ExcelFilename workbook (this may take some time, please be patient)." -f green
if ( $Verbose )
{
	Write-Host "`tRestoring workbook configuration"

	if ( $AutoCalc -eq -4105 )
	{
		Write-Host "`t`tAutomatic Calculation has been reenabled."
		
		$ExcelWorkBook.Parent.Calculation = -4105
	}

	if ( $AutoSave )
	{
		Write-Host "`t`tAuto Save has been reenabled."
		$ExcelWorkBook.AutoSaveOn = $true
	}
}

$ExcelWorkBook.Save()
$ExcelWorkBook.Close($true)
$ExcelObj.Quit()

#
# Release COM objects to free up resources
#
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelObj) | Out-Null

#
# Calculate run time
#
$EndTime = (Get-Date -Format "HH:mm:ss")
$Duration = New-Timespan -start $StartTime -end $EndTime
$DurationFormatted = ""

if ( $Duration.Hours -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Hours.ToString() + ":"
}
else
{
	$DurationFormatted += $Duration.Hours.ToString() + ":"
}

if ( $Duration.Minutes -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Minutes.ToString() + ":"
}
else
{
	$DurationFormatted += $Duration.Minutes.ToString() + ":"
}

if ( $Duration.Seconds -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Seconds.ToString()
}
else
{
	$DurationFormatted += $Duration.Seconds.ToString()
}


if ( ! $NoOpen )
{
	Write-Host "$EndTime - Preparation complete.  Opening $ExcelFilename." -f Green
	Write-Host "Duration: $DurationFormatted" -f Green
	Write-Host "Please complete the configuration, save and exit the spreadsheet and then run the BulkCQsProvisioning script."
	Write-Host -NoNewLine "Press any key to continue..."
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

	Invoke-item "$ExcelFullPathFilename"
}
else
{
	Write-Host "$EndTime - Preparation complete. " -f Green
	Write-Host "Duration: $DurationFormatted" -f Green
}
