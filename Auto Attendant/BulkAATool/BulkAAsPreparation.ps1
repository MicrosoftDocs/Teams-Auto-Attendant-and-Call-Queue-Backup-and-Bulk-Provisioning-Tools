# Version: 1.0.4
# Date: 2025.09.22

#
# Changelog: https://github.com/MicrosoftDocs/Teams-Auto-Attendant-and-Call-Queue-Backup-and-Bulk-Provisioning-Tools/blob/main/Auto%20Attendant/CHANGELOG.md

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
		[string] $AAID,
		[Parameter(Mandatory=$true, Position=4)]
		[string] $AAName,
		[Parameter(Mandatory=$true, Position=5)]
		[string] $CFID, 
		[Parameter(Mandatory=$true, Position=6)]
		[string] $CFName
	)

	# Setup Directory to AudioFiles
	$currentDIR = "\\?\" + (Get-Location).Path
	$audioFilesDIR = $currentDIR + "\AudioFiles\"

	# Does AudioFiles directory exist - make it if not
	if ( ! (Test-Path -Path $audioFilesDIR) )
	{
		$null = New-Item -Path $currentDIR -Name "AudioFiles" -ItemType Directory
	}
	
	# Setup Directory For Auto Attendant
	$autoAttendantDIR = $audioFilesDIR + $AAID + "_" + $AAName + "\"
	#$autoAttendantDIR = $audioFilesDIR + $AAID + "\"
	

	# Does Auto Attenant directory exist - make it if not
	if ( ! (Test-Path -Path $autoAttendantDIR) )
	{
		$DIRName = $AAID + "_" + $AAName
		$null = New-Item -Path $audioFilesDIR -Name $DIRName -ItemType Directory
	}
	
	# Setup Directory For Auto Attendant Call Flow
	$autoAttendantCallFlowDIR = $autoAttendantDIR + $CFID + "_" + $CFName + "\"
	

	# Does Auto Attendant Call Flow directory exist - make it if not
	if ( ! (Test-Path -Path $autoAttendantCallFlowDIR) )
	{
		$DIRName = $CFID + "_" + $CFName
		$null = New-Item -Path $autoAttendantDIR -Name $DIRName -ItemType Directory
	}
	
	$AudioExportPathDIR = $autoAttendantCallFlowDIR + $fileID + "_" + $fileName
	$tempDIR = "C:\temp\" + $fileID + "_" + $fileName

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
									}
		Default						{	return ""	}
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
$AACount = [int]0
$CQCount = [int]0

if ( $args -ne "" )
{
	for ( $i = 0; $i -lt $args.length; $i++ )
	{
		switch ( $args[$i] )
		{
			"-aacount"   			{ 	if ( ( $args[$i+1] -eq $null ) -or ( !($args[$i+1] -match "^[\d\.]+$" ) ) )
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
			"-cqcount"   			{  	if ( ( $args[$i+1] -eq $null ) -or ( !($args[$i+1] -match "^[\d\.]+$" ) ) )
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
			"-download"				{ 	$Download = $true }
			"-excelfile" 			{ 	$ExcelFilename = $args[$i+1]
										$i++
									}
			"-help"      			{ 	$Help = $true }	   
			"-noautoattendants"		{ 	$NoAutoAttendants = $true }
			"-noholidays"			{ 	$NoHolidays = $true }
			"-nocallqueues"			{ 	$NoCallQueues = $true }
			"-nophonenumbers"		{ 	$NoPhoneNumbers = $true }
			"-noresourceaccounts" 	{ 	$NoResourceAccounts = $true }
			"-nousers"				{ 	$NoUsers = $true }
			"-noteams"				{ 	$NoTeams = $true }
			"-noopen"				{ 	$NoOpen = $true }
			"-verbose"   			{ 	$Verbose = $true }
			"-view"      			{ 	$View = $true }	  
			Default      			{ 	$ArgError = $true
										$arg = $args[$i]
									}
		}
	}
	
	if ( $Download )
	{
		$NoResourceAccounts = $false
		$NoCallQueues = $false
		$NoPhoneNumbers = $false
		$NoTeams = $false
		$NoUsers = $false
		$NoAutoAttendants = $false
		$NoHolidays = $false
		
		$Verbose = $true
	}
}

#
# Check for arg errors or invalid combinations
#
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

if ( ! $NoHolidays -and $NoAutoAttendants )
{
	Write-Host "Conflicting parameters. -NoAutoAttendants specified and Holidays requested. Holidays requires AutoAttendants. -NoHolidays has been set." -f Yellow
	$NoHolidays = $true
}

if ( $NoResourceAccounts -and $NoAutoAttendants -and $NoHolidays -and $NoCallQueues -and $NoPhoneNumbers -and $NoTeams -and $NoUsers )
{
	Write-Host "All options disabled. Nothing to do. Processing has been halted." -f Red
	$ArgError = $true
}

if ( ( $Help ) -or ( $ArgError ) )
{
	Write-Host ""
	Write-Host "The following options are avaialble:"
	Write-Host "`t-AACount <n> - provide to limit processing to the first AACount Auto Attendants otherwise all in tenant."
	Write-Host "`t-CQCount <n> - provide to limit processing to the first CQCount Call Queues otherwise all in tenant."
	Write-Host "`t-Download - download all Auto Attendant configuration, including audio file. WARNING - may take a long time."
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkAAs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)."	
	Write-Host "`t-NoAutoAttendants - do not download existing auto attendant information."
	Write-Host "`t-NoHolidays - do not download existing holiday information."	
	Write-Host "`t-NoCallQueues - do not download existing call queue information."
	Write-Host "`t-NoPhoneNumbers - do not download existing Voice Applications phone number information."
	Write-Host "`t-NoResourceAccounts - do not download existing resource account information."
	Write-Host "`t-NoUsers - do not download existing EV enabled user information."	
	Write-Host "`t-NoTeams - do not download existing Teams information."
	Write-Host "`t-NoOpen - do not open spreadsheet when script is finished."
	Write-Host "`t-Verbose - provides extra messaging during the process."
	Write-Host "`t-View - watch the spreadsheet as the script modifies it."
	exit
}

$StartTime = (Get-Date -Format "HH:mm:ss")
Write-Host "$StartTime - Starting BulkAAsPreparation." -f Green

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Module Min Supported Versions
#
$MicrosoftTeamsMinVersion = [version]"6.7.0"
$MicrosoftGraphMinVersion = [version]"2.24.0"
$ImportExcelMinVersion = [version]"7.8.0"

#
# BulkAAs.xlms Required Version
#
$ExcelSpreadsheetRequiredVersion = "1.0.4"
$ExcelSpreadsheetVersionSheet = "Data"
$ExcelSpreadsheetVersionRowRef = 2
$ExcelSpreadsheetVersionColRef = 91

#
# Set range variables - used for clearing existing data
#
$Range_ResourceAccounts = "A2:A2001"
$Range_AutoAttendants = "A2:A2001"
$Range_AutoAttendantsDownload = "A5:ZZ20002"
$Range_Holidays = "A2:CW501"
$Range_CallQueues = "A2:A2001"
$Range_PhoneNumbers = "A2:A2001"
$Range_TeamsChannels = "A2:A2001"
$Range_Users = "A2:A2001"

#
# Set row ranges and colours for AutoAttendantsDownload formatting
#
# Colours: 
# https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
# https://www.excelsupersite.com/what-are-the-56-colorindex-colors-in-excel/
#
$Range_AutoAttendantsDownload_Start = "A"

$Range_AutoAttendantsDownload_Base_End = "BY"
$Range_AutoAttendantsDownload_Base_Colour = 45

$Range_AutoAttendantsDownload_BusinessHours_End = "BL"
$Range_AutoAttendantsDownload_BusinessHours_Colour = 40

$Range_AutoAttendantsDownload_CallFlow_End = "HZ"
$Range_AutoAttendantsDownload_CallFlow_Base_Colour = 50
$Range_AutoAttendantsDownload_CallFlow_Default_Colour = 34
$Range_AutoAttendantsDownload_CallFlow_AfterHours_Colour = 35
$Range_AutoAttendantsDownload_CallFlow_Holidays_Colour = 42
$Range_AutoAttendantsDownload_AutoAttendant_Finished_Colour = 48

#
# Set types, declare global variables and assign defaults
#
$global:AudioFileExportPath = $null
[int]$RowOffset = 0
[int]$ScrollWindow = 20


#
# Check that required modules are installed - install if not
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
   Import-Module -Name ImportExcel -MinimumVersion 7.8.0
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
   $ExcelFilename = "BulkAAs.xlsm"
}
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename

Write-Host "Opening the $ExcelFilename worksheet (this may take some time, please be patient)." -f Green

#
# check if supplied filename exists
#
if ( !( Test-Path -Path $ExcelFullPathFilename ) )
{
	Write-Host "`tERROR: $ExcelFilename does not exist." -f Red
	exit
}

#
# Check if file is already open
#
Write-Host "`tChecking that $ExcelFilename is not already open."
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
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_ResourceAccounts).Value = ""

	Write-Host "Retrieving list of Resource Accounts." -f Green

	$ResourceAccounts = @(Get-CsOnlineApplicationInstance | Where {$_.ApplicationId -ne ""} | Sort-Object ApplicationId, DisplayName)
	
	if ( $Verbose )
	{	
		Write-Host ( "`tTotal number of Resource Accounts: {0,4}" -f $ResourceAccounts.length )
	}
	
	Write-Host "`tRetrieving list of Resource Account User Details (this takes some time, please be patient)."

	$ResourceAccountUserDetails = @()
	
	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		$ResourceAccountUserDetails += (Get-CsOnlineUser -Identity $ResourceAccounts[$i].ObjectId)
	}

	Write-Host "Processing list of Resource Accounts." -f Green

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
				"11cd3e2e-fccb-42ad-ad00-878b93575e07"	{	# Call Queue
															if ( $Verbose )
															{
																Write-Host ( "`t({0,4}/{1,4}) [RA-CQ] Resource Account ({2,-36}): {3,-50}" -f ($i + 1), $ResourceAccounts.length, $ResourceAccounts[$i].ObjectId, $ResourceAccounts[$i].DisplayName)
															}
															
															# request will generate a "Correlation id for this request" message when the RA is not assigned to anything, also generates error so redirecting that to null
															$ResourceAccountPriority = ((Get-CsOnlineApplicationInstanceAssociation -identity $ResourceAccounts[$i].ObjectId).CallPriority 2> $null )
															$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ("[RA-CQ] - " + $ResourceAccounts[$i].DisplayName + "~" + $ResourceAccounts[$i].ObjectId + "~" +$ResourceAccounts[$i].PhoneNumber + "~" + $ResourceAccountUserDetails[$i].UsageLocation  + "~" + $ResourceAccountPriority )
														}
				default									{	# Other - if it's blank we don't want it
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
	# Blank out existing
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

			$CallQueues = @(Get-CsCallQueue -Skip $j -First 100 3> $null )
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
}
else
{
	Write-Host "Downloading Call Queues skipped." -f Yellow
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

	$PhoneNumbers = @(Get-CsPhoneNumberAssignment -CapabilitiesContain "VoiceApplicationAssignment" -PstnAssignmentStatus "Unassigned")
	
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

		$ExcelWorkSheet.Cells.Item($RowOffset++, 1) = ($PhoneNumbers[$i].TelephoneNumber + "~" + $PhoneNumbers[$i].NumberType + "~" + $PhoneNumbers[$i].IsoSubdivision + "~" + $PhoneNumbers[$i].IsoCountryCode)

		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading voice applications phone numbers skipped." -f Yellow
}

#
# Team Channels
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Teams")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoTeams )
{
	
	#
	# Blank out remaining rows
	#
	$ExcelWorkSheet.Range($Range_TeamsChannels).Value = ""

	Write-Host "Retrieving list of existing Teams and Channels." -f Green

	$Teams = @(Get-Team | Sort-Object DisplayName)
	
	if ( $Verbose )
	{
		Write-Host "`tTotal number of Teams: " $Teams.length
	}

	Write-Host "Processing list of Teams." -f Green

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
			if ( $TeamsChannels.length -eq 1 )
			{
				# Write-Host ([string]$row + " : " + $Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id + "~" + $TeamsChannels.DisplayName)
				$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($Teams[$i].GroupId + "~" + $Teams[$i].DisplayName + "~" + $TeamsChannels.Id + "~" + $TeamsChannels.DisplayName)
			}
			else
			{
				# Write-Host ([string]$row + " : " + $Teams.GroupId[$i] + "~" + $Teams.DisplayName[$i] + "~" + $TeamsChannels.Id[$j] + "~" + $TeamsChannels.DisplayName[$j])
				$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($Teams[$i].GroupId + "~" + $Teams[$i].DisplayName + "~" + $TeamsChannels[$j].Id + "~" + $TeamsChannels[$j].DisplayName)
			}
			ExcelScroll $RowOffset $ScrollWindow
		}
	}
	ExcelScroll 2 0
}
else
{
	Write-Host "Downloading Teams skipped." -f Yellow
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
# Auto Attendants
#
if ( ! $NoAutoAttendants )
{
	Write-Host "Retrieving list of Auto Attendants." -f Green

	$HolidayScheduleData = @()
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
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_AutoAttendants).Value = ""
	
	$RowOffset = 2
	ExcelScroll $RowOffset 0

	for ($k=0; $k -lt  $AutoAttendants.length; $k++)
	{
		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}/{1,4}) Auto Attendant ({2,-36}) : {3,-50}" -f ($k + 1), $AutoAttendants.length, $AutoAttendants[$k].Identity, $AutoAttendants[$k].Name )
		}

		$AssignedResourceAccounts = $AutoAttendants[$k].ApplicationInstances
		$ExcelWorkSheet.Cells.Item($RowOffset++,1) = ($AutoAttendants[$k].Name + "~" + $AutoAttendants[$k].Identity + "~" + ($AssignedResourceAccounts -join ","))

		if ( ! $NoHolidays )
		{
			# get holidays 
			$CallHandlingAssociations = @($AutoAttendants[$k].CallHandlingAssociations)
			if ( $CallHandlingAssociations.Type -match "Holiday" )
			{
				for ( $l=0; $l -lt $CallHandlingAssociations.length; $l++ )
				{
					switch ( $CallHandlingAssociations[$l].Type )
					{
						"Holiday"	{ 	$HolidayScheduleID = $CallHandlingAssociations[$l].ScheduleId
						
										$Schedules = $AutoAttendants[$k].Schedules
										for ( $m=0; $m -lt $Schedules.length; $m++ )
										{
											if ( $Schedules[$m].Id -eq $HolidayScheduleID )
											{
												$HolidayScheduleData += [PSCustomObject]@{ScheduleID = $HolidayScheduleID; ScheduleName = $Schedules[$m].Name; `
												StartDateTime00 = $Schedules[$m].FixedSchedule.DateTimeRanges[0].Start; EndDateTime00 = $Schedules[$m].FixedSchedule.DateTimeRanges[0].End; `
												StartDateTime01 = $Schedules[$m].FixedSchedule.DateTimeRanges[1].Start; EndDateTime01 = $Schedules[$m].FixedSchedule.DateTimeRanges[1].End; `
												StartDateTime02 = $Schedules[$m].FixedSchedule.DateTimeRanges[2].Start; EndDateTime02 = $Schedules[$m].FixedSchedule.DateTimeRanges[2].End; `
												StartDateTime03 = $Schedules[$m].FixedSchedule.DateTimeRanges[3].Start; EndDateTime03 = $Schedules[$m].FixedSchedule.DateTimeRanges[3].End; `
												StartDateTime04 = $Schedules[$m].FixedSchedule.DateTimeRanges[4].Start; EndDateTime04 = $Schedules[$m].FixedSchedule.DateTimeRanges[4].End; `
												StartDateTime05 = $Schedules[$m].FixedSchedule.DateTimeRanges[5].Start; EndDateTime05 = $Schedules[$m].FixedSchedule.DateTimeRanges[5].End; `
												StartDateTime06 = $Schedules[$m].FixedSchedule.DateTimeRanges[6].Start; EndDateTime06 = $Schedules[$m].FixedSchedule.DateTimeRanges[6].End; `
												StartDateTime07 = $Schedules[$m].FixedSchedule.DateTimeRanges[7].Start; EndDateTime07 = $Schedules[$m].FixedSchedule.DateTimeRanges[7].End; `
												StartDateTime08 = $Schedules[$m].FixedSchedule.DateTimeRanges[8].Start; EndDateTime08 = $Schedules[$m].FixedSchedule.DateTimeRanges[8].End; `
												StartDateTime09 = $Schedules[$m].FixedSchedule.DateTimeRanges[9].Start; EndDateTime09 = $Schedules[$m].FixedSchedule.DateTimeRanges[9].End; `
												
												StartDateTime10 = $Schedules[$m].FixedSchedule.DateTimeRanges[10].Start; EndDateTime10 = $Schedules[$m].FixedSchedule.DateTimeRanges[10].End; `
												StartDateTime11 = $Schedules[$m].FixedSchedule.DateTimeRanges[11].Start; EndDateTime11 = $Schedules[$m].FixedSchedule.DateTimeRanges[11].End; `
												StartDateTime12 = $Schedules[$m].FixedSchedule.DateTimeRanges[12].Start; EndDateTime12 = $Schedules[$m].FixedSchedule.DateTimeRanges[12].End; `
												StartDateTime13 = $Schedules[$m].FixedSchedule.DateTimeRanges[13].Start; EndDateTime13 = $Schedules[$m].FixedSchedule.DateTimeRanges[13].End; `
												StartDateTime14 = $Schedules[$m].FixedSchedule.DateTimeRanges[14].Start; EndDateTime14 = $Schedules[$m].FixedSchedule.DateTimeRanges[14].End; `
												StartDateTime15 = $Schedules[$m].FixedSchedule.DateTimeRanges[15].Start; EndDateTime15 = $Schedules[$m].FixedSchedule.DateTimeRanges[15].End; `
												StartDateTime16 = $Schedules[$m].FixedSchedule.DateTimeRanges[16].Start; EndDateTime16 = $Schedules[$m].FixedSchedule.DateTimeRanges[16].End; `
												StartDateTime17 = $Schedules[$m].FixedSchedule.DateTimeRanges[17].Start; EndDateTime17 = $Schedules[$m].FixedSchedule.DateTimeRanges[17].End; `
												StartDateTime18 = $Schedules[$m].FixedSchedule.DateTimeRanges[18].Start; EndDateTime18 = $Schedules[$m].FixedSchedule.DateTimeRanges[18].End; `
												StartDateTime19 = $Schedules[$m].FixedSchedule.DateTimeRanges[19].Start; EndDateTime19 = $Schedules[$m].FixedSchedule.DateTimeRanges[19].End; `

												StartDateTime20 = $Schedules[$m].FixedSchedule.DateTimeRanges[20].Start; EndDateTime20 = $Schedules[$m].FixedSchedule.DateTimeRanges[20].End; `
												StartDateTime21 = $Schedules[$m].FixedSchedule.DateTimeRanges[21].Start; EndDateTime21 = $Schedules[$m].FixedSchedule.DateTimeRanges[21].End; `
												StartDateTime22 = $Schedules[$m].FixedSchedule.DateTimeRanges[22].Start; EndDateTime22 = $Schedules[$m].FixedSchedule.DateTimeRanges[22].End; `
												StartDateTime23 = $Schedules[$m].FixedSchedule.DateTimeRanges[23].Start; EndDateTime23 = $Schedules[$m].FixedSchedule.DateTimeRanges[23].End; `
												StartDateTime24 = $Schedules[$m].FixedSchedule.DateTimeRanges[24].Start; EndDateTime24 = $Schedules[$m].FixedSchedule.DateTimeRanges[24].End; `
												StartDateTime25 = $Schedules[$m].FixedSchedule.DateTimeRanges[25].Start; EndDateTime25 = $Schedules[$m].FixedSchedule.DateTimeRanges[25].End; `
												StartDateTime26 = $Schedules[$m].FixedSchedule.DateTimeRanges[26].Start; EndDateTime26 = $Schedules[$m].FixedSchedule.DateTimeRanges[26].End; `
												StartDateTime27 = $Schedules[$m].FixedSchedule.DateTimeRanges[27].Start; EndDateTime27 = $Schedules[$m].FixedSchedule.DateTimeRanges[27].End; `
												StartDateTime28 = $Schedules[$m].FixedSchedule.DateTimeRanges[28].Start; EndDateTime28 = $Schedules[$m].FixedSchedule.DateTimeRanges[28].End; `
												StartDateTime29 = $Schedules[$m].FixedSchedule.DateTimeRanges[29].Start; EndDateTime29 = $Schedules[$m].FixedSchedule.DateTimeRanges[29].End; `

												StartDateTime30 = $Schedules[$m].FixedSchedule.DateTimeRanges[30].Start; EndDateTime30 = $Schedules[$m].FixedSchedule.DateTimeRanges[30].End; `
												StartDateTime31 = $Schedules[$m].FixedSchedule.DateTimeRanges[31].Start; EndDateTime31 = $Schedules[$m].FixedSchedule.DateTimeRanges[31].End; `
												StartDateTime32 = $Schedules[$m].FixedSchedule.DateTimeRanges[32].Start; EndDateTime32 = $Schedules[$m].FixedSchedule.DateTimeRanges[32].End; `
												StartDateTime33 = $Schedules[$m].FixedSchedule.DateTimeRanges[33].Start; EndDateTime33 = $Schedules[$m].FixedSchedule.DateTimeRanges[33].End; `
												StartDateTime34 = $Schedules[$m].FixedSchedule.DateTimeRanges[34].Start; EndDateTime34 = $Schedules[$m].FixedSchedule.DateTimeRanges[34].End; `
												StartDateTime35 = $Schedules[$m].FixedSchedule.DateTimeRanges[35].Start; EndDateTime35 = $Schedules[$m].FixedSchedule.DateTimeRanges[35].End; `
												StartDateTime36 = $Schedules[$m].FixedSchedule.DateTimeRanges[36].Start; EndDateTime36 = $Schedules[$m].FixedSchedule.DateTimeRanges[36].End; `
												StartDateTime37 = $Schedules[$m].FixedSchedule.DateTimeRanges[37].Start; EndDateTime37 = $Schedules[$m].FixedSchedule.DateTimeRanges[37].End; `
												StartDateTime38 = $Schedules[$m].FixedSchedule.DateTimeRanges[38].Start; EndDateTime38 = $Schedules[$m].FixedSchedule.DateTimeRanges[38].End; `
												StartDateTime39 = $Schedules[$m].FixedSchedule.DateTimeRanges[39].Start; EndDateTime39 = $Schedules[$m].FixedSchedule.DateTimeRanges[39].End; `

												StartDateTime40 = $Schedules[$m].FixedSchedule.DateTimeRanges[40].Start; EndDateTime40 = $Schedules[$m].FixedSchedule.DateTimeRanges[40].End; `
												StartDateTime41 = $Schedules[$m].FixedSchedule.DateTimeRanges[41].Start; EndDateTime41 = $Schedules[$m].FixedSchedule.DateTimeRanges[41].End; `
												StartDateTime42 = $Schedules[$m].FixedSchedule.DateTimeRanges[42].Start; EndDateTime42 = $Schedules[$m].FixedSchedule.DateTimeRanges[42].End; `
												StartDateTime43 = $Schedules[$m].FixedSchedule.DateTimeRanges[43].Start; EndDateTime43 = $Schedules[$m].FixedSchedule.DateTimeRanges[43].End; `
												StartDateTime44 = $Schedules[$m].FixedSchedule.DateTimeRanges[44].Start; EndDateTime44 = $Schedules[$m].FixedSchedule.DateTimeRanges[44].End; `
												StartDateTime45 = $Schedules[$m].FixedSchedule.DateTimeRanges[45].Start; EndDateTime45 = $Schedules[$m].FixedSchedule.DateTimeRanges[45].End; `
												StartDateTime46 = $Schedules[$m].FixedSchedule.DateTimeRanges[46].Start; EndDateTime46 = $Schedules[$m].FixedSchedule.DateTimeRanges[46].End; `
												StartDateTime47 = $Schedules[$m].FixedSchedule.DateTimeRanges[47].Start; EndDateTime47 = $Schedules[$m].FixedSchedule.DateTimeRanges[47].End; `
												StartDateTime48 = $Schedules[$m].FixedSchedule.DateTimeRanges[48].Start; EndDateTime48 = $Schedules[$m].FixedSchedule.DateTimeRanges[48].End; `
												StartDateTime49 = $Schedules[$m].FixedSchedule.DateTimeRanges[49].Start; EndDateTime49 = $Schedules[$m].FixedSchedule.DateTimeRanges[49].End }
											}
										}
									}
						Default		{ continue }
					}
				}
			}
		}
		ExcelScroll $RowOffset $ScrollWindow
	}
	ExcelScroll 2 0
	
	if ( $Download )
	{
		$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("AutoAttendantsDownload")

		if ( $View )
		{
			$ExcelWorkSheet.Activate()
		}

		#
		# Blank out existing rows
		#
		$ExcelWorkSheet.Range($Range_AutoAttendantsDownload).Value = ""
		$ExcelWorkSheet.Range($Range_AutoAttendantsDownload).Interior.ColorIndex = 2
		
		# Update formatting of first 3 rows
		$BaseRow1Range = "A1:" + $Range_AutoAttendantsDownload_Base_End + "1"
		$BaseRow2Range = "A2:" + $Range_AutoAttendantsDownload_BusinessHours_End + "2"
		$BaseRow3Range = "A3:" + $Range_AutoAttendantsDownload_CallFlow_End + "3"
		
		$ExcelWorkSheet.Range($BaseRow1Range).Interior.ColorIndex = $Range_AutoAttendantsDownload_Base_Colour 
		$ExcelWorkSheet.Range($BaseRow2Range).Interior.ColorIndex = $Range_AutoAttendantsDownload_BusinessHours_Colour
		$ExcelWorkSheet.Range($BaseRow3Range).Interior.ColorIndex = $Range_AutoAttendantsDownload_CallFlow_Base_Colour
		
		Write-Host "Processing list of Auto Attendants for download." -f Green
		
		$RowOffset = 4
		ExcelScroll $RowOffset 0
				
		for ( $i = 0; $i -lt $AutoAttendants.length; $i++ )
		{
			$RowOffset += 1
			$ColumnOffset = 2
			
			$AA = $AutoAttendants[$i]

			if ( $Verbose )
			{
				Write-Host ( "`t({0,4}/{1,4}) Auto Attendant ({2,-36}) : {3,-50}" -f ($i + 1), $AutoAttendants.length, $AA.Identity, $AA.Name )
				Write-Host ( "`t`tBase Configuration" )
			}

			# First Row - Base Config
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Base"
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.LanguageId
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.VoiceId
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.TimeZoneId
			
			if ( $AA.Operator -ne $null )
			{
				switch ( $AA.Operator.Type )
				{
					"User"					{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User"
											}
					"ConfigurationEndpoint"	{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndpoint"
											}
					"ApplicationEndpoint"	{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndpoint"
											}
					"ExternalPstn"			{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ExternalPstn"
											}
					default					{	$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Operator.Type }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Operator.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $AA.Operator.Id "User")
				
				# These fields aren't used by Operator Transfer right now - leaving them in for future
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Operator.EnableTranscription
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Operator.EnableSharedVoicemailSystemPromptSuppression
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Operator.CallPriority
			}
			else
			{
				$ColumnOffset += 6
			}
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.VoiceResponseEnabled
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.UserNameExtension
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.MainlineAttendantEnabled
			
			for ( $j = 0; $j -lt $AA.AuthorizedUsers.length; $j++ )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.AuthorizedUsers[$j].Guid
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $AA.AuthorizedUsers[$j].Guid "User")
			}
			$ColumnOffset += ( 30 - ($j * 2) ) # Max 15 auth users x 2 fields per user
			
			
			for ( $j = 0; $j -lt $AA.HideAuthorizedUsers.length; $j++ )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.HideAuthorizedUsers[$j].Guid
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $AA.HideAuthorizedUsers[$j].Guid "User")
			}
			$ColumnOffset += ( 30 - ($j * 2) ) # Max 15 hidden auth users x 2 fields per user
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = ($AA.ApplicationInstances).length
			
			$ColumnOffset += 9 # Number of spare columns 
			
			for ( $j = 0; $j -lt $AA.ApplicationInstances.length; $j++ )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.ApplicationInstances[$j]
			}
			
			# First Row - Base Config - Format
			$Range_AutoAttendantsDownload_Base = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_Base_End + $RowOffset
			$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_Base).Interior.ColorIndex = $Range_AutoAttendantsDownload_Base_Colour 



			# Second Row - Business Hours
			$RowOffset += 1
			$ColumnOffset = 2

			if ( $Verbose )
			{
				Write-Host ( "`t`tBusiness Hours" )
			}
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "BusinessHours"
			
			for ( $j = 0; $j -lt $AA.CallHandlingAssociations.length; $j++ )
			{
				if ( $AA.CallHandlingAssociations[$j].Type -eq "AfterHours" )
				{
					$AfterHoursScheduleID = $AA.CallHandlingAssociations[$j].ScheduleId
					
					$AfterHoursScheduleIndex = $null
					
					for ( $k = 0; $k -lt $AA.Schedules.length; $k++ )
					{
						if ( $AA.Schedules[$k].Id -eq $AfterHoursScheduleID )
						{
							$AfterHoursScheduleIndex = $k
							Break
						}
					}
					Break
				}
			}
			
			if ( $AfterHoursScheduleIndex -ne $null )
			{
				$AfterHoursSchedule = $AA.Schedules[$AfterHoursScheduleIndex]
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursSchedule.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursSchedule.Name
				
				switch ( $AfterHoursSchedule.Type )
				{
					"Fixed"				{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Fixed" }
					
					"WeeklyRecurrence"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "WeeklyRecurrence" }
					default 			{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursSchedule.Type }
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursSchedule.WeeklyRecurrentSchedule.ComplementEnabled
				
				# Monday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Hours -lt 10 )
						{
							$Monday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Monday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Minutes -lt 10 )
						{
							$Monday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Monday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].Start.Minutes.ToString()
						}
						$Monday_Start = $Monday_Start_Hour + ":" + $Monday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Hours -lt 10 )
						{
							$Monday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Monday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Minutes -lt 10 )
						{
							$Monday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Monday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.MondayHours[$j].End.Minutes.ToString()
						}
						$Monday_End = $Monday_End_Hour + ":" + $Monday_End_Minute
					}
					else
					{
						$Monday_Start = "SM-" + $j.ToString()
						$Monday_End = "EM-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Monday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Monday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Tuesday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Hours -lt 10 )
						{
							$Tuesday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Tuesday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Minutes -lt 10 )
						{
							$Tuesday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Tuesday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].Start.Minutes.ToString()
						}
						$Tuesday_Start = $Tuesday_Start_Hour + ":" + $Tuesday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Hours -lt 10 )
						{
							$Tuesday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Tuesday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Minutes -lt 10 )
						{
							$Tuesday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Tuesday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.TuesdayHours[$j].End.Minutes.ToString()
						}
						$Tuesday_End = $Tuesday_End_Hour + ":" + $Tuesday_End_Minute
					}
					else
					{
						$Tuesday_Start = "ST-" + $j.ToString()
						$Tuesday_End = "ET-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Tuesday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Tuesday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Wednesday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Hours -lt 10 )
						{
							$Wednesday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Wednesday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Minutes -lt 10 )
						{
							$Wednesday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Wednesday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].Start.Minutes.ToString()
						}
						$Wednesday_Start = $Wednesday_Start_Hour + ":" + $Wednesday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Hours -lt 10 )
						{
							$Wednesday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Wednesday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Minutes -lt 10 )
						{
							$Wednesday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Wednesday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.WednesdayHours[$j].End.Minutes.ToString()
						}
						$Wednesday_End = $Wednesday_End_Hour + ":" + $Wednesday_End_Minute
					}
					else
					{
						$Wednesday_Start = "SW-" + $j.ToString()
						$Wednesday_End = "EW-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Wednesday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Wednesday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Thursday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Hours -lt 10 )
						{
							$Thursday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Thursday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Minutes -lt 10 )
						{
							$Thursday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Thursday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].Start.Minutes.ToString()
						}
						$Thursday_Start = $Thursday_Start_Hour + ":" + $Thursday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Hours -lt 10 )
						{
							$Thursday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Thursday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Minutes -lt 10 )
						{
							$Thursday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Thursday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.ThursdayHours[$j].End.Minutes.ToString()
						}
						$Thursday_End = $Thursday_End_Hour + ":" + $Thursday_End_Minute
					}
					else
					{
						$Thursday_Start = "STh-" + $j.ToString()
						$Thursday_End = "ETh-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Thursday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Thursday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Friday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Hours -lt 10 )
						{
							$Friday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Friday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Minutes -lt 10 )
						{
							$Friday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Friday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].Start.Minutes.ToString()
						}
						$Friday_Start = $Friday_Start_Hour + ":" + $Friday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Hours -lt 10 )
						{
							$Friday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Friday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Minutes -lt 10 )
						{
							$Friday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Friday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.FridayHours[$j].End.Minutes.ToString()
						}
						$Friday_End = $Friday_End_Hour + ":" + $Friday_End_Minute
					}
					else
					{
						$Friday_Start = "SF-" + $j.ToString()
						$Friday_End = "EF-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Friday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Friday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Saturday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Hours -lt 10 )
						{
							$Saturday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Saturday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Minutes -lt 10 )
						{
							$Saturday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Saturday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].Start.Minutes.ToString()
						}
						$Saturday_Start = $Saturday_Start_Hour + ":" + $Saturday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Hours -lt 10 )
						{
							$Saturday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Saturday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Minutes -lt 10 )
						{
							$Saturday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Saturday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.SaturdayHours[$j].End.Minutes.ToString()
						}
						$Saturday_End = $Saturday_End_Hour + ":" + $Saturday_End_Minute
					}
					else
					{
						$Saturday_Start = "SSa-" + $j.ToString()
						$Saturday_End = "ESa-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Saturday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Saturday_End
				}
				$ColumnOffset += 2 * ( 4 - $j )

				# Sunday
				for ( $j = 0; $j -lt $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours.length; $j++)
				{
					if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Hours -ne $null )
					{
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Hours -lt 10 )
						{
							$Sunday_Start_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Hours.ToString()
						}
						else
						{
							$Sunday_Start_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Minutes -lt 10 )
						{
							$Sunday_Start_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Minutes.ToString()
						}
						else
						{
							$Sunday_Start_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].Start.Minutes.ToString()
						}
						$Sunday_Start = $Sunday_Start_Hour + ":" + $Sunday_Start_Minute

						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Hours -lt 10 )
						{
							$Sunday_End_Hour = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Hours.ToString()
						}
						else
						{
							$Sunday_End_Hour = $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Hours.ToString()
						}
						
						if ( $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Minutes -lt 10 )
						{
							$Sunday_End_Minute = "0" + $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Minutes.ToString()
						}
						else
						{
							$Sunday_End_Minute = $AfterHoursSchedule.WeeklyRecurrentSchedule.SundayHours[$j].End.Minutes.ToString()
						}
						$Sunday_End = $Sunday_End_Hour + ":" + $Sunday_End_Minute
					}
					else
					{
						$Sunday_Start = "SSu-" + $j.ToString()
						$Sunday_End = "ESu-" + $j.ToString()
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Sunday_Start
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $Sunday_End
				}
			}
			$ColumnOffset += 2 * ( 4 - $j )

			# Second Row - Business Hours - Format
			$Range_AutoAttendantsDownload_BusinessHours = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_BusinessHours_End + $RowOffset
			$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_BusinessHours).Interior.ColorIndex = $Range_AutoAttendantsDownload_BusinessHours_Colour
			

							
			# Third Row - Default Call Flow
			$RowOffset += 1
			$ColumnOffset = 2

			if ( $Verbose )
			{
				Write-Host ( "`t`tDefault Call Flow" )
			}
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "DefaultCallFlow"
			
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Id
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Name
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.ForceListenMenuEnabled

			switch ( $AA.DefaultCallFlow.Greetings.ActiveType )
			{
				"TextToSpeech"	{
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Greetings.TextToSpeechPrompt
									$ColumnOffset += 2
								}
				"AudioFile"		{ 
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
									$ColumnOffset += 1
									
									if ( ( $AA.DefaultCallFlow.Greetings.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AA.DefaultCallFlow.Greetings.AudioFilePrompt.Id.length -ne 0 ) )
									{
										if ( AudioFileExport "Default Call Flow Greeting" $AA.DefaultCallFlow.Greetings.AudioFilePrompt.FileName $AA.DefaultCallFlow.Greetings.AudioFilePrompt.Id $AA.Identity $AA.Name $AA.DefaultCallFlow.Id $AA.DefaultCallFlow.Name)
										{
											$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Greetings.AudioFilePrompt.Id
											$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AA.DefaultCallFlow.Greetings.AudioFilePrompt.FileName) | Out-Null
										}
										else
										{
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Greetings.AudioFilePrompt.Id
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Greetings.AudioFilePrompt.FileName
										}
									}
									else
									{
										$ColumnOffset += 2
									}
								}
				default			{ 
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Greetings.ActiveType
									$ColumnOffset += 3
								}
			}

			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.Name
			$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.DialByNameEnabled
			
			if ( $AA.DefaultCallFlow.Menu.DirectorySearchMethod -eq 1 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ByName"
			}
			else
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.DirectorySearchMethod
			}

			switch ( $AA.DefaultCallFlow.Menu.Prompts.ActiveType )
			{
				"TextToSpeech"	{
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.Prompts.TextToSpeechPrompt
									$ColumnOffset += 2
								}
				"AudioFile"		{ 
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
									$ColumnOffset += 1
									
									if ( ( $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.Id.length -ne 0 ) )
									{
										if ( AudioFileExport "Default Call Flow Menu Prompt" $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.FileName $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.Id $AA.Identity $AA.Name $AA.DefaultCallFlow.Id $AA.DefaultCallFlow.Name)
										{
											$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath											
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.Id
											$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.FileName) | Out-Null
										}
										else
										{
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.Id
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Menu.Prompts.AudioFilePrompt.FileName
										}
									}
									else
									{
										$ColumnOffset += 2
									}
								}
				default			{ 
									$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.Prompts.ActiveType
									$ColumnOffset += 3
								}
			}

			for ( $j = 0; $j -lt $AA.DefaultCallFlow.Menu.MenuOptions.length; $j++ )
			{
				switch ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].DtmfResponse )
				{
					"Automatic"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Automatic" }
					"Tone0"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone0" }
					"Tone1"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone1" }
					"Tone2"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone2" }
					"Tone3"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone3" }
					"Tone4"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone4" }
					"Tone5"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone5" }
					"Tone6"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone6" }
					"Tone7"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone7" }
					"Tone8"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone8" }
					"Tone9"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone9" }
					"TonePound"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TonePound" }
					"ToneStar"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ToneStar" }
					default		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].DtmfResponse }
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].VoiceResponses
				
				switch ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].Action )
				{
					"DisconnectCall"			{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "DisconnectCall" }
					"TransferCallToTarget"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToTarget" }
					"TransferCallToOperator"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToOperator" }
					"Announcement"				{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Announcement" }
					default						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].Action }
				}
				
				switch ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.Type )
				{
					"User"					{ 
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" 
											}
					"ConfigurationEndpoint"	{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndpoint"
											}
					"ApplicationEndpoint"	{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndpoint"
											}
					"ExternalPstn"			{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ExternalPstn"
											}
					"SharedVoicemail"			{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
											}
					default 				{ 
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.Type 
											}
				}
				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.Id $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.Type)
				
				# These fields aren't used by all menu options right now - leaving them in for future
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.EnableTranscription
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.EnableSharedVoicemailSystemPromptSuppression
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].CallTarget.CallPriority
				
				switch ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.ActiveType )
				{
					"TextToSpeech"	{
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.TextToSpeechPrompt
										$ColumnOffset += 2
									}
					"AudioFile"		{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
										$ColumnOffset += 1
										
										if ( ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id.length -ne 0 ) )
										{
											if ( AudioFileExport "Default Call Flow Menu MenuOptions[$j] Announcement" $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id $AA.Identity $AA.Name $AA.DefaultCallFlow.Id $AA.DefaultCallFlow.Name)
											{
												$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id
												$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName) | Out-Null
											}
											else
											{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName
											}
										}
										else
										{
											$ColumnOffset += 2
										}
									}
					default			{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].Prompt.ActiveType
										$ColumnOffset += 3
									}
				}

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].Description
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].MainlineAttendantTarget
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].AgentTargetType
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].AgentTarget
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.DefaultCallFlow.Menu.MenuOptions[$j].AgentTargetTagTemplateId
			}
				
			# Third Row - Default Call Flow - Format
			$Range_AutoAttendantsDownload_CallFlow = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_CallFlow_End + $RowOffset
			$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_CallFlow).Interior.ColorIndex = $Range_AutoAttendantsDownload_CallFlow_Default_Colour
		


			# Forth Row - After Hours
			$AfterHoursCallFlowID = $null
			
			for ( $j = 0; $j -lt $AA.CallHandlingAssociations.length; $j++ )
			{
				if ( $AA.CallHandlingAssociations[$j].Type -eq "AfterHours" )
				{
					$AfterHoursCallFlowID = $AA.CallHandlingAssociations[$j].CallFlowId
					Break
				}
			}
			
			if ( $AfterHoursCallFlowID -ne $null )
			{
				$RowOffset += 1
				$ColumnOffset = 2

				for ( $j = 0; $j -lt $AA.CallFlows.length; $j++ )
				{
					if ( $AA.CallFlows[$j].Id -eq $AfterHoursCallFlowID )
					{
						$AfterHoursCallFlow = $AA.CallFlows[$j]
						Break
					}
				}
				
				if ( $Verbose )
				{
					Write-Host ( "`t`tAfter Hours Call Flow" )
				}
						
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AfterHoursCallFlow"

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Id
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Name
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.ForceListenMenuEnabled

				switch ( $AfterHoursCallFlow.Greetings.ActiveType )
				{
					"TextToSpeech"	{
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Greetings.TextToSpeechPrompt
										$ColumnOffset += 2
									}
					"AudioFile"		{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
										$ColumnOffset += 1
										
										if ( ( $AfterHoursCallFlow.Greetings.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AfterHoursCallFlow.Greetings.AudioFilePrompt.Id.length -ne 0 ) )
										{
											if ( AudioFileExport "After Hours Call Flow Greeting" $AfterHoursCallFlow.Greetings.AudioFilePrompt.FileName $AfterHoursCallFlow.Greetings.AudioFilePrompt.Id $AA.Identity $AA.Name $AfterHoursCallFlow.Id $AfterHoursCallFlow.Name)
											{
												$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath												
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Greetings.AudioFilePrompt.Id
												$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AfterHoursCallFlow.Greetings.AudioFilePrompt.FileName) | Out-Null
											}
											else
											{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Greetings.AudioFilePrompt.Id
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Greetings.AudioFilePrompt.FileName
											}
										}
										else
										{
											$ColumnOffset += 2
										}
									}
					default			{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Greetings.ActiveType
										$ColumnOffset += 3
									}
				}

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.Name
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.DialByNameEnabled
				
				if ( $AfterHoursCallFlow.Menu.DirectorySearchMethod -eq 1 )
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ByName"
				}
				else
				{
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.DirectorySearchMethod
				}

				switch ( $AfterHoursCallFlow.Menu.Prompts.ActiveType )
				{
					"TextToSpeech"	{
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.Prompts.TextToSpeechPrompt
										$ColumnOffset += 2
									}
					"AudioFile"		{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
										$ColumnOffset += 1
										
										if ( ( $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.Id.length -ne 0 ) )
										{
											if ( AudioFileExport "After Hours Menu Prompt" $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.FileName $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.Id $AA.Identity $AA.Name $AfterHoursCallFlow.Id $AfterHoursCallFlow.Name)
											{
												$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.Id
												$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.FileName) | Out-Null
											}
											else
											{
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.Id
												$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Menu.Prompts.AudioFilePrompt.FileName
											}
										}
										else
										{
											$ColumnOffset += 2
										}
									}
					default			{ 
										$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.Prompts.ActiveType
										$ColumnOffset += 3
									}
				}

				for ( $j = 0; $j -lt $AfterHoursCallFlow.Menu.MenuOptions.length; $j++ )
				{
					switch ( $AfterHoursCallFlow.Menu.MenuOptions[$j].DtmfResponse )
					{
						"Automatic"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Automatic" }
						"Tone0"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone0" }
						"Tone1"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone1" }
						"Tone2"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone2" }
						"Tone3"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone3" }
						"Tone4"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone4" }
						"Tone5"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone5" }
						"Tone6"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone6" }
						"Tone7"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone7" }
						"Tone8"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone8" }
						"Tone9"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone9" }
						"TonePound"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TonePound" }
						"ToneStar"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ToneStar" }
						default		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].DtmfResponse }
					}
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].VoiceResponses
					
					switch ( $AfterHoursCallFlow.Menu.MenuOptions[$j].Action )
					{
						"DisconnectCall"			{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "DisconnectCall" }
						"TransferCallToTarget"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToTarget" }
						"TransferCallToOperator"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToOperator" }
						"Announcement"				{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Announcement" }
						default						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].Action }
					}
					
					switch ( $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.Type )
					{
						"User"					{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" 
												}
						"ConfigurationEndpoint"	{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndpoint"
												}
						"ApplicationEndpoint"	{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndpoint"
												}
						"ExternalPstn"			{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ExternalPstn"
												}
						"SharedVoicemail"			{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
												}
						default 				{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.Type 
												}
					}
					
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.Id
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.Id $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.Type)
					
					# These fields aren't used by all menu options right now - leaving them in for future
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.EnableTranscription
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.EnableSharedVoicemailSystemPromptSuppression
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].CallTarget.CallPriority
					
					switch ( $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.ActiveType )
					{
						"TextToSpeech"	{
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.TextToSpeechPrompt
											$ColumnOffset += 2
										}
						"AudioFile"		{ 
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
											$ColumnOffset += 1
											
											if ( ( $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName.length -ne 0 ) -and ( $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id.length -ne 0 ) )
											{
												if ( AudioFileExport "After Hours Menu MenuOptions[$j] Announcement" $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id $AA.Identity $AA.Name $AfterHoursCallFlow.Id $AfterHoursCallFlow.Name)
												{
													$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id
													$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName) | Out-Null
												}
												else
												{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.Id
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.AudioFilePrompt.FileName
												}
											}
											else
											{
												$ColumnOffset += 2
											}
										}
						default			{ 
											$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].Prompt.ActiveType
											$ColumnOffset += 3
										}
					}

					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].Description
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].MainlineAttendantTarget
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].AgentTargetType
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].AgentTarget
					$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursCallFlow.Menu.MenuOptions[$j].AgentTargetTagTemplateId
				}
			}

			# Forth Row - After Hours - Format
			$Range_AutoAttendantsDownload_CallFlow = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_CallFlow_End + $RowOffset
			$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_CallFlow).Interior.ColorIndex = $Range_AutoAttendantsDownload_CallFlow_AfterHours_Colour

			
			
			# Fifth Row - Holiday Schedule
			# Sixth Row - Holiday Call Flow
			# Repeated as a pair for each holiday
			
			for ( $j = 0; $j -lt $AA.CallHandlingAssociations.length; $j++ )
			{
				if ( $AA.CallHandlingAssociations[$j].Type -eq "Holiday" )
				{
					$HolidaysScheduleID = $AA.CallHandlingAssociations[$j].ScheduleID
					$HolidaysCallFlowID = $AA.CallHandlingAssociations[$j].CallFlowId
					
					for ( $k = 0; $k -lt $AA.Schedules.length; $k++ )
					{
						if ( $AA.Schedules[$k].Id -eq $HolidaysScheduleID )
						{
							$RowOffset += 1
							$ColumnOffset = 2
					
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "HolidaysSchedule"
					
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Schedules[$k].Id
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Schedules[$k].Name
							
							switch ( $AA.Schedules[$k].Type )
							{
								"Fixed"				{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Fixed" }
								"WeeklyRecurrence"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "WeeklyRecurrence" }
								default 			{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AfterHoursSchedule.Type }
							}
							
							
							# Fifth+ Row - Holiday Schedule
							$Range_AutoAttendantsDownload_CallFlow = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_CallFlow_End + $RowOffset
							$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_CallFlow).Interior.ColorIndex = $Range_AutoAttendantsDownload_CallFlow_Holidays_Colour
							
							Break
						}
					}
					
					for ( $k = 0; $k -lt $AA.CallFlows.length; $k++ )
					{
						if ( $AA.CallFlows[$k].Id -eq $HolidaysCallFlowID )
						{
							$HolidaysCallFlow = $AA.CallFlows[$k]
							
							$RowOffset += 1
							$ColumnOffset = 2
							
							if ( $Verbose )
							{
								Write-Host ( "`t`tHoliday Call Flow : $($HolidaysCallFlow.Name)" )
							}
			
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Identity
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $AA.Name
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "HolidaysCallFlow"
							
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Id
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Name
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.ForceListenMenuEnabled

							switch ( $HolidaysCallFlow.Greetings.ActiveType )
							{
								"TextToSpeech"	{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Greetings.TextToSpeechPrompt
													$ColumnOffset += 2
												}
								"AudioFile"		{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
													$ColumnOffset += 1
													
													if ( ( $HolidaysCallFlow.Greetings.AudioFilePrompt.FileName.length -ne 0 ) -and ( $HolidaysCallFlow.Greetings.AudioFilePrompt.Id.length -ne 0 ) )
													{
														if ( AudioFileExport "Holiday Call Flow Greeting" $HolidaysCallFlow.Greetings.AudioFilePrompt.FileName $HolidaysCallFlow.Greetings.AudioFilePrompt.Id $AA.Identity $AA.Name $HolidaysCallFlow.Id $HolidaysCallFlow.Name)
														{
															$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Greetings.AudioFilePrompt.Id
															$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $HolidaysCallFlow.Greetings.AudioFilePrompt.FileName) | Out-Null
														}
														else
														{
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Greetings.AudioFilePrompt.Id
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Greetings.AudioFilePrompt.FileName
														}
													}
													else
													{
														$ColumnOffset += 2
													}
												}
								default			{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Greetings.ActiveType
													$ColumnOffset += 3
												}
							}

							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.Name
							$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.DialByNameEnabled
							
							if ( $HolidaysCallFlow.Menu.DirectorySearchMethod -eq 1 )
							{
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ByName"
							}
							else
							{
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.DirectorySearchMethod
							}

							switch ( $HolidaysCallFlow.Menu.Prompts.ActiveType )
							{
								"TextToSpeech"	{
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.Prompts.TextToSpeechPrompt
													$ColumnOffset += 2
												}
								"AudioFile"		{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
													$ColumnOffset += 1
													
													if ( ( $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.FileName.length -ne 0 ) -and ( $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.Id.length -ne 0 ) )
													{
														if ( AudioFileExport "Holiday Menu Prompt" $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.FileName $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.Id $AA.Identity $AA.Name $HolidaysCallFlow.Id $HolidaysCallFlow.Name)
														{
															$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.Id
															$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.FileName) | Out-Null
														}
														else
														{
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.Id
															$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Menu.Prompts.AudioFilePrompt.FileName
														}
													}
													else
													{
														$ColumnOffset += 2
													}
												}
								default			{ 
													$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.Prompts.ActiveType
													$ColumnOffset += 3
												}
							}

							for ( $l = 0; $l -lt $HolidaysCallFlow.Menu.MenuOptions.length; $l++ )
							{
								switch ( $HolidaysCallFlow.Menu.MenuOptions[$l].DtmfResponse )
								{
									"Automatic"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Automatic" }
									"Tone0"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone0" }
									"Tone1"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone1" }
									"Tone2"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone2" }
									"Tone3"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone3" }
									"Tone4"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone4" }
									"Tone5"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone5" }
									"Tone6"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone6" }
									"Tone7"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone7" }
									"Tone8"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone8" }
									"Tone9"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Tone9" }
									"TonePound"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TonePound" }
									"ToneStar"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ToneStar" }
									default		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].DtmfResponse }
								}
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].VoiceResponses
								
								switch ( $HolidaysCallFlow.Menu.MenuOptions[$l].Action )
								{
									"DisconnectCall"			{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "DisconnectCall" }
									"TransferCallToTarget"		{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToTarget" }
									"TransferCallToOperator"	{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TransferCallToOperator" }
									"Announcement"				{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "Announcement" }
									default						{ $ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].Action }
								}
								
								switch ( $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.Type )
								{
									"User"					{ 
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "User" 
															}
									"ConfigurationEndpoint"	{
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ConfigurationEndpoint"
															}
									"ApplicationEndpoint"	{
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ApplicationEndpoint"
															}
									"ExternalPstn"			{
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ExternalPstn"
															}
									"SharedVoicemail"			{
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "SharedVoicemail"
															}
									default 				{ 
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.Type 
															}
								}
								
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.Id
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = (TargetIDLookup $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.Id $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.Type)
								
								# These fields aren't used by all menu options right now - leaving them in for future
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.EnableTranscription
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.EnableSharedVoicemailSystemPromptSuppression
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].CallTarget.CallPriority
								
								switch ( $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.ActiveType )
								{
									"TextToSpeech"	{
														$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "TextToSpeech"
														$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.TextToSpeechPrompt
														$ColumnOffset += 2
													}
									"AudioFile"		{ 
														$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "AudioFile"
														$ColumnOffset += 1
														
														if ( ( $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.FileName.length -ne 0 ) -and ( $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.Id.length -ne 0 ) )
														{
															if ( AudioFileExport "Holiday Menu MenuOptions[$j] Announcement" $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.FileName $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.Id $AA.Identity $AA.Name $HolidaysCallFlow.Id $HolidaysCallFlow.Name)
															{
																$AudioFileExportPathURL = "file://" + $global:AudioFileExportPath
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.Id
																$ExcelWorkSheet.Hyperlinks.Add($ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++), $AudioFileExportPathURL, "", "Play file", $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.FileName) | Out-Null
															}
															else
															{
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.Id
																$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = "ERROR-" + $HolidaysCallFlow.Menu.MenuOptions[$l].Prompt.AudioFilePrompt.FileName
															}
														}
														else
														{
															$ColumnOffset += 2
														}
													}
									default			{ 
														$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$j].Prompt.ActiveType
														$ColumnOffset += 3
													}
								}

								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].Description
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].MainlineAttendantTarget
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].AgentTargetType
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].AgentTarget
								$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidaysCallFlow.Menu.MenuOptions[$l].AgentTargetTagTemplateId
							}

							# Sixth+ Row - Holiday CallFlow - Format
							$Range_AutoAttendantsDownload_CallFlow = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_CallFlow_End + $RowOffset
							$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_CallFlow).Interior.ColorIndex = $Range_AutoAttendantsDownload_CallFlow_Holidays_Colour
							
							Break
							
						} # If holiday call flow match
					} # for loop checking call flow 
				} # If call handling associations
			} # for loop checking call handling associations
			
			# Blank Row -  Format
			$RowOffset += 1
			$Range_AutoAttendantsDownload_CallFlow = $Range_AutoAttendantsDownload_Start + $RowOffset + ":" + $Range_AutoAttendantsDownload_CallFlow_End + $RowOffset
			$ExcelWorkSheet.Range($Range_AutoAttendantsDownload_CallFlow).Interior.ColorIndex = $Range_AutoAttendantsDownload_AutoAttendant_Finished_Colour
			
			ExcelScroll $RowOffset $RowOffset
		}
	} # Download
	ExcelScroll 5 0
}
else
{
	Write-Host "Downloading Auto Attendants skipped." -f Yellow
}

#
# Auto Attendant Holidays
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Holidays")

if ( $View )
{
   $ExcelWorkSheet.Activate()
}

if ( ! $NoHolidays )
{
		
	#
	# Blank out existing
	#
	$ExcelWorkSheet.Range($Range_Holidays).Value = ""

	Write-Host "Retrieving list of assigned Auto Attendant Holidays." -f Green

	if ( $Verbose )
	{
		Write-Host ( "`tTotal number of holidays: {0,4}" -f $HolidayScheduleData.length )
	}

	Write-Host "Processing list of Auto Attendant Holidays." -f Green

	$RowOffset = 2	
	ExcelScroll $RowOffset 0
	
	$HolidayScheduleIDsProcessed = @()
		
	for ( $i = 0; $i -lt $HolidayScheduleData.length; $i++ )
	{
		if ( ($HolidayScheduleIDsProcessed -match $HolidayScheduleData[$i].ScheduleId).length -eq 0 )
		{
			$HolidayScheduleIDsProcessed += $HolidayScheduleData[$i].ScheduleId
		
			if ( $HolidayScheduleData[$i].StartDateTime00 -ne $null )
			{
				if ( $Verbose )
				{                  
					Write-Host "`tHoliday : " $HolidayScheduleData[$i].ScheduleName
				}
				
				$ColumnOffset = 1

				$ExcelWorkSheet.Cells.Item($RowOffset,  $ColumnOffset++) = $HolidayScheduleData[$i].ScheduleId
				$ExcelWorkSheet.Cells.Item($RowOffset,  $ColumnOffset++) = "[E] $($HolidayScheduleData[$i].ScheduleName)"

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime00
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime00

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime01
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime01

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime02
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime02

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime03
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime03

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime04
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime04

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime05
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime05

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime06
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime06

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime07
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime07

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime08
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime08

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime09
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime09

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime10
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime10

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime11
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime11

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime12
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime12

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime13
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime13

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime14
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime14

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime15
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime15

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime16
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime16

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime17
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime17

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime18
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime18

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime19
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime19

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime20
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime20

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime21
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime21

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime22
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime22

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime23
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime23

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime24
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime24

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime25
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime25

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime26
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime26

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime27
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime27

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime28
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime28

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime29
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime29

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime30
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime30

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime31
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime31

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime32
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime32

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime33
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime33

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime34
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime34

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime35
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime35

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime36
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime36

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime37
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime37

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime38
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime38

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime39
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime39

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime40
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime40

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime41
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime41

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime42
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime42

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime43
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime43

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime44
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime44

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime45
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime45

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime46
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime46

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime47
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime47

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime48
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime48

				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].StartDateTime49
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData[$i].EndDateTime49

				$RowOffset++
				ExcelScroll $RowOffset $ScrollWindow
			} # if not null
			
		} # if not matched
		else
		{
			if ( $Verbose )
			{                  
				Write-Host "`tHoliday : " $HolidayScheduleData[$i].ScheduleName "- Duplicate - Not Processed" -f Yellow
			}
		}
	} # for loop
	ExcelScroll 2 0
	
	if ( $Verbose )
	{
		Write-Host " "
		Write-Host "`tHolidays duplicated : " ( $HolidayScheduleData.length - $HolidayScheduleIDsProcessed.length )
		Write-Host "`tHolidays processed  : " $HolidayScheduleIDsProcessed.length
	}
}
else
{
	Write-Host "Downloading Auto Attendant Holidays skipped." -f Yellow
}

#
# Save and close the Excel file
#
if ( $View )
{
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Config-BusinessHours")
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
	Write-Host "Please complete the configuration, save and exit the spreadsheet and then run the BulkAAProvisioning script."
	Write-Host -NoNewLine "Press any key to continue..."
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

	Invoke-item "$ExcelFullPathFilename"
}
else
{
	Write-Host "$EndTime - Preparation complete. " -f Green
	Write-Host "Duration: $DurationFormatted" -f Green
}

