# Version: 1.0.2
# Date: 2025.04.0x
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
			"-aacount"   			{ $AACount = [int]$args[$i+1]
									  $i++
									}
			"-cqcount"   			{ $CQCount = [int]$args[$i+1] 
									  $i++
									}
			"-excelfile" 			{ $ExcelFilename = $args[$i+1]
									  $i++
									}
			"-help"      			{ $Help = $true }	   
			"-noresourceaccounts" 	{ $NoResourceAccounts = $true }
			"-noautoattendants"		{ $NoAutoAttendants = $true
								      $NoHolidays = $true
									}
			"-noholidays"			{ $NoHolidays = $true }
			"-nocallqueues"			{ $NoCallQueues = $true }
			"-nophonenumbers"		{ $NoPhoneNumbers = $true }
			"-nousers"				{ $NoUsers = $true }
			"-noteams"				{ $NoTeams = $true }
			"-noopen"				{ $NoOpen = $true }
			"-verbose"   			{ $Verbose = $true }
			"-view"      			{ $View = $true }	  
			Default      			{ $ArgError = $true
									  $arg = $args[$i]
									  Write-Host "Unknown argument passed: $arg" }   
		}
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
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkAAs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)"	
	Write-Host "`t-NoResourceAccounts - do not download existing resource account information"
	Write-Host "`t-NoAutoAttendants - do not download existing auto attendant information"
	Write-Host "`t-NoHolidays - do not download existing holiday information"	
	Write-Host "`t-NoCallQueues - do not download existing call queue information"
	Write-Host "`t-NoPhoneNumbers - do not download existing Voice Applications phone number information"	
	Write-Host "`t-NoUsers - do not download existing EV enabled user information"	
	Write-Host "`t-NoTeams - do not download existing Teams information"
	Write-Host "`t-NoOpen - do not open spreadsheet when script is finished"
	Write-Host "`t-Verbose - provides extra messaging during the process"
	Write-Host "`t-View - watch the spreadsheet as the script modifies it"
	exit
}


Write-Host "Starting BulkAAsPreparation."

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Set range variables
#
$Range_ResourceAccounts = "A2:A2001"
$Range_AutoAttendnats = "A2:A2001"
$Range_Holidays = "A2:CW501"
$Range_CallQueues = "A2:A2001"
$Range_PhoneNumbers = "A2:A2001"
$Range_TeamsChannels = "A2:A2001"
$Range_Users = "A2:A2001"

#
# Check that required modules are installed - install if not
#
Write-Host "Checking for MicrosoftTeams module 6.7.0 or later."
$Version = ( (Get-InstalledModule -Name MicrosoftTeams -MinimumVersion "6.7.0").Version 2> $null )
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
   $ExcelFilename = "BulkAAs.xlsm"
}
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename

Write-Host "Accessing the $ExcelFilename worksheet (this may take some time, please be patient)."

#
# check if supplied filename exists
#
if ( !( Test-Path -Path $ExcelFullPathFilename ) )
{
	Write-Host "ERROR: $ExcelFilename does not exist."
	exit
}

$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("$ExcelFullPathFilename")

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

	Write-Host "Getting list of Resource Accounts."

	$ResourceAccounts = @(Get-CsOnlineApplicationInstance | Sort-Object ApplicationId, DisplayName)

	$j = 2
	for ( $i = 0; $i -lt $ResourceAccounts.length; $i++)
	{
		#
		# Make sure resource account is not Deleted
		#
		$ResourceAccountUserDetails = (Get-CsOnlineUser -Identity $ResourceAccounts.ObjectId[$i])
		
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

if ( ! $NoAutoAttendants )
{
	#
	# Blank out existing rows
	#
	$ExcelWorkSheet.Range($Range_AutoAttendnats).Value = ""

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
		elseif ( $j -le ( $AACount - $j ) )
		{
			$First = 100
		}
		else
		{
			$First = $AACount % 100
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

		$HolidayScheduleMeta = @()
		$HolidayScheduleData = @()
		$AutoAttendants = @(Get-CsAutoAttendant -Skip $j -First $First)

		for ($k=0; $k -lt  $AutoAttendants.length; $k++)
		{
			if ( $Verbose )
			{
				Write-Host ( "`t({0,4}) Auto Attendant : {1,-50}" -f ($k + $j + 1), $AutoAttendants.Name[$k] )
			}

			$Row = $k + $j + 2
			$AssignedResourceAccounts = ( (Get-CsAutoAttendant -Identity $AutoAttendants.Identity[$k]).ApplicationInstances 3> $null )
			$ExcelWorkSheet.Cells.Item($Row,1) = ($AutoAttendants.Name[$k] + "~" + $AutoAttendants.Identity[$k] + "~" + ($AssignedResourceAccounts -join ","))

			if ( ! $NoHolidays )
			{
				# get holidays 
				$CallHandlingAssociations = @($AutoAttendants[$k].CallHandlingAssociations)
				if ( $CallHandlingAssociations.Type -match "Holiday" )
				{
					for ( $l=0; $l -lt $CallHandlingAssociations.length; $l++ )
					{
						switch ( $CallHandlingAssociations.Type[$l] )
						{
							"Holiday"	{ 	$HolidayScheduleID = $CallHandlingAssociations[$l].ScheduleId
							
											$Schedules = @($AutoAttendants[$k].Schedules)
											for ( $m=0; $m -lt $Schedules.length; $m++ )
											{
												if ( $Schedules[$m].Id -eq $HolidayScheduleID )
												{
													$HolidayScheduleName = $Schedules[$m].Name
												}
											}
											$HolidayScheduleMeta += [PSCustomObject]@{ScheduleID = $HolidayScheduleID; ScheduleName = $HolidayScheduleName}
										}
							Default		{ continue }
						}
					}
					
					# not optimal but only way I could make this work right now
					$HolidaySchedule = (((Export-CsAutoAttendantHolidays -Identity $AutoAttendants.Identity[$k]) 3> `$null) | ConvertFrom-CSV)	
					if ( $HolidaySchedule.length -gt 0 )
					{
						$HolidayScheduleData += @([System.Text.Encoding]::UTF8.GetString(((Export-CsAutoAttendantHolidays -Identity $AutoAttendants.Identity[$k]) 3> `$null)) | ConvertFrom-CSV)	
					}
				}
			}
		}
	}
}
else
{
	Write-Host "Downloading Auto Attendants skipped."	
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

	Write-Host "Getting list of existing Holidays."

	$RowOffset = 2	
		
	for ( $i = 0; $i -lt $HolidayScheduleData.length; $i++ )
	{
		if ( $HolidayScheduleData.StartDateTime1[$i] -ne $null )
		{
			if ( $Verbose )
			{                  
				Write-Host "`tHoliday : " $HolidayScheduleData.Name[$i]
			}
			
			$ColumnOffset = 1

			$HolidayScheduleMetaIndex = [Array]::IndexOf($HolidayScheduleMeta.ScheduleName, $HolidayScheduleData.Name[$i])
			$ExcelWorkSheet.Cells.Item($RowOffset,  $ColumnOffset++) = $HolidayScheduleMeta[$HolidayScheduleMetaIndex].ScheduleId
			$ExcelWorkSheet.Cells.Item($RowOffset,  $ColumnOffset++) = "[E] $($HolidayScheduleData.Name[$i])"

			# using loop here in order to use break statement when a StartDateTime is null
			# not happy with this approach but it works, reducing processing time down from 1 minute per holiday to seconds when there are only a few holiday dates
			# definitely open to suggestions on improvements
			while ( 1 -eq 1 )
			{
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime1[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime1[$i]	

				if ( $HolidayScheduleData.StartDateTime2[$i] -eq $null )
				{
					break
				}				
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime2[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime2[$i]	

				if ( $HolidayScheduleData.StartDateTime3[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime3[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime3[$i]

				if ( $HolidayScheduleData.StartDateTime4[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime4[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime4[$i]

				if ( $HolidayScheduleData.StartDateTime5[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime5[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime5[$i]

				if ( $HolidayScheduleData.StartDateTime6[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime6[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime6[$i]

				if ( $HolidayScheduleData.StartDateTime7[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime7[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime7[$i]

				if ( $HolidayScheduleData.StartDateTime8[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime8[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime8[$i]

				if ( $HolidayScheduleData.StartDateTime9[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime9[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime9[$i]

				if ( $HolidayScheduleData.StartDateTime10[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime10[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime10[$i]

				if ( $HolidayScheduleData.StartDateTime11[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime11[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime11[$i]

				if ( $HolidayScheduleData.StartDateTime12[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime12[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime12[$i]

				if ( $HolidayScheduleData.StartDateTime13[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime13[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime13[$i]

				if ( $HolidayScheduleData.StartDateTime14[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime14[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime14[$i]

				if ( $HolidayScheduleData.StartDateTime15[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime15[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime15[$i]

				if ( $HolidayScheduleData.StartDateTime16[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime16[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime16[$i]

				if ( $HolidayScheduleData.StartDateTime17[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime17[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime17[$i]

				if ( $HolidayScheduleData.StartDateTime18[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime18[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime18[$i]

				if ( $HolidayScheduleData.StartDateTime19[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime19[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime19[$i]

				if ( $HolidayScheduleData.StartDateTime20[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime20[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime20[$i]

				if ( $HolidayScheduleData.StartDateTime21[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime21[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime21[$i]

				if ( $HolidayScheduleData.StartDateTime22[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime22[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime22[$i]

				if ( $HolidayScheduleData.StartDateTime23[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime23[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime23[$i]

				if ( $HolidayScheduleData.StartDateTime24[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime24[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime24[$i]

				if ( $HolidayScheduleData.StartDateTime25[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime25[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime25[$i]

				if ( $HolidayScheduleData.StartDateTime26[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime26[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime26[$i]

				if ( $HolidayScheduleData.StartDateTime27[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime27[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime27[$i]

				if ( $HolidayScheduleData.StartDateTime28[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime28[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime28[$i]

				if ( $HolidayScheduleData.StartDateTime29[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime29[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime29[$i]

				if ( $HolidayScheduleData.StartDateTime30[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime30[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime30[$i]

				if ( $HolidayScheduleData.StartDateTime31[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime31[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime31[$i]

				if ( $HolidayScheduleData.StartDateTime32[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime32[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime32[$i]

				if ( $HolidayScheduleData.StartDateTime33[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime33[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime33[$i]

				if ( $HolidayScheduleData.StartDateTime34[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime34[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime34[$i]

				if ( $HolidayScheduleData.StartDateTime35[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime35[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime35[$i]

				if ( $HolidayScheduleData.StartDateTime36[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime36[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime36[$i]

				if ( $HolidayScheduleData.StartDateTime37[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime37[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime37[$i]

				if ( $HolidayScheduleData.StartDateTime38[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime38[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime38[$i]

				if ( $HolidayScheduleData.StartDateTime39[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime39[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime39[$i]

				if ( $HolidayScheduleData.StartDateTime40[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime40[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime40[$i]

				if ( $HolidayScheduleData.StartDateTime41[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime41[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime41[$i]

				if ( $HolidayScheduleData.StartDateTime42[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime42[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime42[$i]

				if ( $HolidayScheduleData.StartDateTime43[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime43[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime43[$i]

				if ( $HolidayScheduleData.StartDateTime44[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime44[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime44[$i]

				if ( $HolidayScheduleData.StartDateTime45[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime45[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime45[$i]

				if ( $HolidayScheduleData.StartDateTime46[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime46[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime46[$i]

				if ( $HolidayScheduleData.StartDateTime47[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime47[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime47[$i]

				if ( $HolidayScheduleData.StartDateTime48[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime48[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime48[$i]

				if ( $HolidayScheduleData.StartDateTime49[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime49[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime49[$i]

				if ( $HolidayScheduleData.StartDateTime50[$i] -eq $null )
				{
					break
				}
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.StartDateTime50[$i]
				$ExcelWorkSheet.Cells.Item($RowOffset, $ColumnOffset++) = $HolidayScheduleData.EndDateTime50[$i]
				
				break
			}
			$RowOffset++
		}
	}
}
else
{
	Write-Host "Downloading Auto Attendant holidays skipped."		
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

		if ( $CQCount -le 100 )
		{
			$First = $CQCount
		}
		elseif ( $j -le ( $CQCount - $j ) )
		{
			$First = 100
		}
		else
		{
			$First = $CQCount % 100
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

		$CallQueues = @(Get-CsCallQueue -Skip $j -First 100 3> $null )

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
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($PhoneNumbers.TelephoneNumber[$i] + "~" + $PhoneNumbers.NumberType[$i] + "~" + $PhoneNumbers.IsoSubdivision[$i] + "~" + $PhoneNumbers.IsoCountryCode[$i])

		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}) Phone Number : {1,-50}" -f ($i + 1), $PhoneNumbers.TelephoneNumber[$i] )
		}
	}
}
else
{
	Write-Host "Downloading Voice Applications phone numbers skipped."	
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

	Write-Host "Getting list of existing Teams and Channels."

	$Teams = @(Get-Team | Sort-Object DisplayName)

	$row = 1
	for ($i=0; $i -lt $Teams.length; $i++)
	{
		$TeamsChannels = @(Get-TeamChannel -GroupId $Teams.GroupId[$i] | Where {$_.MembershipType -EQ "Standard"} | Sort-Object DisplayName)

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
}
else
{
	Write-Host "Downloading Teams skipped."	
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

	Write-Host "Getting list of enterprise voice enabled users."

	$Users = @(Get-CsOnlineUser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias)

	for ( $i = 0; $i -lt $Users.length; $i++)
	{
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($Users.UserPrincipalName[$i] + "~" + $Users.Identity[$i])

		if ( $Verbose )
		{
			Write-Host ( "`t({0,4}) User : {1,-50}" -f ($i + 1), $Users.UserPrincipalName[$i] )
		}
	}
}
else
{
	Write-Host "Downloading EV enabled users skipped."	
}

#
# Save and close the Excel file
#
if ( $View )
{
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Config-BusinessHours")
   $ExcelWorkSheet.Activate()
}

$ExcelWorkBook.Save()
$ExcelWorkBook.Close($true)
$ExcelObj.Quit()


if ( ! $NoOpen )
{
	Write-Host "Preparation complete.  Opening $ExcelFilename.  Please complete the configuration, save and exit the spreadsheet and then run the BulkCQsConfig script."
	Write-Host -NoNewLine "Press any key to continue..."
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

	Invoke-item "$ExcelFullPathFilename"
}
else
{
	Write-Host "Preparation complete."
}

