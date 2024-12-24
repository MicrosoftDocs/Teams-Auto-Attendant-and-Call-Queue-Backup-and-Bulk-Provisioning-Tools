# processing arguments
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
			"-aacount"   			{ $AACount = $args[$i+1] }
			"-cqcount"   			{ $CQCount = $args[$i+1] }
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
			Default      			{ Write-Host "Unknown argument passed: $args[$i]" }   
		}
	}
}

if ( $Help )
{
	Write-Host "The following options are avaialble:"
	Write-Host "`t-AACount <n> - the number of Auto Attendants in tenant, only needed if greater than 100"
	Write-Host "`t-CQCount <n> - the number of Call Queues in tenant, only needed if greater than 100"
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkAAs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)"	
	Write-Host "`t-NoResourceAccount - do not download existing resource account information"
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


Write-Host "Starting BulkAAsPrep."
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

if ( Test-Path -Path ".\List-PhoneNumbers-VA.csv" )
{
   Remove-Item -Path ".\List-PhoneNumbers-VA.csv" | Out-Null
}

if ( Test-Path -Path ".\List-EV-Users.csv" )
{
   Remove-Item -Path ".\List-EV-Users.csv" | Out-Null
}

if ( Test-Path -Path ".\List-Teams.csv" )
{
   Remove-Item -Path ".\List-Teams.csv" | Out-Null
}

if ( Test-Path -path ".\List-Holidays.csv" )
{
   Remove-Item -Path ".\List-Holidays.csv" | Out-Null
}


#
# Check that required modules are installed - install if not
#
Write-Host "Checking for MicrosoftTeams module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "MicrosoftTeams" } )
{
   Write-Host "Connecting to Microsoft Teams."
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
else
{
   Write-Host "Module MicrosoftTeams does not exist - installing."
   Install-Module -Name MicrosoftTeams -Force -AllowClobber

   Write-Host "Connecting to Microsoft Teams."
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

Write-Host "Checking for Microsoft.Graph module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "Microsoft.Graph" } )
{
   Write-Host "Connecting to Microsoft Graph."
   Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null

   try
   { 
      # Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null

   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
   }
   try
   { 
      # Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}
else
{
   Write-Host "Module MgGraph does not exist - installing."
   Install-Module -Name Microsoft.Graph -Force -AllowClobber
   Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
   try
   { 
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}

Write-Host "Checking for ImportExcel module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "ImportExcel" } )
{
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel
}
else
{
   Write-Host "Module ImportExcel - installing."
   Install-Module -Name ImportExcel -Force -AllowClobber
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel
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
		$ResourceAccountSoftDeletionTimeStamp = (get-csonlineuser -identity $ResourceAccounts.ObjectId[$i]).SoftDeletionTimeStamp
	
		if ( $ResourceAccountSoftDeletionTimeStamp.length -eq 0 )
		{	
			if ( $ResourceAccounts.ApplicationId[$i] -eq "ce933385-9390-45d1-9512-c8d228074e07" )
			{
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-AA] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" + $ResourceAccounts.PhoneNumber[$i])
			}
			else
			{
				$ExcelWorkSheet.Cells.Item($j,1) = ("[RA-CQ] - " + $ResourceAccounts.DisplayName[$i] + "~" + $ResourceAccounts.ObjectId[$i] + "~" +$ResourceAccounts.PhoneNumber[$i])
			}
			$j++

			if ( $Verbose )
			{
				Write-Host "`tResource Account Added : " $ResourceAccounts.DisplayName[$i]
			}
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
		Get-CsAutoAttendant -Skip $j 3> `$null | Export-csv -Path .\List-AA-$j.csv
	}

	if ( Test-Path -Path ".\`$null" )
	{
		Remove-Item -Path ".\`$null" | Out-Null
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
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($AutoAttendants.Name[$i] + "~" + $AutoAttendants.Identity[$i])

		if ( ! $NoHolidays )
		{
			# get holidays for later processing
			$bytes = ( Export-CsAutoAttendantHolidays -Identity $AutoAttendants.Identity[$i] 3> `$null )
			[System.IO.File]::WriteAllBytes("$PSScriptRoot\List-Holidays-$i.csv", $bytes)
		}
		
		if ( $Verbose )
		{
			Write-Host "`tAuto Attendant : " $AutoAttendants.Name[$i]
		}

	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
	
	if ( Test-Path -Path ".\`$null" )
	{
		Remove-Item -Path ".\`$null" | Out-Null
	}

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
# Auto Attendant Holidays
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Holidays")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoHolidays )
{
	Write-Host "Getting list of existing Holidays."

	$k = 2	
	if ( Test-Path ".\List-Holidays-*.csv" )
	{
		Get-Content Template-List-Holidays.csv, List-Holidays-*.csv | Set-Content List-Holidays.csv

		for ( $i=0; $i -lt $AutoAttendants.length; $i++ )
		{
			Remove-Item -Path ".\List-Holidays-$i.csv" | Out-Null
		}
	

		$Holidays = @(Import-csv -Path ".\List-Holidays.csv")
		
		
		for ( $i = 0; $i -lt $Holidays.length; $i++ )
		{
			if ( $Holidays.StartDateTime1[$i] -ne $null -AND $Holidays.StartDateTime1[$i] -ne "StartDateTime1" )
			{
				if ( $Verbose )
				{
					Write-Host "`tHoliday : " $Holidays.Name[$i]
				}

				$ExcelWorkSheet.Cells.Item($k,  1) = "[E] $($Holidays.Name[$i])"

				# using loop here in order to use break statement when a StartDateTime is null
				# not happy with this approach but it works, reducing processing time down from 1 minute per holiday to seconds when there are only a few holiday dates
				# definitely open to suggestions on improvements
				while ( 1 -eq 1 )
				{
					$ExcelWorkSheet.Cells.Item($k,  2) = $Holidays.StartDateTime1[$i]
					$ExcelWorkSheet.Cells.Item($k,  3) = $Holidays.EndDateTime1[$i]	

					if ( $Holidays.StartDateTime2[$i] -eq $null )
					{
						break
					}				
					$ExcelWorkSheet.Cells.Item($k,  4) = $Holidays.StartDateTime2[$i]
					$ExcelWorkSheet.Cells.Item($k,  5) = $Holidays.EndDateTime2[$i]	

					if ( $Holidays.StartDateTime3[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k,  6) = $Holidays.StartDateTime3[$i]
					$ExcelWorkSheet.Cells.Item($k,  7) = $Holidays.EndDateTime3[$i]

					if ( $Holidays.StartDateTime4[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k,  8) = $Holidays.StartDateTime4[$i]
					$ExcelWorkSheet.Cells.Item($k,  9) = $Holidays.EndDateTime4[$i]

					if ( $Holidays.StartDateTime5[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 10) = $Holidays.StartDateTime5[$i]
					$ExcelWorkSheet.Cells.Item($k, 11) = $Holidays.EndDateTime5[$i]

					if ( $Holidays.StartDateTime6[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 12) = $Holidays.StartDateTime6[$i]
					$ExcelWorkSheet.Cells.Item($k, 13) = $Holidays.EndDateTime6[$i]

					if ( $Holidays.StartDateTime7[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 14) = $Holidays.StartDateTime7[$i]
					$ExcelWorkSheet.Cells.Item($k, 15) = $Holidays.EndDateTime7[$i]

					if ( $Holidays.StartDateTime8[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 16) = $Holidays.StartDateTime8[$i]
					$ExcelWorkSheet.Cells.Item($k, 17) = $Holidays.EndDateTime8[$i]

					if ( $Holidays.StartDateTime9[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 18) = $Holidays.StartDateTime9[$i]
					$ExcelWorkSheet.Cells.Item($k, 19) = $Holidays.EndDateTime9[$i]

					if ( $Holidays.StartDateTime10[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 20) = $Holidays.StartDateTime10[$i]
					$ExcelWorkSheet.Cells.Item($k, 21) = $Holidays.EndDateTime10[$i]

					if ( $Holidays.StartDateTime11[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 22) = $Holidays.StartDateTime11[$i]
					$ExcelWorkSheet.Cells.Item($k, 23) = $Holidays.EndDateTime11[$i]

					if ( $Holidays.StartDateTime12[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 24) = $Holidays.StartDateTime12[$i]
					$ExcelWorkSheet.Cells.Item($k, 25) = $Holidays.EndDateTime12[$i]

					if ( $Holidays.StartDateTime13[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 26) = $Holidays.StartDateTime13[$i]
					$ExcelWorkSheet.Cells.Item($k, 27) = $Holidays.EndDateTime13[$i]

					if ( $Holidays.StartDateTime14[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 28) = $Holidays.StartDateTime14[$i]
					$ExcelWorkSheet.Cells.Item($k, 29) = $Holidays.EndDateTime14[$i]

					if ( $Holidays.StartDateTime15[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 30) = $Holidays.StartDateTime15[$i]
					$ExcelWorkSheet.Cells.Item($k, 31) = $Holidays.EndDateTime15[$i]

					if ( $Holidays.StartDateTime16[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 32) = $Holidays.StartDateTime16[$i]
					$ExcelWorkSheet.Cells.Item($k, 33) = $Holidays.EndDateTime16[$i]

					if ( $Holidays.StartDateTime17[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 34) = $Holidays.StartDateTime17[$i]
					$ExcelWorkSheet.Cells.Item($k, 35) = $Holidays.EndDateTime17[$i]

					if ( $Holidays.StartDateTime18[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 36) = $Holidays.StartDateTime18[$i]
					$ExcelWorkSheet.Cells.Item($k, 37) = $Holidays.EndDateTime18[$i]

					if ( $Holidays.StartDateTime19[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 38) = $Holidays.StartDateTime19[$i]
					$ExcelWorkSheet.Cells.Item($k, 39) = $Holidays.EndDateTime19[$i]

					if ( $Holidays.StartDateTime20[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 40) = $Holidays.StartDateTime20[$i]
					$ExcelWorkSheet.Cells.Item($k, 41) = $Holidays.EndDateTime20[$i]

					if ( $Holidays.StartDateTime21[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 42) = $Holidays.StartDateTime21[$i]
					$ExcelWorkSheet.Cells.Item($k, 43) = $Holidays.EndDateTime21[$i]

					if ( $Holidays.StartDateTime22[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 44) = $Holidays.StartDateTime22[$i]
					$ExcelWorkSheet.Cells.Item($k, 45) = $Holidays.EndDateTime22[$i]

					if ( $Holidays.StartDateTime23[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 46) = $Holidays.StartDateTime23[$i]
					$ExcelWorkSheet.Cells.Item($k, 47) = $Holidays.EndDateTime23[$i]

					if ( $Holidays.StartDateTime24[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 48) = $Holidays.StartDateTime24[$i]
					$ExcelWorkSheet.Cells.Item($k, 49) = $Holidays.EndDateTime24[$i]

					if ( $Holidays.StartDateTime25[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 50) = $Holidays.StartDateTime25[$i]
					$ExcelWorkSheet.Cells.Item($k, 51) = $Holidays.EndDateTime25[$i]

					if ( $Holidays.StartDateTime26[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 52) = $Holidays.StartDateTime26[$i]
					$ExcelWorkSheet.Cells.Item($k, 53) = $Holidays.EndDateTime26[$i]

					if ( $Holidays.StartDateTime27[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 54) = $Holidays.StartDateTime27[$i]
					$ExcelWorkSheet.Cells.Item($k, 55) = $Holidays.EndDateTime27[$i]

					if ( $Holidays.StartDateTime28[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 56) = $Holidays.StartDateTime28[$i]
					$ExcelWorkSheet.Cells.Item($k, 57) = $Holidays.EndDateTime28[$i]

					if ( $Holidays.StartDateTime29[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 58) = $Holidays.StartDateTime29[$i]
					$ExcelWorkSheet.Cells.Item($k, 59) = $Holidays.EndDateTime29[$i]

					if ( $Holidays.StartDateTime30[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 60) = $Holidays.StartDateTime30[$i]
					$ExcelWorkSheet.Cells.Item($k, 61) = $Holidays.EndDateTime30[$i]

					if ( $Holidays.StartDateTime31[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 62) = $Holidays.StartDateTime31[$i]
					$ExcelWorkSheet.Cells.Item($k, 63) = $Holidays.EndDateTime31[$i]

					if ( $Holidays.StartDateTime32[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 64) = $Holidays.StartDateTime32[$i]
					$ExcelWorkSheet.Cells.Item($k, 65) = $Holidays.EndDateTime32[$i]

					if ( $Holidays.StartDateTime33[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 66) = $Holidays.StartDateTime33[$i]
					$ExcelWorkSheet.Cells.Item($k, 67) = $Holidays.EndDateTime33[$i]

					if ( $Holidays.StartDateTime34[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 68) = $Holidays.StartDateTime34[$i]
					$ExcelWorkSheet.Cells.Item($k, 69) = $Holidays.EndDateTime34[$i]

					if ( $Holidays.StartDateTime35[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 70) = $Holidays.StartDateTime35[$i]
					$ExcelWorkSheet.Cells.Item($k, 71) = $Holidays.EndDateTime35[$i]

					if ( $Holidays.StartDateTime36[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 72) = $Holidays.StartDateTime36[$i]
					$ExcelWorkSheet.Cells.Item($k, 73) = $Holidays.EndDateTime36[$i]

					if ( $Holidays.StartDateTime37[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 74) = $Holidays.StartDateTime37[$i]
					$ExcelWorkSheet.Cells.Item($k, 75) = $Holidays.EndDateTime37[$i]

					if ( $Holidays.StartDateTime38[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 76) = $Holidays.StartDateTime38[$i]
					$ExcelWorkSheet.Cells.Item($k, 77) = $Holidays.EndDateTime38[$i]

					if ( $Holidays.StartDateTime39[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 78) = $Holidays.StartDateTime39[$i]
					$ExcelWorkSheet.Cells.Item($k, 79) = $Holidays.EndDateTime39[$i]

					if ( $Holidays.StartDateTime40[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 80) = $Holidays.StartDateTime40[$i]
					$ExcelWorkSheet.Cells.Item($k, 81) = $Holidays.EndDateTime40[$i]

					if ( $Holidays.StartDateTime41[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 82) = $Holidays.StartDateTime41[$i]
					$ExcelWorkSheet.Cells.Item($k, 83) = $Holidays.EndDateTime41[$i]

					if ( $Holidays.StartDateTime42[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 84) = $Holidays.StartDateTime42[$i]
					$ExcelWorkSheet.Cells.Item($k, 85) = $Holidays.EndDateTime42[$i]

					if ( $Holidays.StartDateTime43[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 86) = $Holidays.StartDateTime43[$i]
					$ExcelWorkSheet.Cells.Item($k, 87) = $Holidays.EndDateTime43[$i]

					if ( $Holidays.StartDateTime44[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 88) = $Holidays.StartDateTime44[$i]
					$ExcelWorkSheet.Cells.Item($k, 89) = $Holidays.EndDateTime44[$i]

					if ( $Holidays.StartDateTime45[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 90) = $Holidays.StartDateTime45[$i]
					$ExcelWorkSheet.Cells.Item($k, 91) = $Holidays.EndDateTime45[$i]

					if ( $Holidays.StartDateTime46[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 92) = $Holidays.StartDateTime46[$i]
					$ExcelWorkSheet.Cells.Item($k, 93) = $Holidays.EndDateTime46[$i]

					if ( $Holidays.StartDateTime47[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 94) = $Holidays.StartDateTime47[$i]
					$ExcelWorkSheet.Cells.Item($k, 95) = $Holidays.EndDateTime47[$i]

					if ( $Holidays.StartDateTime48[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 96) = $Holidays.StartDateTime48[$i]
					$ExcelWorkSheet.Cells.Item($k, 97) = $Holidays.EndDateTime48[$i]

					if ( $Holidays.StartDateTime49[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k, 98) = $Holidays.StartDateTime49[$i]
					$ExcelWorkSheet.Cells.Item($k, 99) = $Holidays.EndDateTime49[$i]

					if ( $Holidays.StartDateTime50[$i] -eq $null )
					{
						break
					}
					$ExcelWorkSheet.Cells.Item($k,100) = $Holidays.StartDateTime50[$i]
					$ExcelWorkSheet.Cells.Item($k,101) = $Holidays.EndDateTime50[$i]
					
					break
				}
				$k++
			}
		}
	}
	
	#
	# Blank out remaining rows
	#
	#$Range = "A" + ($i + 2) + ":CW501"
	$Range = "A" + ($k) + ":CW501"
	$ExcelWorkSheet.Range($Range).value = ""
	
	if ( Test-Path -Path ".\List-Holidays.csv" )
	{
		Remove-Item -Path ".\List-Holidays.csv" | Out-Null
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
   $ExcelWorkSheet.activate()
}

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
		Get-CsCallQueue -Skip $j 3> `$null | Export-csv -Path .\List-CQ-$j.csv
	}

	if ( Test-Path -Path ".\`$null" )
	{
		Remove-Item -Path ".\`$null" | Out-Null
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
		Remove-Item -Path ".\List-CQ-$j.csv" | Out-Null
	}

	$CallQueues = @(Import-csv -Path .\List-CQs.csv)

	for ($i=0; $i -lt  $CallQueues.length; $i++)
	{
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($CallQueues.Name[$i] + "~" + $CallQueues.Identity[$i])

		if ( $Verbose )
		{
			Write-Host "`tCall Queue : " $CallQueues.Name[$i]
		}
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
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
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($PhoneNumbers.TelephoneNumber[$i] + "~" + $PhoneNumbers.NumberType[$i] + "~" + $PhoneNumbers.IsoSubdivision[$i] + "~" + $PhoneNumbers.IsoCountryCode[$i])

		if ( $Verbose )
		{
			Write-Host "`tPhone Number : " $PhoneNumbers.TelephoneNumber[$i]
		}
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
		$ExcelWorkSheet.Cells.Item($i + 2,1) = ($Users.UserPrincipalName[$i] + "~" + $Users.Identity[$i])

		if ( $Verbose )
		{
			Write-Host "`tUser : " $Users.UserPrincipalName[$i]
		}
	}

	#
	# Blank out remaining rows
	#
	$Range = "A" + ($i + 2) + ":A2001"
	$ExcelWorkSheet.Range($Range).value = ""
}
else
{
	Write-Host "Downloading EV enabled users skipped."	
}

if ( Test-Path -Path ".\List-EV-Users.csv" )
{
   Remove-Item -Path ".\List-EV-Users.csv" | Out-Null
}

#
# Team Channels
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Teams")

if ( $View )
{
   $ExcelWorkSheet.activate()
}

if ( ! $NoTeams )
{
	Write-Host "Getting list of existing Teams and Channels."
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
	Write-Host "Downloading Teams skipped."	
}

if ( Test-Path -Path ".\List-Teams.csv" )
{
   Remove-Item -Path ".\List-Teams.csv" | Out-Null
}

#
# Save and close the Excel file
#
if ( $View )
{
   $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Config-BusinessHours")
   $ExcelWorkSheet.activate()
}

$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)
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

