Write-Host "Starting BulkAuthUsersPrep."
Write-Host "Cleaning up from any previous runs."

if ( Test-Path -path ".\List-Policies.csv" )
{
   Remove-Item -Path ".\List-Policies.csv" | Out-Null
}

if ( Test-Path -path ".\List-EV-Users.csv" )
{
   Remove-Item -Path ".\List-EV-Users.csv" | Out-Null
}

if ( Test-Path -path ".\List-AAs.csv" )
{
   Remove-Item -Path ".\List-AAs.csv" | Out-Null
}

if ( Test-Path -path ".\List-CQs.csv" )
{
   Remove-Item -Path ".\List-CQs.csv" | Out-Null
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
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
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
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
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
      Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
   }
   try
   { 
      Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
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




$ExcelFilename = "BulkAuthUsers.xlsx"
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename
Write-Host "Accessing the $ExcelFilename worksheet (this may take some time, please be patient)."

$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open($ExcelFullPathFilename)


#
# Policies
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Policies")
Write-Host "Retrieving existing policies."
(Get-CsTeamsVoiceApplicationsPolicy).Identity | Out-File -FilePath .\List-Policies.csv

$Policies = @(Import-csv -Path .\List-Policies.csv)

for ( $i = 0; $i -lt $Policies.length; $i++)
{
   $ExcelWorkSheet.Cells.Item($i + 2,1) = $Policies[$i].Global
}

if ( Test-Path -path ".\List-Policies.csv" )
{
   Remove-Item -Path ".\List-Policies.csv" | Out-Null
}

#
# Users
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Users")
Write-Host "Getting list of enterprise voice enabled users."
#get-csonlineuser | where {$_.EnterpriseVoiceEnabled} | where {$_.AccountEnabled} | Sort-Object Alias | Export-csv -Path .\List-EV-Users.csv
get-csonlineuser -Filter {EnterpriseVoiceEnabled -eq $true -and AccountEnabled -eq $true} | Sort-Object Alias | Export-csv -Path .\List-EV-Users.csv

$Users = @(Import-csv -Path .\List-EV-Users.csv)


for ( $i = 0; $i -lt $Users.length; $i++)
{
   $ExcelWorkSheet.Cells.Item($i + 2,1) = ($Users.UserPrincipalName[$i] + "~" + $Users.Identity[$i])
}

if ( Test-Path -path ".\List-EV-Users.csv" )
{
   Remove-Item -Path ".\List-EV-Users.csv" | Out-Null
}


#
# Auto Attendants
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Auto Attendants")
Write-Host "Getting list of Auto Attendants."

$arguments = (Get-PSCallStack).Arguments
$parms = $arguments[0] -split ","

$parm1 = $parms[0] -split "-"
$parm2 = $parms[1] -split "}"

$parm3 = $parms[2] -split "-"
$parm4 = $parms[3] -split "}"

$aaCount = 0
$cqCount = 0

if ( $parm1 -eq "-AACount" )
{
   $aaCount = [int]$parm2[0]
   $cqCount = [int]$parm4[0]
}
else
{
   $aaCount = [int]$parm4[0]
   $cqCount = [int]$parm2[0]
}

if ( $aaCount -gt 0 )
{
   $loops = [int] [Math]::Truncate($aaCount / 100) + 1
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
   Remove-Item -Path ".\List-AA-$j.csv"
}

$AutoAttendants = @(Import-csv -Path .\List-AAs.csv)

for ($i=0; $i -lt  $AutoAttendants.length; $i++)
{
   $ExcelWorkSheet.Cells.Item($i + 2,1) = ($AutoAttendants.Name[$i] + "~" + $AutoAttendants.Identity[$i])
}

if ( Test-Path -path ".\List-AAs.csv" )
{
   Remove-Item -Path ".\List-AAs.csv"
}

#
# Call Queues
#
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Call Queues")
Write-Host "Getting list of Call Queues."

if ( $cqCount -gt 0 )
{
   $loops = [int] [Math]::Truncate($cqCount / 100) + 1
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
   Remove-Item ".\List-CQ-$j.csv"
}

$CallQueues = @(Import-csv -Path .\List-CQs.csv)

for ($i=0; $i -lt  $CallQueues.length; $i++)
{
   $ExcelWorkSheet.Cells.Item($i + 2,1) = ($CallQueues.Name[$i] + "~" + $CallQueues.Identity[$i])
}

if ( Test-Path -path ".\List-CQs.csv" )
{
   Remove-Item -Path ".\List-CQs.csv"
}

#
# Save and close the Excel file
#
$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)
$ExcelObj.Quit()


Write-Host "Preparation complete.  Opening $ExcelFilename.  Please complete the configuration, save and exit the spreadsheet and then run the BulkAuthUserConfig script."
Write-Host -NoNewLine "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')


Invoke-item $ExcelFullPathFilename