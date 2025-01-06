Write-Host "Starting BulkAuthUsersConfig."
Write-Host "Cleaning up from any previous runs."

if ( Test-Path -path ".\PS-Policy.csv" )
{
   Remove-Item -Path ".\PS-Policy.csv" | Out-Null
}

if ( Test-Path -path ".\PS-AA.csv" )
{
   Remove-Item -Path ".\PS-AA.csv" | Out-Null
}

if ( Test-Path -path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
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


#
# Policy To User Assignments
#
Write-Host "Processing Policy To User Assignments."
$ExcelWorkSheetName = "PS-Policy"
$ExcelCSVFilename = ".\PS-Policy.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCSVFilename -NoTypeInformation

$PSPolicy = @(Import-csv -Path $ExcelCSVFilename)

for ($i=0; $i -lt  $PSPolicy.length; $i++)
{
   $PolicyName = $PSPolicy.PolicyName[$i]
   $PolicyUserList = $PSPolicy.UserList[$i]

   if ($PolicyName -ne "")
   {
      Write-Host "`tPolicy: $PolicyName"
      Write-Host "`t`tUsers:"

      if ($PolicyUserList -ne "")
      {
         $PolicyUsers = $PolicyUserList -split ","

         for ($j=0; $j -lt $PolicyUsers.length; $j++)
         {
            $PolicyUser = $PolicyUsers[$j]

            Write-Host "`t`t`t$PolicyUser"

            Grant-CsTEamsVoiceApplicationsPolicy -Identity "$PolicyUser" -PolicyName "$PolicyName"
         }
         Write-Host " "
      }
   }
}
Write-Host "Completed Policy To User Assignments."

if ( Test-Path -path ".\PS-Policy.csv" )
{
   Remove-Item -Path ".\PS-Policy.csv" | Out-Null
}


#
# User to Auto Attendant Assignments
#
Write-Host "Processing User to Auto Attendant Assignments."
$ExcelWorkSheetName = "PS-AA"
$ExcelCSVFilename = ".\PS-AA.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCSVFilename -NoTypeInformation

$PSAA = @(Import-csv -Path $ExcelCSVFilename)

for ($i=0; $i -lt  $PSAA.length; $i++)
{
   $AAGUID = $PSAA.AAGUID[$i]
   $AAAuthUserList = $PSAA.AuthUserList[$i]

   if ($AAGUID -ne "")
   {
      Write-Host "`tProcessing: $AAGUID"
      Write-Host "`t`tAuth Users: $AAAuthUserList"

      $AuthorizedUsers = $AAAuthUserList -split ","

      for ($j=0; $j -lt $AuthorizedUsers.length; $j++)
      {
         $AuthorizedUsers[$j] = [System.GUID]::Parse($AuthorizedUsers[$j])
      }
      
      $AA = (get-csautoattendant -identity "$AAGUID")

      $AA.AuthorizedUsers = $AuthorizedUsers

      set-csautoattendant $AA
   }
}
Write-Host "Completed User to Auto Attendant Assignments."

if ( Test-Path -path ".\PS-AA.csv" )
{
   Remove-Item -Path ".\PS-AA.csv" | Out-Null
}


#
# User to Call Queue Assignments
#
Write-Host "Processing User to Call Queue Assignments."
$ExcelWorkSheetName = "PS-CQ"
$ExcelCSVFilename = ".\PS-CQ.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCSVFilename -NoTypeInformation

$PSCQ = @(Import-csv -Path $ExcelCSVFilename)

for ($i=0; $i -lt  $PSCQ.length; $i++)
{
   $CQGUID = $PSCQ.CQGUID[$i]
   $CQAuthUserList = $PSCQ.AuthUserList[$i]

   if ($CQGUID -ne "")
   {
      Write-Host "`tProcessing: $CQGUID"
      Write-Host "`t`tAuth Users: $CQAuthUserList"

      $AuthorizedUsers = $CQAuthUserList -split ","

      for ($j=0; $j -lt $AuthorizedUsers.length; $j++)
      {
         $AuthorizedUsers[$j] = [System.GUID]::Parse($AuthorizedUsers[$j])
      }

      set-cscallqueue -identity $CQGUID -authorizedUsers $AuthorizedUsers | Out-Null
   }
}
Write-Host "Completed User to Call Queue Assignments."

if ( Test-Path -path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
}

Write-Host "Bulk Auth User Configuration Completed."
