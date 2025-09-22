# Download the BulkAATool

Download the files and sub-directories in the **BulkAATool** folder to a location on your workstation.

Alternatively, from the main **Code** page, select *<> Code* and *Download ZIP* to download the entire repository as a zip file and then unzip the file.

# Preparation Instructions

This step will download the existing auto attendants, call queues, hoildays, phone numbers, resource accounts, Teams channels and, user configurations in the tenant so they can be referenced when provisioning new auto attendants.

1. Open a PowerShell 5.x window
   - Issue the command: $PsVersionTable.PSVersion if not sure
1. In the PowerShell window, run the "BulkAAsPreparation.ps1" script with any of the optional parameters below.	
   - This will prepare and open the BulkAAs spreadsheet.
   

## BulkAAsPreparation.ps1 command line options

| Option              | Description                                        |
|:--------------------|----------------------------------------------------|
| -AACount n          | All Auto Attendants are processed by default.  Use `-AACount n` to restrict the processing to the first n Auto Attendants                  |         
| -CQCount n          | All Call Queues are processed by default.  Use -CQCount n` to restrict the processing to the first n Call Queues                           |
| -Download           | Download all Auto Attendant configuration, including audio files.                                                                          |
| -ExcelFile filename | Specify an alternative Excel spreadsheet to use. Must be in the same directory as the BulkAAsPreparation.ps1 file<br>Default: BulkAAs.xlsm |
| -Help               | This help message.                                                                                                                         |
| -NoResourceAccounts | Do not download existing resource account information.                                                                                     |
| -NoAutoAttendants   | Do not download existing auto attendant information. <br>*Automaticaly enables*  **-NoHolidays**                                           |
| -NoHolidays         | Do not download existing auto attendant holiday information.                                                                               |
| -NoCallQueues       | Do not download existing call queue information.                                                                                           |
| -NoPhoneNumbers     | Do not download existing voice application phone numbers.                                                                                  |
| -NoUsers            | Do not download existing EV enabled users.                                                                                                 |
| -NoTeams            | Do not download existing teams information.                                                                                                |
| -NoOpen             | Do not open the spreadsheet when the BulkAAsPreparation.ps1 script is finished.                                                            |
| -Verbose            | Watch the spreadsheet get filled with information as the BulkAAsPreparation.psl1 script runs.<br>*Automaticaly disables*  **-NoOpen**      | 

# Provisioning Instructions

1. Open the BulkAAs.xlsm spreadsheet and enable macros if they have been disabled.
1. Complete the follows tabs:
   
   - Config-BusinessHours
   - Config-Holidays
   - Config-Base
   - Config-BusinessHoursMenu
   - Config-AfterHoursMenu
   - Config-HolidaysMenu
  
1. Save the BulkAAs.xlsm spreadsheet and close Excel.
1. Place any referenced prompt files in the AudioFiles sub-directory.
1. Open a PowerShell 5.x window
   - Issue the command: $PsVersionTable.PSVersion if not sure
1. In the PowerShell window, run the "BulkAAsProvisioning.ps1" script

## BulkAAsProvisioning.ps1 command line options

| Option                     | Description                                        |
|:---------------------------|----------------------------------------------------|
| -ExcelFile filename        | Specify an alternative Excel spreadsheet to use. Must be in the same directory as the BulkAAsProvisioning.ps1 file<br>Default: BulkAAs.xlsm |
| -Help                      | This help message.                                                                                                                          |
| -NoResourceAccounts        | Do not perform any resource account related steps. <br>*Automaticaly enables*  **-NoResourceAccountCreation**, **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**  |
| -NoResourceAccountCreation | Do not provision any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**  |
| -NoResourceAccountLicensing| Do not license any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountPhoneNumbers**                                     |
| -Verbose                   | Detailed output.                                                                                                                            |
