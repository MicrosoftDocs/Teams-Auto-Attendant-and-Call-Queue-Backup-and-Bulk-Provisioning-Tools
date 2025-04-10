# Auto Attendant Bulk Provisioning


## Assumptions

1.	PowerShell v5.x is installed on your computer.
    - Issue the command: $PSVersionTable.PSVersion
    - PowerShell 7.x is NOT supported
1.	Enterprise Voice Enabled users have been created.
1.	There are enough spare Phone Resource Account Phone Licenses exist to assign to new Resource Accounts 
    - Provisioning of resource accounts and assigning licenses can be bypassed if necessary.	

## Limitations

| Maximum Existing Items      | Maxium Create Items   |
|:----------------------------|:----------------------|
| Resource Accounts: 2000     | Auto Attendants: 2000 |
| Phone Numbers: 2000         | Business Hours: 50    |
| Teams/Channels: 2000        | Holidays: 50          |
| Users: 2000                 |                       |
| Auto Attendants: 2000       |                       |
| Call Queues: 2000           |                       |
| Holidays: 500               |                       |

Existing Items limits can be increased/removed by updated the spreadsheet accordingly.

## Warnings

Only update information on the following spreadsheet tabs:
  - Config-BusinessHours
  - Config-Holidays
  - Config-Base
  - Config-BusinessHoursMenu
  - Config-AfterHoursMenu
  - Config-HolidaysMenu

Changing any of the grey shaded areas on these or any other tabs may result in warnings, failures or inaccurate backups or provisioning.

## Requirements

Open the PowerShell window as an administrator.

>[!CAUTION]
>Testing has only been done with a Teams Global Administrator account.  Less privileged accounts should work provided they have sufficient access.  

To perform resource accounts related activities, when prompted, login with an account that has the necessary permissions:  [Manage resource accounts for service numbers - Microsoft Teams | Microsoft Learn](https://learn.microsoft.com/microsoftteams/manage-resource-accounts#assign-permissions-for-managing-a-resource-account)

### Microsoft Graph Scopes Requested

The BulkAAsProvisioning.ps1 PowerShell scripts request the following Microsoft Graph scopes:
  - Organization.Read.All
  - User.ReadWrite.All

Note: At the current time these permissions are requested even if Resource Account creating and licensing is bypassed. This will be addressed in a future version of the script.

## Known Issues

### BulkAAsPreparation.ps1

- Can only retrieve holiday schedules that are assigned to an Auto Attendant

### BulkAAsProvisioning.ps1

- It is not possible to self-reference a Resource Account or Auto Autoattendant created in the spreadsheet
- It is not possible to assign multiple resource accounts to an Auto Attendant
- Special characters in text prompts may cause errors

### BulkAAs spreadsheet

- It is possible to self-reference an Auto Attendant created in the spreadsheet however this will cause a failure when the BulkAAsProvisioning.ps1 script is run
- Dial Scope is only honoured on the Config-BusinessHoursMenu tab. Configuration on other tabs will be ignored
- It is highly likely there are some conditional formatting errors. Please report these so they can be addressed.

## Roadmap

### BulkAAsPreparation.ps1

- Address known issues
- Get all holiday sets
- Download Auto Attendant configurations

### BulkAAsProvisioning.ps1

- Address known issues
- Support text prompts with special characters
- Support dial-by-name/number properly
- Remove duplicate holiday sets
- Self-reference an Auto Attendant created in the spreadsheet
- Self-reference a Resource Account created in the spreadsheet
- Do not load Graph library if licensing resource accounts is disabled

### BulkAAs spreadsheet

- Address known issues
- Support dial-by-name/number properly
- Detect/alert on same voice command being used for different menu options
