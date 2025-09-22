# Auto Attendant Bulk Provisioning


## Assumptions

1.	PowerShell v5.x is installed on your computer.
    - Issue the command: $PSVersionTable.PSVersion
    - PowerShell 7.x is NOT supported
1.	Enterprise Voice Enabled users have been created.
1.	There are enough spare Phone Resource Account Phone Licenses exist to assign to new Resource Accounts 
    - Provisioning of resource accounts and assigning licenses can be bypassed if necessary.	

## Required PowerShell Modules
Missing or outdated modules will be automatically installed or updated.

- MicrosoftTeams - 6.7.0 or later
- Microsoft.Graph - 2.24.0
- ImportExcel - 7.8.0

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

The BulkAAsProvisioning.ps1 PowerShell script requests the following Microsoft Graph scopes:
  - Organization.Read.All
  - User.ReadWrite.All

Note: These permissions are requested only if Resource Account creation and licensing is being done

## Known Issues

### BulkAAsPreparation.ps1 - v1.0.4

- No known issues

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
- Download Auto Attendant configurations - delivered 2025.09.22 with v1.0.4

### BulkAAsProvisioning.ps1

- Address known issues
- Support text prompts with special characters
- Support dial-by-name/number properly
- Remove duplicate holiday sets
- Self-reference an Auto Attendant created in the spreadsheet
- Self-reference a Resource Account created in the spreadsheet

### BulkAAs spreadsheet

- Address known issues
- Support dial-by-name/number properly
- Detect/alert on same voice command being used for different menu options
