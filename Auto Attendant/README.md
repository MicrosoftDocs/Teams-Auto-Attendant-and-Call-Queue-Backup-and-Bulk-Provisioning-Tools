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

The PowerShell scripts request the following Microsoft Graph scopes"
  - Organization.Read.All
  - User.ReadWrite.All

Note: At the current time these permissions are requested even if Resource Account creating and licensing is bypassed. This will be addressed in a future version of the script.

## Known Issues

### BulkAAsPreparation.ps1

- Can only retrieve holiday schedules that are assigned to an Auto Attendant
  - This is a system limitation and will not be addressed

### BulkAAsProvisioning.ps1

- It is not possible to self-reference a Resource Account or Auto Autoattendant created in the spreadsheet
- It is not possible to assign multiple resource accounts to an Auto Attendant

### BulkAAs spreadsheet

- Dial Scope should be a Config-Base item however it is shown on Config-BusinessHoursMenu, Config-AfterHoursMenu and, Config-HolidaysMenu.
  - Config-BusinessHoursMenu is the only configuration that will be provisioned. Configuration on other tabs will be ignored
  - A future version will move this to Config-Base
 
- It is highly likely there are some conditional formatting errors. Please report these so they can be addressed.

## Roadmap

### BulkAAsPreparation.ps1

- Update to support MicrosoftTeams 6.9.0
- Remove use of temporary spreadsheets
- Add counters to items in verbose mode
- Stop processing when an invalid parameter is passed

### BulkAAsProvisioning.ps1

- Reference an Auto Attendant created in the spreadsheet

### BulkAAs spreadsheet

- Move Dial Scope to the Config-Base tab
