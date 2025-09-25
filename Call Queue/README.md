# Call Queue Bulk Provisioning and Backup

## Assumptions

1.	PowerShell v5.x is installed on your computer.
    - Issue the command: $PSVersionTable.PSVersion
    - PowerShell 7.x is NOT supported
1.	Enterprise Voice Enabled users have been created.
1.	There are enough spare Phone Resource Account Phone Licenses exist to assign to new Resource Accounts 
    - Provisioning of resource accounts and assigning licenses can be bypassed if necessary.	

## Required PowerShell Modules
Missing or outdated modules will be automatically installed or updated.

- MicrosoftTeams - 7.0.0 or later
- Microsoft.Graph - 2.24.0 or later
- ImportExcel - 7.8.0

## Limitations

| Maximum Existing Items      | Maxium Create Items   |
|:----------------------------|:----------------------|
| Resource Accounts: 2000     | Call Queues: 2000     |
| Auto Attendants: 2000       |                       |
| Call Queues: 2000           |                       |
| Phone Numbers: 2000         |                       |
| Teams<ul><li>Channels: 2000</li><li>Schedule Groups: 2000</li></ul>        |                       |
| Users: 2000                 |                       |

Existing Items limits can be increased/removed by updating the spreadsheet accordingly.

## Warnings

Only update information on the following spreadsheet tabs:
  - Config-CallQueue

Changing any of the grey shaded areas on these or any other tabs may result in warnings, failures, inaccurate backups, or provisioning errors.

## Requirements

Open the PowerShell window as an administrator.

>[!CAUTION]
>Testing has only been done with a Teams Global Administrator account.  Less privileged accounts should work provided they have sufficient access.  

To perform resource accounts related activities, when prompted, login with an account that has the necessary permissions:  [Manage resource accounts for service numbers - Microsoft Teams | Microsoft Learn](https://learn.microsoft.com/microsoftteams/manage-resource-accounts#assign-permissions-for-managing-a-resource-account)

## Microsoft Graph Scopes Requested

### BulkCQsPreparation.ps1
  - Schedule.Read.All

Note: This permission is requested only if Teams Schedule Groups are being downloaded.

### BulkCQsProvisioning.ps1
  - Organization.Read.All
  - User.ReadWrite.All

Note: These permissions are requested only if Resource Account creation and licensing is being done.

## Known Issues

### BulkCQsPreparation.ps1

- No known issues

### BulkCQsProvisioning.ps1

- Can't assign a phone number to a new resource account
- It is not possible to assign multiple resource accounts to a Call Queue

### BulkCQs spreadsheet

- **Config-CallQueue** tab
  - No known issues

- It is highly likely there are some conditional formatting errors. Please report these so they can be addressed.

## Roadmap

### BulkCQsPreparation.ps1

- Address known issues
- Investigate Resource Account priority lookup when resource account type is AA
- Introduce configurable limits

### BulkCQsProvisioning.ps1

- Address known issues
- Assign a phone number to a new resource account

### BulkCQs spreadsheet

- Address known issues
