# Call Queue Bulk Provisioning and Backup

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
| Resource Accounts: 2000     | Call Queues: 2000     |
| Teams/Channels: 2000        |                       |
| Users: 2000                 |                       |
| Call Queues: 2000           |                       |

Existing Items limits can be increased/removed by updated the spreadsheet accordingly.

## Warnings

Only update information on the following spreadsheet tabs:
  - Config-CallQueue

Changing any of the grey shaded areas on these or any other tabs may result in warnings, failures, inaccurate backups, or provisioning.

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

### BulkCQsPreparation.ps1

- Voice applications phone numbers are not downloaded

### BulkCQsProvisioning

- Can't assigne a phone number to a new resource account
- It is not possible to assign multiple resource accounts to a Call Queue

### BulkCQs spreadsheet

- It is highly likely there are some conditional formatting errors. Please report these so they can be addressed.
