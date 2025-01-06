# Authorized Users Bulk Provisioning

## Assumptions

1.	PowerShell v5.x is installed on your computer.
    - Issue the command: $PSVersionTable.PSVersion
    - PowerShell 7.x is NOT supported
1.  Auto Attendants and Call Queues have been created.
1.	Enterprise Voice Enabled users have been created.
1.	Voice Applications Policies have been created.

## Limitations

| Maximum Existing Items            | Maxium Create Items               |
|:----------------------------------|:----------------------------------|
| Voice Applications Policies: 2000 | Policy to User: 2000              |
| Users: 2000                       | Auth User to Auto Attendant: 2000 |
| Auto Attendants: 2000             | Auth User to Call Queue: 2000     |
| Call Queues: 2000                 |                       |

Existing Items limits can be increased/removed by updated the spreadsheet accordingly.

## Warnings

Only update information on the following spreadsheet tabs:
  - Config-PolicyToUser
  - Config-AA-AuthorizedUsers
  - Config-CQ-AuthorizedUsers
  
Changing any of the grey shaded areas on these or any other tabs may result in warnings, failures or inaccurate backups or provisioning.

## Requirements

Open the PowerShell window as an administrator.

>[!CAUTION]
>Testing has only been done with a Teams Global Administrator account.  Less privileged accounts should work provided they have sufficient access.  

### Microsoft Graph Scopes Requested

The PowerShell scripts request the following Microsoft Graph scopes"
  - Organization.Read.All
  - User.ReadWrite.All

## Known Issues

### BulkAuthUsersPreparation.ps1

- No known issues.

### BulkAuthUsersProvisioning.ps1

- No known issues.

### BulkAuthUsers spreadsheet

- Does not detect assigning multiple voice applications policies to the same user.

