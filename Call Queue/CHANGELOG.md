# Changelog

## BulkCQsPreparation.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.xx | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 6.9.1</li><li>Max: 6.9.1</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph<ul><li>Min: 2.24.0</li></ul> | - Support for Microsoft Shifts<br>- Fixed issue with -AACount, -CQCount when < 100<br>- Conflicting parameters now stop processing<br>- Additional method for checking version of PowerShell modules as Version.Major/Version.Minor not always returned<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0 | - Minor bug fixes        |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing<br>- Added counters on verbose output<br>- Eliminated use of temporary spreadsheets<br>- Suppressed CQ warning messages |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.8.0</li></ul>ImportExcel: 7.8.0  | Initial release               |

Notes:
1. Use -NoTeamsScheduleGroups to avoid loading Microsoft.Graph module

## BulkCQsProvisioning.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.xx | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 6.9.1</li><li>Max: 6.9.1</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul> |  - Support for Microsoft Shifts<br>- Additional method for checking version of PowerShell modules as Version.Major/Version.Minor not always returned<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting<br>- Microsoft.Graph module no longer loaded if not needed |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Minor bug fixes |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | Initial release  |


## BulkCQs.xlsm

| Date       | Version | Supported | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|
| 2025.04.xx | 1.0.1  | Yes       | Microsoft Shifts support     |
| 2025.01.23 | 1.0.0  | No        | Initial release              |
