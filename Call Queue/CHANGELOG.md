# Changelog

## BulkCQsPreparation.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.22 | 1.0.4  | Yes      | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph<sup>1</sup><ul><li>Min: 2.24.0</li></ul> | - Fixed bug that resulted in erasing downloaded call queue information |
| 2025.04.21 | 1.0.3  | No       | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph<sup>1</sup><ul><li>Min: 2.24.0</li></ul> | - Support for Microsoft Shifts *(VoiceApps preview only)*<br>- Fixed issue with -AACount, -CQCount when < 100<br>- Conflicting parameters now stop processing<br>- Updated method for checking version of PowerShell modules<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul> | - Minor bug fixes        |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing<br>- Added counters on verbose output<br>- Eliminated use of temporary spreadsheets<br>- Suppressed CQ warning messages |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul> | Initial release               |

Notes:
1. Use -NoTeamsScheduleGroups to avoid loading Microsoft.Graph module

## BulkCQsProvisioning.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.21 | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul> |  - Support for Microsoft Shifts *(VoiceApps preview only)*<br>- Updated method for checking version of PowerShell modules<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting<br>- Microsoft.Graph module no longer loaded if not needed |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Minor bug fixes |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | Initial release  |


## BulkCQs.xlsm

| Date       | Version | Supported | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|
| 2025.04.21 | 1.0.1  | Yes       | - Microsoft Shifts support *(VoiceApps preview only)*<br>- Updated data validation and conditional formatting method<br>- Started support for variable ranges   |
| 2025.01.23 | 1.0.0  | No        | Initial release              |
