# Changelog

## BulkCQsPreparation.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.28 | 1.0.5  | Yes      | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph<sup>1</sup><ul><li>Min: 2.24.0</li></ul> | - Compliance Recording for Call Queues<sup>2</sup><br>- Changed Teams-Channels/Teams-SchedulingGroups retrieval and Excel logic<br>- Corrected all array references<br>-Streamlined data retrieval for AAs, CQs and CQ download<br>- Audio file being downloaded now shown with -Verbose<br>- Check if Excel file already open<br>- Turn auto save and auto calculation off and restore at end |
| 2025.04.22 | 1.0.4  | Yes      | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph<sup>1</sup><ul><li>Min: 2.24.0</li></ul> | - Fixed bug that resulted in erasing downloaded call queue information |
| 2025.04.21 | 1.0.3  | No       | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph<sup>1</sup><ul><li>Min: 2.24.0</li></ul> | - Microsoft Shifts<br>- Fixed issue with -AACount, -CQCount when < 100<br>- Conflicting parameters now stop processing<br>- Updated method for checking version of PowerShell modules<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul> | - Minor bug fixes        |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing<br>- Added counters on verbose output<br>- Eliminated use of temporary spreadsheets<br>- Suppressed CQ warning messages |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul> | Initial release               |

Notes:
1. Use -NoTeamsScheduleGroups to avoid loading Microsoft.Graph module.
2. VoiceApps preview customers only

## BulkCQsProvisioning.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.28 | 1.0.4  | Yes       | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul> | - Compliance Recording for Call Queues<sup>2</sup> |
| 2025.04.21 | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 7.0.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul> | - Microsoft Shifts<br>- Updated method for checking version of PowerShell modules<br>- Disconnect from Microsoft.Graph before reconnect<br>- Updated Help as Callback is no GA<br>- Updated output formatting<br>- Microsoft.Graph module no longer loaded if not needed |
| 2025.04.10 | 1.0.2  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Minor bug fixes |
| 2025.04.01 | 1.0.1  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | - Support MicrosoftTeams  6.9.0<br>- Invalid paramater now stops processing |
| 2025.01.23 | 1.0.0  | No        | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel:<ul><li>Min: 7.8.0</li></ul>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | Initial release  |

Notes
1. Use -NoResourceAccounts or -NoResourceAccountCreation or -NoResourceAccountLicensing to avoid loading Microsoft.Graph modules
2. VoiceApps preview customers only


## BulkCQs.xlsm

| Date       | Version | Supported | Description                                               |
|:-----------|:-------|:---------:|:---------------------------------------------------|
| 2025.08.18 | 1.0.4  | Yes       | - Microsoft Shifts General Availability            |
| 2025.05.22 | 1.0.3  | Yes       | - Call Priorities support                          |
| 2025.04.28 | 1.0.2  | No        | - Compliance Recording For Call Queues<sup>1</sup> |
| 2025.04.21 | 1.0.1  | No        | - Microsoft Shifts<sup>1</sup><br>- Updated data validation and conditional formatting method<br>- Started support for variable ranges   |
| 2025.01.23 | 1.0.0  | No        | Initial release                                    |

Notes:
1. VoiceApps preview customers only
