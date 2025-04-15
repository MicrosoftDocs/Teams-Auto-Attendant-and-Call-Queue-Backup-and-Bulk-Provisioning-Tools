# Changelog

## BulkAAsPreparation.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.15 | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0  | - Fixed issue with -AACount, -CQCount when < 100<br>- Conflicting parameters not stop processing<br>- Additional method for check version of PowerShell modules as Version.Major/Version.Minor not always returned |
| 2025.04.10 | 1.0.2  | Yes       |  MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0  |  - Support MicrosoftTeams  6.9.0<br>- Invalid parameter now stops processing<br>- Added counters on verbose output<br>- Eliminated use of temporary spreadsheets<br>- Suppressed CQ warning messages         |
|            | 1.0.1  | N/A       | N/A                          | Internal testing release                                  |
| 2025.01.23 | 1.0.0  | No        |  MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel: 7.8.0  | Initial release         |


## BulkAAsProvisioning.ps1

| Date       | Version | Supported | Supported PowerShell Modules | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|:----------------------------------------------------------|
| 2025.04.15 | 1.0.3  | Yes       | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul> | - Microsoft.Graph module no longer loaded if not needed<br>- Additional method for check version of PowerShell modules as Version.Major/Version.Minor not always returned |
| 2025.04.10 | 1.0.2  | Yes       | MicrosoftTeams:<ul><li>Min: 6.7.0</li><li>Max: 6.9.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  |  - Support MicrosoftTeams  6.9.0<br>- Invalid parameter now stops processing<br>- Improved holiday processing       |
|            | 1.0.1  | N/A       | N/A                          | Internal testing release                                  |
| 2025.01.23 | 1.0.0  | Yes       | MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel: 7.8.0<br>Microsoft.Graph:<ul><li>Min: 2.24.0</li></ul>  | Initial release    |


## BulkAAs.xlsm

| Date       | Version | Supported | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|
| 2025.04.10 | 1.0.2  | Yes       | - Support for Voice app & Resource account routing options to match TAC<br>- Multiple bug fixes |
|            | 1.0.0  | N/A       | Internal testing release     |
| 2025.01.23 | 1.0.0  | No        | Initial release              |
