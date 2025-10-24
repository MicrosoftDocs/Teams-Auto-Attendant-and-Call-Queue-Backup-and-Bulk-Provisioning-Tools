# Changelog

## BulkAAsPreparation.ps1

### Supported PowerShell Modules
- MicrosoftTeams: Min: 6.7.0
- ImportExcel:Min: 7.8.0

| Date       | Version | Supported | Description                                               |
|:-----------|:--------|:---------:|:----------------------------------------------------------|
| 2025.09.22 | 1.0.4   | Yes       | - Version checking for BulkAAs.xlsm<br>- Added -Download option<br>- Changed how -AACount, -CQCount work<br>- Scrolling in -View mode |
| 2025.04.15 | 1.0.3   | Yes       | - Fixed issue with -AACount, -CQCount when < 100<br>- Conflicting parameters now stop processing<br>- Additional method for checking version of PowerShell modules as Version.Major/Version.Minor not always returned |
| 2025.04.10 | 1.0.2   | No        |  - Support MicrosoftTeams  6.9.0<br>- Invalid parameter now stops processing<br>- Added counters on verbose output<br>- Eliminated use of temporary spreadsheets<br>- Suppressed CQ warning messages         |
|            | 1.0.1   | N/A       | N/A                          | Internal testing release                                  |
| 2025.01.23 | 1.0.0   | No        |  MicrosoftTeams:<ul><li>Min: 6.7.0</li></ul>ImportExcel: 7.8.0  | Initial release         |


## BulkAAsProvisioning.ps1

### Supported PowerShell Modules
- MicrosoftTeams: Min: 6.7.0
- ImportExcel:Min: 7.8.0
- Microsoft.Graph<sup>1</sup>Min: 2.24.0

| Date       | Version | Supported | Description                                               |
|:-----------|:--------|:---------:|:----------------------------------------------------------|
| 2025.04.15 | 1.0.3   | Yes       | - Microsoft.Graph module no longer loaded if not needed<br>- Additional method for checking version of PowerShell modules as Version.Major/Version.Minor not always returned |
| 2025.04.10 | 1.0.2   | Yes       |  - Support MicrosoftTeams  6.9.0<br>- Invalid parameter now stops processing<br>- Improved holiday processing       |
|            | 1.0.1   | N/A       | Internal testing release                                  |
| 2025.01.23 | 1.0.0   | Yes       | Initial release                                           |


## BulkAAs.xlsm

| Date       | Version | Supported | Description                                               |
|:-----------|:-------|:---------:|:-----------------------------|
| 2025.09.22 | 1.0.4  | Yes       | - Support for-Download option<br>- Reordered tabs<br>- Froze header rows |
| 2025.05.22 | 1.0.3  | Yes       | - Call priorities support    |
| 2025.04.10 | 1.0.2  | No        | - Support for Voice app & Resource account routing options to match TAC<br>- Multiple bug fixes |
|            | 1.0.1  | N/A       | Internal testing release     |
| 2025.01.23 | 1.0.0  | No        | Initial release              |
