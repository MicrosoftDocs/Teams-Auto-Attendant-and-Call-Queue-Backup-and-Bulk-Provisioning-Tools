# Download the BulkAuthUsersTool

Download the files and sub-directories in the **BulkAuthUsersTool** folder to a location on your workstation.

Alternatively, from the main **Code** page, select *<> Code* and *Download ZIP* to download the entire repository as a zip file and then unzip the file.

# Preparation Instructions

This step will download the existing voice applications policies, users, auto attendant and, call queue configurations in the tenant so they can be referenced when provisioning authorized users.

1. Login to Teams Admin Center and get the number of Auto Attendants and Call Queues configured in your tenant:

   ![Screenshot showing the Teams Admin Center summary table headers for Auto Attendants and Call Queues.](/media/TAC-Number-AA-CQ.png)

1. Open a PowerShell 5.x window
   - Issue the command: $PsVersionTable.PSVersion if not sure
1. In the PowerShell window, run the "BulkAuthUsersPreparation.ps1" script.	
   - This will prepare and open the BulkAuthUser spreadsheet.
   - If your tenant has more than 100 Auto Attendants or Call Queues use the -AACount or -CQCount options as outlined below.

## BulkAAsPreparation.ps1 command line options

| Option              | Description                                        |
|:--------------------|----------------------------------------------------|
| -AACount n          | Replace n with the number of Auto Attendants from Step 1. <br>*Only use when the number of Auto Attendants is greater than 100.*           |         
| -CQCount n          | Replace n with the number of Call Queues from Step 1. <br>*Only use when the number of Call Queues is greater than 100*                    |


# Provisioning Instructions

1. Open the BulkAuthUsers.xlsx spreadsheet.
1. Complete the follows tabs:
   
   - Config-PolicyToUser
   - Config-AA-AuthorizedUsers
   - Config-CQ-AuthorizedUsers
     
1. Save the BulkAuthUserss.xlsm spreadsheet and close Excel.
1. Open a PowerShell 5.x window
   - Issue the command: $PsVersionTable.PSVersion if not sure
1. In the PowerShell window, run the "BulkAuthUsersProvisioning.ps1" script

## BulkAAsProvisioning.ps1 command line options

None.                                                                                                                          |
