# Download the BulkCQTool

Download the files and sub-directories in the **BulkCQTool** folder to a location on your workstation.

Alternatively, from the main **Code** page, select *<> Code* and *Download ZIP* to download the entire repository as a zip file and then unzip the file.

# Preparation Instructions

This step will download the existing resource account, auto attendant, call queue, Teams channels, and user configurations in the tenant so they can be referenced when provisioning new call queues.

1. Login to Teams Admin Center and get the number of Auto Attendants and Call Queues configured in your tenant:

   ![Screenshot showing the Teams Admin Center summary table headers for Auto Attendants and Call Queues.](/media/TAC-Number-AA-CQ.png)

1. Open a PowerShell 5.x window
   - Issue the command: $PsVersionTable.PSVersion if not sure
   - "Run as administrator" is suggested as the script will try to install the [required PowerShell modules](./README.md#required-powershell-modules) if they aren't present.
  1. In the PowerShell window, run the "BulkCQsPreparation.ps1" script.	
   - This will prepare and open the BulkCQs spreadsheet.
   - If your tenant has more than 100 Auto Attendants or Call Queues use the -AACount or -CQCount options as outlined below.

## BulkCQsPreparation.ps1 command line options

| Option              | Description                                        |
|:--------------------|----------------------------------------------------|
| -AACount n          | Replace n with the number of Auto Attendants from Step 1. <br>*Only use when the number of Auto Attendants is greater than 100.*           |         
| -CQCount n          | Replace n with the number of Call Queues from Step 1. <br>*Only use when the number of Call Queues is greater than 100*                    |
| -Download           | Download all Call Queue data, including audio files.                                                                                       |
| -ExcelFile filename | Specify an alternative Excel spreadsheet to use. Must be in the same directory as the BulkAAsPreparation.ps1 file<br>Default: BulkCQs.xlsm |
| -Help               | This help message.                                                                                                                         |
| -NoResourceAccounts | Do not download existing resource account information.                                                                                     |
| -NoAutoAttendants   | Do not download existing auto attendant information.                                                                                       |
| -NoCallQueues       | Do not download existing call queue information.                                                                                           |
| -NoUsers            | Do not download existing EV enabled users.                                                                                                 |
| -NoTeamsChannels    | Do not download existing teams information.                                                                                                |
| -NoOpen             | Do not open the spreadsheet when the BulkCQsPreparation.ps1 script is finished.                                                            |
| -Verbose            | Watch the spreadsheet get filled with information as the BulkAAsPreparation.psl1 script runs.<br>*Automaticaly disables*  **-NoOpen**      | 

## -Download notes

- All prompt downloads for a call queue will be in the AudioFiles directory, in a sub-directories by the call queue ID. This is due to the fact that call queue names are not unique.
- All audio file names will be prefixed with the unique file id and underscore. This is due to the fact that the same file name used within the same call queue may not actually have the same content.

# Filling In The Spreadsheet

Open the BulkCQs.xlsm, and enable macros if they have been disabled.

1. Complete the follows tabs:
   
   - Config-CallQueue

| Field                           | Description                                                        |
|:--------------------------------|--------------------------------------------------------------------|
| Action                          | Select New to create a new call queue                              |
| NewCallQueueName                | This is the name that will be assigned to the call queue           |
| ResourceAccount                 | Existing: Assign an existing Resource Account to the call queue<br>New: Create a new resource account and assign to the call queue<br>Blank: Do not perform any resource account actitivies for this call queue |
| ExistingResourceAccountName     | Select the existing Resource Account to assign to the call queue<br>Note: Only available if `ResourceAccount` field is set to **Existing**                  |
| NewResourceAccountPrincipalName | The Resource Account UPN<br>Note: Only available if `ResourceAccount` field is set to *New*                                                               |
| NewResourceAccountDisplayName   | The Resource Account display name<br>Note: Onlyavailable if `ResourceAccount` field is set to *New*                                                       |
| NewResourceAccountLocation      | The Resource Account location. This will restrict which phone numbers can be assigned.<br>Note: Onlyavailable if `ResourceAccount` field is set to *New*  |
| NewResourceAccountPhoneNumber   | The Resource Account phone number.<br>Note: Onlyavailable if `ResourceAccount` field is set to *New*                                                      |
| NewResourceAccountPriority      | *Only availble for VoiceApps TAP customers at this time*                                                                                                    |
| OutboundCLID01-OutboundCLID04   | The outbound calling line IDs that can be used by agents when making outbound calls                                                                         |
| Language                        | The language for all Text To Speech (TTS) and system prompts                                                                                                |
| ServiceLevelThreshold           | The time threshold to be used for calculating the service level for real-time displays                                                                      |
| GreetingOptions                 | No Greeting (Default): No greeting message<br>Play an audio file: Use an audio file for the greeting<br>Add a greeting message: Use TTS for the greeting message |
| Greeting                        | Name of audio file or the text message for TTS                                                                                                              |
| MusicOnHold                     | Play default music (Default): Play the system default music for callers waiting in queue<br>Play and audio file: Use an audio file for the music            |
| MusicOnHoldAudioFilename        | Name of the audio file to use for Music on Hold<br>Note: Only available if `MusicOnHold` is set to *Play an audio file*                                   |
| RoutingMethod                   | The [agent routing](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#step-4-select-your-agent-routing-options) method that will be used to present calls to agents |
| PresenceBasedRouting            | Is [presence-based call routing](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#presence-based-call-routing) on or off |
| AllowOptOut                     | Agents [can opt in/out of taking calls](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#call-agents-can-opt-out-of-taking-calls) |
| AgentAlertTime                  | [Agent alert time](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#agent-alert-time)                       |
| OverflowThreshold               | The maximum number of simultaneous calls that can be in queue at one time before [Overflow](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=call-exception-handling#overflow-set-how-to-handle-call-overflow) occurs |
| OverflowAction                  | Calls above the `OverflowThreshold` will be:<br>Disconnect (Default): Disconnected<br>Redirect - Person in organization: Redirected to a Teams user in the tenant<br>Redirect - Voice app: Redirected to another Auto Attendant or Call Queue through a resource account or direction<sup>1</sup><br>Redirect - External phone number: Redirected to the PSTN<sup>2</sup><br>Redirect - Voicemail personal: Redirected to a Teams user's voicemail<br>Redirect - Voicemail (shared): Redirected to a shared voicemail |
| OverflowActionTarget            | The target of the `OverflowAction`<br>Note: Only available if `OverflowAction` is not *Disconnect (Default)*   NEED MORE DETAIL HERE                                             |
| OverflowActionTargetCountry     | The country to overflow calls to<br>Note: Only available if `OverflowAction` is *Redirect - External phone number* and `OverflowActionTarget` is *Calling Plan*   |
| OverflowActionTargetNumber      | The phone number of overflow calls to<br>Note: Only available if `OverflowAction` is *Redirect - External phone number*   |
| OverflowActionCallPriority      | *Only availble for VoiceApps TAP customers at this time* |
| OverflowTreatment               | The message that is played when overflow occurs<br>No Greeting (Default): No message will be played<br>Play an audio file: Use an audio file for the message<br>Add a greeting message: Use TTS for the message |
| OverflowTreatmentPrompt         | Name of audio file or the text message for TTS                                                                                                              |
| OverflowSharedVoicemailSystemPromptSuppression | Supress the system greeting message for shared voicemail<br>Note: Only available when `OverflowAction` is *Redirect - Voicemail (shared)*    |
| OverflowSharedVoicemailTranscription | Enable voicemail transcription for shared voicemail<br>Note: Only available when `OverflowAction` is *Redirect - Voicemail (shared)*    |

- TimeoutThreshold
- TimeoutAction
- TimeoutActionTarget
- TimeoutActionTargetCountry
- TimeoutActionTargetNumber
- TimeoutActionCallPriority
- TimeoutTreatment
- TimeoutTreatmentPrompt
- TimeoutSharedVoicemailSystemPromptSuppression
- TimeoutSharedVoicemailTranscription
- NoAgentsApplyTo
- NoAgentAction
- NoAgentActionTarget
- NoAgentActionTargetCountry
- NoAgentActionTargetNumber
- NoAgentActionCallPriority
- NoAgentTreatment
- NoAgentTreatmentPrompt
- NoAgentSharedVoicemailSystemPromptSuppression
- NoAgentSharedVoicemailTranscription
- IsCallbackEnabled
- CallbackRequestDTMF
- WaitTimeBeforeOfferingCallbackInSecond
- NumberOfCallsInQueueBeforeOfferingCallback
- CallToAgentRatioThresholdBeforeOfferingCallback
- CallbackOfferTreatment
- CallbackOfferPrompt
- CallbackEmailNotificationTarget
- Team-Channel
- DistributionList01-DistributionList04
- Agent01-Agent20



# Provisioning Instructions

1. Make sure the BulkCQs.xlsm spreadsheet is closed.
1. Make sure any referenced prompt files are in the AudioFiles sub-directory.
1. Open a PowerShell 5.x window
   
   - Issue the command: $PSVersionTable.PSVersion if not sure
     
1. In the PowerShell window, run the "BulkCQsProvisioning.ps1" script

## BulkCQsProvisioning.ps1 command line options

| Option                     | Description                                        |
|:---------------------------|----------------------------------------------------|
| -ExcelFile filename        | Specify an alternative Excel spreadsheet to use. Must be in the same directory as the BulkAAsProvisioning.ps1 file<br>Default: BulkAAs.xlsm |
| -Help                      | This help message.                                                                                                                          |
| -NoResourceAccounts        | Do not perform any resource account related steps. <br>*Automaticaly enables*  **-NoResourceAccountCreation**, **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**  |
| -NoResourceAccountCreation | Do not provision any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**  |
| -NoResourceAccountLicensing| Do not license any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountPhoneNumbers**                                     |
| -Verbose                   | Detailed output.                                                                                                                            |
