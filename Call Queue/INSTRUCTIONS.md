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

| Option              | Description                                                                                                                                |
|:--------------------|--------------------------------------------------------------------------------------------------------------------------------------------|
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
| -NoTeamsScheduleGroups    | Do not download existing teams schedule groups                                                                                       |
| -NoOpen             | Do not open the spreadsheet when the BulkCQsPreparation.ps1 script is finished.                                                            |
| -Verbose            | Watch the spreadsheet get filled with information as the BulkAAsPreparation.psl1 script runs.<br>*Automaticaly disables*  **-NoOpen**      | 

## -Download notes

- All prompt downloads for a call queue will be in the AudioFiles directory, in a sub-directories by the call queue ID. This is due to the fact that call queue names are not unique.
- All audio file names will be prefixed with the unique file id and underscore. This is due to the fact that the same file name used within the same call queue may not actually have the same content.

# Preparing The Spreadsheet

Open the BulkCQs.xlsm, and enable macros if they have been disabled.

1. Complete the follows tabs: **Config-CallQueue**

| Field                           | Description                                                        |
|:--------------------------------|--------------------------------------------------------------------|
| Action                          | Select *New* to create a new call queue                              |
| NewCallQueueName                | This is the name that will be assigned to the call queue           |
| ResourceAccount                 | <ul><li>*Existing:* Assign an existing Resource Account to the call queue</li><li>*New:* Create a new resource account and assign to the call queue</li><li>*Blank:* Do not perform any resource account actitivies for this call queue</li></ul> |
| ExistingResourceAccountName     | Select the existing Resource Account to assign to the call queue<br>Note: Only available if `ResourceAccount`  is set to *Existing*                  |
| NewResourceAccountPrincipalName | The Resource Account UPN<br>Note: Only available if `ResourceAccount`  is set to *New*                                                               |
| NewResourceAccountDisplayName   | The Resource Account display name<br>Note: Only available if `ResourceAccount`  is set to *New*                                                       |
| NewResourceAccountLocation      | The Resource Account location. This will restrict which phone numbers can be assigned.<br>Note: Only available if `ResourceAccount`  is set to *New*  |
| NewResourceAccountPhoneNumber   | The Resource Account phone number.<br>Note: Only available if `ResourceAccount`  is set to *New*                                                      |
| NewResourceAccountPriority      | *Only availble for VoiceApps TAP customers at this time*                                                                                                    |
| OutboundCLID01-OutboundCLID04   | The outbound [calling IDs](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=general-info#assign-a-calling-id-optional) that can be used by agents when making outbound calls<sup>1</sup>  |
| Language                        | The [language](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=general-info#set-the-call-queue-language) for all Text To Speech (TTS) and system prompts |
| ServiceLevelThreshold           | The [service level threshold](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=general-info#set-the-service-level-threshold) time used for calculating the service level for real-time displays  |
| GreetingOptions                 | <ul><li>*No Greeting (Default):* No greeting message</li><li>*Play an audio file:* Use an audio file for the greeting</li><li>*Add a greeting message:* Use TTS for the greeting message</li></ul> |
| Greeting                        | Name of audio file or the text message for TTS                                                                                                              |
| MusicOnHold                     | <ul><li>*Play default music (Default):* Play the system default music for callers waiting in queue</li><li>*Play and audio file:* Use an audio file for the music</li></ul>            |
| MusicOnHoldAudioFilename        | Name of the audio file to use for Music on Hold<br>Note: Only available if `MusicOnHold` is set to *Play an audio file*                                   |
| RoutingMethod                   | The [agent routing](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#step-4-select-your-agent-routing-options) method that will be used to present calls to agents |
| PresenceBasedRouting            | Is [presence-based call routing](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#presence-based-call-routing) on or off |
| AllowOptOut                     | Agents [can opt in/out of taking calls](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#call-agents-can-opt-out-of-taking-calls) |
| AgentAlertTime                  | [Agent alert time](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=agent-selection#agent-alert-time)                       |
| OverflowThreshold               | The maximum number of simultaneous calls that can be in queue at one time before [Overflow](https://learn.microsoft.com/microsoftteams/create-a-phone-system-call-queue?tabs=call-exception-handling#overflow-set-how-to-handle-call-overflow) occurs |
| OverflowAction                  | Action to take once the `OverflowThreshold` is reached:<br><ul><li>*Disconnect (Default):* Disconnected</li><li>*Redirect - Person in organization:* Redirected to a Teams user in the tenant</li><li>*Redirect - Voice app:* Redirected to another Auto Attendant or Call Queue through a resource account or directly<sup>2</sup></li><li>*Redirect - External phone number:* Redirected to the PSTN<sup>3</sup></li><li>*Redirect - Voicemail personal:* Redirected to a Teams user's voicemail</li><li>*Redirect - Voicemail (shared):* Redirected to a shared voicemail</li></ul> |
| OverflowActionTarget            | The target of the `OverflowAction`<br>Note:<br>1. Only available if `OverflowAction` is not *Disconnect (Default)*<br>2. When `OverflowAction` is *Redirect - External phone number* the values here will be:<br><ul><li>*Calling Plan:* Microsoft numbers are being used</li><li>*Direct Routing:* Direct Routing is being used</li><li>*Operator Connect:* Operator Connect numbers are being used</li></ul>                                                       |
| OverflowActionTargetCountry     | The country calls will be routed to<br>Note: Only available if `OverflowAction` is *Redirect - External phone number* and `OverflowActionTarget` is *Calling Plan*   |
| OverflowActionTargetNumber      | The phone number calls will be routed to<br>Note: Only available if `OverflowAction` is *Redirect - External phone number*   |
| OverflowActionCallPriority      | *Only availble for VoiceApps TAP customers at this time* |
| OverflowTreatment               | The message that is played when overflow occurs<br><ul><li>*No Greeting (Default):* No message will be played</li><li>*Play an audio file:* Use an audio file for the message</li><li>*Add a greeting message:* Use TTS for the message</li></ul> |
| OverflowTreatmentPrompt         | Name of audio file or the text message for TTS                                                                                                              |
| OverflowSharedVoicemailSystemPromptSuppression | Supress the system greeting message for shared voicemail<br>Note: Only available when `OverflowAction` is *Redirect - Voicemail (shared)*    |
| OverflowSharedVoicemailTranscription | Enable voicemail transcription for shared voicemail<br>Note: Only available when `OverflowAction` is *Redirect - Voicemail (shared)*    |
| TimeoutThreshold                | The maximum amount fo time a caller can remain in queue before [Timeout](https://learn.microsoft.commicrosoftteams/create-a-phone-system-call-queue?tabs=call-exception-handling#call-timeout-set-how-to-handle-call-timeouts) occurs |
| TimeoutAction                   | Action to take once the `TimeoutThreshold` is reached:<br><ul><li>*Disconnect (Default):* Disconnected</li><li>*Redirect - Person in organization:* Redirected to a Teams user in the tenant</li><li>*Redirect - Voice app:* Redirected to another Auto Attendant or Call Queue through a resource account or directly<sup>2</sup></li><li>*Redirect - External phone number:* Redirected to the PSTN<sup>3</sup></li><li>*Redirect - Voicemail personal:* Redirected to a Teams user's voicemail</li><li>*Redirect - Voicemail (shared):* Redirected to a shared voicemail</li></ul> |
| TimeoutActionTarget             | The target of the `TimeoutAction`<br>Note:<br>1. Only available if `TimeoutAction` is not *Disconnect (Default)*<br>2. When `TimeoutAction` is *Redirect - External phone number* the values here will be:<br><ul><li>*Calling Plan:* Microsoft numbers are being used</li><li>*Direct Routing:* Direct Routing is being used</li><li>*Operator Connect:* Operator Connect numbers are being used</li></ul>                                                       |
| TimeoutActionTargetCountry      | The country calls will be routed to<br>Note: Only available if `TimeoutAction` is *Redirect - External phone number* and `TimeoutActionTarget` is *Calling Plan*   |
| TimeoutActionTargetNumber       | The phone number calls will be routed to<br>Note: Only available if `TimeoutAction` is *Redirect - External phone number*   |
| TimeoutActionCallPriority       | *Only availble for VoiceApps TAP customers at this time* |
| TimeoutTreatment                | The message that is played when timeout occurs<br><ul><li>*No Greeting (Default):* No message will be played</li><li>*Play an audio file:* Use an audio file for the message</li><li>*Add a greeting message:* Use TTS for the message</li></ul> |
| TimeoutTreatmentPrompt          | Name of audio file or the text message for TTS                                                                                               |
| TimeoutSharedVoicemailSystemPromptSuppression | Supress the system greeting message for shared voicemail<br>Note: Only available when `TimeoutAction` is *Redirect - Voicemail (shared)*    |
| TimeoutSharedVoicemailTranscription | Enable voicemail transcription for shared voicemail<br>Note: Only available when `TimeoutAction` is *Redirect - Voicemail (shared)*    |
| NoAgentsApplyTo                 | The No agents configuration applies to:<br><ul><li>*All calls (Default):* Calls already in queue and new calls arriving to the queue</li><li>*New calls:* Only new calls that arrive once the No Agents condition occurs, existing calls in queue remain in queue |
| NoAgentAction                    | Action to take once the No Agents condition occurs:<br><ul><li>*Queue call (Default):* The No Agents treatment is ignored and calls are queued<ul><li>*Disconnect:* Disconnected</li><li>*Redirect - Person in organization:* Redirected to a Teams user in the tenant</li><li>*Redirect - Voice app:* Redirected to another Auto Attendant or Call Queue through a resource account or directly<sup>2</sup></li><li>*Redirect - External phone number:* Redirected to the PSTN<sup>3</sup></li><li>*Redirect - Voicemail personal:* Redirected to a Teams user's voicemail</li><li>*Redirect - Voicemail (shared):* Redirected to a shared voicemail</li></ul> |
| NoAgentActionTarget             | The target of the `NoAgentAction`<br>Note:<br>1. Only available if `NoAgentAction` is not *Disconnect (Default)*<br>2. When `NoAgentAction` is *Redirect - External phone number* the values here will be:<br><ul><li>*Calling Plan:* Microsoft numbers are being used</li><li>*Direct Routing:* Direct Routing is being used</li><li>*Operator Connect:* Operator Connect numbers are being used</li></ul>                                                       |
| NoAgentActionTargetCountry      | The country calls will be routed to<br>Note: Only available if `NoAgentAction` is *Redirect - External phone number* and `NoAgentActionTarget` is *Calling Plan*   |
| NoAgentActionTargetNumber       | The phone number calls will be routed to<br>Note: Only available if `NoAgentAction` is *Redirect - External phone number*   |
| NoAgentActionCallPriority       | *Only availble for VoiceApps TAP customers at this time* |
| NoAgentTreatment                | The message that is played when no agents occurs<br><ul><li>*No Greeting (Default):* No message will be played</li><li>*Play an audio file:* Use an audio file for the message</li><li>*Add a greeting message:* Use TTS for the message</li></ul> |
| NoAgentTreatmentPrompt          | Name of audio file or the text message for TTS                                                                                               |
| NoAgentSharedVoicemailSystemPromptSuppression | Supress the system greeting message for shared voicemail<br>Note: Only available when `NoAgentAction` is *Redirect - Voicemail (shared)*    |
| NoAgentSharedVoicemailTranscription | Enable voicemail transcription for shared voicemail<br>Note: Only available when `NoAgentAction` is *Redirect - Voicemail (shared)*    |
| IsCallbackEnabled                   | <ul><li>*Yes:* Callback is enabled</li><li>*No:* Default: Callback is not enabled |
| CallbackRequestDTMF                 | The key a caller need to press to request a callback. This should match what callers are told in the `CallbackOfferPrompt`<br>Note: Only available if `IsCallbackEnabled` is *Yes*   |
| WaitTimeBeforeOfferingCallbackInSecond | The number of seconds a caller must want before becoming *eligibe* for callback<br>Note: Only available if `IsCallbackEnabled` is *Yes*  |
| NumberOfCallsInQueueBeforeOfferingCallback | The number of calls that must be in queue before new callers become *eligibe* for callback<br>Note: Only available if `IsCallbackEnabled` is *Yes*  |
| CallToAgentRatioThresholdBeforeOfferingCallback | The ratio of calls to agents that must be exceeded before new callers become *eligibe* for callback<br>Note: Only available if `IsCallbackEnabled` is *Yes*  |
| CallbackOfferTreatment     | The type of message that is played tp offer callback<br><ul><li>*Play an audio file:* Use an audio file for the message</li><li>*Add a greeting message:* Use TTS for the message</li></ul>Note: Only available if `IsCallbackEnabled` is *Yes*  |
| CallbackOfferPrompt        | Name of audio file or the text message for TTS<br>Note: Only available if `IsCallbackEnabled` is *Yes*                                                   |
| CallbackEmailNotificationTarget | The distribution list to send emails to about callbacks that timeout<br>Note: Only available if `IsCallbackEnabled` is *Yes* 
| Team-Channel               | The Teams Channel to assign to the queue |
| Team-ScheduleGroup         | The Teams Schedule Group to the queue |
| DistributionList01-DistributionList04 | The distribution lists of agents to assign to the queue |
| Agent01-Agent20            | The agents to assign to the queue |

Notes: 
1. If you're using a resource account for calling line ID purposes in Call queues, the resource account must have a Teams Phone Resource Account license and one of the following assigned:
    - A [Calling Plan](https://learn.microsoft.com/microsoftteams/calling-plans-for-office-365) license and a phone number assigned.
    - An [Operator Connect](https://learn.microsoft.com/microsoftteams/operator-connect-plan) phone number assigned.
    - An [online voice routing policy](https://learn.microsoft.com/microsoftteams/manage-voice-routing-policies).
      - Phone number assignment is optional when using Direct Routing.
1. In Teams Admin Center, redirects to Voice apps and Resource accounts are shown as separate options. In this spreadsheet, all redirects are to Voice apps. The [prefix] in front of the Auto Attendant or Call Queue name indicate how the transfer will be done:
    - [RA-AA]: Transfer will be via the Resource Account assigned to the Auto Attendant
    - [RA-CQ]: Transfer will be via the Resource Account assigned to the Call Queue
    - [AA]: Transfer will be directly to the Auto Attendant
    - [CQ]: Transfer will be directly to the Call Queue
    - No prefix: Transfer will be directly to the Call Queue that is being created 
1. In addition to the Teams Phone Resource Account license, when a nested auto attendant or call queue transfers calls to an external number, the last resource account that was part of the call flow must also have one of the following assigned:
      - A [Calling Plan](https://learn.microsoft.com/microsoftteams/calling-plans-for-office-365) license and a phone number.
      - An [Operator Connect](https://learn.microsoft.com/microsoftteams/operator-connect-plan) phone number.
      - An [online voice routing policy](https://learn.microsoft.com/microsoftteams/manage-voice-routing-policies).
        - Phone number assignment is optional when using Direct Routing.

>[!IMPORTANT]
> Entries that become highlighted in red after entering data indicate errors<sup>*</sup>. These need to be addressed before running the provisioning script.
>
> \* - with the exception of [known issues](./README.md#known-issues)

# Provisioning Instructions

1. Make sure the BulkCQs.xlsm spreadsheet is closed.
1. Make sure any referenced prompt files are in the AudioFiles sub-directory.
1. Open a PowerShell 5.x window
   
   - Issue the command: $PSVersionTable.PSVersion if not sure
     
1. In the PowerShell window, run the "BulkCQsProvisioning.ps1" script

## BulkCQsProvisioning.ps1 command line options

| Option                     | Description                                                                                                                                                                         |
|:---------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| -ExcelFile filename        | Specify an alternative Excel spreadsheet to use. Must be in the same directory as the BulkAAsProvisioning.ps1 file<br>Default: BulkAAs.xlsm                                         |
| -Help                      | This help message.                                                                                                                                                                  |
| -NoCallBack                | Do not perform any callback related stemps                                                                                                                                          |
| -NoResourceAccounts        | Do not perform any resource account related steps. <br>*Automaticaly enables*  **-NoResourceAccountCreation**, **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**  |
| -NoResourceAccountCreation | Do not provision any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountLicensing**, **-NoResourceAccountPhoneNumbers**                                          |
| -NoResourceAccountLicensing| Do not license any new resource accounts.<br>*Automaticaly enables*  **-NoResourceAccountPhoneNumbers**                                                                             |
| -NoResourceAccountPhoneNumbers | Do not assign a phone number to a new resource account
| -Verbose                   | Detailed output.                                                                                                                                                                    |
