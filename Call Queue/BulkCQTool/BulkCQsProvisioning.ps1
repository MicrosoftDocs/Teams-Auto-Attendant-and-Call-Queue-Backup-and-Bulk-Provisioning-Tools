# Version: 1.0.7
# Date: 2025.10.24
#

# Changelog: https://github.com/MicrosoftDocs/Teams-Auto-Attendant-and-Call-Queue-Backup-and-Bulk-Provisioning-Tools/blob/main/Call%20Queue/CHANGELOG.md

# PowerShell Streams
#
#Stream #	Description			Write Cmdlet		Variable				Default
#1			Success stream		Write-Output
#2			Error stream		Write-Error			$ErrorActionPreference	Continue
#3			Warning stream		Write-Warning		$WarningPreference		Continue
#4			Verbose stream		Write-Verbose		$VerbosePrefernce		SilentlyContinue
#5			Debug stream		Write-Debug			$DebugPreference		SilentlyContinue
#6			Information stream	Write-Information	$InformationPreference	SilentlyContinue
#
#Preference Variable Options
# Use name or value
#
#Name				Value
#Break				6
#Suspend			5
#Ignore				4
#Inquire			3
#Continue			2
#Stop				1
#SilentlyContinue	0
#



###########################
#  AudioFileImport
###########################
function AudioFileImport
{
	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[string] $fileName
	)

    $currentDIR = (Get-Location).Path
    $currentDIR += "\AudioFiles\$fileName"
    $content = [System.IO.File]::ReadAllBytes($currentDIR) 
    $audioFileID = (Import-CsOnlineAudioFile -ApplicationID HuntGroup -FileName $fileName -Content $content).ID

    return $audioFileID
}

#
#  VerboseOutput
#
function VerboseOutput
{
	Write-Host "`t-------------------------------------------------------------------"
	Write-Host "`tAction:  $Action`tName: $Name"
	Write-Host "`t-------------------------------------------------------------------"

	Write-Host "`tExistingResourceAccoutName                     : $ExistingResourceAccountName"
	Write-Host "`tNewResourceAccountPrincipalName                : $NewResourceAccountPrincipalName"
	Write-Host "`tNewResourceAccountDisplayName                  : $NewResourceAccountDisplayName"
	Write-Host "`tNewResourceAccountLocation                     : $NewResourceAccountLocation"
	Write-Host "`tNewResourceAccountPriority                     : $NewResourceAccountPriority"
	Write-Host "`tNewResourceAccountPhoneNumber                  : $NewResourceAccountPhoneNumber"
	
	Write-Host "`tOutboundCLID                                   : $OboResourceAccountIDs"
	Write-Host "`tLanguage                                       : $Language"
	Write-Host "`tServiceLevelThreshold                          : $ServiceLevelThreshold"
    Write-Host "`tSCHTemplateId                                  : $SCHTemplateId"
   
	Write-Host "`tComplianceRecordingTemplateIDs                 : $ComplianceRecordingTemplateIDs"
	Write-Host "`tCR4CQGreetingOption                            : $CR4CQGreetingOption"
	Write-Host "`tCR4CQGreeting                                  : $CR4CQGreeting"
	Write-Host "`tCR4CQFailureGreetingOption                     : $CR4CQFailureGreetingOption"
	Write-Host "`tCR4CQFailureGreeting                           : $CR4CQFailureGreeting"
   
	Write-Host "`tGreetingOption                                 : $GreetingOption"
	Write-Host "`tGreeting                                       : $Greeting"
	Write-Host "`tMusicOnHoldOption                              : $MusicOnHoldOption"
	Write-Host "`tMusicOnHold                                    : $MusicOnHold"
	Write-Host "`tRoutingMethod                                  : $RoutingMethod"
	Write-Host "`tPresenceBasedRouting                           : $PresenceBasedRouting"
	Write-Host "`tAllowOptOut                                    : $AllowOptOut"
	Write-Host "`tAgentAlertTime                                 : $AgentAlertTime"

	Write-Host "`tOverflowThreshold                              : $OverflowThreshold"
	Write-Host "`tOverflowAction                                 : $OverflowAction"
	Write-Host "`tOverflowActionTarget                           : $OverflowActionTarget"
	Write-Host "`tOverflowActionCallPriority                     : $OverflowActionCallPriority"
	Write-Host "`tOverflowTreatment                              : $OverflowTreatment"
	Write-Host "`tOverflowTreatmentPrompt                        : $OverflowTreatmentPrompt"
	Write-Host "`tOverflowSharedVoicemailSystemPromptSuppression : $OverflowSharedVoicemailSystemPromptSuppression"
	Write-Host "`tOverflowSharedVoicemailTranscription           : $OverflowSharedVoicemailTranscription"

	Write-Host "`tTimeoutThreshold                               : $TimeoutThreshold"
	Write-Host "`tTimeoutAction                                  : $TimeoutAction"
	Write-Host "`tTimeoutActionTarget                            : $TimeoutActionTarget"
	Write-Host "`tTimeoutActionCallPriority                      : $TimeoutActionCallPriority"
	Write-Host "`tTimeoutTreatmentTreatment                      : $TimeoutTreatment"
	Write-Host "`tTimeoutTreatmentPrompt                         : $TimeoutTreatmentPrompt"
	Write-Host "`tTimeoutSharedVoicemailSystemPromptSuppression  : $TimeoutSharedVoicemailSystemPromptSuppression"
	Write-Host "`tTimeoutSharedVoicemailTranscription            : $TimeoutSharedVoicemailTranscription"

	Write-Host "`tNoAgentNewCallsOnly                            : $NoAgentNewCallsOnly"
	Write-Host "`tNoAgentAction                                  : $NoAgentAction"
	Write-Host "`tNoAgentActionTarget                            : $NoAgentActionTarget"
	Write-Host "`tNoAgentActionCallPriority                      : $NoAgentActionCallPriority"
	Write-Host "`tNoAgentTreatment                               : $NoAgentTreatment"
	Write-Host "`tNoAgentTreatmentPrompt                         : $NoAgentTreatmentPrompt"
	Write-Host "`tNoAgentSharedVoicemailSystemPromptSuppression  : $NoAgentSharedVoicemailSystemPromptSuppression"
	Write-Host "`tNoAgentsharedVoicemailTranscription            : $NoAgentSharedVoicemailTranscription"

	Write-Host "`tIsCallbackEnabled                              : $IsCallbackEnabled"
	Write-Host "`tCallBackRequestDTMF                            : $CallbackRequestDTMF"
	Write-Host "`tWaitTimeBeforeOfferingCallbackInSecond         : $WaitTimeBeforeOfferingCallbackInSecond"
	Write-Host "`tNumberOfCallsInQueueBeforeOfferingCallback     : $NumberOfCallsInQueueBeforeOfferingCallback"
	Write-Host "`tCallToAgentRatioThresholdBeforeOfferingCallback: $CallToAgentRatioThresholdBeforeOfferingCallback"
	Write-Host "`tCallbackOfferTreatment                         : $CallbackOfferTreatment"
	Write-Host "`tCallbackOfferPrompt                            : $CallbackOfferPrompt"
	Write-Host "`tCallbackEmailNotificationTarget                : $CallbackEmailNotificationTarget"

	Write-Host "`tTeamGroupID                                    : $TeamGroupID"
	Write-Host "`tTeamChannelID                                  : $TeamChannelID"
	Write-Host "`tTeamChannelName                                : $TeamChannelName"

	Write-Host "`tDistribtuionList                               : $DistributionLists"

	Write-Host "`tAgents                                         : $Users"
}


#
# NewSharedCallQueueHistoryTemplate
#
function NewSCHTemplate
{
	if ( $Verbose )
	{
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`tAction: New Shared Call Queue History Template"
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`tNewSCHTemplateName                     : $NewSCHTemplateName"
		Write-Host "`tNewSCHTemplateDescription              : $NewSCHTemplateDescription"
		Write-Host "`tNewSCHTemplateIncomingMissedCalls      : $NewSCHTemplateIncomingMissedCalls"
		Write-Host "`tNewSCHTemplateAnsweredAndOutboundCalls : $NewSCHTemplateAnsweredAndOutboundCalls"
	}
	else
	{
		Write-Host "`tAction: New Shared Call Queue History Template : $NewSCHTemplateName"
	}
	
	$command = "New-CsSharedCallQueueHistoryTemplate -Name $NewSCHTemplateName -Description $NewSCHTemplateDescription "
	
	#
	# IncomingMissedCalls
	# Default: None (as in None, not nothing)
	#
	if ( $NewSCHTemplateIncomingMissedCalls -ne "" )
	{
		$command += "-IncomingMissedCalls $NewSCHTemplateIncomingMissedCalls "
	}
	
	#
	# AnsweredAndOutboundCalls
	# Default: None (as in None, not nothing)
	#
	if ( $NewSCHTemplateAnsweredAndOutboundCalls -ne "" )
	{
		$command += "-AnsweredAndOutboundCalls $NewSCHTemplateAnsweredAndOutboundCalls "
	}

	if ( $Verbose )
	{
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`t$command"
		Write-Host "`t-------------------------------------------------------------------"
	}

	$command = '$Object = ( ' + $command
	$command += " )"
	
	Write-Host "`tCreating Shared Call Queue History Template : $NewSCHTemplateName " -f Green
	Invoke-Expression $command 

	return $Object
}


#
# New Resource account
#
function NewResourceAccount
{
	if ( $Verbose )
	{
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`tAction: New Resource Account"
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`tNewResourceAccountPrincipalName : $NewResourceAccountPrincipalName"
		Write-Host "`tNewResourceAccountDisplayName   : $NewResourceAccountDisplayName"
	}
	else
	{
		Write-Host "`tAction: New Resource Account : $NewResourceAccountPrincipalName"
	}
	# MicrosoftTeams
	$command = "New-CsOnlineApplicationInstance -UserPrincipalName $NewResourceAccountPrincipalName -DisplayName $NewResourceAccountDisplayName -ApplicationID 11cd3e2e-fccb-42ad-ad00-878b93575e07"

	if ( $Verbose )
	{
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`t$command"
		Write-Host "`t-------------------------------------------------------------------"
	}

	$command = '$Object = ( ' + $command
	$command += " )"

	Write-Host "`tCreating Resource Account : $NewResourceAccountPrincipalName " -f Green
	Invoke-Expression $command 
	
	return $Object
}

	
#
#  NewCallQueue
#
function NewCallQueue 
{
	if ( $Verbose )
	{
		VerboseOutput
	}
	else
	{
		Write-Host "`tAction:  $Action`tName: $Name"
	}

	$command = "New-CsCallQueue -Name $Name "


	#
	#  Calling ID
	#  Default: None
	#
	if ( $OutboundCLID.count -gt 0 )
	{
		$command += "-OboResourceAccountIds @($OboResourceAccountIDs) "
	}

   
	#
	#  Language
	#  Default: en-US (set in spreadhseet)
	#
	$command += "-LanguageId $Language "
 

	#
	#  Service Level Threshold
	#  Default: null (set in spreadsheet)
	#
	if ( $ServiceLevelThreshold -eq "null" )
	{
		$command += "-ServiceLevelThresholdResponseTimeInSecond $" + "null "
	}
	else
	{
		$command += "-ServiceLevelThresholdResponseTimeInSecond $ServiceLevelThreshold "
	}


	#
	# Shared Call History
	#
	if ( $SCHTemplateId -ne "" )
	{
		if ( $SCHTemplateId.Substring(0,4) -eq "NEW:" )
		{
			$NewSCHTemplateName = $SCHTemplateId.Substring(4)
		
			$NewSCHTemplateId = ""
			for ( $i = 0; $i -lt $NewSCHTemplates.length; $i++ )
			{
				if ( $NewSCHTemplateName -eq $NewSCHTemplates[$i].Name )
				{
					$NewSCHTemplateId = $NewSCHTemplates[$i].Id
				}
			}
		
			if ( $NewSCHTemplateId -ne "" )
			{
				$command += "-SharedCallQueueHistoryTemplateId $NewSCHTemplateId "
			}
		}
		else
		{
			$command += "-SharedCallQueueHistoryTemplateId $SCHTemplateId "
		}
	}
   	

    #
    #  Compliance Recording Template
    #  Default: NONE (set in spreadsheet)
    #
    if ( $ComplianceRecordingTemplateIDs.length -gt 0 )
    {
		$command += "-ComplianceRecordingForCallQueueTemplateId @($ComplianceRecordingTemplateIDs) "
	  
		if ( $CR4CQGreetingOption -eq "FILE" )
		{
			$audioFileID = AudioFileImport $CR4CQGreeting

			$command += "-CustomAudioFileAnnouncementForCR $audioFileID "
		}
		  
		if ( $CR4CQGreetingOption -eq "TEXT" )
		{
			$command += "-TextAnnouncementForCR $CR4CQGreeting "
		}
	  
		if ( $CR4CQFailureGreetingOption -eq "FILE" )
		{
			$audioFileID = AudioFileImport $CR4CQFailureGreeting

			$command += "-CustomAudioFileAnnouncementForCRFailure $audioFileID "
		}
		  
		if ( $CR4CQFailureGreetingOption -eq "TEXT" )
		{
			$command += "-TextAnnouncementForCRFailure $CR4CQFailureGreeting "
		}

	}
   
   
	#
	#  Welcome Greeting   
	#  Default: None (set in spreadsheet)
	#  Pref: Audio file
	#
	if ( $GreetingOption -eq "FILE" )
	{
		$audioFileID = AudioFileImport $Greeting

		$command += "-WelcomeMusicAudioFileId $audioFileID "
	}
	if ( $GreetingOption -eq "TEXT" )
	{
		$command += "-WelcomeTextToSpeechPrompt $Greeting "
	}


	#
	#  Music On Hold      
	#  Default: Default Music (set in spreadsheet)
	#
	if ( $MusicOnHoldOption -eq "FILE" )
	{
		$audioFileID = AudioFileImport $MusicOnHold

		$command += "-MusicOnHoldAudioFileId $audioFileID "
	}
	else
	{
		$command += "-UseDefaultMusicOnHold " + "$" + "true "
	}


	#
	#  Conference Mode
	#  TO BE DEPRECATED
	#
	$command += "-ConferenceMode $" + "true "


	#
	#  Routing Method
	#  Default: Round Robin (set in spreadsheet)
	#
	$command += "-RoutingMethod $RoutingMethod "
 

	#
	#  Presence Based Routing 
	#  Default: True (set in spreadsheet)
	#
	$command += "-PresenceBasedRouting $" + $PresenceBasedRouting + " "


	#
	#  Allow Agent Opt Out
	#  Default: True (set in spreadsheet)
	#
	$command += "-AllowOptOut $" + $AllowOptOut + " "


	#
	#  Agent Alert Timer
	#  Default: 20 seconds (set in spreadsheet)
	#
	$command += "-AgentAlertTime $AgentAlertTime "


	#
	#  Overflow Exception 
	#
	#
	#  Overflow Threshold
	#  Default: 50 (set in spreadsheet)
	#
	$command += "-OverflowThreshold $OverflowThreshold "


	#
	#  Overflow Action
	#  Default: DisconnectWithBusy (set in spreadsheet)
	#
	switch ( $OverflowAction )
	{
		"DisconnectWithBusy"
		{
			$command += "-OverflowAction DisconnectWithBusy "

			if ( $OverflowTreatment -eq "FILE" )
			{
				$audioFileID = AudioFileImport $OverflowTreatmentPrompt

				$command += "-OverflowDisconnectAudioFilePrompt $audioFileID "
			}
			if ( $OverflowTreatment-eq "TEXT" )
			{
				$command += "-OverflowDisconnectTextToSpeechPrompt $OverflowTreatmentPrompt "
			}
		}

		"Forward-Person"
		{
			if ( $OverflowActionTarget -ne "" )
			{
				$command += "-OverflowAction Forward -OverflowActionTarget $OverflowActionTarget "

				if ( $OverflowTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $OverflowTreatmentPrompt

					$command += "-OverflowRedirectPersonAudioFilePrompt $audioFileID "
				}
				if ( $OverflowTreatment -eq "TEXT" )
				{
					$command += "-OverflowRedirectPersonTextToSpeechPrompt $OverflowTreatmentPrompt "
				}
			}
		}
      
		"Forward-VoiceApps"
		{
			if ( $OverflowActionTarget -ne "" )
			{
				if ( $OverflowActionTarget.Substring(0, 4) -eq "NEW:" )
				{
					$referenceCQName =  $OverflowActionTarget.Substring(4)
					# $OverflowActionTarget = (Get-CsCallQueue -NameFilter "$referenceCQName").Identity
					
					for ( $i = 0; $i -lt $NewCallQueues.length; $i++ )
					{
						if ( $referenceCQName -eq $NewCallQueues[$i].Name )
						{
							$OverflowActionTarget = $NewCallQueues[$i].Identity
							break
						}
					}
				}
				
				if ( $OverflowActionTarget -ne "" )
				{
					$command += "-OverflowAction Forward -OverflowActionTarget $OverflowActionTarget "

					if ( $OverflowTreatment -eq "FILE" )
					{
						$audioFileID = AudioFileImport $OverflowTreatmentPrompt

						$command += "-OverflowRedirectVoiceAppAudioFilePrompt $audioFileID "
					}
					if ( $OverflowTreatment -eq "TEXT" )
					{
						$command += "-OverflowRedirectVoiceAppTextToSpeechPrompt $OverflowTreatmentPrompt "
					}

					# will be blank if not flighted
					if ( $OverflowActionCallPriority -ne "" )
					{
						$command += "-OverflowActionCallPriority $OverflowActionCallPriority "
					}
				}
			}
		}

		"Forward-External"
		{
			if ( $OverflowActionTarget -ne "" )
			{
				$command += "-OverflowAction Forward -OverflowActionTarget $OverflowActionTarget "

				if ( $OverflowTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $OverflowTreatmentPrompt

					$command += "-OverflowRedirectPhoneNumberAudioFilePrompt $audioFileID "
				}
				if ( $OverflowTreatment -eq "TEXT" )
				{
					$command += "-OverflowRedirectPhoneNumberTextToSpeechPrompt $OverflowTreatmentPrompt "
				}
			}
		}

		"Voicemail"
		{
			if ( $OverflowActionTarget -ne "" )
			{
				$command += "-OverflowAction Voicemail -OverflowActionTarget $OverflowActionTarget "

				if ( $OverflowTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $OverflowTreatmentPrompt

					$command += "-OverflowRedirectVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $OverflowTreatment -eq "TEXT" )
				{
					$command += "-OverflowRedirectVoicemailTextToSpeechPrompt $OverflowTreatmentPrompt "
				}
			}
		}

		"SharedVoicemail"
		{
			if ( $OverflowActionTarget -ne "" )
			{
				$command += "-OverflowAction SharedVoicemail -OverflowActionTarget $OverflowActionTarget "

				if ( $OverflowTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $OverflowTreatmentPrompt

					$command += "-OverflowSharedVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $OverflowTreatment -eq "TEXT" )
				{
					$command += "-OverflowSharedVoicemailTextToSpeechPrompt $OverflowTreatmentPrompt "
				}

				$command += "-EnableOverflowSharedVoicemailSystemPromptSuppression $" + $OverflowSharedVoicemailSystemPromptSuppression + " "
				$command += "-EnableOverflowSharedVoicemailTranscription $" + $OverflowSharedVoicemailTranscription + " "
			}
		}
	}


	#
	#  Timeout Exception 
	#
	#
	#  Timeout Threshold
	#  Default: 1200 seconds (set in spreadsheet)
	#
	$command += "-TimeoutThreshold $TimeoutThreshold "


	#
	#  Timeout Action
	#  Default: Disconnect (set in spreadsheet)
	#
	switch ( $TimeoutAction )
	{
		"Disconnect"
		{
			$command += "-TimeoutAction Disconnect "

			if ( $TimeoutTreatment -eq "FILE" )
			{
				$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

				$command += "-TimeoutDisconnectAudioFilePrompt $audioFileID "
			}
			if ( $TimeoutTreatment -eq "TEXT")
			{
				$command += "-TimeoutDisconnectTextToSpeechPrompt $TimeoutTreatmentPrompt "
			}
		}

		"Forward-Person"
		{
			if ( $TimeoutActionTarget -ne "" )
			{
				$command += "-TimeoutAction Forward -TimeoutActionTarget $TimeoutActionTarget "

				if ( $TimeoutTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

					$command += "-TimeoutRedirectPersonAudioFilePrompt $audioFileID "
				}
				if ( $TimeoutTreatment -eq "TEXT" )
				{
					$command += "-TimeoutRedirectPersonTextToSpeechPrompt $TimeoutTreatmentPrompt "
				}
			}
		}
      
		"Forward-VoiceApps"
		{
			if ( $TimeoutActionTarget -ne "" )
			{
				if ( $TimeoutActionTarget.Substring(0, 4) -eq "NEW:" )
				{
					$referenceCQName =  $TimeoutActionTarget.Substring(4)
					#$TimeoutActionTarget = (Get-CsCallQueue -NameFilter "$referenceCQName").Identity
					
					for ( $i = 0; $i -lt $NewCallQueues.length; $i++ )
					{
						if ( $referenceCQName -eq $NewCallQueues[$i].Name )
						{
							$TimeoutActionTarget = $NewCallQueues[$i].Identity
							break
						}
					}
				}
 
				if ( $TimeoutActionTarget -ne "" )
				{
					$command += "-TimeoutAction Forward -TimeoutActionTarget $TimeoutActionTarget "

					if ( $TimeoutTreatment -eq "FILE" )
					{
						$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

						$command += "-TimeoutRedirectVoiceAppAudioFilePrompt $audioFileID "
					}
					if ( $TimeoutTreatment -eq "TEXT" )
					{
						$command += "-TimeoutRedirectVoiceAppTextToSpeechPrompt $TimeoutTreatmentPrompt "
					}

					# will be blank if not flighted
					if ( $TimeoutActionCallPriority -ne "" )
					{
						$command += "-TimeoutActionCallPriority $TimeoutActionCallPriority "
					}
				}
			}
		}

		"Forward-External"
		{
			if ( $TimeoutActionTarget -ne "" )
			{
				$command += "-TimeoutAction Forward -TimeoutActionTarget $TimeoutActionTarget "

				if ( $TimeoutTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

					$command += "-TimeoutRedirectPhoneNumberAudioFilePrompt $audioFileID "
				}
				if ( $TimeoutTreatment -eq "TEXT" )
				{
					$command += "-TimeoutRedirectPhoneNumberTextToSpeechPrompt $TimeoutTreatmentPrompt "
				}
			}
		}

		"Voicemail"
		{
			if ( $TimeoutActionTarget -ne "" )
			{
				$command += "-TimeoutAction Voicemail -TimeoutActionTarget $TimeoutActionTarget "

				if ( $TimeoutTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

					$command += "-TimeoutRedirectVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $TimeoutTreatment -eq "TEXT" )
				{
					$command += "-TimeoutRedirectVoicemailTextToSpeechPrompt $TimeoutTreatmentPrompt "
				}
			}
		}

		"SharedVoicemail"
		{
			if ( $TimeoutActionTarget -ne "" )
			{
				$command += "-TimeoutAction SharedVoicemail -TimeoutActionTarget $TimeoutActionTarget "

				if ( $TimeoutTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $TimeoutTreatmentPrompt

					$command += "-TimeoutSharedVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $TimeoutTreatment -eq "TEXT")
				{
					$command += "-TimeoutSharedVoicemailTextToSpeechPrompt $TimeoutTreatmentPrompt "
				}

				$command += "-EnableTimeoutSharedVoicemailSystemPromptSuppression $" + $TimeoutSharedVoicemailSystemPromptSuppression + " "
				$command += "-EnableTimeoutSharedVoicemailTranscription $" + $TimeoutSharedVoicemailTranscription + " "
			}
		}
	}

	#
    #  No Agents Exception 
    #
    #  Apply To
    #  Default: AllCalls
    #
    # no agents
	$command += "-NoAgentApplyTo $NoAgentNewCallsOnly "


    #
	#  No Agents Action
	#  Default: Queue
	#
	switch ( $NoAgentAction )
	{
		"Disconnect"
		{
			$command += "-NoAgentAction Disconnect "

			if ( $NoAgentTreatment -eq "FILE" )
			{
				$audioFileID = AudioFileImport $NoAgentTreatmentPrompt

				$command += "-NoAgentDisconnectAudioFilePrompt $audioFileID "
			}
			if ( $NoAgentTreatment -eq "TEXT" )
			{
				$command += "-NoAgentDisconnectTextToSpeechPrompt $NoAgentTreatmentPrompt "
			}
		}

		"Forward-Person"
		{
			if ( $NoAgentActionTarget -ne "" )
			{
				$command += "-NoAgentAction Forward -NoAgentActionTarget $NoAgentActionTarget "

				if ( $NoAgentTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $NoAgentTreatmentPrompt

					$command += "-NoAgentRedirectPersonAudioFilePrompt $audioFileID "
				}
				if ( $NoAgentTreatment -eq "TEXT" )
				{
					$command += "-NoAgentRedirectPersonTextToSpeechPrompt $NoAgentTreatmentPrompt "
				}
			}
		}
      
		"Forward-VoiceApps"
		{
			if ( $NoAgentActionTarget -ne "" )
			{
				if ( $NoAgentActionTarget.Substring(0, 4) -eq "NEW:" )
				{
					$referenceCQName =  $NoAgentActionTarget.Substring(4)
					#$NoAgentActionTarget = (Get-CsCallQueue -NameFilter "$referenceCQName").Identity
				
					for ( $i = 0; $i -lt $NewCallQueues.length; $i++ )
					{
						if ( $referenceCQName -eq $NewCallQueues[$i].Name )
						{
							$NoAgentActionTarget = $NewCallQueues[$i].Identity
							break
						}
					}
				}
 
				if ( $NoAgentActionTarget -ne "" )
				{
					$command += "-NoAgentAction Forward -NoAgentActionTarget $NoAgentActionTarget "

					if ( $NoAgentTreatment -eq "FILE" )
					{
						$audioFileID = AudioFileImport $NoAgentTreatmentPrompt

						$command += "-NoAgentRedirectVoiceAppAudioFilePrompt $audioFileID "
					}
					if ( $NoAgentTreatment -eq "TEXT" )
					{
						$command += "-NoAgentRedirectVoiceAppTextToSpeechPrompt $NoAgentTreatmentPrompt "
					}

					# will be blank if not flighted
					if ( $NoAgentActionCallPriority -ne "" )
					{
						$command += "-NoAgentActionCallPriority $NoAgentActionCallPriority "
					}
				}
			}
		}

		"Forward-External"
		{
			if ( $NoAgentActionTarget -ne "" )
			{
				$command += "-NoAgentAction Forward -NoAgentActionTarget $NoAgentActionTarget "

				if ( $NoAgentTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $NoAgentTreatmentPrompt

					$command += "-NoAgentRedirectPhoneNumberAudioFilePrompt $audioFileID "
				}
				if ( $NoAgentTreatment -eq "TEXT" )
				{
					$command += "-NoAgentRedirectPhoneNumberTextToSpeechPrompt $NoAgentTreatmentPrompt "
				}
			}
		}

		"Voicemail"
		{
			if ( $NoAgentActionTarget -ne "" )
			{
				$command += "-NoAgentAction Voicemail -NoAgentActionTarget $NoAgentActionTarget "

				if ( $NoAgentTreatment -eq "FILE" )
				{
				$audioFileID = AudioFileImport $NoAgentTreatmentFile

				$command += "-NoAgentRedirectVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $NoAgentTreatment -eq "TEXT" )
				{
					$command += "-NoAgentRedirectVoicemailTextToSpeechPrompt $NoAgentTreatmentPrompt "
				}
			}
		}

		"SharedVoicemail"
		{
			if ( $NoAgentActionTarget -ne "" )
			{
				$command += "-NoAgentAction SharedVoicemail -NoAgentActionTarget $NoAgentActionTarget "

				if ( $NoAgentTreatment -eq "FILE" )
				{
					$audioFileID = AudioFileImport $NoAgentTreatmentPrompt

					$command += "-NoAgentSharedVoicemailAudioFilePrompt $audioFileID "
				}
				if ( $NoAgentTreatment -eq "TEXT" )
				{
					$command += "-NoAgentSharedVoicemailTextToSpeechPrompt $NoAgentTreatmentPrompt "
				}

				$command += "-EnableNoAgentSharedVoicemailSystemPromptSuppression $" + $NoAgentSharedVoicemailSystemPromptSuppression + " "
				$command += "-EnableNoAgentSharedVoicemailTranscription $" + $NoAgentSharedVoicemailTranscription + " "
			}
		}
	}


	#
	#  Callback
	#  Default: False/Disabled
	#
	if ( ! $NoCallback )
	{
		if ( $IsCallbackEnabled -eq "true" -and $CallbackRequestDTMF -ne "" -and $CallbackEmailNotificationTarget -ne "" -and $CallbackOfferPrompt -ne "" )
		{
            $conditionSet = $false

            if ( $WaitTimeBeforeOfferingCallbackInSecond -ne "" )
            {
				if ( $WaitTimeBeforeOfferingCallbackInSecond.ToLower() -eq "null")
                {
                    $command += ("-WaitTimeBeforeOfferingCallbackInSecond $" + "null ")
                }
                else
                {
                    $command += "-WaitTimeBeforeOfferingCallbackInSecond $WaitTimeBeforeOfferingCallbackInSecond "
                }
                $conditionSet = $true
            }

			if ( $NumberOfCallsInQueueBeforeOfferingCallback -ne "" )
			{
				if ( $NumberOfCallsInQueueBeforeOfferingCallback.ToLower() -eq "null" )
				{
					$command += ("-NumberOfCallsInQueueBeforeOfferingCallback $" + "null ")
				}
				else
				{
					$command += "-NumberOfCallsInQueueBeforeOfferingCallback $NumberOfCallsInQueueBeforeOfferingCallback "
				}
				$conditionSet = $true
			}

            if ( $CallToAgentRatioThresholdBeforeOfferingCallback -ne "")
            {
                if ( $CallToAgentRatioThresholdBeforeOfferingCallback.ToLower() -eq "null" )
                {
                    $command += ("-CallToAgentRatioThresholdBeforeOfferingCallback $" + "null ")
                }
                else
                {
                    $command += "-CallToAgentRatioThresholdBeforeOfferingCallback $CallToAgentRatioThresholdBeforeOfferingCallback "
                }
                $conditionSet = $true
            }

            if ( $conditionSet )
            {
                if ( $CallbackOfferTreatment -eq "FILE" )
                {
                    $audioFileID = AudioFileImport $CallbackOfferPrompt

                    $command += "-CallbackOfferAudioFilePromptResourceId $audioFileID "
                }
                else
                {
                    $command += "-CallbackOfferTextToSpeechPrompt $CallbackOfferPrompt "
                }

                $command += "-IsCallbackEnabled $" + $IsCallbackEnabled + " -CallbackRequestDTMF $CallbackRequestDTMF -CallbackEmailNotificationTarget $CallbackEmailNotificationTarget "
            }
		}
		else
		{
			$command += "-IsCallbackEnabled $" + "false "
		}
	}
	else
	{
		Write-Host "`tCallback configuration skipped."
	}


	#
	#  Queue Membership
	#  Default: None
	#
	if ( $TeamGroupID -ne "" -And $TeamChannelID -ne "" -And $TeamChannelName -ne "" )
	{
		switch ( $TeamChannelID.substring(0,3) )
		{
			"19:"	{	$teamOwner = (Get-TeamChannelUser -GroupId $TeamGroupID -DisplayName $TeamChannelName -Role Owner).UserId

						if ( $teamOwner.count -gt 1 )
						{
							$command += "-ChannelId $TeamChannelID -ChannelUserObjectId " + $teamOwner[0] + " -DistributionList $TeamGroupID "
						}
						else
						{
							$command += "-ChannelId $TeamChannelID -ChannelUserObjectId $teamOwner -DistributionList $TeamGroupID "
						}
					}
			"TAG"	{	$command += "-ShiftsTeamId $TeamGroupID -ShiftsSchedulingGroupId $TeamChannelID " }				
		}
	}
	else
	{
		if ( $DL.count -gt 0 )
		{
			$command += "-DistributionLists @($DistributionLists) "
		}

		if ( $Agents.count -gt 0 )
		{
			$command += "-Users @($Users) "
		}
	}
   
   
   
	#
	#  Create Call Queue
	#

	if ( $Verbose )
	{
		Write-Host "`t-------------------------------------------------------------------"
		Write-Host "`t$command"
		Write-Host "`t-------------------------------------------------------------------"
	}

	Write-Host "`tCreating Call Queue : $Name" -f Green

	$command = '$Object = ( ' + $command
	$command += " )"
	
	Invoke-Expression $command 


	#
	#  Assign Resource Account
	#
	if ( ! $NoResourceAccounts )
	{
		if ( $ExistingResourceAccountName -ne "" )
		{
			Write-Host "`tAssigning Resource Account" -f Green
			# MicrosoftTeams
			New-CsOnlineApplicationInstanceAssociation -Identities @($ExistingResourceAccountName) -ConfigurationID $Object.Identity -ConfigurationType CallQueue | Out-Null
		}
		else
		{
			if ( ! $NoResourceAccountCreation )
			{
				if ( $NewResourceAccountPrincipalName -ne "" )
				{
					if ( $NewResourceAccountDisplayName -eq "" )
					{
						$NewResourceAccountDisplayName = $Name # Name of call queue
					}

					if ( $NewResourceAccountLocation -eq "" )
					{	
						$NewResourceAccountLocation = "US"
					}

					Write-Host "`tUpdating Resource Account ($NewResourceAccountPrincipalName)" -f Green
					for ( $i = 0; $i -lt $NewResourceAccounts.length; $i++ )
					{
						if ( $NewResourceAccountPrincipalName -eq $NewResourceAccounts[$i].UserPrincipalName )
						{
							$applicationInstanceID = (Get-CsOnlineUser -Identity $NewResourceAccounts[$i].ObjectId).Identity 2> $null

							if ( $applicationInstanceID.length -eq 0 )
							{
								$j = 1
								do
								{										
									for ($k = 0; $k -lt 30; $k++)
									{
										Write-Host -NoNewLine "`t`tWaiting 30 seconds for resource account creation to sync. (Attempt $j of 10) [$($k+1)]`r"
										Start-Sleep -Seconds 1
									}
									Write-Host " "
									$j++
						
									$applicationInstanceID = (Get-CsOnlineUser -Identity $NewResourceAccounts[$i].ObjectId).Identity 2> $null
								}
								until ( $applicationInstanceID.length -gt 0 -or $j -gt 10 )
								Write-Host " "
							}
							
							if ( $j -lt 10 )
							{
								$j = 1
								
								Write-Host "`tAssigning Location" -f Green
								# Module: Microsoft.Graph.Users
								Update-MgUser -UserId $NewResourceAccountPrincipalName -Id $NewResourceAccounts[$i].ObjectId -UsageLocation $NewResourceAccountLocation
								
								if ( ! $NoResourceAccountLicensing )
								{
									Write-Host "`tAttempting to license Resource Account" -f Green
									# Module: Microsoft.Graph.Identity.DirectoryManagement
									$skuID = (Get-MgSubscribedSKU | Where {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}).SkuId
									# Module: Microsoft.Graph.Users.Actions
									#Set-MgUserLicense -UserId $applicationInstanceID -AddLicenses @{SkuId = $skuID} -RemoveLicenses @() | Out-Null
									Set-MgUserLicense -UserId $NewResourceAccounts[$i].ObjectId -AddLicenses @{SkuId = $skuID} -RemoveLicenses @() | Out-Null
								}
								else
								{
									Write-Host "`tResource Account Licensing is disabled" -f Yellow
								}
								
								if ( ! $NoResourceAccountPhoneNumbers )
								{
									if ( $NewResourceAccountPhoneNumber -ne "" )
									{						
										$ProvisionedPlan = @((Get-CsOnlineuser -Identity $NewResourceAccounts[$i].ObjectId).ProvisionedPlan)

										if ( ! $ProvisionedPlan.Capability -contains "MCOEV_VIRTUALUSER" )
										{
											$j = 1
											do
											{
												for ($k = 0; $k -lt 30; $k++)
												{
													Write-Host -NoNewLine "`t`tWaiting 30 seconds for resource account licensing to sync. (Attempt $j of 10) [$($k+1)]`r"
													Start-Sleep -Seconds 1
												}
												Write-Host " "
												$j++
						
												$ProvisionedPlan = @((Get-CsOnlineuser -Identity $NewResourceAccounts[$i].ObjectId).ProvisionedPlan)
											}
											until ($ProvisionedPlan.Capability -contains "MCOEV_VIRTUALUSER" -OR $j -gt 10 )
											Write-Host " "
										}

										if( $j -lt 10 )
										{
											Write-Host "`tAssigning Phone number ($NewResourceAccountPhoneNumber)" -f Green
											Set-CsPhoneNumberAssignment -Identity $NewResourceAccounts[$i].ObjectId -PhoneNumber $NewResourceAccountPhoneNumber -PhoneNumberType CallingPlan
										}
										else
										{
											Write-Host "`tUnable to assign phone number ($NewResourceAccountPhoneNumber) - couldn't confirm licensing" -f Red
										}
									}
									else
									{
										Write-Host "`tResourceAccountPhoneNumber is blank" -f Red
									}
								}
								else
								{
									Write-Host "`tResource Account Phone Number Assignment is disabled" -f Yellow
								} # NoResourceAccountPhoneNumbers
								
								Write-Host "`tAssigning Resource Account" -f Green
								if ( $NewResourceAccountPriority -ne "" )
								{
									# New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue -CallPriority $NewResourceAccountPriority| Out-Null
									# Module: MicrosoftTeams	
									New-CsOnlineApplicationInstanceAssociation -Identities @($NewResourceAccounts[$i].ObjectId) -ConfigurationID $Object.Identity -ConfigurationType CallQueue -CallPriority $NewResourceAccountPriority | Out-Null
								}
								else
								{
									# New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue | Out-Null
									# Module: MicrosoftTeams
									New-CsOnlineApplicationInstanceAssociation -Identities @($NewResourceAccounts[$i].ObjectId) -ConfigurationID $Object.Identity -ConfigurationType CallQueue | Out-Null
								}
							}
							else
							{
								Write-Host "`tUnable to continue with provisioning resource account ($NewResourceAccountPhoneNumber) - couldn't confirm existence." -f Red
							} # Can't confirm Resource Account existence
						} # New Resource Account Not Found
					} # For loop
				}
				else
				{
					Write-Host "`tResourceAccountPrincipalName is blank" -f Red
				}
			}
			else
			{
				Write-Host "`tResource Account Creation is disabled" -f Yellow
			}
		}
	}
	else
	{
		Write-Host "`tResource Account Processing is disabled" -f Yellow
	}

	Write-Host "`tCall Queue Created" -f Green
	return $Object
}



#
# PowerShell Streams
#
#Stream #	Description			Write Cmdlet		Variable				Default
#1			Success stream		Write-Output
#2			Error stream		Write-Error			$ErrorActionPreference	Continue
#3			Warning stream		Write-Warning		$WarningPreference		Continue
#4			Verbose stream		Write-Verbose		$VerbosePrefernce		SilentlyContinue
#5			Debug stream		Write-Debug			$DebugPreference		SilentlyContinue
#6			Information stream	Write-Information	$InformationPreference	SilentlyContinue
#
#Preference Variable Options
# Use name or value
#
#Name				Value
#Break				6
#Suspend			5
#Ignore				4
#Inquire			3
#Continue			2
#Stop				1
#SilentlyContinue	0
#

#
# Main 
#
#
# Confirm running in PowerShell v5.x
#
if ( $PSVersionTable.PSVersion.Major -ne 5 )
{
	Write-Host "This script is only supported in PowerShell v5.x" -f Red
	exit
}

#
# Process arguments
#
$args = @()
$arguments = (Get-PSCallStack).Arguments
$arguments = $arguments -replace '[{}]', ''
$arguments = $arguments[0].ToLower()
$arguments = $arguments -split ", "
$args += $arguments

for ( $i = 0; $i -lt $args.length; $i++ )
{	
	switch ( $args[$i] )
	{
		"-help"             				{ 	$Help = $true }
		"-excelfile"        				{ 
												$ExcelFilename = $args[$i+1]
												$i++
											}
		"-nocallback"       	 			{ 	$NoCallback = $true }
		"-noresourceaccounts" 				{	$NoResourceAccounts = $true
												$NoResourceAccountCreation = $true
												$NoResourceAccountLicensing = $true
												$NoResourceAccountPhoneNumbers = $true
											}
	   
		"-noresourceaccountcreation"		{	$NoResourceAccountCreation = $true
												$NoResourceAccountLicensing = $true
												$NoResourceAccountPhoneNumbers = $true
											}
		"-noresourceaccountlicensing" 		{ 	$NoResourceAccountLicensing = $true
												$NoResourceAccountPhoneNumbers = $true
											}
		
		"-noresourceaccountphonenumbers"	{ 	$NoResourceAccountPhoneNumbers = $true }


		"-verbose"            				{ 	$Verbose = $true }
		Default      						{ 	$ArgError = $true
												$arg = $args[$i]
											}
	}
}

if ( $ArgError )
{
	Write-Host "An unknown argument was encountered: $arg" -f Red
}

if ( ( $Help ) -or ( $ArgError ) )
{
	Write-Host "The following options are avaialble:"
	Write-Host "`t-Help - shows the options that are available (this help message)"
	Write-Host "`t-ExcelFile - the Excel file to use.  Default is BulkCQs.xlsm"
	Write-Host "`t-NoCallback - will not configure the callback component"
	Write-Host "`t-NoResourceAccounts - will not configure the resource account to the call queue"
	Write-Host "`t-NoResourceAccountCreation - don't create or license new Resource Accounts"
	Write-Host "`t-NoResourceAccountLicensing - don't assign a license to new Resource Accounts"
    Write-Host "`t-NoResourceAccountPhoneNumbers - don't assign a phone number to a new Resource Accounts"
	Write-Host "`t-Verbose - provides extra messaging during the process"
	exit
}

$StartTime = (Get-Date -Format "HH:mm:ss")
Write-Host "$StartTime - Starting BulkCQsProvisioning." -f Green
Write-Host "`tCleaning up from any previous runs."

if ( Test-Path -Path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
}

if ( Test-Path -Path ".\PS-SCH.csv" )
{
   Remove-Item -Path ".\PS-SCH.csv" | Out-Null
}

if ( Test-Path -Path ".\Data.csv" )
{
   Remove-Item -Path ".\Data.csv" | Out-Null
}

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768

#
# Module Min Supported Versions
#
$MicrosoftTeamsMinVersion = [version]"7.1.0"
$MicrosoftGraphMinVersion = [version]"2.24.0"
$ImportExcelMinVersion = [version]"7.8.0"

#
# BulkCQs.xlms Required Version
#
$ExcelSpreadsheetRequiredVersion = "1.0.6"
$ExcelSpreadsheetVersionSheet = "Data"


#
# Check that required modules are installed - install if not
#
Write-Host "Checking for MicrosoftTeams module $MicrosoftTeamsMinVersion or later." -f Green
$Version = ( (Get-InstalledModule -Name MicrosoftTeams -MinimumVersion "$MicrosoftTeamsMinVersion").Version 2> $null )

if ( $Version -match "-preview" )
{
	$Version = $Version.Replace("-preview", "")
}

$Version = [version]$Version

Write-Host "`tMicrosoftTeams Version: $Version"
if ( $Version -ge $MicrosoftTeamsMinVersion )
{
	Write-Host "`tConnecting to Microsoft Teams."
	Import-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion
   
	try
	{ 
		Get-CsTenant -ErrorAction Stop | Out-Null
	} 
	catch [System.UnauthorizedAccessException] 
	{ 
		Connect-MicrosoftTeams | Out-Null
	}
	try
	{ 
		Get-CsTenant -ErrorAction Stop | Out-Null
	} 
	catch [System.UnauthorizedAccessException] 
	{ 
		Write-Error "`tNot signed into Microsoft Teams!" -f Red
		exit
	}
	Write-Host "`tConnected to Microsoft Teams." -f Green
}
else
{
	Write-Host "The MicrosoftTeams module is not installed or does not meet the minimum requirements - installing." -f Yellow
	Install-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion -Force -AllowClobber

	Write-Host "`tConnecting to Microsoft Teams."
	Import-Module -Name MicrosoftTeams -MinimumVersion $MicrosoftTeamsMinVersion
   
	try
	{ 
		Get-CsTenant -ErrorAction Stop | Out-Null
	} 
	catch [System.UnauthorizedAccessException] 
	{ 
		Connect-MicrosoftTeams | Out-Null
	}
	try
	{ 
		Get-CsTenant -ErrorAction Stop | Out-Null
	} 
	catch [System.UnauthorizedAccessException] 
	{ 
		Write-Error "Not signed into Microsoft Teams!" -f Red
		exit
	}
	Write-Host "Connected to Microsoft Teams." -f Green
}


if ( (! $NoResourceAccounts ) -or (! $NoResourceAccountCreation) -or (! $NoResourceAccountLicensing) )
{
	Write-Host "Checking for Microsoft.Graph module $MicrosoftGraphMinVersion or later." -f Green
	$Version = ( (Get-InstalledModule -Name Microsoft.Graph -MinimumVersion "$MicrosoftGraphMinVersion").Version 2> $null)
	
	if ( $Version -match "-preview" )
	{
		$Version = $Version.Replace("-preview", "")
	}

	$Version = [version]$Version
	
	Write-Host "`tMicrosoft.Graph Version: $Version"
	if ( $Version -ge $MicrosoftGraphMinVersion )
	{
		Write-Host "`tConnecting to Microsoft Graph."
	   
		Disconnect-MgGraph | Out-Null
		Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null

		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
		}
		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Write-Error "Not signed into Microsoft Graph!" -f Red
			exit
		}
		Write-Host "`tConnected to Microsoft Graph." -f Green
	}
	else
	{
		Write-Host "The Microsoft.Graph module is not installed or does not meet the minimum requirements - installing." -f Yellow
		Install-Module -Name Microsoft.Graph -MinimumVersion $MicrosoftGraphMinVersion -Force -AllowClobber

		Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null

		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
		}
		try
		{ 
			Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
		} 
		catch [System.UnauthorizedAccessException] 
		{ 
			Write-Error "Not signed into Microsoft Graph!" -f Red
			exit
		}
		Write-Host "`tConnected to Microsoft Graph." -f Green
	}
}

Write-Host "Checking for ImportExcel module $ImportExcelMinVersion or later." -f Green
$Version = ( (Get-InstalledModule -Name ImportExcel -MinimumVersion "$ImportExcelMinVersion").Version 2> $null )

if ( $Version -match "-preview" )
{
	$Version = $Version.Replace("-preview", "")
}

$Version = [version]$Version

Write-Host "`tImportExcel Version: $Version"
if ( $Version -ge $ImportExcelMinVersion )
{
	Write-Host "`tImporting ImportExcel." -f Green
	Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
}
else
{
	Write-Host "The ImportExcel module is not installed or does not meet the minimum requirements - installing." -f Yellow
	Install-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion -Force -AllowClobber
   
	Write-Host "`tImporting ImportExcel." -f Green
	Import-Module -Name ImportExcel -MinimumVersion $ImportExcelMinVersion
}

#
# setup filename
#
if ( $ExcelFilename -eq $null )
{
	$ExcelFilename = "BulkCQs.xlsm"
}
$ExcelFullPathFilename = $PSScriptRoot + "\" + $ExcelFilename

#
# check if supplied filename exists
#
if ( !( Test-Path -Path $ExcelFullPathFilename ) )
{
	Write-Host "ERROR: $ExcelFilename does not exist." -f Red
	exit
}

#
# Checking Spreadsheet Version
#
Write-Host "Checking that $ExcelFilename is the correct version." -f Green

$ExcelDataCSVFilename = ".\Data.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelSpreadsheetVersionSheet

if ( $data.length -lt 1 )
{
	Write-Host "`tERROR: No data retrieved from $ExcelWorkSheetName" -f Red
	exit
}
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelDataCSVFilename -NoTypeInformation


$DataConfig = @(Import-csv -Path $ExcelDataCSVFilename)
$ExcelSpreadsheetVersion = $DataConfig[0].Version

if ( $ExcelSpreadsheetVersion -ne $ExcelSpreadsheetRequiredVersion )
{
	Write-Host "`tThe $ExcelFileName version ($ExcelSpreadsheetVersion) does not match the required version ($ExcelSpreadsheetRequiredVersion)."  -f Red
	exit
}


#
# Call Queue configuration
#
Write-Host "Starting Call Queue Configuration." -f Green

# Shared Call History
$ExcelWorkSheetName = "PS-SCH"
$ExcelSCHCSVFilename = ".\PS-SCH.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName

if ( $data.length -lt 1 )
{
	Write-Host "ERROR: No data retrieved from $ExcelWorkSheetName" -f Red
	exit
}
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelSCHCSVFilename -NoTypeInformation

#
# Process for Shared Call History
#
$PSCQConfig = @(Import-csv -Path $ExcelSCHCSVFilename)

Write-Host "Checking for new Shared Call Queue History Templates" -f Green

$NewSCHTemplates = @()
for ($i=0; $i -lt  $PSCQConfig.length; $i++)
{
	if ( $PSCQConfig[$i].NewSCHTemplateName -ne 0 )
	{
		$NewSCHTemplateName = '"' + $PSCQConfig[$i].NewSCHTemplateName + '"'
		$NewSCHTemplateDescription = '"' + $PSCQConfig[$i].NewSCHTemplateDescription + '"'
		$NewSCHTemplateIncomingMissedCalls = $PSCQConfig[$i].NewSCHTemplateIncomingMissedCalls
		$NewSCHTemplateAnsweredAndOutboundCalls = $PSCQConfig[$i].NewSCHTemplateAnsweredAndOutboundCalls

		$NewSCHTemplates += NewSCHTemplate
	}
}

if ( $Verbose )
{	
	Write-Host ( "`tTotal number of new Shared Call History Templates: {0,4}" -f $NewSCHTemplates.length  )
}


# Call Queues
$ExcelWorkSheetName = "PS-CQ"
$ExcelCQCSVFilename = ".\PS-CQ.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName

if ( $data.length -lt 1 )
{
	Write-Host "ERROR: No data retrieved from $ExcelWorkSheetName" -f Red
	exit
}
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCQCSVFilename -NoTypeInformation


#
# Process for Resource Accounts 
#
$PSCQConfig = @(Import-csv -Path $ExcelCQCSVFilename)

if ( ! $NoResourceAccountCreation )
{
	Write-Host "Checking for new Resource Accounts" -f Green
	
	$NewResourceAccounts = @()
	for ($i=0; $i -lt  $PSCQConfig.length; $i++)
	{
		if ( $PSCQConfig[$i].NewResourceAccountPrincipalName -ne "" )
		{
			$NewResourceAccountPrincipalName = $PSCQConfig[$i].NewResourceAccountPrincipalName
			$NewResourceAccountDisplayName = '"' + $PSCQConfig[$i].NewResourceAccountDisplayName + '"'

			$NewResourceAccounts += NewResourceAccount
		}
	}
}

if ( $Verbose )
{	
	Write-Host ( "`tTotal number of new Resource Accounts: {0,4}" -f $NewResourceAccounts.length  )
}


#
# Process for Call Queues
#
Write-Host "Checking for new Call Queues" -f Green
$NewCallQueues = @()
for ($i=0; $i -lt  $PSCQConfig.length; $i++)
{
	$Action = $PSCQConfig[$i].Action
	if ( $Action -eq "New" )
	{
		$Name   = '"' + $PSCQConfig[$i].Name + '"'
		$ExistingResourceAccountName = $PSCQConfig[$i].ExistingResourceAccountName
		$NewResourceAccountPrincipalName = $PSCQConfig[$i].NewResourceAccountPrincipalName
		$NewResourceAccountDisplayName = '"' + $PSCQConfig[$i].NewResourceAccountDisplayName + '"'
		$NewResourceAccountLocation = $PSCQConfig[$i].NewResourceAccountLocation
		$NewResourceAccountPhoneNumber = $PSCQConfig[$i].NewResourceAccountPhoneNumber
		$NewResourceAccountPriority = $PSCQConfig[$i].NewResourceAccountPriority

		$OutboundCLID = @()
		if ( $PSCQConfig[$i].OutboundCLID01 -ne "" )
		{
			$OutboundCLID += '"' + $PSCQConfig[$i].OutboundCLID01 + '"'
		}

		if ( $PSCQConfig[$i].OutboundCLID02 -ne "" )
		{
			$OutboundCLID += '"' + $PSCQConfig[$i].OutboundCLID02 + '"'
		}

		if ( $PSCQConfig[$i].OutboundCLID03 -ne "" )
		{
			$OutboundCLID += '"' + $PSCQConfig[$i].OutboundCLID03 + '"'
		}

		if ( $PSCQConfig[$i].OutboundCLID04 -ne "" )
		{
			$OutboundCLID += '"' + $PSCQConfig[$i].OutboundCLID04 + '"'
		}
		$OboResourceAccountIDs = $OutboundCLID -join ","

		$Language = $PSCQConfig[$i].Language
		$ServiceLevelThreshold = $PSCQConfig[$i].ServiceLevelThreshold

		$SCHTemplateId = $PSCQConfig[$i].SharedCallQueueHistoryTemplateId

		$Future01 = $PSCQConfig[$i].Future01
		$Future02 = $PSCQConfig[$i].Future02
		$Future03 = $PSCQConfig[$i].Future03
		$Future04 = $PSCQConfig[$i].Future04
		$Future05 = $PSCQConfig[$i].Future05
	  
		$ComplianceRecordingTemplate = @()
		if ( ( $PSCQConfig[$i].CR4CQ01 -ne "NOTENABLED" ) -and  ( $PSCQConfig[$i].CR4CQ01 -ne "NONE" ) )
		{
			$ComplianceRecordingTemplate += '"' + $PSCQConfig[$i].CR4CQ01 + '"'
		}
	  
		if ( ( $PSCQConfig[$i].CR4CQ02 -ne "NOTENABLED" ) -and ( $PSCQConfig[$i].CR4CQ02 -ne "NONE" ) )
		{
			$ComplianceRecordingTemplate += '"' + $PSCQConfig[$i].CR4CQ02 + '"'
		}
		$ComplianceRecordingTemplateIDs = $ComplianceRecordingTemplate -join ","
	  
		$CR4CQGreetingOption = $PSCQConfig[$i].CR4CQGreetingOption
		switch ( $CR4CQGreetingOption )
		{
			"FILE"  { $CR4CQGreeting = $PSCQConfig[$i].CR4CQGreeting }
			"TEXT"  { $CR4CQGreeting = '"' + $PSCQConfig[$i].CR4CQGreeting + '"' }
			Default { $CR4CQGreeting = "" }
		}

		$CR4CQFailureGreetingOption = $PSCQConfig[$i].CR4CQFailureGreetingOption
		switch ( $CR4CQFailureGreetingOption )
		{
			"FILE"  { $CR4CQFailureGreeting = $PSCQConfig[$i].CR4CQFailureGreeting }
			"TEXT"  { $CR4CQFailureGreeting = '"' + $PSCQConfig[$i].CR4CQFailureGreeting + '"' }
			Default { $CR4CQFailureGreeting = "" }
		}
	  
		$GreetingOption = $PSCQConfig[$i].GreetingOption
		switch ( $GreetingOption )
		{
			"FILE"  { $Greeting = $PSCQConfig[$i].Greeting }
			"TEXT"  { $Greeting = '"' + $PSCQConfig[$i].Greeting + '"' }
			Default { $Greeting= "" }
		}

		$MusicOnHoldOption = $PSCQConfig[$i].MusicOnHoldOption
		$MusicOnHOld = $PSCQConfig[$i].MusicOnHold
		$RoutingMethod = $PSCQConfig[$i].RoutingMethod
		$PresenceBasedRouting = $PSCQConfig[$i].PresenceBasedRouting
		$AllowOptOut = $PSCQConfig[$i].AllowOptOut
		$AgentAlertTime = $PSCQConfig[$i].AgentAlertTime

		$OverflowThreshold = $PSCQConfig[$i].OverflowThreshold
		$OverflowAction = $PSCQConfig[$i].OverflowAction
		$OverflowActionTarget = $PSCQConfig[$i].OverflowActionTarget
		$OverflowActionCallPriority = $PSCQConfig[$i].OverflowActionCallPriority

		$OverflowTreatment = $PSCQConfig[$i].OverflowTreatment
		switch ( $OverflowTreatment )
		{
			"FILE"  { $OverflowTreatmentPrompt = $PSCQConfig[$i].OverflowTreatmentPrompt }
			"TEXT"  { $OverflowTreatmentPrompt = '"' + $PSCQConfig[$i].OverflowTreatmentPrompt + '"' }
			Default { $OverflowTreatmentPrompt = "" }
		}

		$OverflowSharedVoicemailSystemPromptSuppression = $PSCQConfig[$i].OverflowSharedVoicemailSystemPromptSuppression
		$OverflowSharedVoicemailTranscription = $PSCQConfig[$i].OverflowSharedVoicemailTranscription

		$TimeoutThreshold = $PSCQConfig[$i].TimeoutThreshold
		$TimeoutAction = $PSCQConfig[$i].TimeoutAction
		$TimeoutActionTarget = $PSCQConfig[$i].TimeoutActionTarget
		$TimeoutActionCallPriority = $PSCQConfig[$i].TimeoutActionCallPriority

		$TimeoutTreatment = $PSCQConfig[$i].TimeoutTreatment
		switch ( $TimeoutTreatment )
		{
			"FILE"  { $TimeoutTreatmentPrompt = $PSCQConfig[$i].TimeoutTreatmentPrompt }
			"TEXT"  { $TimeoutTreatmentPrompt = '"' + $PSCQConfig[$i].TimeoutTreatmentPrompt + '"' }
			Default { $TimeoutTreatmentPrompt = "" }
		}

		$TimeoutSharedVoicemailSystemPromptSuppression = $PSCQConfig[$i].TimeoutSharedVoicemailSystemPromptSuppression
		$TimeoutSharedVoicemailTranscription = $PSCQConfig[$i].TimeoutSharedVoicemailTranscription

		$NoAgentNewCallsOnly = $PSCQConfig[$i].NoAgentNewCallsOnly
		$NoAgentAction = $PSCQConfig[$i].NoAgentAction
		$NoAgentActionTarget = $PSCQConfig[$i].NoAgentActionTarget
		$NoAgentActionCallPriority = $PSCQConfig[$i].NoAgentActionCallPriority

		$NoAgentTreatment = $PSCQConfig[$i].NoAgentTreatment
		switch ( $NoAgentTreatment )
		{
			"FILE"  { $NoAgentTreatmentPrompt = $PSCQConfig[$i].NoAgentTreatmentPrompt }
			"TEXT"  { $NoAgentTreatmentPrompt = '"' + $PSCQConfig[$i].NoAgentTreatmentPrompt + '"' }
			Default { $NoAgentTreatmentPrompt = "" }
		}

		$NoAgentSharedVoicemailSystemPromptSuppression = $PSCQConfig[$i].NoAgentSharedVoicemailSystemPromptSuppression
		$NoAgentSharedVoicemailTranscription = $PSCQConfig[$i].NoAgentSharedVoicemailTranscription

		$IsCallbackEnabled = $PSCQConfig[$i].IsCallbackEnabled
		$CallbackRequestDTMF = $PSCQConfig[$i].CallbackRequestDTMF
		$WaitTimeBeforeOfferingCallbackInSecond = $PSCQConfig[$i].WaitTimeBeforeOfferingCallbackInSecond
		$NumberOfCallsInQueueBeforeOfferingCallback = $PSCQConfig[$i].NumberOfCallsInQueueBeforeOfferingCallback
		$CallToAgentRatioThresholdBeforeOfferingCallback = $PSCQConfig[$i].CallToAgentRatioThresholdBeforeOfferingCallback

		$CallbackOfferTreatment = $PSCQConfig[$i].CallbackOfferTreatment
		switch ( $CallbackOfferTreatment )
		{
			"FILE"  { $CallbackOfferPrompt = $PSCQConfig[$i].CallbackOfferPrompt }
			"TEXT"  { $CallbackOfferPrompt = '"' + $PSCQConfig[$i].CallbackOfferPrompt + '"' }
			Default { $CallbackOfferTreatment = "" }
		}

		if ( $PSCQConfig[$i].CallbackEmailNotificationTarget -ne "" )
		{
			$CallbackEmailNotificationTarget = [System.GUID]::Parse($PSCQConfig[$i].CallbackEmailNotificationTarget)
		}

		$TeamGroupID = $PSCQConfig[$i].TeamGroupID
		$TeamChannelID = $PSCQConfig[$i].TeamChannelID
		$TeamChannelName = $PSCQConfig[$i].TeamChannelName

		$DL = @()
		if ( $PSCQConfig[$i].DistributionList01 -ne "" )
		{
			$DL += '"' + $PSCQConfig[$i].DistributionList01 + '"'
		}

		if ( $PSCQConfig[$i].DistributionList02 -ne "" )
		{
			$DL += '"' + $PSCQConfig[$i].DistributionList02 + '"'
		}

		if ( $PSCQConfig[$i].DistributionList03 -ne "" )
		{
			$DL += '"' + $PSCQConfig[$i].DistributionList03 + '"'
		}

		if ( $PSCQConfig[$i].DistributionList04 -ne "" )
		{
			$DL += '"' + $PSCQConfig[$i].DistributionList04 + '"'
		}
		$DistributionLists = $DL -join ","


		$Agents = @()
		if ( $PSCQConfig[$i].Agent01 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent01 + '"'
		}

		if ( $PSCQConfig[$i].Agent02 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent02 + '"'
		}

		if ( $PSCQConfig[$i].Agent03 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent03 + '"'
		}

		if ( $PSCQConfig[$i].Agent04 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent04 + '"'
		}

		if ( $PSCQConfig[$i].Agent05 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent05 + '"'
		}

		if ( $PSCQConfig[$i].Agent06 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent06 + '"'
		}

		if ( $PSCQConfig[$i].Agent07 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent07 + '"'
		}

		if ( $PSCQConfig[$i].Agent08 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent08 + '"'
		}

		if ( $PSCQConfig[$i].Agent09 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent09 + '"'
		}

		if ( $PSCQConfig[$i].Agent10 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent10 + '"'
		}

		if ( $PSCQConfig[$i].Agent11 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent11 + '"'
		}

		if ( $PSCQConfig[$i].Agent12 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent12 + '"'
		}

		if ( $PSCQConfig[$i].Agent13 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent13 + '"'
		}

		if ( $PSCQConfig[$i].Agent14 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent14 + '"'
		}

		if ( $PSCQConfig[$i].Agent15 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent15 + '"'
		}

		if ( $PSCQConfig[$i].Agent16 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent16 + '"'
		}

		if ( $PSCQConfig[$i].Agent17 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent17 + '"'
		}

		if ( $PSCQConfig[$i].Agent18 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent18 + '"'
		}

		if ( $PSCQConfig[$i].Agent19 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent19 + '"'
		}

		if ( $PSCQConfig[$i].Agent20 -ne "" )
		{
			$Agents += '"' + $PSCQConfig[$i].Agent20 + '"'
		}
		$Users = $Agents -join ","

		$NewCallQueues += NewCallQueue
	}
}


#
# Calculate run time
#
$EndTime = (Get-Date -Format "HH:mm:ss")
$Duration = New-Timespan -start $StartTime -end $EndTime
$DurationFormatted = ""

if ( $Duration.Hours -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Hours.ToString() + ":"
}
else
{
	$DurationFormatted += $Duration.Hours.ToString() + ":"
}

if ( $Duration.Minutes -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Minutes.ToString() + ":"
}
else
{
	$DurationFormatted += $Duration.Minutes.ToString() + ":"
}

if ( $Duration.Seconds -lt 10 )
{
	$DurationFormatted += "0" + $Duration.Seconds.ToString()
}
else
{
	$DurationFormatted += $Duration.Seconds.ToString()
}

Write-Host "$EndTime - Completed Call Queue Configuration." -f Green

if ( $Verbose )
{
	Write-Host ( "`tTotal number of new call queues: {0,4}" -f $NewCallQueues.length )
}
Write-Host "Duration: $DurationFormatted" -f Green

#
# Cleanup files
#
if ( Test-Path -Path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
}

if ( Test-Path -Path ".\PS-SCH.csv" )
{
   Remove-Item -Path ".\PS-SCH.csv" | Out-Null
}

if ( Test-Path -Path ".\Data.csv" )
{
   Remove-Item -Path ".\Data.csv" | Out-Null
}

exit



	