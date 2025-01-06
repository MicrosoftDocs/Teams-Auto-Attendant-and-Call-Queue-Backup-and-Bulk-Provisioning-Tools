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
   Write-Host "Action:  $Action`tName: $Name"
   Write-Host "-----------------------------------------------------------------"

   Write-Host "`tExistingResourceAccoutName : $ExistingResourceAccountName"
   Write-Host "`tNewResourceAccountPrincipalName : $NewResourceAccountPrincipalName"
   Write-Host "`tNewResourceAccountDisplayName : $NewResourceAccountDisplayName"
   Write-Host "`tNewResourceAccountLocation : $NewResourceAccountLocation"
   Write-Host "`tNewResourceAccountPriority : $NewResourceAccountPriority"
   Write-Host "`tOutboundCLID : $OboResourceAccountIDs"
   Write-Host "`tLanguage : $Language"
   Write-Host "`tServiceLevelThreshold : $ServiceLevelThreshold"
   Write-Host "`tGreetingOption : $GreetingOption"
   Write-Host "`tGreeting : $Greeting"
   Write-Host "`tMusicOnHoldOption : $MusicOnHoldOption"
   Write-Host "`tMusicOnHold : $MusicOnHold"
   Write-Host "`tRoutingMethod : $RoutingMethod"
   Write-Host "`tPresenceBasedRouting : $PresenceBasedRouting"
   Write-Host "`tAllowOptOut : $AllowOptOut"
   Write-Host "`tAgentAlertTime : $AgentAlertTime"

   Write-Host "`tOverflowThreshold : $OverflowThreshold"
   Write-Host "`tOverflowAction : $OverflowAction"
   Write-Host "`tOverflowActionTarget : $OverflowActionTarget"
   Write-Host "`tOverflowActionCallPriority : $OverflowActionCallPriority"
   Write-Host "`tOverflowTreatment : $OverflowTreatment"
   Write-Host "`tOverflowTreatmentPrompt : $OverflowTreatmentPrompt"
   Write-Host "`tOverflowSharedVoicemailSystemPromptSuppression : $OverflowSharedVoicemailSystemPromptSuppression"
   Write-Host "`tOverflowSharedVoicemailTranscription : $OverflowSharedVoicemailTranscription"

   Write-Host "`tTimeoutThreshold : $TimeoutThreshold"
   Write-Host "`tTimeoutAction : $TimeoutAction"
   Write-Host "`tTimeoutActionTarget : $TimeoutActionTarget"
   Write-Host "`tTimeoutActionCallPriority : $TimeoutActionCallPriority"
   Write-Host "`tTimeoutTreatmentTreatment : $TimeoutTreatment"
   Write-Host "`tTimeoutTreatmentPrompt : $TimeoutTreatmentPrompt"
   Write-Host "`tTimeoutSharedVoicemailSystemPromptSuppression : $TimeoutSharedVoicemailSystemPromptSuppression"
   Write-Host "`tTimeoutSharedVoicemailTranscription : $TimeoutSharedVoicemailTranscription"

   Write-Host "`tNoAgentNewCallsOnly : $NoAgentNewCallsOnly"
   Write-Host "`tNoAgentAction : $NoAgentAction"
   Write-Host "`tNoAgentActionTarget : $NoAgentActionTarget"
   Write-Host "`tNoAgentActionCallPriority : $NoAgentActionCallPriority"
   Write-Host "`tNoAgentTreatment : $NoAgentTreatment"
   Write-Host "`tNoAgentTreatmentPrompt  : $NoAgentTreatmentPrompt"
   Write-Host "`tNoAgentSharedVoicemailSystemPromptSuppression : $NoAgentSharedVoicemailSystemPromptSuppression"
   Write-Host "`tNoAgentsharedVoicemailTranscription : $NoAgentSharedVoicemailTranscription"

   Write-Host "`tIsCallbackEnabled : $IsCallbackEnabled"
   Write-Host "`tCallBackRequestDTMF : $CallbackRequestDTMF"
   Write-Host "`tWaitTimeBeforeOfferingCallbackInSecond : $WaitTimeBeforeOfferingCallbackInSecond"
   Write-Host "`tNumberOfCallsInQueueBeforeOfferingCallback : $NumberOfCallsInQueueBeforeOfferingCallback"
   Write-Host "`tCallToAgentRatioThresholdBeforeOfferingCallback : $CallToAgentRatioThresholdBeforeOfferingCallback"
   Write-Host "`tCallbackOfferTreatment : $CallbackOfferTreatment"
   Write-Host "`tCallbackOfferPrompt : $CallbackOfferPrompt"
   Write-Host "`tCallbackEmailNotificationTarget : $CallbackEmailNotificationTarget"

   Write-Host "`tTeamGroupID : $TeamGroupID"
   Write-Host "`tTeamChannelID : $TeamChannelID"
   Write-Host "`tTeamChannelName : $TeamChannelName"

   Write-Host "`tDistribtuionList : $DistributionLists"

   Write-Host "`tAgents: $Users"

   Write-Host "-----------------------------------------------------------------"
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
      Write-Host "Action:  $Action`tName: $Name"
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

      $teamOwner = (Get-TeamChannelUser -GroupId $TeamGroupID -DisplayName $TeamChannelName -Role Owner).UserId

      if ( $teamOwner.count -gt 1 )
      {
         $command += "-ChannelId $TeamChannelID -ChannelUserObjectId " + $teamOwner[0] + " -DistributionList $TeamGroupID "
      }
      else
      {
         $command += "-ChannelId $TeamChannelID -ChannelUserObjectId $teamOwner -DistributionList $TeamGroupID "
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

   $command += "| Out-Null"

   if ( $Verbose )
   {
      Write-Host "-----------------------------------------------------------------"
      Write-Host $command
      Write-Host "-----------------------------------------------------------------"
  }

   Write-Host "Creating Call Queue : $Name"

   Invoke-Expression $command


	#
	#  Assign Resource Account
	#
	if ( ! $NoResourceAccounts )
	{
		# $callqueueID = Invoke-Expression "(Get-CsCallQueue -name $Name).Identity"
		$CallQueue = Invoke-Expression "Get-CsCallQueue -NameFilter $Name"
		
		if ( $CallQueue.length -eq 1 )
		{
			if ( $ExistingResourceAccountName -ne "" )
			{
				Write-Host "Assigning Resource Account"
				# New-CsOnlineApplicationInstanceAssociation -Identities @($ExistingResourceAccountName) -ConfigurationID $callQueueID -ConfigurationType CallQueue | Out-Null
				New-CsOnlineApplicationInstanceAssociation -Identities @($ExistingResourceAccountName) -ConfigurationID $CallQueue.Identity -ConfigurationType CallQueue | Out-Null
			}
			else
			{
				if ( ! $NoResourceAccountCreation )
				{
					if ( $NewResourceAccountPrincipalName -ne "" )
					{
						if ( $NewResourceAccountDisplayName -eq "" )
						{
							$NewResourceAccountDisplayName = $Name
						}

						Write-Host "Creating Resource Account ($NewResourceAccountPrincipalName)"
						
						New-CsOnlineApplicationInstance -UserPrincipalName $NewResourceAccountPrincipalName -DisplayName $NewResourceAccountDisplayName -ApplicationID "11cd3e2e-fccb-42ad-ad00-878b93575e07" | Out-Null
			
						$j = 1
						do
						{
							Write-Host -NoNewLine "`tResource Account created, pausing for 10 seconds to allow sync to occur. (Attempt $j of 10) "
							for ($i = 0; $i -lt 10; $i++)
							{
								Write-Host -NoNewline "."
								Start-Sleep -Seconds 1
							}
							$j++

							$applicationInstanceID = (Get-CsOnlineUser -Identity $NewResourceAccountPrincipalName).Identity 2> `$null
						}
						until ( $applicationInstanceID.length -gt 0 -or $j -gt 10 )
						Write-Host " "

						if ( $NewResourceAccountLocation -eq "" )
						{	
							$NewResourceAccountLocation = "US"
						}

						Write-Host "`tAssigning Location"
						Update-MgUser -UserId $NewResourceAccountPrincipalName -Id $applicationInstanceID -UsageLocation $NewResourceAccountLocation

						if ( ! $NoResourceAccountLicensing )
						{
							Write-Host "`tAttempting to license Resource Account"
							$skuID = (Get-MgSubscribedSKU | Where {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}).SkuId
							Set-MgUserLicense -UserId $applicationInstanceID -AddLicenses @{SkuId = $skuID} -RemoveLicenses @() | Out-Null
						}
						else
						{
							Write-Host "`tResource Account Licensing is disabled"
						}

						Write-Host "`tAssigning Resource Account"
						if ( $NewResourceAccountPriority -ne "" )
						{
							# New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue -CallPriority $NewResourceAccountPriority| Out-Null
							New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $CallQueue.Identity -ConfigurationType CallQueue -CallPriority $NewResourceAccountPriority| Out-Null
						}
						else
						{
							# New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $callQueueID -ConfigurationType CallQueue | Out-Null
							New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $CallQueue.Identity -ConfigurationType CallQueue | Out-Null
						}
					}
					else
					{
						Write-Host "`tResourceAccountPrincipalName is blank"
					}
				}
				else
				{
					Write-Host "`tResource Account Creation is disabled"
				}
			}
		}
		else
		{
			Write-Host "`tUnable to process Resource Account configuration for $Name Call Queue as more than one Call Queue with that name exists."
		} 		
	}
	else
	{
		Write-Host "`tResource Account Processing is disabled"
	}

	Write-Host "`tCall Queue Created"
}










#
# Main 
#
# processing arguments
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
		"-help"             			{ $Help = $true }
		"-excelfile"        			{ 
											$ExcelFilename = $args[$i+1]
											$i++
										}
		"-nocallback"       	 		{ $NoCallback = $true }
		"-noresourceaccounts" 			{ $NoResourceAccounts = $true }
	   
		"-noresourceaccountcreation"	{ $NoResourceAccountCreation = $true }
		"-noresourceaccountlicensing" 	{ $NoResourceAccountLicensing = $true }


       "-verbose"            			{ $Verbose = $true }
       Default               			{ Write-Host "Unknown argument passed: $args[$i]" }
   }
}

if ( $Help )
{
   Write-Host "The following options are avaialble:"
   Write-Host "`t-Help - shows the options that are available (this help message)"
   Write-Host "`t-NoCallback - will not configure the callback component - if your tenant is not flighted for callback then enable this option"
   Write-Host "`t-NoResourceAccount - will not configure the resource account to the call queue"
   Write-Host "`t-Verbose - provides extra messaging during the process"
   exit
}

Write-Host "Starting BulkCQsConfig."
Write-Host "Cleaning up from any previous runs."

if ( Test-Path -Path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
}


#
# Check that required modules are installed - install if not
#
Write-Host "Checking for MicrosoftTeams module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "MicrosoftTeams" } )
{
   Write-Host "Connecting to Microsoft Teams."
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}
else
{
   Write-Host "Module MicrosoftTeams does not exist - installing."
   Install-Module -Name MicrosoftTeams -Force -AllowClobber

   Write-Host "Connecting to Microsoft Teams."
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MicrosoftTeams | Out-Null
   }
   try
   { 
      Get-CsTenant -ErrorAction Stop 2>&1> $null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}

Write-Host "Checking for Microsoft.Graph module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "Microsoft.Graph" } )
{
   Write-Host "Connecting to Microsoft Graph."
   Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null

   try
   { 
      # Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null

   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
   }
   try
   { 
      # Get-MgSubscribedSKU -ErrorAction Stop 2>&1> $null
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}
else
{
   Write-Host "Module MgGraph does not exist - installing."
   Install-Module -Name Microsoft.Graph -Force -AllowClobber
   Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome | Out-Null
   try
   { 
      Get-MgSubscribedSKU -ErrorAction Stop | Out-Null
   } 
   catch [System.UnauthorizedAccessException] 
   { 
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}

Write-Host "Checking for ImportExcel module."
if ( Get-InstalledModule | Where-Object { $_.Name -eq "ImportExcel" } )
{
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel
}
else
{
   Write-Host "Module ImportExcel - installing."
   Install-Module -Name ImportExcel -Force -AllowClobber
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel
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
	Write-Host "ERROR: $ExcelFilename does not exist."
	exit
}

#
# Call Queue configuration
#
Write-Host "Starting Call Queue Configuration."
$ExcelWorkSheetName = "PS-CQ"
$ExcelCSVFilename = ".\PS-CQ.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName

if ( $data.length -lt 1 )
{
	Write-Host "ERROR: No data retrieved from $ExcelWorkSheetName"
	exit
}
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCSVFilename -NoTypeInformation


$PSCQConfig = @(Import-csv -Path $ExcelCSVFilename)

for ($i=0; $i -lt  $PSCQConfig.length; $i++)
{
   $Action = $PSCQConfig.Action[$i]
   if ( $Action -eq "New" )
   {
      $Name   = '"' + $PSCQConfig.Name[$i] + '"'
      $ExistingResourceAccountName = $PSCQConfig.ExistingResourceAccountName[$i]
      $NewResourceAccountPrincipalName = $PSCQConfig.NewResourceAccountPrincipalName[$i]
      $NewResourceAccountDisplayName = $PSCQConfig.NewResourceAccountDisplayName[$i]
      $NewResourceAccountLocation = $PSCQConfig.NewResourceAccountLocation[$i]
	  $NewResourceAccountPriority = $PSCQConfig.NewResourceAccountPriority[$i]

      $OutboundCLID = @()
      if ( $PSCQConfig.OutboundCLID01[$i] -ne "" )
      {
         $OutboundCLID += '"' + $PSCQConfig.OutboundCLID01[$i] + '"'
      }

      if ( $PSCQConfig.OutboundCLID02[$i] -ne "" )
      {
         $OutboundCLID += '"' + $PSCQConfig.OutboundCLID02[$i] + '"'
      }

      if ( $PSCQConfig.OutboundCLID03[$i] -ne "" )
      {
         $OutboundCLID += '"' + $PSCQConfig.OutboundCLID03[$i] + '"'
      }

      if ( $PSCQConfig.OutboundCLID04[$i] -ne "" )
      {
         $OutboundCLID += '"' + $PSCQConfig.OutboundCLID04[$i] + '"'
      }
      $OboResourceAccountIDs = $OutboundCLID -join ","

      $Language = $PSCQConfig.Language[$i]
      $ServiceLevelThreshold = $PSCQConfig.ServiceLevelThreshold[$i]

      $GreetingOption = $PSCQConfig.GreetingOption[$i]
      switch ( $GreetingOption )
      {
         "FILE"  { $Greeting = $PSCQConfig.Greeting[$i] }
         "TEXT"  { $Greeting = '"' + $PSCQConfig.Greeting[$i] + '"' }
         Default { $Greeting= "" }
      }

      $MusicOnHoldOption = $PSCQConfig.MusicOnHoldOption[$i]
      $MusicOnHOld = $PSCQConfig.MusicOnHold[$i]
      $RoutingMethod = $PSCQConfig.RoutingMethod[$i]
      $PresenceBasedRouting = $PSCQConfig.PresenceBasedRouting[$i]
      $AllowOptOut = $PSCQConfig.AllowOptOut[$i]
      $AgentAlertTime = $PSCQConfig.AgentAlertTime[$i]

      $OverflowThreshold = $PSCQConfig.OverflowThreshold[$i]
      $OverflowAction = $PSCQConfig.OverflowAction[$i]
      $OverflowActionTarget = $PSCQConfig.OverflowActionTarget[$i]
      $OverflowActionCallPriority = $PSCQConfig.OverflowActionCallPriority[$i]

      $OverflowTreatment = $PSCQConfig.OverflowTreatment[$i]
      switch ( $OverflowTreatment )
      {
         "FILE"  { $OverflowTreatmentPrompt = $PSCQConfig.OverflowTreatmentPrompt[$i] }
         "TEXT"  { $OverflowTreatmentPrompt = '"' + $PSCQConfig.OverflowTreatmentPrompt[$i] + '"' }
         Default { $OverflowTreatmentPrompt = "" }
      }

      $OverflowSharedVoicemailSystemPromptSuppression = $PSCQConfig.OverflowSharedVoicemailSystemPromptSuppression[$i]
      $OverflowSharedVoicemailTranscription = $PSCQConfig.OverflowSharedVoicemailTranscription[$i]

      $TimeoutThreshold = $PSCQConfig.TimeoutThreshold[$i]
      $TimeoutAction = $PSCQConfig.TimeoutAction[$i]
      $TimeoutActionTarget = $PSCQConfig.TimeoutActionTarget[$i]
      $TimeoutActionCallPriority = $PSCQConfig.TimeoutActionCallPriority[$i]

      $TimeoutTreatment = $PSCQConfig.TimeoutTreatment[$i]
      switch ( $TimeoutTreatment )
      {
         "FILE"  { $TimeoutTreatmentPrompt = $PSCQConfig.TimeoutTreatmentPrompt[$i] }
         "TEXT"  { $TimeoutTreatmentPrompt = '"' + $PSCQConfig.TimeoutTreatmentPrompt[$i] + '"' }
         Default { $TimeoutTreatmentPrompt = "" }
      }

      $TimeoutSharedVoicemailSystemPromptSuppression = $PSCQConfig.TimeoutSharedVoicemailSystemPromptSuppression[$i]
      $TimeoutSharedVoicemailTranscription = $PSCQConfig.TimeoutSharedVoicemailTranscription[$i]

      $NoAgentNewCallsOnly = $PSCQConfig.NoAgentNewCallsOnly[$i]
      $NoAgentAction = $PSCQConfig.NoAgentAction[$i]
      $NoAgentActionTarget = $PSCQConfig.NoAgentActionTarget[$i]
      $NoAgentActionCallPriority = $PSCQConfig.NoAgentActionCallPriority[$i]

      $NoAgentTreatment = $PSCQConfig.NoAgentTreatment[$i]
      switch ( $NoAgentTreatment )
      {
         "FILE"  { $NoAgentTreatmentPrompt = $PSCQConfig.NoAgentTreatmentPrompt[$i] }
         "TEXT"  { $NoAgentTreatmentPrompt = '"' + $PSCQConfig.NoAgentTreatmentPrompt[$i] + '"' }
         Default { $NoAgentTreatmentPrompt = "" }
      }

      $NoAgentSharedVoicemailSystemPromptSuppression = $PSCQConfig.NoAgentSharedVoicemailSystemPromptSuppression[$i]
      $NoAgentSharedVoicemailTranscription = $PSCQConfig.NoAgentSharedVoicemailTranscription[$i]

      $IsCallbackEnabled = $PSCQConfig.IsCallbackEnabled[$i]
      $CallbackRequestDTMF = $PSCQConfig.CallbackRequestDTMF[$i]
      $WaitTimeBeforeOfferingCallbackInSecond = $PSCQConfig.WaitTimeBeforeOfferingCallbackInSecond[$i]
      $NumberOfCallsInQueueBeforeOfferingCallback = $PSCQConfig.NumberOfCallsInQueueBeforeOfferingCallback[$i]
      $CallToAgentRatioThresholdBeforeOfferingCallback = $PSCQConfig.CallToAgentRatioThresholdBeforeOfferingCallback[$i]

      $CallbackOfferTreatment = $PSCQConfig.CallbackOfferTreatment[$i]
      switch ( $CallbackOfferTreatment )
      {
         "FILE"  { $CallbackOfferPrompt = $PSCQConfig.CallbackOfferPrompt[$i] }
         "TEXT"  { $CallbackOfferPrompt = '"' + $PSCQConfig.CallbackOfferPrompt[$i] + '"' }
         Default { $CallbackOfferTreatment = "" }
      }

      if ( $PSCQConfig.CallbackEmailNotificationTarget[$i] -ne "" )
      {
         $CallbackEmailNotificationTarget = [System.GUID]::Parse($PSCQConfig.CallbackEmailNotificationTarget[$i])
      }

      $TeamGroupID = $PSCQConfig.TeamGroupID[$i]
      $TeamChannelID = $PSCQConfig.TeamChannelID[$i]
      $TeamChannelName = $PSCQConfig.TeamChannelName[$i]

      $DL = @()
      if ( $PSCQConfig.DistributionList01[$i] -ne "" )
      {
         $DL += '"' + $PSCQConfig.DistributionList01[$i] + '"'
      }

      if ( $PSCQConfig.DistributionList02[$i] -ne "" )
      {
         $DL += '"' + $PSCQConfig.DistributionList02[$i] + '"'
      }

      if ( $PSCQConfig.DistributionList03[$i] -ne "" )
      {
         $DL += '"' + $PSCQConfig.DistributionList03[$i] + '"'
      }

      if ( $PSCQConfig.DistributionList04[$i] -ne "" )
      {
         $DL += '"' + $PSCQConfig.DistributionList04[$i] + '"'
      }
      $DistributionLists = $DL -join ","


      $Agents = @()
      if ( $PSCQConfig.Agent01[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent01[$i] + '"'
      }

      if ( $PSCQConfig.Agent02[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent02[$i] + '"'
      }

      if ( $PSCQConfig.Agent03[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent03[$i] + '"'
      }

      if ( $PSCQConfig.Agent04[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent04[$i] + '"'
      }

      if ( $PSCQConfig.Agent05[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent05[$i] + '"'
      }

      if ( $PSCQConfig.Agent06[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent06[$i] + '"'
      }

      if ( $PSCQConfig.Agent07[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent07[$i] + '"'
      }

      if ( $PSCQConfig.Agent08[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent08[$i] + '"'
      }

      if ( $PSCQConfig.Agent09[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent09[$i] + '"'
      }

      if ( $PSCQConfig.Agent10[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent10[$i] + '"'
      }

      if ( $PSCQConfig.Agent11[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent11[$i] + '"'
      }

      if ( $PSCQConfig.Agent12[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent12[$i] + '"'
      }

      if ( $PSCQConfig.Agent13[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent13[$i] + '"'
      }

      if ( $PSCQConfig.Agent14[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent14[$i] + '"'
      }

      if ( $PSCQConfig.Agent15[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent15[$i] + '"'
      }

      if ( $PSCQConfig.Agent16[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent16[$i] + '"'
      }

      if ( $PSCQConfig.Agent17[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent17[$i] + '"'
      }

      if ( $PSCQConfig.Agent18[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent18[$i] + '"'
      }

      if ( $PSCQConfig.Agent19[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent19[$i] + '"'
      }

      if ( $PSCQConfig.Agent20[$i] -ne "" )
      {
         $Agents += '"' + $PSCQConfig.Agent20[$i] + '"'
      }
      $Users = $Agents -join ","

      NewCallQueue
   }
}

Write-Host "Completed Call Queue Configuration."

if ( Test-Path -Path ".\PS-CQ.csv" )
{
   Remove-Item -Path ".\PS-CQ.csv" | Out-Null
}

exit



