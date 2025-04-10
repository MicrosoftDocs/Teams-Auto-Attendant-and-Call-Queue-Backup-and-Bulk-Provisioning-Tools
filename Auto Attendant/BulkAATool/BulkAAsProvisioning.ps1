# Version: 1.0.2
# Date: 2025.04.10
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
#  AudioFileImport
#
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
    $audioFileID = Import-CsOnlineAudioFile -ApplicationID OrgAutoAttendant -FileName $fileName -Content $content

    return $audioFileID
}

#
# CheckFileExists
#
function CheckFileExists
{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $fileName
    )
	
	$currentDIR = (Get-Location).Path
    $currentDIR += "\AudioFiles\$fileName"

	if ( Test-Path -Path $currentDIR )
	{
		return $true
	}
	else
	{
		return $false
	}
}

#
# VerboseOutputMenuOption
#
function VerboseOutputMenuOption ([string]$MenuType, [string]$Key, [string]$Redirect, [string]$VoiceCommand, [string]$RedirectTarget, [string]$RedirectPriority, [string]$RedirectComment, [string]$SharedVoicemailTranscription, [string]$SharedVoicemailSuppress, [string]$CallPriority)
{	
	if ( $Key -eq "0" )
	{
		Write-Host ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}`t{4,-30}" -f "Key", "$($MenuType)_x_VoiceCommand",     "$($MenuType)_x_Redirect",                   "$($MenuType)_x_RedirectTarget",          "Additional Information")
		Write-Host ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}`t{4,-30}" -f "---", "--------------------", "---------------------", "------------------------------------","----------------------")
	}

	switch ( $Redirect )
	{
		"Operator"					{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}
										
		"User"						{				
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}

		"ApplicationEndpoint"		{
										if ( $CallPriority )
										{
											Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}`t{4,-11}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget, "Priority: $RedirectPriority")
										}
										else
										{
											Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
										}
									}

		"ConfigurationEndpoint"		{
										if ( $CallPriority )
										{
											Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}`t{4,-30}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget, "Priority: $RedirectPriority")
										}
										else
										{
											Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
										}
									}

		"SharedVoicemail"			{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}`t{4,-30}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget, "Transcript/Suppress: $SharedVoicemailTranscription / $SharedVoicemailSuppress")
									}

		"ExternalPSTN"				{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}

		"FILE"						{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}

		"TEXT"						{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}

		Default						{
										Write-Host -NoNewLine ("`t`t{0,-3}`t{1,-20}`t{2,-20}`t{3,-36}" -f " $Key ", $VoiceCommand, $Redirect, $RedirectTarget)
									}										
	}
	Write-Host -ForegroundColor red ("`t{0,-30}" -f $RedirectComment)
}



#
#  VerboseOutput
#
function VerboseOutput
{
	Write-Host "-----------------------------------------------------------------"
    Write-Host "Action:  $Action`tName: $Name"
    Write-Host "-----------------------------------------------------------------"

    Write-Host "`tResourceAccount : $ResourceAccount"
    Write-Host "`tExistingResourceAccoutName : $ExistingResourceAccountName"
    Write-Host "`tNewResourceAccountPrincipalName : $NewResourceAccountPrincipalName"
    Write-Host "`tNewResourceAccountDisplayName : $NewResourceAccountDisplayName"
    Write-Host "`tNewResourceAccountLocation : $NewResourceAccountLocation"
	
	if ( $NewResourceAccountPhoneNumber_Comment -eq $null )
	{
		Write-Host "`tNewResourceAccountPhoneNumber : $NewResourceAccountPhoneNumber"
	}
	else
	{
		Write-Host -NoNewLine "`tNewResourceAccountPhoneNumber : $NewResourceAccountPhoneNumber <---"
		Write-Host -ForegroundColor red $NewResourceAccountPhoneNumber_Comment
	}
	
    Write-Host "`tOperator : $Operator"
	
	if ( $OperatorTarget -eq "ERROR" )
	{
		Write-Host -NoNewLine  "`tOperatorTarget: "
		Write-Host -ForegroundColor red $Operator_Comment
	}
	else
	{
		if ( $CallPriority )
		{
			Write-Host "`tOperatorTarget: $OperatorTarget / Priority: $Operator_Call_Priority"
		}
		else
		{
			Write-Host "`tOperatorTarget: $OperatorTarget"
		}
	}
	
    Write-Host "`tTimeZone : $TimeZone"
	
	if ( $LanguageGender )
	{
		Write-Host "`tLanguage / Voice : $Language / $LanguageGenderId"
	}
	else
	{
		Write-Host "`tLanguage : $Language"	
	}
	
    Write-Host "`tVoiceInputs : $VoiceInputs"
    Write-Host "`tHours24 : $Hours24"

    Write-Host "`tBusiness Hours (Default) Menu"
	    Write-Host "`t`tBusinessHours : $BusinessHours"
        Write-Host "`t`tB_DirectorySearch : $B_DirectorySearch"
		
		if ( $B_DialScopeInclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tB_DialScopeInclude : "
			Write-Host -ForegroundColor red "$B_DialScopeInclude_Comment"
		}
		else
		{
			Write-Host "`t`tB_DialScopeInclude : $B_DialScopeInclude"
		}
		
		if ( $B_DialScopeExclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tB_DialScopeExclude : "
			Write-Host -ForegroundColor red "$B_DialScopeExclude_Comment"
		}
		else
		{
			Write-Host "`t`tB_DialScopeExclude : $B_DialScopeExclude"
		}	
		
        Write-Host "`t`tB_Force : $B_Force"
        Write-Host "`t`tB_GreetingOption : $B_GreetingOption"
        Write-Host -NoNewLine "`t`tB_Greeting : $B_Greeting"
		Write-Host -ForegroundColor red "`t$B_Greeting_Comment"

		switch ( $B_GreetingRouting )
		{
			"ApplicationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting`t/ Priority: $B_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting"
											}											
										}
			"ConfigurationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting`t/ Priority: $B_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting"
											}											

										}
			"Disconnect"				{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting"				
										}
			"ExternalPSTN"				{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting"				
										}
			"User"						{
												Write-Host "`t`tB_GreetingRouting : $B_GreetingRouting"				
										}
		}
				
		if ( $B_GreetingRoutingTarget -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tB_GreetingRoutingTarget: "
			Write-Host -ForegroundColor red $B_GreetingRouting_Comment
		}
		else
		{
			Write-Host "`t`tB_GreetingRoutingTarget: $B_GreetingRoutingTarget"
		}

        Write-Host "`t`tB_MenuGreetingOption : $B_MenuGreetingOption"
        Write-Host "`t`tB_MenuGreeting : $B_MenuGreeting"

		VerboseOutputMenuOption "B" "0" $B_0_Redirect $B_0_VoiceCommand $B_0_RedirectTarget $B_0_Redirect_Call_Priority $B_0_Redirect_Comment $B_0_SharedVoicemailTranscription $B_0_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "1" $B_1_Redirect $B_1_VoiceCommand $B_1_RedirectTarget $B_1_Redirect_Call_Priority $B_1_Redirect_Comment $B_1_SharedVoicemailTranscription $B_1_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "2" $B_2_Redirect $B_2_VoiceCommand $B_2_RedirectTarget $B_2_Redirect_Call_Priority $B_2_Redirect_Comment $B_2_SharedVoicemailTranscription $B_2_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "3" $B_3_Redirect $B_3_VoiceCommand $B_3_RedirectTarget $B_3_Redirect_Call_Priority $B_3_Redirect_Comment $B_3_SharedVoicemailTranscription $B_3_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "4" $B_4_Redirect $B_4_VoiceCommand $B_4_RedirectTarget $B_4_Redirect_Call_Priority $B_4_Redirect_Comment $B_4_SharedVoicemailTranscription $B_4_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "5" $B_5_Redirect $B_5_VoiceCommand $B_5_RedirectTarget $B_5_Redirect_Call_Priority $B_5_Redirect_Comment $B_5_SharedVoicemailTranscription $B_5_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "6" $B_6_Redirect $B_6_VoiceCommand $B_6_RedirectTarget $B_6_Redirect_Call_Priority $B_6_Redirect_Comment $B_6_SharedVoicemailTranscription $B_6_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "7" $B_7_Redirect $B_7_VoiceCommand $B_7_RedirectTarget $B_7_Redirect_Call_Priority $B_7_Redirect_Comment $B_7_SharedVoicemailTranscription $B_7_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "8" $B_8_Redirect $B_8_VoiceCommand $B_8_RedirectTarget $B_8_Redirect_Call_Priority $B_8_Redirect_Comment $B_8_SharedVoicemailTranscription $B_8_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "9" $B_9_Redirect $B_9_VoiceCommand $B_9_RedirectTarget $B_9_Redirect_Call_Priority $B_9_Redirect_Comment $B_9_SharedVoicemailTranscription $B_9_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "*" $B_Star_Redirect $B_Star_VoiceCommand $B_Star_RedirectTarget $B_Star_Redirect_Call_Priority $B_Star_Redirect_Comment $B_Star_SharedVoicemailTranscription $B_Star_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "B" "#" $B_Pound_Redirect $B_Pound_VoiceCommand $B_Pound_RedirectTarget $B_Pound_Redirect_Call_Priority $B_Pound_Redirect_Comment $B_Pound_SharedVoicemailTranscription $B_Pound_SharedVoicemailSuppress $CallPriority
		

    Write-Host "`tAfter Hours Menu"
        Write-Host "`t`tA_DirectorySearch : $A_DirectorySearch"
		
		if ( $A_DialScopeInclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tA_DialScopeInclude : "
			Write-Host -ForegroundColor red "$A_DialScopeInclude_Comment"
		}
		else
		{
			Write-Host "`t`tA_DialScopeInclude : $A_DialScopeInclude"
		}
		
		if ( $A_DialScopeExclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tA_DialScopeExclude : "
			Write-Host -ForegroundColor red "$A_DialScopeExclude_Comment"
		}
		else
		{
			Write-Host "`t`tA_DialScopeExclude : $A_DialScopeExclude"
		}	
		
        Write-Host "`t`tA_Force : $A_Force"
        Write-Host "`t`tA_GreetingOption : $A_GreetingOption"
		
		Write-Host -NoNewLine "`t`tA_Greeting : $A_Greeting"
		Write-Host -ForegroundColor red "`t$A_Greeting_Comment"

		switch ( $A_GreetingRouting )
		{
			"ApplicationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting`t/ Priority: $A_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting"
											}											
										}
			"ConfigurationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting`t/ Priority: $A_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting"
											}											

										}
			"Disconnect"				{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting"				
										}
			"ExternalPSTN"				{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting"				
										}
			"User"						{
												Write-Host "`t`tA_GreetingRouting : $A_GreetingRouting"				
										}
		}
		
		if ( $A_GreetingRoutingTarget -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tA_GreetingRoutingTarget: "
			Write-Host -ForegroundColor red $A_GreetingRouting_Comment
		}
		else
		{
			Write-Host "`t`tA_GreetingRoutingTarget: $A_GreetingRoutingTarget"
		}

        Write-Host "`t`tA_MenuGreetingOption : $A_MenuGreetingOption"
        Write-Host "`t`tA_MenuGreeting : $A_MenuGreeting"

		VerboseOutputMenuOption "A" "0" $A_0_Redirect $A_0_VoiceCommand $A_0_RedirectTarget $A_0_Redirect_Call_Priority $A_0_Redirect_Comment $A_0_SharedVoicemailTranscription $A_0_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "1" $A_1_Redirect $A_1_VoiceCommand $A_1_RedirectTarget $A_1_Redirect_Call_Priority $A_1_Redirect_Comment $A_1_SharedVoicemailTranscription $A_1_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "2" $A_2_Redirect $A_2_VoiceCommand $A_2_RedirectTarget $A_2_Redirect_Call_Priority $A_2_Redirect_Comment $A_2_SharedVoicemailTranscription $A_2_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "3" $A_3_Redirect $A_3_VoiceCommand $A_3_RedirectTarget $A_3_Redirect_Call_Priority $A_3_Redirect_Comment $A_3_SharedVoicemailTranscription $A_3_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "4" $A_4_Redirect $A_4_VoiceCommand $A_4_RedirectTarget $A_4_Redirect_Call_Priority $A_4_Redirect_Comment $A_4_SharedVoicemailTranscription $A_4_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "5" $A_5_Redirect $A_5_VoiceCommand $A_5_RedirectTarget $A_5_Redirect_Call_Priority $A_5_Redirect_Comment $A_5_SharedVoicemailTranscription $A_5_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "6" $A_6_Redirect $A_6_VoiceCommand $A_6_RedirectTarget $A_6_Redirect_Call_Priority $A_6_Redirect_Comment $A_6_SharedVoicemailTranscription $A_6_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "7" $A_7_Redirect $A_7_VoiceCommand $A_7_RedirectTarget $A_7_Redirect_Call_Priority $A_7_Redirect_Comment $A_7_SharedVoicemailTranscription $A_7_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "8" $A_8_Redirect $A_8_VoiceCommand $A_8_RedirectTarget $A_8_Redirect_Call_Priority $A_8_Redirect_Comment $A_8_SharedVoicemailTranscription $A_8_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "9" $A_9_Redirect $A_9_VoiceCommand $A_9_RedirectTarget $A_9_Redirect_Call_Priority $A_9_Redirect_Comment $A_9_SharedVoicemailTranscription $A_9_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "*" $A_Star_Redirect $A_Star_VoiceCommand $A_Star_RedirectTarget $A_Star_Redirect_Call_Priority $A_Star_Redirect_Comment $A_Star_SharedVoicemailTranscription $A_Star_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "A" "#" $A_Pound_Redirect $A_Pound_VoiceCommand $A_Pound_RedirectTarget $A_Pound_Redirect_Call_Priority $A_Pound_Redirect_Comment $A_Pound_SharedVoicemailTranscription $A_Pound_SharedVoicemailSuppress $CallPriority
	  
	  
		Write-Host "`tHolidays Menu"
	    Write-Host "`t`tHolidays : $Holidays"
        Write-Host "`t`tH_DirectorySearch : $H_DirectorySearch"

		if ( $H_DialScopeInclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tH_DialScopeInclude : "
			Write-Host -ForegroundColor red "$H_DialScopeInclude_Comment"
		}
		else
		{
			Write-Host "`t`tH_DialScopeInclude : $H_DialScopeInclude"
		}
		
		if ( $H_DialScopeExclude -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tH_DialScopeExclude : "
			Write-Host -ForegroundColor red "$H_DialScopeExclude_Comment"
		}
		else
		{
			Write-Host "`t`tH_DialScopeExclude : $H_DialScopeExclude"
		}	

        Write-Host "`t`tH_Force : $H_Force"
        Write-Host "`t`tH_GreetingOption : $H_GreetingOption"
		
		Write-Host -NoNewLine "`t`tH_Greeting : $H_Greeting"
		Write-Host -ForegroundColor red "`t$H_Greeting_Comment"

		switch ( $H_GreetingRouting )
		{
			"ApplicationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting`t/ Priority: $H_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting"
											}											
										}
			"ConfigurationEndpoint"		{
											if ( $CallPriority )
											{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting`t/ Priority: $H_GreetingRouting_Call_Priority"
											}
											else
											{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting"
											}											

										}
			"Disconnect"				{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting"				
										}
			"ExternalPSTN"				{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting"				
										}
			"User"						{
												Write-Host "`t`tH_GreetingRouting : $H_GreetingRouting"				
										}
		}

		if ( $H_GreetingRoutingTarget -eq "ERROR" )
		{
			Write-Host -NoNewLine "`t`tH_GreetingRoutingTarget: "
			Write-Host -ForegroundColor red $H_GreetingRouting_Comment
		}
		else
		{
			Write-Host "`t`tH_GreetingRoutingTarget: $H_GreetingRoutingTarget"
		}

        Write-Host "`t`tH_MenuGreetingOption : $H_MenuGreetingOption"
        Write-Host "`t`tH_MenuGreeting : $H_MenuGreeting"

		VerboseOutputMenuOption "H" "0" $H_0_Redirect $H_0_VoiceCommand $H_0_RedirectTarget $H_0_Redirect_Call_Priority $H_0_Redirect_Comment $H_0_SharedVoicemailTranscription $H_0_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "1" $H_1_Redirect $H_1_VoiceCommand $H_1_RedirectTarget $H_1_Redirect_Call_Priority $H_1_Redirect_Comment $H_1_SharedVoicemailTranscription $H_1_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "2" $H_2_Redirect $H_2_VoiceCommand $H_2_RedirectTarget $H_2_Redirect_Call_Priority $H_2_Redirect_Comment $H_2_SharedVoicemailTranscription $H_2_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "3" $H_3_Redirect $H_3_VoiceCommand $H_3_RedirectTarget $H_3_Redirect_Call_Priority $H_3_Redirect_Comment $H_3_SharedVoicemailTranscription $H_3_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "4" $H_4_Redirect $H_4_VoiceCommand $H_4_RedirectTarget $H_4_Redirect_Call_Priority $H_4_Redirect_Comment $H_4_SharedVoicemailTranscription $H_4_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "5" $H_5_Redirect $H_5_VoiceCommand $H_5_RedirectTarget $H_5_Redirect_Call_Priority $H_5_Redirect_Comment $H_5_SharedVoicemailTranscription $H_5_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "6" $H_6_Redirect $H_6_VoiceCommand $H_6_RedirectTarget $H_6_Redirect_Call_Priority $H_6_Redirect_Comment $H_6_SharedVoicemailTranscription $H_6_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "7" $H_7_Redirect $H_7_VoiceCommand $H_7_RedirectTarget $H_7_Redirect_Call_Priority $H_7_Redirect_Comment $H_7_SharedVoicemailTranscription $H_7_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "8" $H_8_Redirect $H_8_VoiceCommand $H_8_RedirectTarget $H_8_Redirect_Call_Priority $H_8_Redirect_Comment $H_8_SharedVoicemailTranscription $H_8_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "9" $H_9_Redirect $H_9_VoiceCommand $H_9_RedirectTarget $H_9_Redirect_Call_Priority $H_9_Redirect_Comment $H_9_SharedVoicemailTranscription $H_9_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "*" $H_Star_Redirect $H_Star_VoiceCommand $H_Star_RedirectTarget $H_Star_Redirect_Call_Priority $H_Star_Redirect_Comment $H_Star_SharedVoicemailTranscription $H_Star_SharedVoicemailSuppress $CallPriority
		VerboseOutputMenuOption "H" "#" $H_Pound_Redirect $H_Pound_VoiceCommand $H_Pound_RedirectTarget $H_Pound_Redirect_Call_Priority $H_Pound_Redirect_Comment $H_Pound_SharedVoicemailTranscription $H_Pound_SharedVoicemailSuppress $CallPriority

    Write-Host "-----------------------------------------------------------------"
}




#
#  NewAutoAttendant
#

function NewAutoAttendant
{
    if ( $Verbose -OR $VerboseStopProcessing )
    {
        VerboseOutput
    }
    else
    {
        Write-Host "Action:  $Action`tName: $Name"
    }
	
	if ( $StopProcessing )
	{
		Write-Host -ForegroundColor red "`tProcessing for $Name can't continue. See above for errors."
		return
	}
		
	if ( $LanguageGender )
	{
		$command = "New-CsAutoAttendant -Name `$Name -LanguageId `$Language -VoiceId `$LanguageGenderId -TimeZoneId `$TimeZone "
	}
	else
	{
		$command = "New-CsAutoAttendant -Name `$Name -LanguageId `$Language -TimeZoneId `$TimeZone "
	}

    if ( $B_DialScopeInclude -ne "" -AND $B_DialScopeExclude -ne "" )
	{
		$B_DialScopeIncludeGroup = New-CsAutoAttendantDialScope -GroupScope -GroupIds ("$B_DialScopeInclude")
		$B_DialScopeExcludeGroup = New-CsAutoAttendantDialScope -GroupScope -GroupIds ("$B_DialScopeExclude")
		$command += "-InclusionScope `$B_DialScopeIncludeGroup -ExclusionScope `$B_DialScopeExcludeGroup "
	}
	elseif ( $B_DialScopeInclude -ne "" )
	{
		$B_DialScopeIncludeGroup = New-CsAutoAttendantDialScope -GroupScope -GroupIds ("$B_DialScopeInclude")
		$command += "-InclusionScope `$B_DialScopeIncludeGroup "
	}
	elseif ( $B_DialScopeExclude -ne "" )
	{
		$B_DialScopeExcludeGroup = New-CsAutoAttendantDialScope -GroupScope -GroupIds ("$B_DialScopeExclude")
		$command += "-ExclusionScope `$B_DialScopeExcludeGroup "	
	}
	else
	{
		# Both are blank - currently a don't care condition
	}


    #
    #  Operator
    #  Default: None
    #
    if ( $Operator -ne "None" )
    {
		if ( ( $Operator -eq "ApplicationEndpoint" -or $Operator -eq "ConfigurationEndpoint" ) -AND $CallPriority )
		{
			$OperatorCallableEntity = New-CsAutoAttendantCallableEntity -Identity $OperatorTarget -Type $Operator -CallPriority $Operator_Call_Priority
		}
		else
		{
			$OperatorCallableEntity = New-CsAutoAttendantCallableEntity -Identity $OperatorTarget -Type $Operator
		}

        $command += "-Operator `$OperatorCallableEntity "
    }


    #
    #  Voice Inputs
    #  Default: Off
    #

    if ( $VoiceInputs -eq "On" )
    {
        $command += "-EnableVoiceResponse "
    }


    #
    #  Default Call Flow (Business Hours Call Flow)
    #
	Write-Host "`tBuilding Default Call Flow"
    switch ( $B_GreetingOption )
    {
        "FILE"  { 
                    $audioFile = AudioFileImport $B_Greeting
                    $B_GreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                    $B_GreetingPromptConfigured = $true
                }
        "TEXT"  {
                    $B_GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_Greeting
                    $B_GreetingPromptConfigured = $true
                } 
        Default {   $B_GreetingPromptConfigured = $false }
    }


    if ( $B_GreetingRouting -eq "Disconnect" )
    {
        $defaultMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
        $defaultMenu = New-CsAutoAttendantMenu -Name "Default Menu" -MenuOptions @($defaultMenuOption)

        if ( $B_GreetingPromptConfigured )
        {
            $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Greetings @($B_GreetingPrompt) -Menu @($defaultMenu)
        }
        else
        {
            $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Menu $defaultMenu
        }
    }
    elseif ( $B_GreetingRouting -eq "User" -OR $B_GreetingRouting -eq "ApplicationEndpoint" -OR $B_GreetingRouting -eq "ConfigurationEndpoint" -OR $B_GreetingRouting -eq "ExternalPSTN" )
    {
		
		if ( ( $B_GreetingRouting -eq "ApplicationEndpoint" -OR $B_GreetingRouting -eq "ConfigurationEndpoint" ) -AND $CallPriority )
        {
			$defaultMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $B_GreetingRoutingTarget -Type $B_GreetingRouting -CallPriority $B_GreetingRouting_Call_Priority
        }
        else
        {
			$defaultMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $B_GreetingRoutingTarget -Type $B_GreetingRouting
		}
		
        $defaultMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $defaultMenuOptionCallableEntity -DtmfResponse Automatic
        $defaultMenu = New-CsAutoAttendantMenu -Name "Default Menu" -MenuOptions @($defaultMenuOption)

        if ( $B_GreetingPromptConfigured )
        {
            $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Greetings @($B_GreetingPrompt) -Menu $defaultMenu
        }
        else
        {
            $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Menu $defaultMenu
        }
    }
    elseif ( $B_GreetingRouting -eq "MENU" )
    {
        switch ( $B_MenuGreetingOption )
        {
            "FILE"  {
                        $audioFile = AudioFileImport $B_MenuGreeting
                        $B_MenuGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                        $B_GreetingPromptConfigured = $true
                    }
            "TEXT"  {
                        $B_MenuGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_MenuGreeting
                        $B_GreetingPromptConfigured = $true
                    } 
            Default {   $B_MenuGreetingPromptConfigured = $false }
        }

        #
        # Option 0
        #
        if ( $B_0_Redirect -ne "NONE" )
        {
            if ( $B_0_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0
            }
            elseif ( $B_0_Redirect -eq "User" -OR $B_0_Redirect -eq "ApplicationEndpoint" -OR $B_0_Redirect -eq "ConfigurationEndpoint" -OR $B_0_Redirect -eq "SharedVoicemail" -OR $B_0_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_0_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_0_SharedVoicemailTranscription -AND $B_0_SharedVoicemailSuppress )
                    {
                        $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_0_SharedVoicemailTranscription )
                    {
                        $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect -EnableTranscription
                    }
                    elseif ( $B_0_SharedVoicemailSuppress )
                    {
                        $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    } 
                    else
                    { 
                         $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect
                    }
                }
                elseif ( ( $B_0_Redirect -eq "ApplicationEndpoint" -OR $B_0_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect -CallPriority $B_0_Redirect_Call_Priority
                }
                else
                {
                    $B_0_Entity = New-CsAutoAttendantCallableEntity -Identity $B_0_RedirectTarget -Type $B_0_Redirect
				}

                if ( $B_0_VoiceCommand -eq "" )
                {
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $B_0_Entity
                }
                else
                {
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $B_0_Entity -VoiceResponses $B_0_VoiceCommand
                }
            }
            elseif ( $B_0_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_0_RedirectTarget
                $B_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                if ( $B_0_VoiceCommand -eq "" )
                {            
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $B_0_MenuOptionPrompt
                }
                else
                {
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $B_0_MenuOptionPrompt -VoiceResponses $B_0_VoiceCommand
                }
            }
            elseif ( $B_0_Redirect -eq "TEXT" )
            {
                $B_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_0_RedirectTarget
 
                if ( $B_0_VoiceCommand -eq "" )
                {            
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $B_0_MenuOptionPrompt
                }
                else
                { 
                    $B_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $B_0_MenuOptionPrompt -VoiceResponses $B_0_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-0-ERROR"
           } 
        }


        #
        # option 1
        #
        if ( $B_1_Redirect -ne "NONE" )
        {
            if ( $B_1_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone1
            }
            elseif ( $B_1_Redirect -eq "User" -OR $B_1_Redirect -eq "ApplicationEndpoint" -OR $B_1_Redirect -eq "ConfigurationEndpoint" -OR $B_1_Redirect -eq "SharedVoicemail" -OR $B_1_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_1_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_1_SharedVoicemailTranscription -AND $B_1_SharedVoicemailSuppress )
                    {
                        $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_1_SharedVoicemailTranscription )
                    {
                        $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect -EnableTranscription
                    }
                    elseif ( $B_1_SharedVoicemailSuppress )
                    {
                        $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect
                    }
                }
                elseif ( ( $B_1_Redirect -eq "ApplicationEndpoint" -OR $B_1_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect -CallPriority $B_1_Redirect_Call_Priority
                }
                else
                {
                    $B_1_Entity = New-CsAutoAttendantCallableEntity -Identity $B_1_RedirectTarget -Type $B_1_Redirect
                }
 
                if ( $B_1_VoiceCommand -eq "" )
                {
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $B_1_Entity
                }
                else
                {
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $B_1_Entity -VoiceResponses $B_1_VoiceCommand
                }
            }
            elseif ( $B_1_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_1_RedirectTarget
                $B_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                if ( $B_1_VoiceCommand -eq "" )
                {            
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $B_1_MenuOptionPrompt
                }
                else
                {
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $B_1_MenuOptionPrompt -VoiceResponses $B_1_VoiceCommand
                }
            }
            elseif ( $B_1_Redirect -eq "TEXT" )
            {
                $B_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_1_RedirectTarget
 
                if ( $B_1_VoiceCommand -eq "" )
                {            
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $B_1_MenuOptionPrompt
                }
                else
                {
                    $B_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $B_1_MenuOptionPrompt -VoiceResponses $B_1_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-1-ERROR"
            }
        }

        #
        # option 2
        #
        if ( $B_2_Redirect -ne "NONE" )
        {
            if ( $B_2_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone2
            }
            elseif ( $B_2_Redirect -eq "User" -OR $B_2_Redirect -eq "ApplicationEndpoint" -OR $B_2_Redirect -eq "ConfigurationEndpoint" -OR $B_2_Redirect -eq "SharedVoicemail" -OR $B_2_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_2_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_2_SharedVoicemailTranscription -AND $B_2_SharedVoicemailSuppress )
                    {
                        $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_2_SharedVoicemailTranscription )
                    {
                        $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect -EnableTranscription
                    } 
                    elseif ( $B_2_SharedVoicemailSuppress )
                    {
                        $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }					
                    else
                    {
                        $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect
                    }
                }
                elseif ( ( $B_2_Redirect -eq "ApplicationEndpoint" -OR $B_2_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect -CallPriority $B_2_Redirect_Call_Priority
                }
                else
                {
                    $B_2_Entity = New-CsAutoAttendantCallableEntity -Identity $B_2_RedirectTarget -Type $B_2_Redirect
                }
 
                if ( $B_2_VoiceCommand -eq "" )
                {
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $B_2_Entity
                }
                else
                {
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $B_2_Entity -VoiceResponses $B_2_VoiceCommand
                }
            }
            elseif ( $B_2_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_2_RedirectTarget
                $B_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                if ( $B_2_VoiceCommand -eq "" )
                {            
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $B_2_MenuOptionPrompt
                }
                else
                {
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $B_2_MenuOptionPrompt -VoiceResponses $B_2_VoiceCommand
                }
            }
            elseif ( $B_2_Redirect -eq "TEXT" )
            {
                $B_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_2_RedirectTarget
 
                if ( $B_2_VoiceCommand -eq "" )
                {            
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $B_2_MenuOptionPrompt
                }
                else
                {
                    $B_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $B_2_MenuOptionPrompt -VoiceResponses $B_2_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-2-ERROR"
            }
	    }


        #
        # option 3
        #
        if ( $B_3_Redirect -ne "NONE" )
        {
            if ( $B_3_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone3
            }
            elseif ( $B_3_Redirect -eq "User" -OR $B_3_Redirect -eq "ApplicationEndpoint" -OR $B_3_Redirect -eq "ConfigurationEndpoint" -OR $B_3_Redirect -eq "SharedVoicemail" -OR $B_3_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_3_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_3_SharedVoicemailTranscription -AND $B_3_SharedVoicemailSuppress )
                    {
                        $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_3_SharedVoicemailTranscription )
                    {
                        $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect -EnableTranscription
                    }
                    elseif ( $B_3_SharedVoicemailSuppress )
                    {
                        $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect
                    }
                }
                elseif ( ( $B_3_Redirect -eq "ApplicationEndpoint" -OR $B_3_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect -CallPriority $B_3_Redirect_Call_Priority
                }
				else
                {
                    $B_3_Entity = New-CsAutoAttendantCallableEntity -Identity $B_3_RedirectTarget -Type $B_3_Redirect
                }

                if ( $B_3_VoiceCommand -eq "" )
                {
                    $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $B_3_Entity
                }
                else
                {
                    $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $B_3_Entity -VoiceResponses $B_3_VoiceCommand
                }
            }
            elseif ( $B_3_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_3_RedirectTarget
                $B_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

               if ( $B_3_VoiceCommand -eq "" )
               {            
                   $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $B_3_MenuOptionPrompt
               }
               else
               {
                   $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $B_3_MenuOptionPrompt -VoiceResponses $B_3_VoiceCommand
               }
            }
            elseif ( $B_3_Redirect -eq "TEXT" )
            {
                $B_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_3_RedirectTarget

                if ( $B_3_VoiceCommand -eq "" )
                {            
                    $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $B_3_MenuOptionPrompt
                }
                else
                {
                    $B_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $B_3_MenuOptionPrompt -VoiceResponses $B_3_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-3-ERROR"
            }
        }



        #
        # option 4
        #
        if ( $B_4_Redirect -ne "NONE" )
        {
            if ( $B_4_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone4
            }
            elseif ( $B_4_Redirect -eq "User" -OR $B_4_Redirect -eq "ApplicationEndpoint" -OR $B_4_Redirect -eq "ConfigurationEndpoint" -OR $B_4_Redirect -eq "SharedVoicemail" -OR $B_4_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_4_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_4_SharedVoicemailTranscription -AND $B_4_SharedVoicemailSuppress )
                    {
                        $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_4_SharedVoicemailTranscription )
                    {
                        $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect -EnableTranscription
                    }
                    elseif ( $B_4_SharedVoicemailSuppress )
                    {
                        $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect
                    }
                }
                elseif ( ( $B_4_Redirect -eq "ApplicationEndpoint" -OR $B_4_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect -CallPriority $B_4_Redirect_Call_Priority
                }
                else
                {
                    $B_4_Entity = New-CsAutoAttendantCallableEntity -Identity $B_4_RedirectTarget -Type $B_4_Redirect
                }

                if ( $B_4_VoiceCommand -eq "" )
                {
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $B_4_Entity
                }
                else
                {
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $B_4_Entity -VoiceResponses $B_4_VoiceCommand
                }
            }
            elseif ( $B_4_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_4_RedirectTarget
                $B_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_4_VoiceCommand -eq "" )
                {            
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $B_4_MenuOptionPrompt
                }
                else
                {
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $B_4_MenuOptionPrompt -VoiceResponses $B_4_VoiceCommand
                }
            }
            elseif ( $B_4_Redirect -eq "TEXT" )
            {
                $B_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_4_RedirectTarget

                if ( $B_4_VoiceCommand -eq "" )
                {            
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $B_4_MenuOptionPrompt
                }
                else
                {
                    $B_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $B_4_MenuOptionPrompt -VoiceResponses $B_4_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-4-ERROR"
            }
        }


        #
        # option 5
        #
        if ( $B_5_Redirect -ne "NONE" )
        {
            if ( $B_5_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone5
            }
            elseif ( $B_5_Redirect -eq "User" -OR $B_5_Redirect -eq "ApplicationEndpoint" -OR $B_5_Redirect -eq "ConfigurationEndpoint" -OR $B_5_Redirect -eq "SharedVoicemail" -OR $B_5_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_5_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_5_SharedVoicemailTranscription -AND $B_5_SharedVoicemailSuppress )
                    {
                        $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_5_SharedVoicemailTranscription )
                    {
                        $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect -EnableTranscription
                    }
                    elseif ( $B_5_SharedVoicemailSuppress )
                    {
                        $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect
                    }
                }
                elseif ( ( $B_5_Redirect -eq "ApplicationEndpoint" -OR $B_5_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect -CallPriority $B_5_Redirect_Call_Priority
                }
                else
                {
                     $B_5_Entity = New-CsAutoAttendantCallableEntity -Identity $B_5_RedirectTarget -Type $B_5_Redirect
                }

                if ( $B_5_VoiceCommand -eq "" )
                {
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $B_5_Entity
                }
                else
                {
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $B_5_Entity -VoiceResponses $B_5_VoiceCommand
                }
            }
            elseif ( $B_5_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_5_RedirectTarget
                $B_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_5_VoiceCommand -eq "" )
                {            
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $B_5_MenuOptionPrompt
                }
                else
                {
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $B_5_MenuOptionPrompt -VoiceResponses $B_5_VoiceCommand
                }
            }
            elseif ( $B_5_Redirect -eq "TEXT" )
            {
                $B_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_5_RedirectTarget

                if ( $B_5_VoiceCommand -eq "" )
                {            
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $B_5_MenuOptionPrompt
                }
                else
                {
                    $B_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $B_5_MenuOptionPrompt -VoiceResponses $B_5_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-5-ERROR"
            }
        }


        #
         # option 6
        #
        if ( $B_6_Redirect -ne "NONE" )
        {
            if ( $B_6_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone6
            }
            elseif ( $B_6_Redirect -eq "User" -OR $B_6_Redirect -eq "ApplicationEndpoint" -OR $B_6_Redirect -eq "ConfigurationEndpoint" -OR $B_6_Redirect -eq "SharedVoicemail" -OR $B_6_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_6_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_6_SharedVoicemailTranscription -AND $B_6_SharedVoicemailSuppress )
                    {
                        $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_6_SharedVoicemailTranscription )
                    {
                        $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect -EnableTranscription
                    }
                    elseif ( $B_6_SharedVoicemailSuppress )
                    {
                        $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect
                    }
                }
                elseif ( ( $B_6_Redirect -eq "ApplicationEndpoint" -OR $B_6_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect -CallPriority $B_6_Redirect_Call_Priority
                }
                else
                {
                    $B_6_Entity = New-CsAutoAttendantCallableEntity -Identity $B_6_RedirectTarget -Type $B_6_Redirect
                }

                if ( $B_6_VoiceCommand -eq "" )
                {
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $B_6_Entity
                }
                else
                {
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $B_6_Entity -VoiceResponses $B_6_VoiceCommand
                }
            }
            elseif ( $B_6_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_6_RedirectTarget
                $B_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_6_VoiceCommand -eq "" )
                {            
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $B_6_MenuOptionPrompt
                }
                else
                {
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $B_6_MenuOptionPrompt -VoiceResponses $B_6_VoiceCommand
                }
            }
            elseif ( $B_6_Redirect -eq "TEXT" )
            {
                $B_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_6_RedirectTarget

                if ( $B_6_VoiceCommand -eq "" )
                {            
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $B_6_MenuOptionPrompt
                }
                else
                {
                    $B_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $B_6_MenuOptionPrompt -VoiceResponses $B_6_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-6-ERROR"
             }
        }


        #
        # option 7
        #
        if ( $B_7_Redirect -ne "NONE" )
        {
            if ( $B_7_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone7
            }
            elseif ( $B_7_Redirect -eq "User" -OR $B_7_Redirect -eq "ApplicationEndpoint" -OR $B_7_Redirect -eq "ConfigurationEndpoint" -OR $B_7_Redirect -eq "SharedVoicemail" -OR $B_7_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_7_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_7_SharedVoicemailTranscription -AND $B_7_SharedVoicemailSuppress )
                    {
                        $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_7_SharedVoicemailTranscription )
                    {
                        $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect -EnableTranscription
                    }
                    elseif ( $B_7_SharedVoicemailSuppress )
                    {
                        $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect
                    }
                }
                elseif ( ( $B_7_Redirect -eq "ApplicationEndpoint" -OR $B_7_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect -CallPriority $B_7_Redirect_Call_Priority
                }
                else
                {
                    $B_7_Entity = New-CsAutoAttendantCallableEntity -Identity $B_7_RedirectTarget -Type $B_7_Redirect
                }

                if ( $B_7_VoiceCommand -eq "" )
                {
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $B_7_Entity
                }
                else
                {
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $B_7_Entity -VoiceResponses $B_7_VoiceCommand
                }
            }
            elseif ( $B_7_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_7_RedirectTarget
                $B_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_7_VoiceCommand -eq "" )
                {            
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $B_7_MenuOptionPrompt
                }
                else
                {
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $B_7_MenuOptionPrompt -VoiceResponses $B_7_VoiceCommand
                }
            }
            elseif ( $B_7_Redirect -eq "TEXT" )
            {
                $B_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_7_RedirectTarget

                if ( $B_7_VoiceCommand -eq "" )
                {            
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $B_7_MenuOptionPrompt
                }
                else
                {
                    $B_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $B_7_MenuOptionPrompt -VoiceResponses $B_7_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-7-ERROR"
            }
        }


        #
        # option 8
        #
        if ( $B_8_Redirect -ne "NONE" )
        {
            if ( $B_8_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone8
            }
            elseif ( $B_8_Redirect -eq "User" -OR $B_8_Redirect -eq "ApplicationEndpoint" -OR $B_8_Redirect -eq "ConfigurationEndpoint" -OR $B_8_Redirect -eq "SharedVoicemail" -OR $B_8_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_8_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_8_SharedVoicemailTranscription -AND $B_8_SharedVoicemailSuppress )
                    {
                        $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_8_SharedVoicemailTranscription )
                    {
                        $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect -EnableTranscription
                    }
                    elseif ( $B_8_SharedVoicemailSuppress )
                    {
                        $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect
                    }
                }
                elseif ( ( $B_8_Redirect -eq "ApplicationEndpoint" -OR $B_8_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect -CallPriority $B_8_Redirect_Call_Priority
                }
                else
                {
                    $B_8_Entity = New-CsAutoAttendantCallableEntity -Identity $B_8_RedirectTarget -Type $B_8_Redirect
                }

                if ( $B_8_VoiceCommand -eq "" )
                {
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $B_8_Entity
                }
                else
                {
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $B_8_Entity -VoiceResponses $B_8_VoiceCommand
                }
            }
            elseif ( $B_8_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_8_RedirectTarget
                $B_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_8_VoiceCommand -eq "" )
                {            
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $B_8_MenuOptionPrompt
                }
                else
                {
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $B_8_MenuOptionPrompt -VoiceResponses $B_8_VoiceCommand
                }
            }
            elseif ( $B_8_Redirect -eq "TEXT" )
            {
                $B_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_8_RedirectTarget

                if ( $B_8_VoiceCommand -eq "" )
                {            
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $B_8_MenuOptionPrompt
                }
                else
                {
                    $B_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $B_8_MenuOptionPrompt -VoiceResponses $B_8_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-8-ERROR"
            }
        }


        #
        # option 9
         #
        if ( $B_9_Redirect -ne "NONE" )
        {
            if ( $B_9_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone9
            }
            elseif ( $B_9_Redirect -eq "User" -OR $B_9_Redirect -eq "ApplicationEndpoint" -OR $B_9_Redirect -eq "ConfigurationEndpoint" -OR $B_9_Redirect -eq "SharedVoicemail" -OR $B_9_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_9_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_9_SharedVoicemailTranscription -AND $B_9_SharedVoicemailSuppress )
                    {
                        $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_9_SharedVoicemailTranscription )
                    {
                        $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect -EnableTranscription
                    }
                    elseif ( $B_9_SharedVoicemailSuppress )
                    {
                        $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect
                    }
                }
                elseif ( ( $B_9_Redirect -eq "ApplicationEndpoint" -OR $B_9_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect -CallPriority $B_9_Redirect_Call_Priority
                }
                else
                {
                    $B_9_Entity = New-CsAutoAttendantCallableEntity -Identity $B_9_RedirectTarget -Type $B_9_Redirect
                }

                if ( $B_9_VoiceCommand -eq "" )
                {
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $B_9_Entity
                }
                else
                {
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $B_9_Entity -VoiceResponses $B_9_VoiceCommand
                }
            }
            elseif ( $B_9_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_9_RedirectTarget
                $B_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_9_VoiceCommand -eq "" )
                {            
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $B_9_MenuOptionPrompt
                }
                else
                {
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $B_9_MenuOptionPrompt -VoiceResponses $B_9_VoiceCommand
                }
            }
            elseif ( $B_9_Redirect -eq "TEXT" )
            {
                $B_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_9_RedirectTarget

                if ( $B_9_VoiceCommand -eq "" )
                {            
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $B_9_MenuOptionPrompt
                }
                else
                {
                    $B_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $B_9_MenuOptionPrompt -VoiceResponses $B_9_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-9-ERROR"
            }
        }


        #
        # option Star
        #
        if ( $B_Star_Redirect -ne "NONE" )
        {
            if ( $B_Star_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse ToneStar
            }
            elseif ( $B_Star_Redirect -eq "User" -OR $B_Star_Redirect -eq "ApplicationEndpoint" -OR $B_Star_Redirect -eq "ConfigurationEndpoint" -OR $B_Star_Redirect -eq "SharedVoicemail" -OR $B_Star_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_Star_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_Star_SharedVoicemailTranscription -AND $B_Star_SharedVoicemailSuppress )
                    {
                        $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_Star_SharedVoicemailTranscription )
                    {
                        $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect -EnableTranscription
                    }
                    elseif ( $B_Star_SharedVoicemailSuppress )
                    {
                        $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect
                    }
                }
                elseif ( ( $B_Star_Redirect -eq "ApplicationEndpoint" -OR $B_Star_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect -CallPriority $B_Star_Redirect_Call_Priority
                }
                else
                {
                    $B_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Star_RedirectTarget -Type $B_Star_Redirect
                }

                if ( $B_Star_VoiceCommand -eq "" )
                {
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $B_Star_Entity
                }
                else
                {
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $B_Star_Entity -VoiceResponses $B_Star_VoiceCommand
                }
            }
            elseif ( $B_Star_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_Star_RedirectTarget
                $B_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_Star_VoiceCommand -eq "" )
                {            
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $B_Star_MenuOptionPrompt
                }
                else
                {
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $B_Star_MenuOptionPrompt -VoiceResponses $B_Star_VoiceCommand
                }
            }
            elseif ( $B_Star_Redirect -eq "TEXT" )
            {
                $B_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_Star_RedirectTarget

                if ( $B_Star_VoiceCommand -eq "" )
                {            
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $B_Star_MenuOptionPrompt
                }
                else
                {
                    $B_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $B_Star_MenuOptionPrompt -VoiceResponses $B_Star_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-Star-ERROR"
            }
        }


        #
        # option Pound
        #
        if ( $B_Pound_Redirect -ne "NONE" )
        {
            if ( $B_Pound_Redirect -eq "Operator" )
            {
                # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse TonePound
            }
            elseif ( $B_Pound_Redirect -eq "User" -OR $B_Pound_Redirect -eq "ApplicationEndpoint" -OR $B_Pound_Redirect -eq "ConfigurationEndpoint" -OR $B_Pound_Redirect -eq "SharedVoicemail" -OR $B_Pound_Redirect -eq "ExternalPSTN" )
            {
                if ( $B_Pound_Redirect -eq "SharedVoicemail" )
                {
                    if ( $B_Pound_SharedVoicemailTranscription -AND $B_Pound_SharedVoicemailSuppress )
                    {
                        $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                    }
                    elseif ( $B_Pound_SharedVoicemailTranscription )
                    {
                        $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect -EnableTranscription
                    }
                    elseif ( $B_Pound_SharedVoicemailSuppress )
                    {
                        $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect -EnableSharedVoicemailSystemPromptSuppression
                    }
                    else
                    {
                        $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect
                    }
                }
                elseif ( ( $B_Pound_Redirect -eq "ApplicationEndpoint" -OR $B_Pound_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
                {
                    $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect -CallPriority $B_Pound_Redirect_Call_Priority
                }
                else
                {
                    $B_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $B_Pound_RedirectTarget -Type $B_Pound_Redirect
                }

                if ( $B_Pound_VoiceCommand -eq "" )
                {
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $B_Pound_Entity
                }
                else
                {
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $B_Pound_Entity -VoiceResponses $B_Pound_VoiceCommand
                }
            }
            elseif ( $B_Pound_Redirect -eq "FILE" )
            {
                $audioFile = AudioFileImport $B_Pound_RedirectTarget
                $B_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                if ( $B_Pound_VoiceCommand -eq "" )
                {            
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $B_Pound_MenuOptionPrompt
                }
                else
                {
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $B_Pound_MenuOptionPrompt -VoiceResponses $B_Pound_VoiceCommand
                }
            }
            elseif ( $B_Pound_Redirect -eq "TEXT" )
            {
                $B_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $B_Pound_RedirectTarget

                if ( $B_Pound_VoiceCommand -eq "" )
                {            
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $B_Pound_MenuOptionPrompt
                }
                else
                {
                    $B_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $B_Pound_MenuOptionPrompt -VoiceResponses $B_Pound_VoiceCommand
                }
            }
            else
            {
                Write-Host "B-Pound-ERROR"
            }
        }
		
		#
        # build default menu
        #
		$defaultMenu = New-CsAutoAttendantMenu -Name "Open Hours Menu" -MenuOptions @($B_0_MenuOption,$B_1_MenuOption,$B_2_MenuOption,$B_3_MenuOption,$B_4_MenuOption,$B_5_MenuOption,$B_6_MenuOption,$B_7_MenuOption,$B_8_MenuOption,$B_9_MenuOption,$B_Star_MenuOption,$B_Pound_MenuOption) -Prompts $B_MenuGreetingPrompt -DirectorySearchMethod $B_DirectorySearch
		
        if ( $B_GreetingPromptConfigured )
        {
            if ( $B_Force )
            {
                $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Greetings @($B_GreetingPrompt) -Menu $defaultMenu -ForceListenMenuEnabled
            }
            else
            {
                $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Greetings @($B_GreetingPrompt) -Menu $defaultMenu
            }
        }
        else
        {
            if ( $B_Force )
            {
                $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Menu $defaultMenu -ForceListenMenuEnabled
            }
            else
            {
                $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default Call Flow" -Menu $defaultMenu
            }
        }
    }
    else
    {
        Write-Host "BLANK - can't build default call flow.  Fix error and start again."
        exit
    }


    $command += "-DefaultCallFlow `$defaultCallFlow "


    #
    # business hours check
    #
    if ( ! $Hours24 )
    {
		Write-Host "`tBuilding After Hours Call Flow"
        $BH = $BusinessHours -split ","

        #
        # breakup business hours by day of week
        #
		Write-Host "`t`tProcessing Business Hours Schedule"
        for ( $i=0; $i -lt $BH.length; $i++ )
        {
            switch ( $BH[$i] )
            {
                "Monday"    {
                                $BH_Monday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Monday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Tuesday"      {  
			                    $BH_Tuesday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Tuesday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Wednesday"    {  
			                    $BH_Wednesday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Wednesday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Thursday"     {  
			                    $BH_Thursday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Thursday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Friday"       {  
			                    $BH_Friday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Friday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Saturday"     {  
			                    $BH_Saturday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Saturday += $BH[$j]
                                    $j += 1
                                }
                            }
             "Sunday"       {  
			                    $BH_Sunday = @()

                                $j = $i + 1
                                while ( $BH[$j] -match ":" )
                                {
                                    $BH_Sunday += $BH[$j]
                                    $j += 1
                                }
                            }
             Default        {
                                # it's a time not a day
                            }
            }
        }


        $BH_Schedule_List =  "New-CsOnlineSchedule -Name 'After Hours Schedule' -WeeklyRecurrentSchedule "

        if ( $BH_Monday.length -gt 0 )
        {
            $BH_Schedule_List += "-MondayHours @("
        }

        for ( $i=0; $i -lt $BH_Monday.length; $i++ )
        {
            if ( $BH_Monday[$i+1] -match "\+1$" )
            {
                $BH_Monday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Monday_Time_1 = New-CsOnlineTimeRange -Start $BH_Monday[$i] -End $BH_Monday[$i+1]
                        $BH_Schedule_List += "`$BH_Monday_Time_1"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tMonday Start Time 1:`t$($BH_Monday[$i])`tEnd Time 1:`t$($BH_Monday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Monday_Time_2 = New-CsonlineTimeRange -Start $BH_Monday[$i] -End $BH_Monday[$i+1]
                        $BH_Schedule_List += ",`$BH_Monday_Time_2"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tMonday Start Time 2:`t$($BH_Monday[$i])`tEnd Time 2:`t$($BH_Monday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Monday_Time_3 = New-CsOnlineTimeRange -Start $BH_Monday[$i] -End $BH_Monday[$i+1]
                        $BH_Schedule_List += ",`$BH_Monday_Time_3"

						if ( $Verbose )
						{
							Write-Host "`t`t`tMonday Start Time 3:`t$($BH_Monday[$i])`tEnd Time 3:`t$($BH_Monday[$i+1])"
						}
                    }
                "6" {  
                        $BH_Monday_Time_4 = New-CsOnlineTimeRange -Start $BH_Monday[$i] -End $BH_Monday[$i+1]
                        $BH_Schedule_List += ",`$BH_Monday_Time_4"

						if ( $Verbose )
						{
							Write-Host "`t`t`tMonday Start Time 4:`t$($BH_Monday[$i])`tEnd Time 4:`t$($BH_Monday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Monday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }


        if ( $BH_Tuesday.length -gt 0 )
        {
            $BH_Schedule_List += "-TuesdayHours @("
        }

        for ( $i=0; $i -lt $BH_Tuesday.length; $i++ )
        {
            if ( $BH_Tuesday[$i+1] -match "\+1$" )
            {
                $BH_Tuesday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Tuesday_Time_1 = New-CsOnlineTimeRange -Start $BH_Tuesday[$i] -End $BH_Tuesday[$i+1]
                        $BH_Schedule_List += "`$BH_Tuesday_Time_1"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tTuesday Start Time 1:`t$($BH_Tuesday[$i])`tEnd Time 1:`t$($BH_Tuesday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Tuesday_Time_2 = New-CsonlineTimeRange -Start $BH_Tuesday[$i] -End $BH_Tuesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Tuesday_Time_2"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tTuesday Start Time 2:`t$($BH_Tuesday[$i])`tEnd Time 2:`t$($BH_Tuesday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Tuesday_Time_3 = New-CsOnlineTimeRange -Start $BH_Tuesday[$i] -End $BH_Tuesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Tuesday_Time_3"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tTuesday Start Time 3:`t$($BH_Tuesday[$i])`tEnd Time 3:`t$($BH_Tuesday[$i+1])"
						}
                    }
                "6" {  
                        $BH_Tuesday_Time_4 = New-CsOnlineTimeRange -Start $BH_Tuesday[$i] -End $BH_Tuesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Tuesday_Time_4"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tTuesday Start Time 4:`t$($BH_Tuesday[$i])`tEnd Time 4:`t$($BH_Tuesday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Tuesday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }


        if ( $BH_Wednesday.length -gt 0 )
        {
            $BH_Schedule_List += "-WednesdayHours @("
        }

        for ( $i=0; $i -lt $BH_Wednesday.length; $i++ )
        {
            if ( $BH_Wednesday[$i+1] -match "\+1$" )
            {
                $BH_Wednesday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Wednesday_Time_1 = New-CsOnlineTimeRange -Start $BH_Wednesday[$i] -End $BH_Wednesday[$i+1]
                        $BH_Schedule_List += "`$BH_Wednesday_Time_1"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tWednesday Start Time 1:`t$($BH_Wednesday[$i])`tEnd Time 1:`t$($BH_Wednesday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Wednesday_Time_2 = New-CsonlineTimeRange -Start $BH_Wednesday[$i] -End $BH_Wednesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Wednesday_Time_2"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tWednesday Start Time 2:`t$($BH_Wednesday[$i])`tEnd Time 2:`t$($BH_Wednesday[$i+1])"
						}
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tWednesday Start Time 2:`t$($BH_Wednesday[$i])`tEnd Time 2:`t$($BH_Wednesday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Wednesday_Time_3 = New-CsOnlineTimeRange -Start $BH_Wednesday[$i] -End $BH_Wednesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Wednesday_Time_3"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tWednesday Start Time 3:`t$($BH_Wednesday[$i])`tEnd Time 3:`t$($BH_Wednesday[$i+1])"
						}
                   }
                "6" {  
                        $BH_Wednesday_Time_4 = New-CsOnlineTimeRange -Start $BH_Wednesday[$i] -End $BH_Wednesday[$i+1]
                        $BH_Schedule_List += ",`$BH_Wednesday_Time_4"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tWednesday Start Time 4:`t$($BH_Wednesday[$i])`tEnd Time 4:`t$($BH_Wednesday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Wednesday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }



        if ( $BH_Thursday.length -gt 0 )
        {
            $BH_Schedule_List += "-ThursdayHours @("
        }

        for ( $i=0; $i -lt $BH_Thursday.length; $i++ )
        {
            if ( $BH_Thursday[$i+1] -match "\+1$" )
            {
                $BH_Thursday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Thursday_Time_1 = New-CsOnlineTimeRange -Start $BH_Thursday[$i] -End $BH_Thursday[$i+1]
                        $BH_Schedule_List += "`$BH_Thursday_Time_1"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tThursday Start Time 1:`t$($BH_Thursday[$i])`tEnd Time 1:`t$($BH_Thursday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Thursday_Time_2 = New-CsonlineTimeRange -Start $BH_Thursday[$i] -End $BH_Thursday[$i+1]
                        $BH_Schedule_List += ",`$BH_Thursday_Time_2"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tThursday Start Time 2:`t$($BH_Thursday[$i])`tEnd Time 2:`t$($BH_Thursday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Thursday_Time_3 = New-CsOnlineTimeRange -Start $BH_Thursday[$i] -End $BH_Thursday[$i+1]
                        $BH_Schedule_List += ",`$BH_Thursday_Time_3"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tThursday Start Time 3:`t$($BH_Thursday[$i])`tEnd Time 3:`t$($BH_Thursday[$i+1])"
						}
                    }
                 "6" {  
                        $BH_Thursday_Time_4 = New-CsOnlineTimeRange -Start $BH_Thursday[$i] -End $BH_Thursday[$i+1]
                        $BH_Schedule_List += ",`$BH_Thursday_Time_4"
						
						if ( $Verbose )
						{
							Write-Host "`t`t`tThursday Start Time 4:`t$($BH_Thursday[$i])`tEnd Time 4:`t$($BH_Thursday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Thursday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }


        if ( $BH_Friday.length -gt 0 )
        {
            $BH_Schedule_List += "-FridayHours @("
        }

        for ( $i=0; $i -lt $BH_Friday.length; $i++ )
        {
            if ( $BH_Friday[$i+1] -match "\+1$" )
            {
                $BH_Friday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Friday_Time_1 = New-CsOnlineTimeRange -Start $BH_Friday[$i] -End $BH_Friday[$i+1]
                        $BH_Schedule_List += "`$BH_Friday_Time_1"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tFriday Start Time 1:`t$($BH_Friday[$i])`tEnd Time 1:`t$($BH_Friday[$i+1])"
						}
                   }
                "2" {  
                        $BH_Friday_Time_2 = New-CsonlineTimeRange -Start $BH_Friday[$i] -End $BH_Friday[$i+1]
                        $BH_Schedule_List += ",`$BH_Friday_Time_2"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tFriday Start Time 2:`t$($BH_Friday[$i])`tEnd Time 2:`t$($BH_Friday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Friday_Time_3 = New-CsOnlineTimeRange -Start $BH_Friday[$i] -End $BH_Friday[$i+1]
                        $BH_Schedule_List += ",`$BH_Friday_Time_3"
  						
						if ( $Verbose )
						{
							Write-Host "`t`t`tFriday Start Time 3:`t$($BH_Friday[$i])`tEnd Time 3:`t$($BH_Friday[$i+1])"
						}
                   }
                "6" {  
                        $BH_Friday_Time_4 = New-CsOnlineTimeRange -Start $BH_Friday[$i] -End $BH_Friday[$i+1]
                        $BH_Schedule_List += ",`$BH_Friday_Time_4"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tFriday Start Time 4:`t$($BH_Friday[$i])`tEnd Time 4:`t$($BH_Friday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Friday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }


        if ( $BH_Saturday.length -gt 0 )
        {
            $BH_Schedule_List += "-SaturdayHours @("
        }

        for ( $i=0; $i -lt $BH_Saturday.length; $i++ )
        {
            if ( $BH_Saturday[$i+1] -match "\+1$" )
            {
                $BH_Saturday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Saturday_Time_1 = New-CsOnlineTimeRange -Start $BH_Saturday[$i] -End $BH_Saturday[$i+1]
                        $BH_Schedule_List += "`$BH_Saturday_Time_1"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSaturday Start Time 1:`t$($BH_Saturday[$i])`tEnd Time 1:`t$($BH_Saturday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Saturday_Time_2 = New-CsonlineTimeRange -Start $BH_Saturday[$i] -End $BH_Saturday[$i+1]
                        $BH_Schedule_List += ",`$BH_Saturday_Time_2"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSaturday Start Time 2:`t$($BH_Saturday[$i])`tEnd Time 2:`t$($BH_Saturday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Saturday_Time_3 = New-CsOnlineTimeRange -Start $BH_Saturday[$i] -End $BH_Saturday[$i+1]
                        $BH_Schedule_List += ",`$BH_Saturday_Time_3"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSaturday Start Time 3:`t$($BH_Saturday[$i])`tEnd Time 3:`t$($BH_Saturday[$i+1])"
						}
                    }
                "6" {  
                        $BH_Saturday_Time_4 = New-CsOnlineTimeRange -Start $BH_Saturday[$i] -End $BH_Saturday[$i+1]
                        $BH_Schedule_List += ",`$BH_Saturday_Time_4"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSaturday Start Time 4:`t$($BH_Saturday[$i])`tEnd Time 4:`t$($BH_Saturday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Saturday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }


        if ( $BH_Sunday.length -gt 0 )
        {
            $BH_Schedule_List += "-SundayHours @("
        }
 
        for ( $i=0; $i -lt $BH_Sunday.length; $i++ )
        {
            if ( $BH_Sunday[$i+1] -match "\+1$" )
            {
                $BH_Sunday[$i+1] = "1.00:00"
            }
            switch ( $i )
            {
                "0" {  
                        $BH_Sunday_Time_1 = New-CsOnlineTimeRange -Start $BH_Sunday[$i] -End $BH_Sunday[$i+1]
                        $BH_Schedule_List += "`$BH_Sunday_Time_1"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSunday Start Time 1:`t$($BH_Sunday[$i])`tEnd Time 1:`t$($BH_Sunday[$i+1])"
						}
                    }
                "2" {  
                        $BH_Sunday_Time_2 = New-CsonlineTimeRange -Start $BH_Sunday[$i] -End $BH_Sunday[$i+1]
                        $BH_Schedule_List += ",`$BH_Sunday_Time_2"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSunday Start Time 2:`t$($BH_Sunday[$i])`tEnd Time 2:`t$($BH_Sunday[$i+1])"
						}
                    }
                "4" {  
                        $BH_Sunday_Time_3 = New-CsOnlineTimeRange -Start $BH_Sunday[$i] -End $BH_Sunday[$i+1]
                        $BH_Schedule_List += ",`$BH_Sunday_Time_3"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSunday Start Time 3:`t$($BH_Sunday[$i])`tEnd Time 3:`t$($BH_Sunday[$i+1])"
						}
                    }
                "6" {  
                        $BH_Sunday_Time_4 = New-CsOnlineTimeRange -Start $BH_Sunday[$i] -End $BH_Sunday[$i+1]
                        $BH_Schedule_List += ",`$BH_Sunday_Time_4"
 						
						if ( $Verbose )
						{
							Write-Host "`t`t`tSunday Start Time 4:`t$($BH_Sunday[$i])`tEnd Time 4:`t$($BH_Sunday[$i+1])"
						}
                    }
            }
        }
        if ( $BH_Sunday.length -gt 0 )
        {
            $BH_Schedule_List += ") "
        }

        $BH_Schedule_List += "-Complement"

        if ( $Verbose )
		{
			Write-Host "`t`tBuilding Schedule"
		}

        $BH_Schedule = Invoke-Expression $BH_Schedule_List

        #
        #  After Hours Call Flow
        #
        switch ( $A_GreetingOption )
        {
            "FILE"  {  
                        $audioFile = AudioFileImport $A_Greeting
                        $A_GreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                        $A_GreetingPromptConfigured = $true
                    }
            "TEXT"  {  
                        $A_GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_Greeting
                        $A_GreetingPromptConfigured = $true
                    } 
            Default {  $A_GreetingPromptConfigured = $false }
        }


        if ( $A_GreetingRouting -eq "Disconnect" )
        {
            $afterhoursMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $afterhoursMenu = New-CsAutoAttendantMenu -Name "After Hours Menu" -MenuOptions @($afterhoursMenuOption)

            if ( $A_GreetingPromptConfigured )
            {
                  $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu @($afterhoursMenu)
            }
            else
            {
                  $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu
            }
			
            if ( $A_GreetingPromptConfigured )
            {
                if ( $A_Force )
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu $afterhoursMenu -ForceListenMenuEnabled
                }
                else
                {
                     $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu $afterhoursMenu
                }
            }
            else
            {
                if ( $A_Force )
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu -ForceListenMenuEnabled
                }
                else
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu
                }
            }

            $afterhoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $BH_Schedule.Id -CallFlowId $afterHoursCallFlow.Id
			
			if ( $Holidays -eq "" )
			{
				$command += "-CallFlows @(`$afterhoursCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation)"
			}
        }
        elseif ( $A_GreetingRouting -eq "User" -OR $A_GreetingRouting -eq "ApplicationEndpoint" -OR $A_GreetingRouting -eq "ConfigurationEndpoint" -OR $A_GreetingRouting -eq "ExternalPSTN" )
        {
			if ( ( $A_GreetingRouting -eq "ApplicationEndpoint" -OR $A_GreetingRouting -eq "ConfigurationEndpoint" ) -AND $CallPriority )
			{
				$afterhoursMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $A_GreetingRoutingTarget -Type $A_GreetingRouting -CallPriority $A_GreetingRouting_Call_Priority
			}
			else
			{
				$afterhoursMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $A_GreetingRoutingTarget -Type $A_GreetingRouting
			}

            $afterhoursMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $afterhoursMenuOptionCallableEntity -DtmfResponse Automatic
            $afterhoursMenu = New-CsAutoAttendantMenu -Name "After Hours Menu" -MenuOptions @($afterhoursMenuOption)

            if ( $A_GreetingPromptConfigured )
            {
                  $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu $afterhoursMenu
            }
            else
            {
                  $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu
            }
			
            $afterhoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $BH_Schedule.Id -CallFlowId $afterHoursCallFlow.Id

            if ( $Holidays -eq "" )
			{
                $command += "-CallFlows @(`$afterhoursCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation)"
			}
        }
        elseif ( $A_GreetingRouting -eq "MENU" )
        {
            switch ( $A_MenuGreetingOption )
            {
                "FILE"  {  
                            $audioFile = AudioFileImport $A_MenuGreeting
                            $A_MenuGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                            $A_GreetingPromptConfigured = $true
                        }
                "TEXT"  {  
                            $A_MenuGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_MenuGreeting
                            $A_GreetingPromptConfigured = $true
                        } 
                Default { $A_MenuGreetingPromptConfigured = $false }
            }

            #
            # Option 0
            #
            if ( $A_0_Redirect -ne "NONE" )
            {
                if ( $A_0_Redirect -eq "Operator" )
                {
                   # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                   $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0
                }
                elseif ( $A_0_Redirect -eq "User" -OR $A_0_Redirect -eq "ApplicationEndpoint" -OR $A_0_Redirect -eq "ConfigurationEndpoint" -OR $A_0_Redirect -eq "SharedVoicemail" -OR $A_0_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_0_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_0_SharedVoicemailTranscription -AND $A_0_SharedVoicemailSuppress )
                        {
                            $A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_0_SharedVoicemailTranscription )
                        {
                            $A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect -EnableTranscription
                        }
                        elseif ( $A_0_SharedVoicemailSuppress )
                        {
                            $A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect
                        }
                    }
					elseif ( ( $A_0_Redirect -eq "ApplicationEndpoint" -OR $A_0_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect -CallPriority $A_0_Redirect_Call_Priority
					}
                    else
                    {
                        $A_0_Entity = New-CsAutoAttendantCallableEntity -Identity $A_0_RedirectTarget -Type $A_0_Redirect
                    }
 
                    if ( $A_0_VoiceCommand -eq "" )
                    {
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $A_0_Entity
                    }
                    else
                    {
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $A_0_Entity -VoiceResponses $A_0_VoiceCommand
                    }
                }
                elseif ( $A_0_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_0_RedirectTarget
                    $A_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                    if ( $A_0_VoiceCommand -eq "" )
                    {                        
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $A_0_MenuOptionPrompt
                    }
                    else
                    {
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $A_0_MenuOptionPrompt -VoiceResponses $A_0_VoiceCommand
                    }
                }
                elseif ( $A_0_Redirect -eq "TEXT" )
                {
                    $A_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_0_RedirectTarget

                    if ( $A_0_VoiceCommand -eq "" )
                    {                        
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $A_0_MenuOptionPrompt
                    }
                    else
                    {
                        $A_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $A_0_MenuOptionPrompt -VoiceResponses $A_0_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-0-ERROR"
                }
            } # A_0_Redirect


            #
            # option 1
            #
            if ( $A_1_Redirect -ne "NONE" )
            {
                if ( $A_1_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone1
                }
                elseif ( $A_1_Redirect -eq "User" -OR $A_1_Redirect -eq "ApplicationEndpoint" -OR $A_1_Redirect -eq "ConfigurationEndpoint" -OR $A_1_Redirect -eq "SharedVoicemail" -OR $A_1_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_1_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_1_SharedVoicemailTranscription -AND $A_1_SharedVoicemailSuppress )
                        {
                            $A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_1_SharedVoicemailTranscription )
                        {
                            $A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect -EnableTranscription
                        }
                        elseif ( $A_1_SharedVoicemailSuppress )
                        {
                            $A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect
                        }
                    }
					elseif ( ( $A_1_Redirect -eq "ApplicationEndpoint" -OR $A_1_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect -CallPriority $A_1_Redirect_Call_Priority
					}
                    else
                    {
                        $A_1_Entity = New-CsAutoAttendantCallableEntity -Identity $A_1_RedirectTarget -Type $A_1_Redirect
                    }
 
                    if ( $A_1_VoiceCommand -eq "" )
                    {
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $A_1_Entity
                    }
                    else
                    {
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $A_1_Entity -VoiceResponses $A_1_VoiceCommand
                    }
                }
                elseif ( $A_1_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_1_RedirectTarget
                    $A_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                    if ( $A_1_VoiceCommand -eq "" )
                    {                        
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $A_1_MenuOptionPrompt
                    }
                    else
                    {
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $A_1_MenuOptionPrompt -VoiceResponses $A_1_VoiceCommand
                    }
                }
                elseif ( $A_1_Redirect -eq "TEXT" )
                {
                    $A_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_1_RedirectTarget
 
                    if ( $A_1_VoiceCommand -eq "" )
                    {                        
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $A_1_MenuOptionPrompt
                    }
                    else
                    {
                        $A_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $A_1_MenuOptionPrompt -VoiceResponses $A_1_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-1-ERROR"
                }
            }

            #
            # option 2
            #
            if ( $A_2_Redirect -ne "NONE" )
            {
                if ( $A_2_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone2
                }
                elseif ( $A_2_Redirect -eq "User" -OR $A_2_Redirect -eq "ApplicationEndpoint" -OR $A_2_Redirect -eq "ConfigurationEndpoint" -OR $A_2_Redirect -eq "SharedVoicemail" -OR $A_2_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_2_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_2_SharedVoicemailTranscription -AND $A_2_SharedVoicemailSuppress )
                        {
                            $A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_2_SharedVoicemailTranscription )
                        {
                            $A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect -EnableTranscription
                        }
                        elseif ( $A_2_SharedVoicemailSuppress )
                        {
                            $A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect
                        }
                    }
					elseif ( ( $A_2_Redirect -eq "ApplicationEndpoint" -OR $A_2_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect -CallPriority $A_2_Redirect_Call_Priority
					}
                    else
                    {
                        $A_2_Entity = New-CsAutoAttendantCallableEntity -Identity $A_2_RedirectTarget -Type $A_2_Redirect
                    }
 
                    if ( $A_2_VoiceCommand -eq "" )
                    {
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $A_2_Entity
                    }
                    else
                    {
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $A_2_Entity -VoiceResponses $A_2_VoiceCommand
                    }
                }
                elseif ( $A_2_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_2_RedirectTarget
                    $A_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_2_VoiceCommand -eq "" )
                    {                        
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $A_2_MenuOptionPrompt
                    }
                    else
                    {
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $A_2_MenuOptionPrompt -VoiceResponses $A_2_VoiceCommand
                    }
                }
                elseif ( $A_2_Redirect -eq "TEXT" )
                {
                    $A_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_2_RedirectTarget
   
                    if ( $A_2_VoiceCommand -eq "" )
                    {                        
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $A_2_MenuOptionPrompt
                    }
                    else
                    {
                        $A_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $A_2_MenuOptionPrompt -VoiceResponses $A_2_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-2-ERROR"
                }
            }


            #
            # option 3
            #
            if ( $A_3_Redirect -ne "NONE" )
            {
                if ( $A_3_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone3
                }
                elseif ( $A_3_Redirect -eq "User" -OR $A_3_Redirect -eq "ApplicationEndpoint" -OR $A_3_Redirect -eq "ConfigurationEndpoint" -OR $A_3_Redirect -eq "SharedVoicemail" -OR $A_3_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_3_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_3_SharedVoicemailTranscription -AND $A_3_SharedVoicemailSuppress )
                        {
                            $A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_3_SharedVoicemailTranscription )
                        {
                            $A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect -EnableTranscription
                        }
                        elseif ( $A_3_SharedVoicemailSuppress )
                        {
                            $A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect
                        }
                    }
					elseif ( ( $A_3_Redirect -eq "ApplicationEndpoint" -OR $A_3_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect -CallPriority $A_3_Redirect_Call_Priority
					}
                    else
                    {
                        $A_3_Entity = New-CsAutoAttendantCallableEntity -Identity $A_3_RedirectTarget -Type $A_3_Redirect
                    }
   
                    if ( $A_3_VoiceCommand -eq "" )
                    {
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $A_3_Entity
                    }
                    else
                    {
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $A_3_Entity -VoiceResponses $A_3_VoiceCommand
                    }
                }
                elseif ( $A_3_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_3_RedirectTarget
                    $A_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_3_VoiceCommand -eq "" )
                    {                        
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $A_3_MenuOptionPrompt
                    }
                    else
                    {
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $A_3_MenuOptionPrompt -VoiceResponses $A_3_VoiceCommand
                    }
                }
                elseif ( $A_3_Redirect -eq "TEXT" )
                {
                    $A_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_3_RedirectTarget
   
                    if ( $A_3_VoiceCommand -eq "" )
                    {                        
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $A_3_MenuOptionPrompt
                    }
                    else
                    {
                        $A_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $A_3_MenuOptionPrompt -VoiceResponses $A_3_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-3-ERROR"
                }
            }


            #
            # option 4
            #
            if ( $A_4_Redirect -ne "NONE" )
            {
                if ( $A_4_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone4
                }
                elseif ( $A_4_Redirect -eq "User" -OR $A_4_Redirect -eq "ApplicationEndpoint" -OR $A_4_Redirect -eq "ConfigurationEndpoint" -OR $A_4_Redirect -eq "SharedVoicemail" -OR $A_4_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_4_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_4_SharedVoicemailTranscription -AND $A_4_SharedVoicemailSuppress )
                        {
                            $A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_4_SharedVoicemailTranscription )
                        {
                            $A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect -EnableTranscription
                        }
                        elseif ( $A_4_SharedVoicemailSuppress )
                        {
                            $A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect
                        }
                    }
					elseif ( ( $A_4_Redirect -eq "ApplicationEndpoint" -OR $A_4_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect -CallPriority $A_4_Redirect_Call_Priority
					}
                    else
                    {
                        $A_4_Entity = New-CsAutoAttendantCallableEntity -Identity $A_4_RedirectTarget -Type $A_4_Redirect
                    }
 
                    if ( $A_4_VoiceCommand -eq "" )
                    {
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $A_4_Entity
                    }
                    else
                    {
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $A_4_Entity -VoiceResponses $A_4_VoiceCommand
                    }
                }
                elseif ( $A_4_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_4_RedirectTarget
                    $A_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_4_VoiceCommand -eq "" )
                    {                        
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $A_4_MenuOptionPrompt
                    }
                    else
                    {
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $A_4_MenuOptionPrompt -VoiceResponses $A_4_VoiceCommand
                    }
                }
                elseif ( $A_4_Redirect -eq "TEXT" )
                {
                    $A_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_4_RedirectTarget
   
                    if ( $A_4_VoiceCommand -eq "" )
                    {                        
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $A_4_MenuOptionPrompt
                    }
                    else
                    {
                        $A_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $A_4_MenuOptionPrompt -VoiceResponses $A_4_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-4-ERROR"
                }
            }


            #
            # option 5
            #
            if ( $A_5_Redirect -ne "NONE" )
            {
                if ( $A_5_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone5
                }
                elseif ( $A_5_Redirect -eq "User" -OR $A_5_Redirect -eq "ApplicationEndpoint" -OR $A_5_Redirect -eq "ConfigurationEndpoint" -OR $A_5_Redirect -eq "SharedVoicemail" -OR $A_5_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_5_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_5_SharedVoicemailTranscription -AND $A_5_SharedVoicemailSuppress )
                        {
                            $A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_5_SharedVoicemailTranscription )
                        {
                            $A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect -EnableTranscription
                        }
                        elseif ( $A_5_SharedVoicemailSuppress )
                        {
                            $A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect
                        }
                    }
					elseif ( ( $A_5_Redirect -eq "ApplicationEndpoint" -OR $A_5_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect -CallPriority $A_5_Redirect_Call_Priority
					}
                    else
                    {
                        $A_5_Entity = New-CsAutoAttendantCallableEntity -Identity $A_5_RedirectTarget -Type $A_5_Redirect
                    }
 
                    if ( $A_5_VoiceCommand -eq "" )
                    {
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $A_5_Entity
                    }
                    else
                    {
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $A_5_Entity -VoiceResponses $A_5_VoiceCommand
                    }
                }
                elseif ( $A_5_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_5_RedirectTarget
                    $A_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                    if ( $A_5_VoiceCommand -eq "" )
                    {                        
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $A_5_MenuOptionPrompt
                    }
                    else
                    {
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $A_5_MenuOptionPrompt -VoiceResponses $A_5_VoiceCommand
                    }
                }
                elseif ( $A_5_Redirect -eq "TEXT" )
                {
                    $A_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_5_RedirectTarget
 
                    if ( $A_5_VoiceCommand -eq "" )
                    {                        
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $A_5_MenuOptionPrompt
                    }
                    else
                    {
                        $A_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $A_5_MenuOptionPrompt -VoiceResponses $A_5_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-5-ERROR"
                }
            }


            #
            # option 6
            #
            if ( $A_6_Redirect -ne "NONE" )
            {
                if ( $A_6_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone6
                }
                elseif ( $A_6_Redirect -eq "User" -OR $A_6_Redirect -eq "ApplicationEndpoint" -OR $A_6_Redirect -eq "ConfigurationEndpoint" -OR $A_6_Redirect -eq "SharedVoicemail" -OR $A_6_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_6_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_6_SharedVoicemailTranscription -AND $A_6_SharedVoicemailSuppress )
                        {
                            $A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_6_SharedVoicemailTranscription )
                        {
                             $A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect -EnableTranscription
                        }
                        elseif ( $A_6_SharedVoicemailSuppress )
                        {
                            $A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect
                        }
                    }
					elseif ( ( $A_6_Redirect -eq "ApplicationEndpoint" -OR $A_6_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect -CallPriority $A_6_Redirect_Call_Priority
					}
                    else
                    {
                        $A_6_Entity = New-CsAutoAttendantCallableEntity -Identity $A_6_RedirectTarget -Type $A_6_Redirect
                    }
 
                    if ( $A_6_VoiceCommand -eq "" )
                    {
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $A_6_Entity
                    }
                    else
                    {
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $A_6_Entity -VoiceResponses $A_6_VoiceCommand
                    }
                }
                elseif ( $A_6_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_6_RedirectTarget
                    $A_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                    if ( $A_6_VoiceCommand -eq "" )
                    {                        
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $A_6_MenuOptionPrompt
                    }
                    else
                    {
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $A_6_MenuOptionPrompt -VoiceResponses $A_6_VoiceCommand
                    }
                }
                elseif ( $A_6_Redirect -eq "TEXT" )
                {
                    $A_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_6_RedirectTarget
   
                    if ( $A_6_VoiceCommand -eq "" )
                    {                        
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $A_6_MenuOptionPrompt
                    }
                    else
                    {
                        $A_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $A_6_MenuOptionPrompt -VoiceResponses $A_6_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-6-ERROR"
                }
            }


            #
            # option 7
            #
            if ( $A_7_Redirect -ne "NONE" )
            {
                if ( $A_7_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone7
                }
                elseif ( $A_7_Redirect -eq "User" -OR $A_7_Redirect -eq "ApplicationEndpoint" -OR $A_7_Redirect -eq "ConfigurationEndpoint" -OR $A_7_Redirect -eq "SharedVoicemail" -OR $A_7_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_7_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_7_SharedVoicemailTranscription -AND $A_7_SharedVoicemailSuppress )
                        {
                            $A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_7_SharedVoicemailTranscription )
                        {
                            $A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect -EnableTranscription
                        }
                        elseif ( $A_7_SharedVoicemailSuppress )
                        {
                            $A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect
                        }
                    }
					elseif ( ( $A_7_Redirect -eq "ApplicationEndpoint" -OR $A_7_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect -CallPriority $A_7_Redirect_Call_Priority
					}
                    else
                    {
                        $A_7_Entity = New-CsAutoAttendantCallableEntity -Identity $A_7_RedirectTarget -Type $A_7_Redirect
                    }

                    if ( $A_7_VoiceCommand -eq "" )
                    {
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $A_7_Entity
                    }
                    else
                    {
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $A_7_Entity -VoiceResponses $A_7_VoiceCommand
                    }
                }
                elseif ( $A_7_Redirect -eq "FILE" )
                 {
                    $audioFile = AudioFileImport $A_7_RedirectTarget
                    $A_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_7_VoiceCommand -eq "" )
                    {                        
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $A_7_MenuOptionPrompt
                    }
                    else
                    {
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $A_7_MenuOptionPrompt -VoiceResponses $A_7_VoiceCommand
                    }
                }
                elseif ( $A_7_Redirect -eq "TEXT" )
                {
                    $A_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_7_RedirectTarget
   
                    if ( $A_7_VoiceCommand -eq "" )
                    {                        
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $A_7_MenuOptionPrompt
                    }
                    else
                    {
                        $A_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $A_7_MenuOptionPrompt -VoiceResponses $A_7_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-7-ERROR"
                }
           }


            #
            # option 8
            #
            if ( $A_8_Redirect -ne "NONE" )
            {
                if ( $A_8_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone8
                }
                elseif ( $A_8_Redirect -eq "User" -OR $A_8_Redirect -eq "ApplicationEndpoint" -OR $A_8_Redirect -eq "ConfigurationEndpoint" -OR $A_8_Redirect -eq "SharedVoicemail" -OR $A_8_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_8_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_8_SharedVoicemailTranscription -AND $A_8_SharedVoicemailSuppress )
                        {
                            $A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_8_SharedVoicemailTranscription )
                        {
                            $A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect -EnableTranscription
                        }
                        elseif ( $A_8_SharedVoicemailSuppress )
                        {
                            $A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect
                        }
                    }
					elseif ( ( $A_8_Redirect -eq "ApplicationEndpoint" -OR $A_8_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect -CallPriority $A_8_Redirect_Call_Priority
					}
                    else
                    {
                        $A_8_Entity = New-CsAutoAttendantCallableEntity -Identity $A_8_RedirectTarget -Type $A_8_Redirect
                    }
 
                    if ( $A_8_VoiceCommand -eq "" )
                    {
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $A_8_Entity
                    }
                    else
                    {
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $A_8_Entity -VoiceResponses $A_8_VoiceCommand
                    }
                }
                elseif ( $A_8_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_8_RedirectTarget
                    $A_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
  
                    if ( $A_8_VoiceCommand -eq "" )
                    {                        
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $A_8_MenuOptionPrompt
                    }
                    else
                    {
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $A_8_MenuOptionPrompt -VoiceResponses $A_8_VoiceCommand
                    }
                }
                elseif ( $A_8_Redirect -eq "TEXT" )
                {
                    $A_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_8_RedirectTarget
   
                    if ( $A_8_VoiceCommand -eq "" )
                    {                        
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $A_8_MenuOptionPrompt
                    }
                    else
                    {
                        $A_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $A_8_MenuOptionPrompt -VoiceResponses $A_8_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-8-ERROR"
                }
            }


            #
            # option 9
            #
            if ( $A_9_Redirect -ne "NONE" )
            {
                if ( $A_9_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone9
                }
                elseif ( $A_9_Redirect -eq "User" -OR $A_9_Redirect -eq "ApplicationEndpoint" -OR $A_9_Redirect -eq "ConfigurationEndpoint" -OR $A_9_Redirect -eq "SharedVoicemail" -OR $A_9_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_9_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_9_SharedVoicemailTranscription -AND $A_9_SharedVoicemailSuppress )
                        {
                            $A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_9_SharedVoicemailTranscription )
                        {
                            $A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect -EnableTranscription
                        }
                        elseif ( $A_9_SharedVoicemailSuppress )
                        {
                            $A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect
                        }
                    }
					elseif ( ( $A_9_Redirect -eq "ApplicationEndpoint" -OR $A_9_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect -CallPriority $A_9_Redirect_Call_Priority
					}
                    else
                    {
                        $A_9_Entity = New-CsAutoAttendantCallableEntity -Identity $A_9_RedirectTarget -Type $A_9_Redirect
                    }
 
                    if ( $A_9_VoiceCommand -eq "" )
                    {
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $A_9_Entity
                    }
                    else
                    {
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $A_9_Entity -VoiceResponses $A_9_VoiceCommand
                    }
                }
                elseif ( $A_9_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_9_RedirectTarget
                    $A_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_9_VoiceCommand -eq "" )
                    {                        
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $A_9_MenuOptionPrompt
                    }
                    else
                    {
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $A_9_MenuOptionPrompt -VoiceResponses $A_9_VoiceCommand
                    }
                }
                elseif ( $A_9_Redirect -eq "TEXT" )
                {
                    $A_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_9_RedirectTarget
   
                    if ( $A_9_VoiceCommand -eq "" )
                    {                        
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $A_9_MenuOptionPrompt
                    }
                    else
                    {
                        $A_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $A_9_MenuOptionPrompt -VoiceResponses $A_9_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-9-ERROR"
                }
            }


            #
            # option Star
            #
            if ( $A_Star_Redirect -ne "NONE" )
            {
                if ( $A_Star_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse ToneStar
                }
                elseif ( $A_Star_Redirect -eq "User" -OR $A_Star_Redirect -eq "ApplicationEndpoint" -OR $A_Star_Redirect -eq "ConfigurationEndpoint" -OR $A_Star_Redirect -eq "SharedVoicemail" -OR $A_Star_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_Star_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_Star_SharedVoicemailTranscription -AND $A_Star_SharedVoicemailSuppress )
                        {
                            $A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_Star_SharedVoicemailTranscription )
                        {
                            $A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect -EnableTranscription
                        }
                        elseif ( $A_Star_SharedVoicemailSuppress )
                        {
                            $A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect
                        }
                    }
					elseif ( ( $A_Star_Redirect -eq "ApplicationEndpoint" -OR $A_Star_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect -CallPriority $A_Star_Redirect_Call_Priority
					}
                    else
                    {
                        $A_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Star_RedirectTarget -Type $A_Star_Redirect
                    }
   
                    if ( $A_Star_VoiceCommand -eq "" )
                    {
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $A_Star_Entity
                    }
                    else
                    {
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $A_Star_Entity -VoiceResponses $A_Star_VoiceCommand
                    }
                }
                elseif ( $A_Star_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_Star_RedirectTarget
                    $A_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_Star_VoiceCommand -eq "" )
                    {                        
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $A_Star_MenuOptionPrompt
                    }
                    else
                    {
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $A_Star_MenuOptionPrompt -VoiceResponses $A_Star_VoiceCommand
                    }
                }
                elseif ( $A_Star_Redirect -eq "TEXT" )
                {
                    $A_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_Star_RedirectTarget
   
                    if ( $A_Star_VoiceCommand -eq "" )
                    {                        
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $A_Star_MenuOptionPrompt
                    }
                    else
                    {
                        $A_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $A_Star_MenuOptionPrompt -VoiceResponses $A_Star_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-Star-ERROR"
                }
            }


            #
            # option Pound
            #
            if ( $A_Pound_Redirect -ne "NONE" )
            {
                if ( $A_Pound_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse TonePound
                }
                elseif ( $A_Pound_Redirect -eq "User" -OR $A_Pound_Redirect -eq "ApplicationEndpoint" -OR $A_Pound_Redirect -eq "ConfigurationEndpoint" -OR $A_Pound_Redirect -eq "SharedVoicemail" -OR $A_Pound_Redirect -eq "ExternalPSTN" )
                {
                    if ( $A_Pound_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $A_Pound_SharedVoicemailTranscription -AND $A_Pound_SharedVoicemailSuppress )
                        {
                            $A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $A_Pound_SharedVoicemailTranscription )
                        {
                            $A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect -EnableTranscription
                        }
                         elseif ( $A_Pound_SharedVoicemailSuppress )
                        {
                            $A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect
                        }
                    }
					elseif ( ( $A_Pound_Redirect -eq "ApplicationEndpoint" -OR $A_Pound_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect -CallPriority $A_Pound_Redirect_Call_Priority
					}
                    else
                    {
                        $A_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $A_Pound_RedirectTarget -Type $A_Pound_Redirect
                    }
 
                    if ( $A_Pound_VoiceCommand -eq "" )
                    {
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $A_Pound_Entity
                    }
                    else
                    {
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $A_Pound_Entity -VoiceResponses $A_Pound_VoiceCommand
                    }
                }
                elseif ( $A_Pound_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $A_Pound_RedirectTarget
                    $A_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $A_Pound_VoiceCommand -eq "" )
                    {                        
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $A_Pound_MenuOptionPrompt
                    }
                    else
                    {
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $A_Pound_MenuOptionPrompt -VoiceResponses $A_Pound_VoiceCommand
                    }
                }
                elseif ( $A_Pound_Redirect -eq "TEXT" )
                {
                    $A_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $A_Pound_RedirectTarget
   
                    if ( $A_Pound_VoiceCommand -eq "" )
                    {                        
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $A_Pound_MenuOptionPrompt
                    }
                    else
                    {
                        $A_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $A_Pound_MenuOptionPrompt -VoiceResponses $A_Pound_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "A-Pound-ERROR"
                }
            }


            #
            # build after hours menu
            #
            $afterhoursMenu = New-CsAutoAttendantMenu -Name "After Hours Menu" -MenuOptions @($A_0_MenuOption,$A_1_MenuOption,$A_2_MenuOption,$A_3_MenuOption,$A_4_MenuOption,$A_5_MenuOption,$A_6_MenuOption,$A_7_MenuOption,$A_8_MenuOption,$A_9_MenuOption,$A_Star_MenuOption,$A_Pound_MenuOption) -Prompts $A_MenuGreetingPrompt -DirectorySearchMethod $A_DirectorySearch

            if ( $A_GreetingPromptConfigured )
            {
                if ( $A_Force )
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu $afterhoursMenu -ForceListenMenuEnabled
                }
                else
                {
                     $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($A_GreetingPrompt) -Menu $afterhoursMenu
                }
            }
            else
            {
                if ( $A_Force )
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu -ForceListenMenuEnabled
                }
                else
                {
                    $afterhoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Menu $afterhoursMenu
                }
            }

            $afterhoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $BH_Schedule.Id -CallFlowId $afterHoursCallFlow.Id

            if ( $Holidays -eq "" )
			{
                $command += "-CallFlows @(`$afterhoursCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation)"
			}
        }  # Menu 
        else
        {
            Write-Host "BLANK - can't build after hours call flow.  Fix error and start again."
            exit
        }
    } # afterhours



    #
    # holidays check
    #
    if ( $Holidays -ne "" )
    {
        Write-Host "`tBuilding Holiday Call Flow"
		
		$Holiday = $Holidays -split ","

        if ( $Verbose )
        {		
		    Write-Host "`t`tHoliday Name : " $Holiday[0]
		}

		if ( $Holiday[0].Substring(0,3) -eq "[N]" )
		{
			#if ( $global:NewHolidaysProvisioned -contains $Holiday[0].SubString(4) )
			
			
			$NewHolidaysProvisionedIndex = [Array]::IndexOf($global:NewHolidaysProvisioned.ScheduleName, $Holiday[0])
			if ( $NewHolidaysProvisionedIndex -ne -1 )
			{
				# becomes an existing holiday
				$Holiday_Schedule = $global:NewHolidaysProvisioned[$NewHolidaysProvisionedIndex].ScheduleId
			}
			else
			{
				$HolidayDateTimeRange = @()
				for ( $i = 1; $i -lt $Holiday.length; $i +=2 )
				{
					if ( $Verbose )
					{
						Write-Host "`t`t`tNew-CsOnlineDateTimeRange -Start `"$($Holiday[$i])`" -End `"$($Holiday[$i+1])`""
					}
					$HolidayDateTimeRange += New-CsOnlineDateTimeRange -Start "$($Holiday[$i])" -End "$($Holiday[$i+1])"
				}
				$Holiday_Schedule = (New-CsOnlineSchedule -Name $Holiday[0].Substring(4) -FixedSchedule -DateTimeRanges @($HolidayDateTimeRange)).Id
				
				$global:NewHolidaysProvisioned += [PSCustomObject]@{ScheduleID = $Holiday_Schedule; ScheduleName = $Holiday[0]}
			}
		}
		else
		{
			# existing holiday		
			$Holiday_Schedule = $Holiday[0]
		}

        #
        #  Holiday Call Flow
        #
        switch ( $H_GreetingOption )
        {
            "FILE"  {  
                        $audioFile = AudioFileImport $H_Greeting
                        $H_GreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                        $H_GreetingPromptConfigured = $true
                    }
            "TEXT"  {  
                        $H_GreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_Greeting
                        $H_GreetingPromptConfigured = $true
                    } 
            Default {  $H_GreetingPromptConfigured = $false }
        }


        if ( $H_GreetingRouting -eq "Disconnect" )
        {
            $holidayMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $holidayMenu = New-CsAutoAttendantMenu -Name "Holiday Menu" -MenuOptions @($holidayMenuOption)

            if ( $H_GreetingPromptConfigured )
            {
                  $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu @($holidayMenu)
            }
            else
            {
                  $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu
            }
			
            if ( $H_GreetingPromptConfigured )
            {
                if ( $H_Force )
                {
                    $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu $holidayMenu -ForceListenMenuEnabled
                }
                else
                {
                     $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu $holidayMenu
                }
            }
            else
            {
                if ( $H_Force )
                {
                    $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu -ForceListenMenuEnabled
                }
                else
                {
                    $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu
                }
            }

            $holidayCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $Holiday_Schedule.Id -CallFlowId $holidayCallFlow.Id
			
			if ( ! $Hours24 )
     		{
			    $command += "-CallFlows @(`$afterhoursCallFlow, `$holidayCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation, `$holidayCallHandlingAssociation)"
		    }
		    else
		    {
			    $command += "-CallFlows @(`$holidayCallFlow) -CallHandlingAssociations @(`$holidayCallHandlingAssociation)"
		    }
        }
        elseif ( $H_GreetingRouting -eq "User" -OR $H_GreetingRouting -eq "ApplicationEndpoint" -OR $H_GreetingRouting -eq "ConfigurationEndpoint" -OR $H_GreetingRouting -eq "ExternalPSTN" )
        {

			if ( ( $H_GreetingRouting -eq "ApplicationEndpoint" -OR $H_GreetingRouting -eq "ConfigurationEndpoint" ) -AND $CallPriority )
			{
				$holidayMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $H_GreetingRoutingTarget -Type $H_GreetingRouting -CallPriority $H_GreetingRouting_Call_Priority
			}
			else
			{
				$holidayMenuOptionCallableEntity = New-CsAutoAttendantCallableEntity -Identity $H_GreetingRoutingTarget -Type $H_GreetingRouting
			}

            $holidayMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -CallTarget $holidayMenuOptionCallableEntity -DtmfResponse Automatic
            $holidayMenu = New-CsAutoAttendantMenu -Name "Holiday Menu" -MenuOptions @($holidayMenuOption)

            if ( $H_GreetingPromptConfigured )
            {
                  $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu $holidayMenu
            }
            else
            {
                  $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu
            }
			
            $holidayCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $Holiday_Schedule.Id -CallFlowId $holidayCallFlow.Id

			if ( ! $Hours24 )
     		{
			    $command += "-CallFlows @(`$afterhoursCallFlow, `$holidayCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation, `$holidayCallHandlingAssociation)"
		    }
		    else
		    {
			    $command += "-CallFlows @(`$holidayCallFlow) -CallHandlingAssociations @(`$holidayCallHandlingAssociation)"
		    }
        }
        elseif ( $H_GreetingRouting -eq "MENU" )
        {
            switch ( $H_MenuGreetingOption )
            {
                "FILE"  {  
                            $audioFile = AudioFileImport $H_MenuGreeting
                            $H_MenuGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
                            $H_GreetingPromptConfigured = $true
                        }
                "TEXT"  {  
                            $H_MenuGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_MenuGreeting
                            $H_GreetingPromptConfigured = $true
                        } 
                Default { $H_MenuGreetingPromptConfigured = $false }
            }

            #
            # Option 0
            #
            if ( $H_0_Redirect -ne "NONE" )
            {
                if ( $H_0_Redirect -eq "Operator" )
                {
                   # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                   $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0
                }
                elseif ( $H_0_Redirect -eq "User" -OR $H_0_Redirect -eq "ApplicationEndpoint" -OR $H_0_Redirect -eq "ConfigurationEndpoint" -OR $H_0_Redirect -eq "SharedVoicemail" -OR $H_0_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_0_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_0_SharedVoicemailTranscription -AND $H_0_SharedVoicemailSuppress )
                        {
                            $H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_0_SharedVoicemailTranscription )
                        {
                            $H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect -EnableTranscription
                        }
                        elseif ( $H_0_SharedVoicemailSuppress )
                        {
                            $H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect
                        }
                    }
					elseif ( ( $H_0_Redirect -eq "ApplicationEndpoint" -OR $H_0_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect -CallPriority $H_0_Redirect_Call_Priority
					}
                    else
                    {
                        $H_0_Entity = New-CsAutoAttendantCallableEntity -Identity $H_0_RedirectTarget -Type $H_0_Redirect
                    }
 
                    if ( $H_0_VoiceCommand -eq "" )
                    {
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $H_0_Entity
                    }
                    else
                    {
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone0 -CallTarget $H_0_Entity -VoiceResponses $H_0_VoiceCommand
                    }
                }
                elseif ( $H_0_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_0_RedirectTarget
                    $H_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                    if ( $H_0_VoiceCommand -eq "" )
                    {                        
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $H_0_MenuOptionPrompt
                    }
                    else
                    {
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $H_0_MenuOptionPrompt -VoiceResponses $H_0_VoiceCommand
                    }
                }
                elseif ( $H_0_Redirect -eq "TEXT" )
                {
                    $H_0_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_0_RedirectTarget

                    if ( $H_0_VoiceCommand -eq "" )
                    {                        
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $H_0_MenuOptionPrompt
                    }
                    else
                    {
                        $H_0_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone0 -Prompt $H_0_MenuOptionPrompt -VoiceResponses $H_0_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-0-ERROR"
                }
            } # H_0_Redirect


            #
            # option 1
            #
            if ( $H_1_Redirect -ne "NONE" )
            {
                if ( $H_1_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone1
                }
                elseif ( $H_1_Redirect -eq "User" -OR $H_1_Redirect -eq "ApplicationEndpoint" -OR $H_1_Redirect -eq "ConfigurationEndpoint" -OR $H_1_Redirect -eq "SharedVoicemail" -OR $H_1_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_1_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_1_SharedVoicemailTranscription -AND $H_1_SharedVoicemailSuppress )
                        {
                            $H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_1_SharedVoicemailTranscription )
                        {
                            $H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect -EnableTranscription
                        }
                        elseif ( $H_1_SharedVoicemailSuppress )
                        {
                            $H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect
                        }
                    }
					elseif ( ( $H_1_Redirect -eq "ApplicationEndpoint" -OR $H_1_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect -CallPriority $H_1_Redirect_Call_Priority
					}
                    else
                     {
                        $H_1_Entity = New-CsAutoAttendantCallableEntity -Identity $H_1_RedirectTarget -Type $H_1_Redirect
                    }
 
                    if ( $H_1_VoiceCommand -eq "" )
                    {
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $H_1_Entity
                    }
                    else
                    {
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $H_1_Entity -VoiceResponses $H_1_VoiceCommand
                    }
                }
                elseif ( $H_1_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_1_RedirectTarget
                    $H_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                    if ( $H_1_VoiceCommand -eq "" )
                    {                        
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $H_1_MenuOptionPrompt
                    }
                    else
                    {
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $H_1_MenuOptionPrompt -VoiceResponses $H_1_VoiceCommand
                    }
                }
                elseif ( $H_1_Redirect -eq "TEXT" )
                {
                    $H_1_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_1_RedirectTarget
 
                    if ( $H_1_VoiceCommand -eq "" )
                    {                        
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $H_1_MenuOptionPrompt
                    }
                    else
                    {
                        $H_1_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone1 -Prompt $H_1_MenuOptionPrompt -VoiceResponses $H_1_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-1-ERROR"
                }
            }

            #
            # option 2
            #
            if ( $H_2_Redirect -ne "NONE" )
            {
                if ( $H_2_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone2
                }
                elseif ( $H_2_Redirect -eq "User" -OR $H_2_Redirect -eq "ApplicationEndpoint" -OR $H_2_Redirect -eq "ConfigurationEndpoint" -OR $H_2_Redirect -eq "SharedVoicemail" -OR $H_2_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_2_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_2_SharedVoicemailTranscription -AND $H_2_SharedVoicemailSuppress )
                        {
                            $H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_2_SharedVoicemailTranscription )
                        {
                            $H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect -EnableTranscription
                        }
                        elseif ( $H_2_SharedVoicemailSuppress )
                        {
                            $H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect
                        }
                    }
					elseif ( ( $H_2_Redirect -eq "ApplicationEndpoint" -OR $H_2_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect -CallPriority $H_2_Redirect_Call_Priority
					}
                    else
                    {
                        $H_2_Entity = New-CsAutoAttendantCallableEntity -Identity $H_2_RedirectTarget -Type $H_2_Redirect
                    }
 
                    if ( $H_2_VoiceCommand -eq "" )
                    {
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $H_2_Entity
                    }
                    else
                    {
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $H_2_Entity -VoiceResponses $H_2_VoiceCommand
                    }
                }
                elseif ( $H_2_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_2_RedirectTarget
                    $H_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_2_VoiceCommand -eq "" )
                    {                        
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $H_2_MenuOptionPrompt
                    }
                    else
                    {
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $H_2_MenuOptionPrompt -VoiceResponses $H_2_VoiceCommand
                    }
                }
                elseif ( $H_2_Redirect -eq "TEXT" )
                {
                    $H_2_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_2_RedirectTarget
   
                    if ( $H_2_VoiceCommand -eq "" )
                    {                        
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $H_2_MenuOptionPrompt
                    }
                    else
                    {
                        $H_2_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $H_2_MenuOptionPrompt -VoiceResponses $H_2_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-2-ERROR"
                }
            }


            #
            # option 3
            #
            if ( $H_3_Redirect -ne "NONE" )
            {
                if ( $H_3_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone3
                }
                elseif ( $H_3_Redirect -eq "User" -OR $H_3_Redirect -eq "ApplicationEndpoint" -OR $H_3_Redirect -eq "ConfigurationEndpoint" -OR $H_3_Redirect -eq "SharedVoicemail" -OR $H_3_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_3_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_3_SharedVoicemailTranscription -AND $H_3_SharedVoicemailSuppress )
                        {
                            $H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_3_SharedVoicemailTranscription )
                        {
                            $H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect -EnableTranscription
                        }
                        elseif ( $H_3_SharedVoicemailSuppress )
                        {
                            $H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect
                        }
                    }
					elseif ( ( $H_3_Redirect -eq "ApplicationEndpoint" -OR $H_3_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect -CallPriority $H_3_Redirect_Call_Priority
					}
                    else
                    {
                        $H_3_Entity = New-CsAutoAttendantCallableEntity -Identity $H_3_RedirectTarget -Type $H_3_Redirect
                    }
   
                    if ( $H_3_VoiceCommand -eq "" )
                    {
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $H_3_Entity
                    }
                    else
                    {
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone3 -CallTarget $H_3_Entity -VoiceResponses $H_3_VoiceCommand
                    }
                }
                elseif ( $H_3_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_3_RedirectTarget
                    $H_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_3_VoiceCommand -eq "" )
                    {                        
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $H_3_MenuOptionPrompt
                    }
                    else
                    {
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $H_3_MenuOptionPrompt -VoiceResponses $H_3_VoiceCommand
                    }
                }
                elseif ( $H_3_Redirect -eq "TEXT" )
                {
                    $H_3_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_3_RedirectTarget
   
                    if ( $H_3_VoiceCommand -eq "" )
                    {                        
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $H_3_MenuOptionPrompt
                    }
                    else
                    {
                        $H_3_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $H_3_MenuOptionPrompt -VoiceResponses $H_3_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-3-ERROR"
                }
            }


            #
            # option 4
            #
            if ( $H_4_Redirect -ne "NONE" )
            {
                if ( $H_4_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone4
                }
                elseif ( $H_4_Redirect -eq "User" -OR $H_4_Redirect -eq "ApplicationEndpoint" -OR $H_4_Redirect -eq "ConfigurationEndpoint" -OR $H_4_Redirect -eq "SharedVoicemail" -OR $H_4_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_4_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_4_SharedVoicemailTranscription -AND $H_4_SharedVoicemailSuppress )
                        {
                            $H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_4_SharedVoicemailTranscription )
                        {
                            $H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect -EnableTranscription
                        }
                        elseif ( $H_4_SharedVoicemailSuppress )
                        {
                            $H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect
                        }
                    }
					elseif ( ( $H_4_Redirect -eq "ApplicationEndpoint" -OR $H_4_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect -CallPriority $H_4_Redirect_Call_Priority
					}
                    else
                    {
                        $H_4_Entity = New-CsAutoAttendantCallableEntity -Identity $H_4_RedirectTarget -Type $H_4_Redirect
                    }
 
                    if ( $H_4_VoiceCommand -eq "" )
                    {
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $H_4_Entity
                    }
                    else
                    {
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone4 -CallTarget $H_4_Entity -VoiceResponses $H_4_VoiceCommand
                    }
                }
                elseif ( $H_4_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_4_RedirectTarget
                    $H_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_4_VoiceCommand -eq "" )
                    {                        
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $H_4_MenuOptionPrompt
                    }
                    else
                    {
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $H_4_MenuOptionPrompt -VoiceResponses $H_4_VoiceCommand
                    }
                }
                elseif ( $H_4_Redirect -eq "TEXT" )
                {
                    $H_4_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_4_RedirectTarget
   
                    if ( $H_4_VoiceCommand -eq "" )
                    {                        
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $H_4_MenuOptionPrompt
                    }
                    else
                    {
                        $H_4_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $H_4_MenuOptionPrompt -VoiceResponses $H_4_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-4-ERROR"
                }
            }


            #
            # option 5
            #
            if ( $H_5_Redirect -ne "NONE" )
            {
                if ( $H_5_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone5
                }
                elseif ( $H_5_Redirect -eq "User" -OR $H_5_Redirect -eq "ApplicationEndpoint" -OR $H_5_Redirect -eq "ConfigurationEndpoint" -OR $H_5_Redirect -eq "SharedVoicemail" -OR $H_5_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_5_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_5_SharedVoicemailTranscription -AND $H_5_SharedVoicemailSuppress )
                        {
                            $H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_5_SharedVoicemailTranscription )
                        {
                            $H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect -EnableTranscription
                        }
                        elseif ( $H_5_SharedVoicemailSuppress )
                        {
                            $H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect
                        }
                    }
					elseif ( ( $H_5_Redirect -eq "ApplicationEndpoint" -OR $H_5_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect -CallPriority $H_5_Redirect_Call_Priority
					}
                    else
                    {
                        $H_5_Entity = New-CsAutoAttendantCallableEntity -Identity $H_5_RedirectTarget -Type $H_5_Redirect
                    }
 
                    if ( $H_5_VoiceCommand -eq "" )
                    {
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $H_5_Entity
                    }
                    else
                    {
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone5 -CallTarget $H_5_Entity -VoiceResponses $H_5_VoiceCommand
                    }
                }
                elseif ( $H_5_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_5_RedirectTarget
                    $H_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
 
                    if ( $H_5_VoiceCommand -eq "" )
                    {                        
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $H_5_MenuOptionPrompt
                    }
                    else
                    {
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $H_5_MenuOptionPrompt -VoiceResponses $H_5_VoiceCommand
                    }
                }
                elseif ( $H_5_Redirect -eq "TEXT" )
                {
                    $H_5_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_5_RedirectTarget
 
                    if ( $H_5_VoiceCommand -eq "" )
                    {                        
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $H_5_MenuOptionPrompt
                    }
                    else
                    {
                        $H_5_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone5 -Prompt $H_5_MenuOptionPrompt -VoiceResponses $H_5_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-5-ERROR"
                }
            }


            #
            # option 6
            #
            if ( $H_6_Redirect -ne "NONE" )
            {
                if ( $H_6_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone6
                }
                elseif ( $H_6_Redirect -eq "User" -OR $H_6_Redirect -eq "ApplicationEndpoint" -OR $H_6_Redirect -eq "ConfigurationEndpoint" -OR $H_6_Redirect -eq "SharedVoicemail" -OR $H_6_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_6_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_6_SharedVoicemailTranscription -AND $H_6_SharedVoicemailSuppress )
                        {
                            $H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_6_SharedVoicemailTranscription )
                        {
                             $H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect -EnableTranscription
                        }
                        elseif ( $H_6_SharedVoicemailSuppress )
                        {
                            $H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect
                        }
                    }
					elseif ( ( $H_6_Redirect -eq "ApplicationEndpoint" -OR $H_6_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect -CallPriority $H_6_Redirect_Call_Priority
					}
                    else
                    {
                        $H_6_Entity = New-CsAutoAttendantCallableEntity -Identity $H_6_RedirectTarget -Type $H_6_Redirect
                    }
 
                    if ( $H_6_VoiceCommand -eq "" )
                    {
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $H_6_Entity
                    }
                    else
                    {
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone6 -CallTarget $H_6_Entity -VoiceResponses $H_6_VoiceCommand
                    }
                }
                elseif ( $H_6_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_6_RedirectTarget
                    $H_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile

                    if ( $H_6_VoiceCommand -eq "" )
                    {                        
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $H_6_MenuOptionPrompt
                    }
                    else
                    {
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $H_6_MenuOptionPrompt -VoiceResponses $H_6_VoiceCommand
                    }
                }
                elseif ( $H_6_Redirect -eq "TEXT" )
                {
                    $H_6_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_6_RedirectTarget
   
                    if ( $H_6_VoiceCommand -eq "" )
                    {                        
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $H_6_MenuOptionPrompt
                    }
                    else
                    {
                        $H_6_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone6 -Prompt $H_6_MenuOptionPrompt -VoiceResponses $H_6_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-6-ERROR"
                }
            }


            #
            # option 7
            #
            if ( $H_7_Redirect -ne "NONE" )
            {
                if ( $H_7_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone7
                }
                elseif ( $H_7_Redirect -eq "User" -OR $H_7_Redirect -eq "ApplicationEndpoint" -OR $H_7_Redirect -eq "ConfigurationEndpoint" -OR $H_7_Redirect -eq "SharedVoicemail" -OR $H_7_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_7_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_7_SharedVoicemailTranscription -AND $H_7_SharedVoicemailSuppress )
                        {
                            $H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_7_SharedVoicemailTranscription )
                        {
                            $H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect -EnableTranscription
                        }
                        elseif ( $H_7_SharedVoicemailSuppress )
                        {
                            $H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect
                        }
                    }
					elseif ( ( $H_7_Redirect -eq "ApplicationEndpoint" -OR $H_7_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect -CallPriority $H_7_Redirect_Call_Priority
					}
                    else
                    {
                        $H_7_Entity = New-CsAutoAttendantCallableEntity -Identity $H_7_RedirectTarget -Type $H_7_Redirect
                    }

                    if ( $H_7_VoiceCommand -eq "" )
                    {
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $H_7_Entity
                    }
                    else
                    {
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone7 -CallTarget $H_7_Entity -VoiceResponses $H_7_VoiceCommand
                    }
                }
                elseif ( $H_7_Redirect -eq "FILE" )
                 {
                    $audioFile = AudioFileImport $H_7_RedirectTarget
                    $H_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_7_VoiceCommand -eq "" )
                    {                        
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $H_7_MenuOptionPrompt
                    }
                    else
                    {
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $H_7_MenuOptionPrompt -VoiceResponses $H_7_VoiceCommand
                    }
                }
                elseif ( $H_7_Redirect -eq "TEXT" )
                {
                    $H_7_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_7_RedirectTarget
   
                    if ( $H_7_VoiceCommand -eq "" )
                    {                        
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $H_7_MenuOptionPrompt
                    }
                    else
                    {
                        $H_7_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone7 -Prompt $H_7_MenuOptionPrompt -VoiceResponses $H_7_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-7-ERROR"
                }
           }


            #
            # option 8
            #
            if ( $H_8_Redirect -ne "NONE" )
            {
                if ( $H_8_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone8
                }
                elseif ( $H_8_Redirect -eq "User" -OR $H_8_Redirect -eq "ApplicationEndpoint" -OR $H_8_Redirect -eq "ConfigurationEndpoint" -OR $H_8_Redirect -eq "SharedVoicemail" -OR $H_8_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_8_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_8_SharedVoicemailTranscription -AND $H_8_SharedVoicemailSuppress )
                        {
                            $H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_8_SharedVoicemailTranscription )
                        {
                            $H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect -EnableTranscription
                        }
                        elseif ( $H_8_SharedVoicemailSuppress )
                        {
                            $H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect
                        }
                    }
					elseif ( ( $H_8_Redirect -eq "ApplicationEndpoint" -OR $H_8_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect -CallPriority $H_8_Redirect_Call_Priority
					}
                    else
                    {
                        $H_8_Entity = New-CsAutoAttendantCallableEntity -Identity $H_8_RedirectTarget -Type $H_8_Redirect
                    }
 
                    if ( $H_8_VoiceCommand -eq "" )
                    {
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $H_8_Entity
                    }
                    else
                    {
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone8 -CallTarget $H_8_Entity -VoiceResponses $H_8_VoiceCommand
                    }
                }
                elseif ( $H_8_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_8_RedirectTarget
                    $H_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
  
                    if ( $H_8_VoiceCommand -eq "" )
                    {                        
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $H_8_MenuOptionPrompt
                    }
                    else
                    {
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $H_8_MenuOptionPrompt -VoiceResponses $H_8_VoiceCommand
                    }
                }
                elseif ( $H_8_Redirect -eq "TEXT" )
                {
                    $H_8_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_8_RedirectTarget
   
                    if ( $H_8_VoiceCommand -eq "" )
                    {                        
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $H_8_MenuOptionPrompt
                    }
                    else
                    {
                        $H_8_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone8 -Prompt $H_8_MenuOptionPrompt -VoiceResponses $H_8_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-8-ERROR"
                }
            }


            #
            # option 9
            #
            if ( $H_9_Redirect -ne "NONE" )
            {
                if ( $H_9_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone9
                }
                elseif ( $H_9_Redirect -eq "User" -OR $H_9_Redirect -eq "ApplicationEndpoint" -OR $H_9_Redirect -eq "ConfigurationEndpoint" -OR $H_9_Redirect -eq "SharedVoicemail" -OR $H_9_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_9_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_9_SharedVoicemailTranscription -AND $H_9_SharedVoicemailSuppress )
                        {
                            $H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_9_SharedVoicemailTranscription )
                        {
                            $H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect -EnableTranscription
                        }
                        elseif ( $H_9_SharedVoicemailSuppress )
                        {
                            $H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect
                        }
                    }
					elseif ( ( $H_9_Redirect -eq "ApplicationEndpoint" -OR $H_9_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect -CallPriority $H_9_Redirect_Call_Priority
					}
                    else
                    {
                        $H_9_Entity = New-CsAutoAttendantCallableEntity -Identity $H_9_RedirectTarget -Type $H_9_Redirect
                    }
 
                    if ( $H_9_VoiceCommand -eq "" )
                    {
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $H_9_Entity
                    }
                    else
                    {
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $H_9_Entity -VoiceResponses $H_9_VoiceCommand
                    }
                }
                elseif ( $H_9_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_9_RedirectTarget
                    $H_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_9_VoiceCommand -eq "" )
                    {                        
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $H_9_MenuOptionPrompt
                    }
                    else
                    {
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $H_9_MenuOptionPrompt -VoiceResponses $H_9_VoiceCommand
                    }
                }
                elseif ( $H_9_Redirect -eq "TEXT" )
                {
                    $H_9_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_9_RedirectTarget
   
                    if ( $H_9_VoiceCommand -eq "" )
                    {                        
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $H_9_MenuOptionPrompt
                    }
                    else
                    {
                        $H_9_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone9 -Prompt $H_9_MenuOptionPrompt -VoiceResponses $H_9_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-9-ERROR"
                }
            }


            #
            # option Star
            #
            if ( $H_Star_Redirect -ne "NONE" )
            {
                if ( $H_Star_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse ToneStar
                }
                elseif ( $H_Star_Redirect -eq "User" -OR $H_Star_Redirect -eq "ApplicationEndpoint" -OR $H_Star_Redirect -eq "ConfigurationEndpoint" -OR $H_Star_Redirect -eq "SharedVoicemail" -OR $H_Star_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_Star_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_Star_SharedVoicemailTranscription -AND $H_Star_SharedVoicemailSuppress )
                        {
                            $H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_Star_SharedVoicemailTranscription )
                        {
                            $H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect -EnableTranscription
                        }
                        elseif ( $H_Star_SharedVoicemailSuppress )
                        {
                            $H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect
                        }
                    }
					elseif ( ( $H_Star_Redirect -eq "ApplicationEndpoint" -OR $H_Star_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect -CallPriority $H_Star_Redirect_Call_Priority
					}
                    else
                    {
                        $H_Star_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Star_RedirectTarget -Type $H_Star_Redirect
                    }
   
                    if ( $H_Star_VoiceCommand -eq "" )
                    {
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $H_Star_Entity
                    }
                    else
                    {
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse ToneStar -CallTarget $H_Star_Entity -VoiceResponses $H_Star_VoiceCommand
                    }
                }
                elseif ( $H_Star_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_Star_RedirectTarget
                    $H_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_Star_VoiceCommand -eq "" )
                    {                        
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $H_Star_MenuOptionPrompt
                    }
                    else
                    {
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $H_Star_MenuOptionPrompt -VoiceResponses $H_Star_VoiceCommand
                    }
                }
                elseif ( $H_Star_Redirect -eq "TEXT" )
                {
                    $H_Star_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_Star_RedirectTarget
   
                    if ( $H_Star_VoiceCommand -eq "" )
                    {                        
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $H_Star_MenuOptionPrompt
                    }
                    else
                    {
                        $H_Star_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse ToneStar -Prompt $H_Star_MenuOptionPrompt -VoiceResponses $H_Star_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-Star-ERROR"
                }
            }


            #
            # option Pound
            #
            if ( $H_Pound_Redirect -ne "NONE" )
            {
                if ( $H_Pound_Redirect -eq "Operator" )
                {
                    # don't need VoiceResponses option as "Operator" will be set automatically if voice response is enabled
                    $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse TonePound
                }
                elseif ( $H_Pound_Redirect -eq "User" -OR $H_Pound_Redirect -eq "ApplicationEndpoint" -OR $H_Pound_Redirect -eq "ConfigurationEndpoint" -OR $H_Pound_Redirect -eq "SharedVoicemail" -OR $H_Pound_Redirect -eq "ExternalPSTN" )
                {
                    if ( $H_Pound_Redirect -eq "SharedVoicemail" )
                    {
                        if ( $H_Pound_SharedVoicemailTranscription -AND $H_Pound_SharedVoicemailSuppress )
                        {
                            $H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect -EnableTranscription -EnableSharedVoicemailSystemPromptSuppression
                        }
                        elseif ( $H_Pound_SharedVoicemailTranscription )
                        {
                            $H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect -EnableTranscription
                        }
                         elseif ( $H_Pound_SharedVoicemailSuppress )
                        {
                            $H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect -EnableSharedVoicemailSystemPromptSuppression
                        }
                        else
                        {
                            $H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect
                        }
                    }
					elseif ( ( $H_Pound_Redirect -eq "ApplicationEndpoint" -OR $H_Pound_Redirect -eq "ConfigurationEndpoint" ) -AND $CallPriority )
					{
						$H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect -CallPriority $H_Pound_Redirect_Call_Priority
					}
                    else
                    {
                        $H_Pound_Entity = New-CsAutoAttendantCallableEntity -Identity $H_Pound_RedirectTarget -Type $H_Pound_Redirect
                    }
 
                    if ( $H_Pound_VoiceCommand -eq "" )
                    {
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $H_Pound_Entity
                    }
                    else
                    {
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse TonePound -CallTarget $H_Pound_Entity -VoiceResponses $H_Pound_VoiceCommand
                    }
                }
                elseif ( $H_Pound_Redirect -eq "FILE" )
                {
                    $audioFile = AudioFileImport $H_Pound_RedirectTarget
                    $H_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $audioFile
   
                    if ( $H_Pound_VoiceCommand -eq "" )
                    {                        
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $H_Pound_MenuOptionPrompt
                    }
                    else
                    {
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $H_Pound_MenuOptionPrompt -VoiceResponses $H_Pound_VoiceCommand
                    }
                }
                elseif ( $H_Pound_Redirect -eq "TEXT" )
                {
                    $H_Pound_MenuOptionPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $H_Pound_RedirectTarget
   
                    if ( $H_Pound_VoiceCommand -eq "" )
                    {                        
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $H_Pound_MenuOptionPrompt
                    }
                    else
                    {
                        $H_Pound_MenuOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse TonePound -Prompt $H_Pound_MenuOptionPrompt -VoiceResponses $H_Pound_VoiceCommand
                    }
                }
                else
                {
                    Write-Host "H-Pound-ERROR"
                }
            }


            #
            # build holiday menu
            #
            $holidayMenu = New-CsAutoAttendantMenu -Name "Holiday Menu" -MenuOptions @($H_0_MenuOption,$H_1_MenuOption,$H_2_MenuOption,$H_3_MenuOption,$H_4_MenuOption,$H_5_MenuOption,$H_6_MenuOption,$H_7_MenuOption,$H_8_MenuOption,$H_9_MenuOption,$H_Star_MenuOption,$H_Pound_MenuOption) -Prompts $H_MenuGreetingPrompt -DirectorySearchMethod $H_DirectorySearch

            if ( $H_GreetingPromptConfigured )
            {
				if ( $H_MenuGreetingOption -eq "TEXT" )
				{
					if ( $H_Force )
					{
						$holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @("""$H_GreetingPrompt""") -Menu $holidayMenu -ForceListenMenuEnabled
					}
					else
					{
						$holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @("""$H_GreetingPrompt""") -Menu $holidayMenu
					}
				}
				else
				{
					if ( $H_Force )
					{
						$holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu $holidayMenu -ForceListenMenuEnabled
					}
					else
					{
						$holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Greetings @($H_GreetingPrompt) -Menu $holidayMenu
					}
				}					
            }
            else
            {
                if ( $H_Force )
                {
                    $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu -ForceListenMenuEnabled
                }
                else
                {
                    $holidayCallFlow = New-CsAutoAttendantCallFlow -Name "Holiday Call Flow" -Menu $holidayMenu
                }
            }

            $holidayCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $Holiday_Schedule -CallFlowId $holidayCallFlow.Id
		
		    if ( ! $Hours24 )
		    {
			    $command += "-CallFlows @(`$afterhoursCallFlow, `$holidayCallFlow) -CallHandlingAssociations @(`$afterhoursCallHandlingAssociation, `$holidayCallHandlingAssociation)"
		    }
		    else
		    {
			    $command += "-CallFlows @(`$holidayCallFlow) -CallHandlingAssociations @(`$holidayCallHandlingAssociation)"
			}
		}
    }


    if ( $Verbose )
    {
       Write-Host "-----------------------------------------------------------------"
       Write-Host $command
       Write-Host "-----------------------------------------------------------------"
    }

    Write-Host "`tCreating Auto Attendant : $Name"

    $command += "| Out-Null"

    Invoke-Expression $command

    #
    #  Assign Resource Account
    #
    if ( ! $NoResourceAccounts )
    {
        $AutoAttendant = @(Get-CsAutoAttendant -NameFilter "$Name")
		
		if ( $AutoAttendant.length -eq 1 )
		{
            if ( $ResourceAccount -eq "Existing" )
	   	    {
                if ( $ExistingResourceAccountName -ne "" ) 
                {
                    Write-Host "`tAssigning Resource Account to $Name"
                    New-CsOnlineApplicationInstanceAssociation -Identities @($ExistingResourceAccountName) -ConfigurationID $AutoAttendant.Identity -ConfigurationType AutoAttendant | Out-Null
                }
	        }
            elseif ( $ResourceAccount -eq "New" )
            {
				if ( ! $NoResourceAccountCreation )
			    {
                    if ( $NewResourceAccountPrincipalName -ne "" )
                    {
                        if ( $NewResourceAccountDisplayName -eq "" )
                        {
                            $NewResourceAccountDisplayName = $Name
                        }

                        Write-Host "`tCreating Resource Account ($NewResourceAccountPrincipalName)"
                        New-CsOnlineApplicationInstance -UserPrincipalName $NewResourceAccountPrincipalName -DisplayName $NewResourceAccountDisplayName -ApplicationID "ce933385-9390-45d1-9512-c8d228074e07" | Out-Null

                        $j = 1
						do
                        {
                            Write-Host -NoNewLine "`t`tResource Account created, pausing for 10 seconds to allow sync to occur.(Attempt $j of 10) "
							for ( $i = 0; $i -lt 10; $i++)
							{
								Write-Host -NoNewLine "."
								Start-Sleep -Seconds 1
							}
							$j++

                            $applicationInstanceID = (Get-CsOnlineUser -Identity $NewResourceAccountPrincipalName).Identity 2> `$null
                        }
                        until ( $applicationInstanceID.length -gt 0 -OR $j -gt 10 )
						Write-Host " "

						if ( Test-Path -Path ".\`$null" )
						{
							Remove-Item -Path ".\`$null" | Out-Null
						}

                        if ( $NewResourceAccountLocation -eq "" )
                        {
                            $NewResourceAccountLocation = "US"
                        }

                        Write-Host "`tAssigning Location ($NewResourceAccountLocation)"
                        Update-MgUser -UserId $NewResourceAccountPrincipalName -Id $applicationInstanceID -UsageLocation $NewResourceAccountLocation
 
                        if ( ! $NoResourceAccountLicensing )
						{
                            Write-Host "`tAssigning license"
                            $skuID = (Get-MgSubscribedSKU | Where {$_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER"}).SkuId
                            Set-MgUserLicense -UserId $applicationInstanceID -AddLicenses @{SkuId = $skuID} -RemoveLicenses @() | Out-Null

							if ( $NewResourceAccountPhoneNumber -ne "" )
							{							
								$j = 1
								do
								{
									Write-Host -NoNewLine "`t`tResource Account licensed, pausing for 10 seconds to allow sync to occur.(Attempt $j of 10) "
									for ( $i = 0; $i -lt 10; $i++)
									{
										Write-Host -NoNewLine "."
										Start-Sleep -Seconds 1
									}
									Write-Host " "
									$j++

									$ProvisionedPlan = @((Get-CsOnlineuser -Identity $NewResourceAccountPrincipalName).ProvisionedPlan)
								}
								until ($ProvisionedPlan.Capability -contains "MCOEV_VIRTUALUSER" -OR $j -gt 10 )
								
								if( $j -lt 10 )
								{
									if ( ! $NoResourceAccountPhoneNumbers )
									{
										Write-Host "`tAssigning Phone number ($NewResourceAccountPhoneNumber)"
										Set-CsPhoneNumberAssignment -Identity $NewResourceAccountPrincipalName -PhoneNumber $NewResourceAccountPhoneNumber -PhoneNumberType CallingPlan
									}
									else
									{
										Write-Host "`tResource Account Phone Number Assignment is disabled"
									}
								}
								else
								{
									Write-Host "`tUnable to assign phone number ($NewResourceAccountPhoneNumber) - couldn't confirm licensing"
								}
							}
						}
						else
						{
							Write-Host "`tResource Account Licensing is disabled"
						}

                        Write-Host "`tAssigning Resource Account"
                        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceID) -ConfigurationID $AutoAttendant.Identity -ConfigurationType AutoAttendant | Out-Null						
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
		    else
		    {
			    Write-Host "`tInvalid Resource account option - skipping"
		    }
        }
        else
        {
            Write-Host "`tUnable to process Resource Account configuration for $Name Auto Attendant as more than one Auto Attendant with that name exists."
        }    
	}
    else
    {
        Write-Host "`tResource Account Processing is disabled."
    }    

    Write-Host "`tAuto Attendant Created"
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
		"-excelfile"                 		{ $ExcelFilename = $args[$i+1]
											  $i++
											}
		"-help"                      		{ $Help = $true }
		"-noresourceaccounts"        		{ $NoResourceAccounts = $true }
		"-noresourceaccountcreation" 		{ $NoResourceAccountCreation = $true }
		"-noresourceaccountlicensing"	 	{ $NoResourceAccountLicensing = $true }
		"-noresourceaccountphonenumbers"	{ $NoResourceAccountPhoneNumbers = $true }
		"-verbose"              	      	{ $Verbose = $true }
		Default      						{ $ArgError = $true
											  $arg = $args[$i]
											  Write-Host "Unknown argument passed: $arg" }   
	}
}

if ( $ArgError )
{
	Write-Host "An unknown argument was encountered. Processing has been halted." -f Red
	Write-Host ""
}

if ( ( $Help ) -or ( $ArgError ) )
{
	Write-Host "The following options are avaialble:"
	Write-Host "`t-ExcelFile - specify an alternative Excel Spreadsheet to use.  Default is BulkAAs.xlsm"
	Write-Host "`t-Help - shows the options that are available (this help message)"
	Write-Host "`t-NoResourceAccounts - don't perform any Resource Account related steps" 
	Write-Host "`t-NoResourceAccountCreation - don't create or license new Resource Accounts"
	Write-Host "`t-NoResourceAccountLicensing - don't assign a license to new Resource Accounts"
    Write-Host "`t-NoResourceAccountPhoneNumbers - don't assign a phone number to a new Resource Accounts --- NOT AVAILABLE"
	Write-Host "`t-Verbose - provides extra messaging during the process"

	exit
}

Write-Host "Starting BulkAAsConfig."
Write-Host "Cleaning up from any previous runs."

if ( Test-Path -Path ".\PS-AA.csv" )
{
   Remove-Item -Path ".\PS-AA.csv" | Out-Null
}

#
# Increase maximum variable and function count (function count for ImportExcel)
#
$MaximumVariableCount = 10000
$MaximumFunctionCount = 32768


#
# Check that required modules are installed - install if not
#
Write-Host "Checking for MicrosoftTeams module 6.7.0 or later."
$Version = ( (Get-InstalledModule -Name MicrosoftTeams -MinimumVersion "6.7.0").Version 2> $null )
if ( ( $Version.Major -ge 6 ) -and ( $Version.minor -ge 7 ) )
{
   Write-Host "Connecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion 6.7.0
   
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
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}
else
{
   Write-Host "Module MicrosoftTeams does not exist - installing."
   Install-Module -Name MicrosoftTeams -MinimumVersion 6.7.0 -Force -AllowClobber

   Write-Host "Connecting to Microsoft Teams."
   Import-Module -Name MicrosoftTeams -MinimumVersion 6.7.0
   
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
      Write-Error "Not signed into Microsoft Teams!" 
      exit
   }
   Write-Host "Connected to Microsoft Teams."
}

#
# Microsoft Graph
#
Write-Host "Checking for Microsoft.Graph module 2.24.0 or later."
$Version = ( (get-installedmodule -Name Microsoft.Graph -MinimumVersion "2.24.0").Version 2> $null)
if ( ( $Version.Major -ge 2 ) -and ( $Version.minor -ge 24 ) )
{
   Write-Host "Connecting to Microsoft Graph."
   
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
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}
else
{
   Write-Host "Module Microsoft.Graph does not exist - installing."
   Install-Module -Name Microsoft.Graph -MinimumVersion 2.24.0 -Force -AllowClobber

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
      Write-Error "Not signed into Microsoft Graph!" 
      exit
   }
   Write-Host "Connected to Microsoft Graph."
}

#
# ImportExcel
#
Write-Host "Checking for ImportExcel module 7.8.0 or later."
$Version = ( (get-installedmodule -Name ImportExcel -MinimumVersion "7.8.0").Version 2> $null )
if ( ( $Version.Major -ge 7 ) -and ( $Version.minor -ge 8 ) )
{
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion 7.8.0
}
else
{
   Write-Host "Module ImportExcel - installing."
   Install-Module -Name ImportExcel -MinimumVersion 7.8.0 -Force -AllowClobber
   
   Write-Host "Importing ImportExcel."
   Import-Module -Name ImportExcel -MinimumVersion 7.8.0
}

#
# setup filename
#
if ( $ExcelFilename -eq $null )
{
   $ExcelFilename = "BulkAAs.xlsm"
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
# Auto Attendant configuration
#
Write-Host "Starting Auto Attendant Configuration."
$ExcelWorkSheetName = "PS-AA"
$ExcelCSVFilename = ".\PS-AA.csv"

# Import the specific tab from the Excel file
$data = Import-Excel -Path $ExcelFullPathFilename -WorksheetName $ExcelWorkSheetName
 
# Export the data to a CSV file
$data | Export-Csv -Path $ExcelCSVFilename -NoTypeInformation

# track new holiday provisioning
$global:NewHolidaysProvisioned = @()
$global:NewHolidaysProvisioned += [PSCustomObject]@{ScheduleID = "0"; ScheduleName = "0"}

# track phone number assignments to resource Accounts
$PhoneNumbersAssignedToResourceAccounts = @()

$PSAAConfig = @(Import-csv -Path $ExcelCSVFilename)

for ( $i = 0; $i -lt $PSAAConfig.length; $i++)
{
	$StopProcessing = $false
	$VerboseStopProcessing = $false

	$NewResourceAccountPhoneNumber_Comment = $null
	
    $Action = $PSAAConfig.Action[$i]
    if ( $Action -eq "New" )
    {
        # $Name   = '"' + $PSAAConfig.Name[$i] + '"'
        $Name   = $PSAAConfig.Name[$i]

        $ResourceAccount = $PSAAConfig.ResourceAccount[$i]
        $ExistingResourceAccountName = $PSAAConfig.ExistingResourceAccountName[$i]
        $NewResourceAccountPrincipalName = $PSAAConfig.NewResourceAccountPrincipalName[$i]
        $NewResourceAccountDisplayName = $PSAAConfig.NewResourceAccountDisplayName[$i]
        $NewResourceAccountLocation = $PSAAConfig.NewResourceAccountLocation[$i]
		$NewResourceAccountPhoneNumber = $PSAAConfig.NewResourceAccountPhoneNumber[$i]
		
		if ( $NewResourceAccountPhoneNumber -ne "ERROR" )
		{
			if ( $PhoneNumbersAssignedToResourceAccounts -contains $NewResourceAccountPhoneNumber )
			{
				$NewResourceAccountPhoneNumber_Comment = "ERROR: Phone number has been previously assigned"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$PhoneNumbersAssignedToResourceAccounts += $NewResourceAccountPhoneNumber
			}
		}

        $Operator = $PSAAConfig.Operator[$i]
        $OperatorTarget = $PSAAConfig.OperatorTarget[$i]
		
		if ( $Operator -match "ApplicationEndpoint" )
        {
			if ( $Operator.length -gt 19 )
			{
				$Operator_Call_Priority = $Operator.Substring(20,1)
				$Operator = $Operator.Substring(0,19)
				
				$CallPriority = $true
			}

			if ( $OperatorTarget -ne "ERROR" )
			{
				switch ( $OperatorTarget.Substring(0,4) )
				{
					"[AA]"	{
								$Operator = "ConfigurationEndpoint"
							}
					"[CQ]"	{
								$Operator = "ConfigurationEndpoint"
							}						
				}
				$OperatorTarget = $OperatorTarget.Substring(5)
			}
        }
						
		if ( $OperatorTarget -eq "ERROR" )
		{
			$Operator_Comment = "ERROR: Check OperatorTarget on Config-Base"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        $TimeZone = $PSAAConfig.TimeZone[$i]

        $Language = $PSAAConfig.Language[$i]
		if ( $Language.length -gt 5 )
		{
			$LanguageGender = $true
			
			if ( $Language.Substring(6,1) -eq "M" )
			{
				$LanguageGenderId = "Male"
			}
			else
			{
				$LanguageGenderId = "Female"
			}
			
			$Language = $Language.Substring(0,5)
		}
		
        $VoiceInputs = $PSAAConfig.VoiceInputs[$i]

        if ( $PSAAConfig.Hours24[$i] -eq "false" )
        {
            $Hours24 = $false
        }
        else
        {
            $Hours24 = $true
        }

        $BusinessHours = $PSAAConfig.BusinessHours[$i]
        $Holidays = $PSAAConfig.Holidays[$i]

 
		$B_DirectorySearch = $PSAAConfig.B_DirectorySearch[$i]
        $B_DialScopeInclude = $PSAAConfig.B_DialScopeInclude[$i]
		$B_DialScopeExclude = $PSAAConfig.B_DialScopeExclude[$i]
		
		if ( $B_DialScopeInclude -eq "ERROR" )
		{
			$B_DialScopeInclude_Comment = "ERROR: Check B_DialScopeInclude on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}
			
		if ( $B_DialScopeExclude -eq "ERROR" )
		{
			$B_DialScopeExclude_Comment = "ERROR: Check B_DialScopeExclude on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}
				


        #
        # Business Hours (default) Menu
        #
		$B_Operator_Redirect = $false
				
        if ( $PSAAConfig.B_Force[$i] -eq "true" )
        {
            $B_Force = $true
        }
        else
        {
            $B_Force = $false
        }

        $B_GreetingOption = $PSAAConfig.B_GreetingOption[$i]
        $B_Greeting = $PSAAConfig.B_Greeting[$i]
		
		if ( $B_GreetingOption -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_Greeting
			
			if ( ! $FileExists )
			{
				$B_Greeting_Comment = "ERROR: $B_Greeting does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}
		
        $B_GreetingRouting = $PSAAConfig.B_GreetingRouting[$i]
        $B_GreetingRoutingTarget = $PSAAConfig.B_GreetingRoutingTarget[$i]

		if ( $B_GreetingRouting -match "ApplicationEndpoint" )
        {
			if ( $B_GreetingRouting.length -gt 19 )
			{
				$B_GreetingRouting_Call_Priority = $B_GreetingRouting.Substring(20,1)
				$B_GreetingRouting = $B_GreetingRouting.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_GreetingRouting -match "ConfigurationEndpoint" )
        {
			if ( $B_GreetingRouting.length -gt 21 )
			{
				$B_GreetingRouting_Call_Priority = $B_GreetingRouting.Substring(22,1)
				$B_GreetingRouting = $B_GreetingRouting.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_GreetingRoutingTarget -eq "ERROR" )
		{
			$B_GreetingRouting_Comment = "ERROR: Check B_GreetingRoutingTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        $B_MenuGreetingOption = $PSAAConfig.B_MenuGreetingOption[$i]
        $B_MenuGreeting = $PSAAConfig.B_MenuGreeting[$i]


		# Business Hours - Option 0
        $B_0_VoiceCommand = $PSAAConfig.B_0_VoiceCommand[$i]
        $B_0_Redirect = $PSAAConfig.B_0_Redirect[$i]
		$B_0_RedirectTarget = $PSAAConfig.B_0_RedirectTarget[$i]
		
		# handle SharedVoicemail:n1:n2
		# n1={0,1} - SharedVoicemailTranscription On/Off
		# n2={0,1} - SharedVoicemailSuppress On/Off
        if ( $B_0_Redirect -match "SharedVoicemail" )
        {
            if ( $B_0_Redirect.Substring(16,1) -eq 0 )
            {
                $B_0_SharedVoicemailTranscription = $false
            }
            else
            {
                $B_0_SharedVoicemailTranscription = $true
            }
			
            if ( $B_0_Redirect.Substring(18,1) -eq 0 )
            {
                $B_0_SharedVoicemailSuppress = $false
            }
            else
            {
                $B_0_SharedVoicemailSuppress = $true
            }

            $B_0_Redirect = $B_0_Redirect.Substring(0,15)
		}

		# 2024.08.23 - call priorities currently flighted, need to support both ApplicationEndpoint and ApplicationEndpoint: 
		# handle ApplicationEndpoint:n
		# n={1-5} - Call Priority
		if ( $B_0_Redirect -match "ApplicationEndpoint")
        {
			if ( $B_0_Redirect.length -gt 19 )
			{
				$B_0_Redirect_Call_Priority = $B_0_Redirect.Substring(20,1)
				$B_0_Redirect = $B_0_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_0_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_0_Redirect.length -gt 21 )
			{
				$B_0_Redirect_Call_Priority = $B_0_Redirect.Substring(22,1)
				$B_0_Redirect = $B_0_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}
		
		#
		# check if FILE exists
		#
		if ( $B_0_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_0_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_0_Redirect_Comment = "ERROR: $B_0_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		#
		# Track if an operator redirect has already been done - only one supported.
		#
		if ( $B_0_Redirect -eq "Operator" )
		{
			# as this is the first check, only need to set to true here if option 0 is redirecting to Operator
			$B_Operator_Redirect = $true
		}

		#
		# Handle a Redirect Target error
		#
		if ( $B_0_RedirectTarget -eq "ERROR" )
		{
			$B_0_Redirect_Comment = "ERROR: Check B_0_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 1
        $B_1_VoiceCommand = $PSAAConfig.B_1_VoiceCommand[$i]
        $B_1_Redirect = $PSAAConfig.B_1_Redirect[$i]
		$B_1_RedirectTarget = $PSAAConfig.B_1_RedirectTarget[$i]
		
        if ( $B_1_Redirect -match "SharedVoicemail" )
        {
            if ( $B_1_Redirect.Substring(16,1) -eq 0 )
            {
                $B_1_SharedVoicemailTranscription = $false
            }
            else
            {
                $B_1_SharedVoicemailTranscription = $true
            }
            if ( $B_1_Redirect.Substring(18,1) -eq 0 )
            {
                $B_1_SharedVoicemailSuppress = $false
            }
            else
            {
                $B_1_SharedVoicemailSuppress = $true
            }

            $B_1_Redirect = $B_1_Redirect.Substring(0,15)
        }

		if ( $B_1_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_1_Redirect.length -gt 19 )
			{
				$B_1_Redirect_Call_Priority = $B_1_Redirect.Substring(20,1)
				$B_1_Redirect = $B_1_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_1_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_1_Redirect.length -gt 21 )
			{
				$B_1_Redirect_Call_Priority = $B_1_Redirect.Substring(22,1)
				$B_1_Redirect = $B_1_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}
        
		if ( $B_1_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_1_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_1_Redirect_Comment = "ERROR: $B_1_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $B_1_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_1_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_1_RedirectTarget -eq "ERROR" )
		{
			$B_1_Redirect_Comment = "ERROR: Check B_1_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 2
        $B_2_VoiceCommand = $PSAAConfig.B_2_VoiceCommand[$i]
        $B_2_Redirect = $PSAAConfig.B_2_Redirect[$i]
        $B_2_RedirectTarget = $PSAAConfig.B_2_RedirectTarget[$i]
		
        if ( $B_2_Redirect -match "SharedVoicemail" )
        {
            if ( $B_2_Redirect.Substring(16,1) -eq 0 )
            {
                $B_2_SharedVoicemailTranscription = $false
            }
            else
            {
                $B_2_SharedVoicemailTranscription = $true
            }
            if ( $B_2_Redirect.Substring(18,1) -eq 0 )
            {
                $B_2_SharedVoicemailSuppress = $false
            }
            else
            {
                $B_2_SharedVoicemailSuppress = $true
            }

            $B_2_Redirect = $B_2_Redirect.Substring(0,15)
        }

		if ( $B_2_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_2_Redirect.length -gt 19 )
			{
				$B_2_Redirect_Call_Priority = $B_2_Redirect.Substring(20,1)
				$B_2_Redirect = $B_2_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_2_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_2_Redirect.length -gt 21 )
			{
				$B_2_Redirect_Call_Priority = $B_2_Redirect.Substring(22,1)
				$B_2_Redirect = $B_2_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_2_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_2_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_2_Redirect_Comment = "ERROR: $B_2_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_2_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_2_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_2_RedirectTarget -eq "ERROR" )
		{
			$B_2_Redirect_Comment = "ERROR: Check B_2_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 3
        $B_3_VoiceCommand = $PSAAConfig.B_3_VoiceCommand[$i]
        $B_3_Redirect = $PSAAConfig.B_3_Redirect[$i]
		$B_3_RedirectTarget = $PSAAConfig.B_3_RedirectTarget[$i]
		
        if ( $B_3_Redirect -match "SharedVoicemail" )
        {
            if ( $B_3_Redirect.Substring(16,1) -eq 0 )
            {
                $B_3_SharedVoicemailTranscription = $false
            }
            else
            {
                $B_3_SharedVoicemailTranscription = $true
            }
            if ( $B_3_Redirect.Substring(18,1) -eq 0 )
            {
                $B_3_SharedVoicemailSuppress = $false
            }
            else
            {
                $B_3_SharedVoicemailSuppress = $true
            }

            $B_3_Redirect = $B_3_Redirect.Substring(0,15)
        }

		if ( $B_3_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_3_Redirect.length -gt 19 )
			{
				$B_3_Redirect_Call_Priority = $B_3_Redirect.Substring(20,1)
				$B_3_Redirect = $B_3_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_3_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_3_Redirect.length -gt 21 )
			{
				$B_3_Redirect_Call_Priority = $B_3_Redirect.Substring(22,1)
				$B_3_Redirect = $B_3_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}
 
		if ( $B_3_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_3_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_3_Redirect_Comment = "ERROR: $B_3_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_3_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_3_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_3_RedirectTarget -eq "ERROR" )
		{
			$B_3_Redirect_Comment = "ERROR: Check B_3_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 4
        $B_4_VoiceCommand = $PSAAConfig.B_4_VoiceCommand[$i]
        $B_4_Redirect = $PSAAConfig.B_4_Redirect[$i]
        $B_4_RedirectTarget = $PSAAConfig.B_4_RedirectTarget[$i]
		
        if ( $B_4_Redirect -match "SharedVoicemail" )
        {
            if ( $B_4_Redirect.Substring(16,1) -eq 0 )
            {
                $B_4_SharedVoicemailTranscription = $false
            }
            else
            {
                $B_4_SharedVoicemailTranscription = $true
            }
            if ( $B_4_Redirect.Substring(18,1) -eq 0 )
            {
                $B_4_SharedVoicemailSuppress = $false
            }
            else
            {
                $B_4_SharedVoicemailSuppress = $true
            }

            $B_4_Redirect = $B_4_Redirect.Substring(0,15)
        }

		if ( $B_4_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_4_Redirect.length -gt 19 )
			{
				$B_4_Redirect_Call_Priority = $B_4_Redirect.Substring(20,1)
				$B_4_Redirect = $B_4_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_4_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_4_Redirect.length -gt 21 )
			{
				$B_4_Redirect_Call_Priority = $B_4_Redirect.Substring(22,1)
				$B_4_Redirect = $B_4_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_4_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_4_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_4_Redirect_Comment = "ERROR: $B_4_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_4_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_4_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_4_RedirectTarget -eq "ERROR" )
		{
			$B_4_Redirect_Comment = "ERROR: Check B_4_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 5
		$B_5_VoiceCommand = $PSAAConfig.B_5_VoiceCommand[$i]
		$B_5_Redirect = $PSAAConfig.B_5_Redirect[$i]
		$B_5_RedirectTarget = $PSAAConfig.B_5_RedirectTarget[$i]
		
		if ( $B_5_Redirect -match "SharedVoicemail" )
		{
			if ( $B_5_Redirect.Substring(16,1) -eq 0 )
			{
				$B_5_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_5_SharedVoicemailTranscription = $true
			}
			if ( $B_5_Redirect.Substring(18,1) -eq 0 )
			{
				$B_5_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_5_SharedVoicemailSuppress = $true
			}

			$B_5_Redirect = $B_5_Redirect.Substring(0,15)
		}

		if ( $B_5_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_5_Redirect.length -gt 19 )
			{
				$B_5_Redirect_Call_Priority = $B_5_Redirect.Substring(20,1)
				$B_5_Redirect = $B_5_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_5_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_5_Redirect.length -gt 21 )
			{
				$B_5_Redirect_Call_Priority = $B_5_Redirect.Substring(22,1)
				$B_5_Redirect = $B_5_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_5_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_5_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_5_Redirect_Comment = "ERROR: $B_5_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_5_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_5_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_5_RedirectTarget -eq "ERROR" )
		{
			$B_5_Redirect_Comment = "ERROR: Check B_5_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 6
		$B_6_VoiceCommand = $PSAAConfig.B_6_VoiceCommand[$i]
		$B_6_Redirect = $PSAAConfig.B_6_Redirect[$i]
		$B_6_RedirectTarget = $PSAAConfig.B_6_RedirectTarget[$i]
		
		if ( $B_6_Redirect -match "SharedVoicemail" )
		{
			if ( $B_6_Redirect.Substring(16,1) -eq 0 )
			{
				$B_6_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_6_SharedVoicemailTranscription = $true
			}
			if ( $B_6_Redirect.Substring(18,1) -eq 0 )
			{
				$B_6_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_6_SharedVoicemailSuppress = $true
			}

			$B_6_Redirect = $B_6_Redirect.Substring(0,15)
		}

		if ( $B_6_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_6_Redirect.length -gt 19 )
			{
				$B_6_Redirect_Call_Priority = $B_6_Redirect.Substring(20,1)
				$B_6_Redirect = $B_6_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_6_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_6_Redirect.length -gt 21 )
			{
				$B_6_Redirect_Call_Priority = $B_6_Redirect.Substring(22,1)
				$B_6_Redirect = $B_6_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_6_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_6_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_6_Redirect_Comment = "ERROR: $B_6_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_6_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_6_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}
		
		if ( $B_6_RedirectTarget -eq "ERROR" )
		{
			$B_6_Redirect_Comment = "ERROR: Check B_6_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 7
		$B_7_VoiceCommand = $PSAAConfig.B_7_VoiceCommand[$i]
		$B_7_Redirect = $PSAAConfig.B_7_Redirect[$i]
		$B_7_RedirectTarget = $PSAAConfig.B_7_RedirectTarget[$i]
		
		if ( $B_7_Redirect -match "SharedVoicemail" )
		{
			if ( $B_7_Redirect.Substring(16,1) -eq 0 )
			{
				$B_7_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_7_SharedVoicemailTranscription = $true
			}
			if ( $B_7_Redirect.Substring(18,1) -eq 0 )
			{
				$B_7_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_7_SharedVoicemailSuppress = $true
			}

			$B_7_Redirect = $B_7_Redirect.Substring(0,15)
		}

		if ( $B_7_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_7_Redirect.length -gt 19 )
			{
				$B_7_Redirect_Call_Priority = $B_7_Redirect.Substring(20,1)
				$B_7_Redirect = $B_7_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_7_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_7_Redirect.length -gt 21 )
			{
				$B_7_Redirect_Call_Priority = $B_7_Redirect.Substring(22,1)
				$B_7_Redirect = $B_7_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_7_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_7_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_7_Redirect_Comment = "ERROR: $B_7_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_7_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_7_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_7_RedirectTarget -eq "ERROR" )
		{
			$B_7_Redirect_Comment = "ERROR: Check B_7_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 8
		$B_8_VoiceCommand = $PSAAConfig.B_8_VoiceCommand[$i]
		$B_8_Redirect = $PSAAConfig.B_8_Redirect[$i]
		$B_8_RedirectTarget = $PSAAConfig.B_8_RedirectTarget[$i]
		
		if ( $B_8_Redirect -match "SharedVoicemail" )
		{
			if ( $B_8_Redirect.Substring(16,1) -eq 0 )
			{
				$B_8_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_8_SharedVoicemailTranscription = $true
			}
			if ( $B_8_Redirect.Substring(18,1) -eq 0 )
			{
				$B_8_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_8_SharedVoicemailSuppress = $true
			}

			$B_8_Redirect = $B_8_Redirect.Substring(0,15)
		}

		if ( $B_8_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_8_Redirect.length -gt 19 )
			{
				$B_8_Redirect_Call_Priority = $B_8_Redirect.Substring(20,1)
				$B_8_Redirect = $B_8_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_8_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_8_Redirect.length -gt 21 )
			{
				$B_8_Redirect_Call_Priority = $B_8_Redirect.Substring(22,1)
				$B_8_Redirect = $B_8_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_8_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_8_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_8_Redirect_Comment = "ERROR: $B_8_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_8_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_8_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_8_RedirectTarget -eq "ERROR" )
		{
			$B_8_Redirect_Comment = "ERROR: Check B_8_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option 9
		$B_9_VoiceCommand = $PSAAConfig.B_9_VoiceCommand[$i]
		$B_9_Redirect = $PSAAConfig.B_9_Redirect[$i]
		$B_9_RedirectTarget = $PSAAConfig.B_9_RedirectTarget[$i]
		
		if ( $B_9_Redirect -match "SharedVoicemail" )
		{
			if ( $B_9_Redirect.Substring(16,1) -eq 0 )
			{
				$B_9_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_9_SharedVoicemailTranscription = $true
			}
			if ( $B_9_Redirect.Substring(18,1) -eq 0 )
			{
				$B_9_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_9_SharedVoicemailSuppress = $true
			}

			$B_9_Redirect = $B_9_Redirect.Substring(0,15)
		}

		if ( $B_9_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_9_Redirect.length -gt 19 )
			{
				$B_9_Redirect_Call_Priority = $B_9_Redirect.Substring(20,1)
				$B_9_Redirect = $B_9_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_9_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_9_Redirect.length -gt 21 )
			{
				$B_9_Redirect_Call_Priority = $B_9_Redirect.Substring(22,1)
				$B_9_Redirect = $B_9_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_9_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_9_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_9_Redirect_Comment = "ERROR: $B_9_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_9_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_9_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_9_RedirectTarget -eq "ERROR" )
		{
			$B_9_Redirect_Comment = "ERROR: Check B_9_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option *
		$B_Star_VoiceCommand = $PSAAConfig.B_Star_VoiceCommand[$i]
		$B_Star_Redirect = $PSAAConfig.B_Star_Redirect[$i]
		$B_Star_RedirectTarget = $PSAAConfig.B_Star_RedirectTarget[$i]
		
		if ( $B_Star_Redirect -match "SharedVoicemail" )
		{
			if ( $B_Star_Redirect.Substring(16,1) -eq 0 )
			{
				$B_Star_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_Star_SharedVoicemailTranscription = $true
			}
			if ( $B_Star_Redirect.Substring(18,1) -eq 0 )
			{
				$B_Star_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_Star_SharedVoicemailSuppress = $true
			}

			$B_Star_Redirect = $B_Star_Redirect.Substring(0,15)
		}

		if ( $B_Star_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_Star_Redirect.length -gt 19 )
			{
				$B_Star_Redirect_Call_Priority = $B_Star_Redirect.Substring(20,1)
				$B_Star_Redirect = $B_Star_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_Star_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_Star_Redirect.length -gt 21 )
			{
				$B_Star_Redirect_Call_Priority = $B_Star_Redirect.Substring(22,1)
				$B_Star_Redirect = $B_Star_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_Star_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_Star_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_Star_Redirect_Comment = "ERROR: $B_Star_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_Star_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_Star_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_Star_RedirectTarget -eq "ERROR" )
		{
			$B_Star_Redirect_Comment = "ERROR: Check B_Star_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Business Hours - Option #
		$B_Pound_VoiceCommand = $PSAAConfig.B_Pound_VoiceCommand[$i]
		$B_Pound_Redirect = $PSAAConfig.B_Pound_Redirect[$i]
		$B_Pound_RedirectTarget = $PSAAConfig.B_Pound_RedirectTarget[$i]
		
		if ( $B_Pound_Redirect -match "SharedVoicemail" )
		{
			if ( $B_Pound_Redirect.Substring(16,1) -eq 0 )
			{
				$B_Pound_SharedVoicemailTranscription = $false
			}
			else
			{
				$B_Pound_SharedVoicemailTranscription = $true
			}
			if ( $B_Pound_Redirect.Substring(18,1) -eq 0 )
			{
				$B_Pound_SharedVoicemailSuppress = $false
			}
			else
			{
				$B_Pound_SharedVoicemailSuppress = $true
			}

			$B_Pound_Redirect = $B_Pound_Redirect.Substring(0,15)
		}

		if ( $B_Pound_Redirect -match "ApplicationEndpoint" )
        {
			if ( $B_Pound_Redirect.length -gt 19 )
			{
				$B_Pound_Redirect_Call_Priority = $B_Pound_Redirect.Substring(20,1)
				$B_Pound_Redirect = $B_Pound_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $B_Pound_Redirect -match "ConfigurationEndpoint")
        {
			if ( $B_Pound_Redirect.length -gt 21 )
			{
				$B_Pound_Redirect_Call_Priority = $B_Pound_Redirect.Substring(22,1)
				$B_Pound_Redirect = $B_Pound_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $B_Pound_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $B_Pound_RedirectTarget
			
			if ( ! $FileExists )
			{
				$B_Pound_Redirect_Comment = "ERROR: $B_Pound_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true				
			}
		}

		if ( $B_Pound_Redirect -eq "Operator" )
		{
			if ( $B_Operator_Redirect )
			{
				$B_Pound_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$B_Operator_Redirect = $true
			}
		}

		if ( $B_Pound_RedirectTarget -eq "ERROR" )
		{
			$B_Pound_Redirect_Comment = "ERROR: Check B_Pound_RedirectTarget on Config-BusinessHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


        #
        # After Hours
        #
		$A_Operator_Redirect = $false
				
        $A_DirectorySearch = $PSAAConfig.A_DirectorySearch[$i]
        $A_DialScopeInclude = $PSAAConfig.A_DialScopeInclude[$i]
        $A_DialScopeExclude = $PSAAConfig.A_DialScopeExclude[$i] 

		if ( $A_DialScopeInclude -eq "ERROR" )
		{
			$A_DialScopeInclude_Comment = "ERROR: Check A_DialScopeInclude on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}
			
		if ( $A_DialScopeExclude -eq "ERROR" )
		{
			$A_DialScopeExclude_Comment = "ERROR: Check A_DialScopeExclude on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        if ( $PSAAConfig.A_Force[$i] -eq "true" )
        {
            $A_Force = $true
        }
        else
        {
            $A_Force = $false
        }

        $A_GreetingOption = $PSAAConfig.A_GreetingOption[$i]
        $A_Greeting = $PSAAConfig.A_Greeting[$i]
		
		if ( $A_GreetingOption -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_Greeting
			
			if ( ! $FileExists )
			{
				$A_Greeting_Comment = "ERROR: $A_Greeting does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}
				
        $A_GreetingRouting = $PSAAConfig.A_GreetingRouting[$i]
        $A_GreetingRoutingTarget = $PSAAConfig.A_GreetingRoutingTarget[$i]

		if ( $A_GreetingRouting -match "ApplicationEndpoint" )
        {
			if ( $A_GreetingRouting.length -gt 19 )
			{
				$A_GreetingRouting_Call_Priority = $A_GreetingRouting.Substring(20,1)
				$A_GreetingRouting = $A_GreetingRouting.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_GreetingRouting -match "ConfigurationEndpoint" )
        {
			if ( $A_GreetingRouting.length -gt 21 )
			{
				$A_GreetingRouting_Call_Priority = $A_GreetingRouting.Substring(22,1)
				$A_GreetingRouting = $A_GreetingRouting.Substring(0,21)
				
				$CallPriority = $true
			}
		}
		
		if ( $A_GreetingRoutingTarget -eq "ERROR" )
		{
			$A_GreetingRouting_Comment = "ERROR: Check A_GreetingRoutingTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        $A_MenuGreetingOption = $PSAAConfig.A_MenuGreetingOption[$i]
        $A_MenuGreeting = $PSAAConfig.A_MenuGreeting[$i]

		# Afer Hours - Option 0
        $A_0_VoiceCommand = $PSAAConfig.A_0_VoiceCommand[$i]
        $A_0_Redirect = $PSAAConfig.A_0_Redirect[$i]
        $A_0_RedirectTarget = $PSAAConfig.A_0_RedirectTarget[$i]		

		# handle SharedVoicemail:n1:n2
		# n1={0,1} - SharedVoicemailTranscription On/Off
		# n2={0,1} - SharedVoicemailSuppress On/Off
        if ( $A_0_Redirect -match "SharedVoicemail" )
        {
            if ( $A_0_Redirect.Substring(16,1) -eq 0 )
            {
                $A_0_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_0_SharedVoicemailTranscription = $true
            }
            if ( $A_0_Redirect.Substring(18,1) -eq 0 )
            {
                $A_0_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_0_SharedVoicemailSuppress = $true
            }

            $A_0_Redirect = $A_0_Redirect.Substring(0,15)
        }

		# 2024.08.23 - call priorities currently flighted, need to support both ApplicationEndpoint and ApplicationEndpoint: 
		# handle ApplicationEndpoint:n
		# n={1-5} - Call Priority
		if ( $A_0_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_0_Redirect.length -gt 19 )
			{
				$A_0_Redirect_Call_Priority = $A_0_Redirect.Substring(20,1)
				$A_0_Redirect = $A_0_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_0_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_0_Redirect.length -gt 21 )
			{
				$A_0_Redirect_Call_Priority = $A_0_Redirect.Substring(22,1)
				$A_0_Redirect = $A_0_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_0_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_0_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_0_Redirect_Comment = "ERROR: $A_0_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		#
		# Track if an operator redirect has already been done - only one supported.
		#
		if ( $A_0_Redirect -eq "Operator" )
		{
			# as this is the first check, only need to set to true here if option 0 is redirecting to Operator
			$A_Operator_Redirect = $true
		}

		if ( $A_0_RedirectTarget -eq "ERROR" )
		{
			$A_0_Redirect_Comment = "ERROR: Check A_0_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 1
        $A_1_VoiceCommand = $PSAAConfig.A_1_VoiceCommand[$i]
        $A_1_Redirect = $PSAAConfig.A_1_Redirect[$i]
        $A_1_RedirectTarget = $PSAAConfig.A_1_RedirectTarget[$i]
		
        if ( $A_1_Redirect -match "SharedVoicemail" )
        {
            if ( $A_1_Redirect.Substring(16,1) -eq 0 )
            {
                $A_1_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_1_SharedVoicemailTranscription = $true
            }
            if ( $A_1_Redirect.Substring(18,1) -eq 0 )
            {
                $A_1_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_1_SharedVoicemailSuppress = $true
            }

            $A_1_Redirect = $A_1_Redirect.Substring(0,15)
        }

		if ( $A_1_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_1_Redirect.length -gt 19 )
			{
				$A_1_Redirect_Call_Priority = $A_1_Redirect.Substring(20,1)
				$A_1_Redirect = $A_1_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_1_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_1_Redirect.length -gt 21 )
			{
				$A_1_Redirect_Call_Priority = $A_1_Redirect.Substring(22,1)
				$A_1_Redirect = $A_1_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_1_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_1_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_1_Redirect_Comment = "ERROR: $A_1_RedirectTarget does not exist"
				$StopProcessing = $true
				$Ver
				boseStopProcessing = $true
			}
		}

		if ( $A_1_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_1_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_1_RedirectTarget -eq "ERROR" )
		{
			$A_1_Redirect_Comment = "ERROR: Check A_1_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 2
        $A_2_VoiceCommand = $PSAAConfig.A_2_VoiceCommand[$i]
        $A_2_Redirect = $PSAAConfig.A_2_Redirect[$i]
        $A_2_RedirectTarget = $PSAAConfig.A_2_RedirectTarget[$i]
		
        if ( $A_2_Redirect -match "SharedVoicemail" )
        {
            if ( $A_2_Redirect.Substring(16,1) -eq 0 )
            {
                $A_2_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_2_SharedVoicemailTranscription = $true
            }
            if ( $A_2_Redirect.Substring(18,1) -eq 0 )
            {
                $A_2_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_2_SharedVoicemailSuppress = $true
            }

            $A_2_Redirect = $A_2_Redirect.Substring(0,15)
        }

		if ( $A_2_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_2_Redirect.length -gt 19 )
			{
				$A_2_Redirect_Call_Priority = $A_2_Redirect.Substring(20,1)
				$A_2_Redirect = $A_2_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_2_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_2_Redirect.length -gt 21 )
			{
				$A_2_Redirect_Call_Priority = $A_2_Redirect.Substring(22,1)
				$A_2_Redirect = $A_2_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_2_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_2_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_2_Redirect_Comment = "ERROR: $A_2_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_2_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_2_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_2_RedirectTarget -eq "ERROR" )
		{
			$A_2_Redirect_Comment = "ERROR: Check A_2_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 3
        $A_3_VoiceCommand = $PSAAConfig.A_3_VoiceCommand[$i]
        $A_3_Redirect = $PSAAConfig.A_3_Redirect[$i]
        $A_3_RedirectTarget = $PSAAConfig.A_3_RedirectTarget[$i]
		
        if ( $A_3_Redirect -match "SharedVoicemail" )
        {
            if ( $A_3_Redirect.Substring(16,1) -eq 0 )
            {
                $A_3_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_3_SharedVoicemailTranscription = $true
            }
            if ( $A_3_Redirect.Substring(18,1) -eq 0 )
            {
                $A_3_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_3_SharedVoicemailSuppress = $true
            }

            $A_3_Redirect = $A_3_Redirect.Substring(0,15)
        }

		if ( $A_3_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_3_Redirect.length -gt 19 )
			{
				$A_3_Redirect_Call_Priority = $A_3_Redirect.Substring(20,1)
				$A_3_Redirect = $A_3_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_3_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_3_Redirect.length -gt 21 )
			{
				$A_3_Redirect_Call_Priority = $A_3_Redirect.Substring(22,1)
				$A_3_Redirect = $A_3_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_3_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_3_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_3_Redirect_Comment = "ERROR: $A_3_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_3_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_3_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_3_RedirectTarget -eq "ERROR" )
		{
			$A_3_Redirect_Comment = "ERROR: Check A_3_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 4
        $A_4_VoiceCommand = $PSAAConfig.A_4_VoiceCommand[$i]
        $A_4_Redirect = $PSAAConfig.A_4_Redirect[$i]
        $A_4_RedirectTarget = $PSAAConfig.A_4_RedirectTarget[$i]
		
        if ( $A_4_Redirect -match "SharedVoicemail" )
        {
            if ( $A_4_Redirect.Substring(16,1) -eq 0 )
            {
                $A_4_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_4_SharedVoicemailTranscription = $true
            }
            if ( $A_4_Redirect.Substring(18,1) -eq 0 )
            {
                $A_4_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_4_SharedVoicemailSuppress = $true
            }

            $A_4_Redirect = $A_4_Redirect.Substring(0,15)
        }

		if ( $A_4_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_4_Redirect.length -gt 19 )
			{
				$A_4_Redirect_Call_Priority = $A_4_Redirect.Substring(20,1)
				$A_4_Redirect = $A_4_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_4_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_4_Redirect.length -gt 21 )
			{
				$A_4_Redirect_Call_Priority = $A_4_Redirect.Substring(22,1)
				$A_4_Redirect = $A_4_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_4_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_4_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_4_Redirect_Comment = "ERROR: $A_4_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_4_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_4_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_4_RedirectTarget -eq "ERROR" )
		{
			$A_4_Redirect_Comment = "ERROR: Check A_4_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 5
        $A_5_VoiceCommand = $PSAAConfig.A_5_VoiceCommand[$i]
        $A_5_Redirect = $PSAAConfig.A_5_Redirect[$i]
        $A_5_RedirectTarget = $PSAAConfig.A_5_RedirectTarget[$i]
		
        if ( $A_5_Redirect -match "SharedVoicemail" )
        {
            if ( $A_5_Redirect.Substring(16,1) -eq 0 )
            {
                $A_5_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_5_SharedVoicemailTranscription = $true
            }
            if ( $A_5_Redirect.Substring(18,1) -eq 0 )
            {
                $A_5_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_5_SharedVoicemailSuppress = $true
            }

            $A_5_Redirect = $A_5_Redirect.Substring(0,15)
        }

		if ( $A_5_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_5_Redirect.length -gt 19 )
			{
				$A_5_Redirect_Call_Priority = $A_5_Redirect.Substring(20,1)
				$A_5_Redirect = $A_5_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_5_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_5_Redirect.length -gt 21 )
			{
				$A_5_Redirect_Call_Priority = $A_5_Redirect.Substring(22,1)
				$A_5_Redirect = $A_5_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_5_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_5_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_5_Redirect_Comment = "ERROR: $A_5_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_5_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_5_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_5_RedirectTarget -eq "ERROR" )
		{
			$A_5_Redirect_Comment = "ERROR: Check A_5_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 6
        $A_6_VoiceCommand = $PSAAConfig.A_6_VoiceCommand[$i]
        $A_6_Redirect = $PSAAConfig.A_6_Redirect[$i]
        $A_6_RedirectTarget = $PSAAConfig.A_6_RedirectTarget[$i]
		
        if ( $A_6_Redirect -match "SharedVoicemail" )
        {
            if ( $A_6_Redirect.Substring(16,1) -eq 0 )
            {
                $A_6_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_6_SharedVoicemailTranscription = $true
            }
            if ( $A_6_Redirect.Substring(18,1) -eq 0 )
            {
                $A_6_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_6_SharedVoicemailSuppress = $true
            }

            $A_6_Redirect = $A_6_Redirect.Substring(0,15)
        }

		if ( $A_6_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_6_Redirect.length -gt 19 )
			{
				$A_6_Redirect_Call_Priority = $A_6_Redirect.Substring(20,1)
				$A_6_Redirect = $A_6_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_6_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_6_Redirect.length -gt 21 )
			{
				$A_6_Redirect_Call_Priority = $A_6_Redirect.Substring(22,1)
				$A_6_Redirect = $A_6_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_6_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_6_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_6_Redirect_Comment = "ERROR: $A_6_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_6_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_6_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_6_RedirectTarget -eq "ERROR" )
		{
			$A_6_Redirect_Comment = "ERROR: Check A_6_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 7
        $A_7_VoiceCommand = $PSAAConfig.A_7_VoiceCommand[$i]
        $A_7_Redirect = $PSAAConfig.A_7_Redirect[$i]
        $A_7_RedirectTarget = $PSAAConfig.A_7_RedirectTarget[$i]
		
        if ( $A_7_Redirect -match "SharedVoicemail" )
        {
            if ( $A_7_Redirect.Substring(16,1) -eq 0 )
            {
                $A_7_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_7_SharedVoicemailTranscription = $true
            }
            if ( $A_7_Redirect.Substring(18,1) -eq 0 )
            {
                $A_7_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_7_SharedVoicemailSuppress = $true
            }

            $A_7_Redirect = $A_7_Redirect.Substring(0,15)
        }

		if ( $A_7_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_7_Redirect.length -gt 19 )
			{
				$A_7_Redirect_Call_Priority = $A_7_Redirect.Substring(20,1)
				$A_7_Redirect = $A_7_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_7_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_7_Redirect.length -gt 21 )
			{
				$A_7_Redirect_Call_Priority = $A_7_Redirect.Substring(22,1)
				$A_7_Redirect = $A_7_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_7_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_7_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_7_Redirect_Comment = "ERROR: $A_7_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_7_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_7_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_7_RedirectTarget -eq "ERROR" )
		{
			$A_7_Redirect_Comment = "ERROR: Check A_7_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 8
        $A_8_VoiceCommand = $PSAAConfig.A_8_VoiceCommand[$i]
        $A_8_Redirect = $PSAAConfig.A_8_Redirect[$i]
        $A_8_RedirectTarget = $PSAAConfig.A_8_RedirectTarget[$i]
		
        if ( $A_8_Redirect -match "SharedVoicemail" )
        {
            if ( $A_8_Redirect.Substring(16,1) -eq 0 )
            {
                $A_8_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_8_SharedVoicemailTranscription = $true
            }
            if ( $A_8_Redirect.Substring(18,1) -eq 0 )
            {
                $A_8_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_8_SharedVoicemailSuppress = $true
            }

           $A_8_Redirect = $A_8_Redirect.Substring(0,15)
        }

		if ( $A_8_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_8_Redirect.length -gt 19 )
			{
				$A_8_Redirect_Call_Priority = $A_8_Redirect.Substring(20,1)
				$A_8_Redirect = $A_8_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_8_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_8_Redirect.length -gt 21 )
			{
				$A_8_Redirect_Call_Priority = $A_8_Redirect.Substring(22,1)
				$A_8_Redirect = $A_8_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_8_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_8_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_8_Redirect_Comment = "ERROR: $A_8_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_8_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_8_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_8_RedirectTarget -eq "ERROR" )
		{
			$A_8_Redirect_Comment = "ERROR: Check A_8_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option 9
        $A_9_VoiceCommand = $PSAAConfig.A_9_VoiceCommand[$i]
        $A_9_Redirect = $PSAAConfig.A_9_Redirect[$i]
        $A_9_RedirectTarget = $PSAAConfig.A_9_RedirectTarget[$i]
		
        if ( $A_9_Redirect -match "SharedVoicemail" )
        {
            if ( $A_9_Redirect.Substring(16,1) -eq 0 )
            {
                $A_9_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_9_SharedVoicemailTranscription = $true
            }
            if ( $A_9_Redirect.Substring(18,1) -eq 0 )
            {
                $A_9_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_9_SharedVoicemailSuppress = $true
            }

            $A_9_Redirect = $A_9_Redirect.Substring(0,15)
        }

		if ( $A_9_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_9_Redirect.length -gt 19 )
			{
				$A_9_Redirect_Call_Priority = $A_9_Redirect.Substring(20,1)
				$A_9_Redirect = $A_9_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_9_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_9_Redirect.length -gt 21 )
			{
				$A_9_Redirect_Call_Priority = $A_9_Redirect.Substring(22,1)
				$A_9_Redirect = $A_9_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_9_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_9_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_9_Redirect_Comment = "ERROR: $A_9_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_9_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_9_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_9_RedirectTarget -eq "ERROR" )
		{
			$A_9_Redirect_Comment = "ERROR: Check A_9_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option *
        $A_Star_VoiceCommand = $PSAAConfig.A_Star_VoiceCommand[$i]
        $A_Star_Redirect = $PSAAConfig.A_Star_Redirect[$i]
        $A_Star_RedirectTarget = $PSAAConfig.A_Star_RedirectTarget[$i]
		
        if ( $A_Star_Redirect -match "SharedVoicemail" )
        {
            if ( $A_Star_Redirect.Substring(16,1) -eq 0 )
            {
                $A_Star_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_Star_SharedVoicemailTranscription = $true
            }
            if ( $A_Star_Redirect.Substring(18,1) -eq 0 )
            {
                $A_Star_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_Star_SharedVoicemailSuppress = $true
            }

            $A_Star_Redirect = $A_Star_Redirect.Substring(0,15)
        }

		if ( $A_Star_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_Star_Redirect.length -gt 19 )
			{
				$A_Star_Redirect_Call_Priority = $A_Star_Redirect.Substring(20,1)
				$A_Star_Redirect = $A_Star_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_Star_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_Star_Redirect.length -gt 21 )
			{
				$A_Star_Redirect_Call_Priority = $A_Star_Redirect.Substring(22,1)
				$A_Star_Redirect = $A_Star_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_Star_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_Star_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_Star_Redirect_Comment = "ERROR: $A_Star_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_Star_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_Star_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_Star_RedirectTarget -eq "ERROR" )
		{
			$A_Star_Redirect_Comment = "ERROR: Check A_Star_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Afer Hours - Option #
        $A_Pound_VoiceCommand = $PSAAConfig.A_Pound_VoiceCommand[$i]
        $A_Pound_Redirect = $PSAAConfig.A_Pound_Redirect[$i]
        $A_Pound_RedirectTarget = $PSAAConfig.A_Pound_RedirectTarget[$i]
		
        if ( $A_Pound_Redirect -match "SharedVoicemail" )
        {
            if ( $A_Pound_Redirect.Substring(16,1) -eq 0 )
            {
                $A_Pound_SharedVoicemailTranscription = $false
            }
            else
            {
                $A_Pound_SharedVoicemailTranscription = $true
            }
            if ( $A_Pound_Redirect.Substring(18,1) -eq 0 )
            {
                $A_Pound_SharedVoicemailSuppress = $false
            }
            else
            {
                $A_Pound_SharedVoicemailSuppress = $true
            }

            $A_Pound_Redirect = $A_Pound_Redirect.Substring(0,15)
        }

		if ( $A_Pound_Redirect -match "ApplicationEndpoint" )
        {
			if ( $A_Pound_Redirect.length -gt 19 )
			{
				$A_Pound_Redirect_Call_Priority = $A_Pound_Redirect.Substring(20,1)
				$A_Pound_Redirect = $A_Pound_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $A_Pound_Redirect -match "ConfigurationEndpoint")
        {
			if ( $A_Pound_Redirect.length -gt 21 )
			{
				$A_Pound_Redirect_Call_Priority = $A_Pound_Redirect.Substring(22,1)
				$A_Pound_Redirect = $A_Pound_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $A_Pound_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $A_Pound_RedirectTarget
			
			if ( ! $FileExists )
			{
				$A_Pound_Redirect_Comment = "ERROR: $A_Pound_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $A_Pound_Redirect -eq "Operator" )
		{
			if ( $A_Operator_Redirect )
			{
				$A_Pound_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$A_Operator_Redirect = $true
			}
		}

		if ( $A_Pound_RedirectTarget -eq "ERROR" )
		{
			$A_Pound_Redirect_Comment = "ERROR: Check A_Pound_RedirectTarget on Config-AfterHoursMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


        #
        # Holidays
        #
		$H_Operator_Redirect = $false
				
        $H_DirectorySearch = $PSAAConfig.H_DirectorySearch[$i]
        $H_DialScopeInclude = $PSAAConfig.H_DialScopeInclude[$i]
        $H_DialScopeExclude = $PSAAConfig.H_DialScopeExclude[$i] 

		if ( $H_DialScopeInclude -eq "ERROR" )
		{
			$H_DialScopeInclude_Comment = "ERROR: Check H_DialScopeInclude on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}
			
		if ( $H_DialScopeExclude -eq "ERROR" )
		{
			$H_DialScopeExclude_Comment = "ERROR: Check H_DialScopeExclude on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        if ( $PSAAConfig.H_Force[$i] -eq "true" )
        {
            $H_Force = $true
        }
        else
        {
            $H_Force = $false
        }

        $H_GreetingOption = $PSAAConfig.H_GreetingOption[$i]
        $H_Greeting = $PSAAConfig.H_Greeting[$i]
		
		if ( $H_GreetingOption -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_Greeting
			
			if ( ! $FileExists )
			{
				$H_Greeting_Comment = "ERROR: $H_Greeting does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}		
		
        $H_GreetingRouting = $PSAAConfig.H_GreetingRouting[$i]
        $H_GreetingRoutingTarget = $PSAAConfig.H_GreetingRoutingTarget[$i]

		if ( $H_GreetingRouting -match "ApplicationEndpoint" )
        {
			if ( $H_GreetingRouting.length -gt 19 )
			{
				$H_GreetingRouting_Call_Priority = $H_GreetingRouting.Substring(20,1)
				$H_GreetingRouting = $H_GreetingRouting.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_GreetingRouting -match "ConfigurationEndpoint" )
        {
			if ( $H_GreetingRouting.length -gt 21 )
			{
				$H_GreetingRouting_Call_Priority = $H_GreetingRouting.Substring(22,1)
				$H_GreetingRouting = $H_GreetingRouting.Substring(0,21)
				
				$CallPriority = $true
			}
		}
		
		if ( $H_GreetingRoutingTarget -eq "ERROR" )
		{
			$H_GreetingRouting_Comment = "ERROR: Check H_GreetingRoutingTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

        $H_MenuGreetingOption = $PSAAConfig.H_MenuGreetingOption[$i]
        $H_MenuGreeting = $PSAAConfig.H_MenuGreeting[$i]

		# Holiday - Option 0
        $H_0_VoiceCommand = $PSAAConfig.H_0_VoiceCommand[$i]
        $H_0_Redirect = $PSAAConfig.H_0_Redirect[$i]
        $H_0_RedirectTarget = $PSAAConfig.H_0_RedirectTarget[$i]		

		# handle SharedVoicemail:n1:n2
		# n1={0,1} - SharedVoicemailTranscription On/Off
		# n2={0,1} - SharedVoicemailSuppress On/Off
        if ( $H_0_Redirect -match "SharedVoicemail" )
        {
            if ( $H_0_Redirect.Substring(16,1) -eq 0 )
            {
                $H_0_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_0_SharedVoicemailTranscription = $true
            }
            if ( $H_0_Redirect.Substring(18,1) -eq 0 )
            {
                $H_0_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_0_SharedVoicemailSuppress = $true
            }

            $H_0_Redirect = $H_0_Redirect.Substring(0,15)
        }

		# 2024.08.23 - call priorities currently flighted, need to support both ApplicationEndpoint and ApplicationEndpoint: 
		# handle ApplicationEndpoint:n
		# n={1-5} - Call Priority
		if ( $H_0_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_0_Redirect.length -gt 19 )
			{
				$H_0_Redirect_Call_Priority = $H_0_Redirect.Substring(20,1)
				$H_0_Redirect = $H_0_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_0_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_0_Redirect.length -gt 21 )
			{
				$H_0_Redirect_Call_Priority = $H_0_Redirect.Substring(22,1)
				$H_0_Redirect = $H_0_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_0_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_0_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_0_Redirect_Comment = "ERROR: $H_0_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		#
		# Track if an operator redirect has already been done - only one supported.
		#
		if ( $H_0_Redirect -eq "Operator" )
		{
			# as this is the first check, only need to set to true here if option 0 is redirecting to Operator
			$H_Operator_Redirect = $true
		}

		if ( $H_0_RedirectTarget -eq "ERROR" )
		{
			$H_0_Redirect_Comment = "ERROR: Check H_0_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 1
        $H_1_VoiceCommand = $PSAAConfig.H_1_VoiceCommand[$i]
        $H_1_Redirect = $PSAAConfig.H_1_Redirect[$i]
        $H_1_RedirectTarget = $PSAAConfig.H_1_RedirectTarget[$i]
		
        if ( $H_1_Redirect -match "SharedVoicemail" )
        {
            if ( $H_1_Redirect.Substring(16,1) -eq 0 )
            {
                $H_1_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_1_SharedVoicemailTranscription = $true
            }
            if ( $H_1_Redirect.Substring(18,1) -eq 0 )
            {
                $H_1_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_1_SharedVoicemailSuppress = $true
            }

            $H_1_Redirect = $H_1_Redirect.Substring(0,15)
        }

		if ( $H_1_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_1_Redirect.length -gt 19 )
			{
				$H_1_Redirect_Call_Priority = $H_1_Redirect.Substring(20,1)
				$H_1_Redirect = $H_1_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_1_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_1_Redirect.length -gt 21 )
			{
				$H_1_Redirect_Call_Priority = $H_1_Redirect.Substring(22,1)
				$H_1_Redirect = $H_1_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_1_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_1_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_1_Redirect_Comment = "ERROR: $H_1_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_1_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_1_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_1_RedirectTarget -eq "ERROR" )
		{
			$H_1_Redirect_Comment = "ERROR: Check H_1_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 2
        $H_2_VoiceCommand = $PSAAConfig.H_2_VoiceCommand[$i]
        $H_2_Redirect = $PSAAConfig.H_2_Redirect[$i]
        $H_2_RedirectTarget = $PSAAConfig.H_2_RedirectTarget[$i]
		
        if ( $H_2_Redirect -match "SharedVoicemail" )
        {
            if ( $H_2_Redirect.Substring(16,1) -eq 0 )
            {
                $H_2_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_2_SharedVoicemailTranscription = $true
            }
            if ( $H_2_Redirect.Substring(18,1) -eq 0 )
            {
                $H_2_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_2_SharedVoicemailSuppress = $true
            }

            $H_2_Redirect = $H_2_Redirect.Substring(0,15)
        }

		if ( $H_2_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_2_Redirect.length -gt 19 )
			{
				$H_2_Redirect_Call_Priority = $H_2_Redirect.Substring(20,1)
				$H_2_Redirect = $H_2_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_2_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_2_Redirect.length -gt 21 )
			{
				$H_2_Redirect_Call_Priority = $H_2_Redirect.Substring(22,1)
				$H_2_Redirect = $H_2_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_2_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_2_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_2_Redirect_Comment = "ERROR: $H_2_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_2_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_2_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_2_RedirectTarget -eq "ERROR" )
		{
			$H_2_Redirect_Comment = "ERROR: Check H_2_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 3
        $H_3_VoiceCommand = $PSAAConfig.H_3_VoiceCommand[$i]
        $H_3_Redirect = $PSAAConfig.H_3_Redirect[$i]
        $H_3_RedirectTarget = $PSAAConfig.H_3_RedirectTarget[$i]
		
        if ( $H_3_Redirect -match "SharedVoicemail" )
        {
            if ( $H_3_Redirect.Substring(16,1) -eq 0 )
            {
                $H_3_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_3_SharedVoicemailTranscription = $true
            }
            if ( $H_3_Redirect.Substring(18,1) -eq 0 )
            {
                $H_3_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_3_SharedVoicemailSuppress = $true
            }

            $H_3_Redirect = $H_3_Redirect.Substring(0,15)
        }

		if ( $H_3_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_3_Redirect.length -gt 19 )
			{
				$H_3_Redirect_Call_Priority = $H_3_Redirect.Substring(20,1)
				$H_3_Redirect = $H_3_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_3_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_3_Redirect.length -gt 21 )
			{
				$H_3_Redirect_Call_Priority = $H_3_Redirect.Substring(22,1)
				$H_3_Redirect = $H_3_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_3_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_3_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_3_Redirect_Comment = "ERROR: $H_3_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_3_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_3_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_3_RedirectTarget -eq "ERROR" )
		{
			$H_3_Redirect_Comment = "ERROR: Check H_3_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 4
        $H_4_VoiceCommand = $PSAAConfig.H_4_VoiceCommand[$i]
        $H_4_Redirect = $PSAAConfig.H_4_Redirect[$i]
        $H_4_RedirectTarget = $PSAAConfig.H_4_RedirectTarget[$i]
		
        if ( $H_4_Redirect -match "SharedVoicemail" )
        {
            if ( $H_4_Redirect.Substring(16,1) -eq 0 )
            {
                $H_4_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_4_SharedVoicemailTranscription = $true
            }
            if ( $H_4_Redirect.Substring(18,1) -eq 0 )
            {
                $H_4_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_4_SharedVoicemailSuppress = $true
            }

            $H_4_Redirect = $H_4_Redirect.Substring(0,15)
        }

		if ( $H_4_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_4_Redirect.length -gt 19 )
			{
				$H_4_Redirect_Call_Priority = $H_4_Redirect.Substring(20,1)
				$H_4_Redirect = $H_4_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_4_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_4_Redirect.length -gt 21 )
			{
				$H_4_Redirect_Call_Priority = $H_4_Redirect.Substring(22,1)
				$H_4_Redirect = $H_4_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_4_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_4_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_4_Redirect_Comment = "ERROR: $H_4_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_4_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_4_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_4_RedirectTarget -eq "ERROR" )
		{
			$H_4_Redirect_Comment = "ERROR: Check H_4_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 5
        $H_5_VoiceCommand = $PSAAConfig.H_5_VoiceCommand[$i]
        $H_5_Redirect = $PSAAConfig.H_5_Redirect[$i]
        $H_5_RedirectTarget = $PSAAConfig.H_5_RedirectTarget[$i]
		
        if ( $H_5_Redirect -match "SharedVoicemail" )
        {
            if ( $H_5_Redirect.Substring(16,1) -eq 0 )
            {
                $H_5_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_5_SharedVoicemailTranscription = $true
            }
            if ( $H_5_Redirect.Substring(18,1) -eq 0 )
            {
                $H_5_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_5_SharedVoicemailSuppress = $true
            }

            $H_5_Redirect = $H_5_Redirect.Substring(0,15)
        }

		if ( $H_5_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_5_Redirect.length -gt 19 )
			{
				$H_5_Redirect_Call_Priority = $H_5_Redirect.Substring(20,1)
				$H_5_Redirect = $H_5_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_5_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_5_Redirect.length -gt 21 )
			{
				$H_5_Redirect_Call_Priority = $H_5_Redirect.Substring(22,1)
				$H_5_Redirect = $H_5_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_5_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_5_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_5_Redirect_Comment = "ERROR: $H_5_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_5_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_5_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_5_RedirectTarget -eq "ERROR" )
		{
			$H_5_Redirect_Comment = "ERROR: Check H_5_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 6
        $H_6_VoiceCommand = $PSAAConfig.H_6_VoiceCommand[$i]
        $H_6_Redirect = $PSAAConfig.H_6_Redirect[$i]
        $H_6_RedirectTarget = $PSAAConfig.H_6_RedirectTarget[$i]
		
        if ( $H_6_Redirect -match "SharedVoicemail" )
        {
            if ( $H_6_Redirect.Substring(16,1) -eq 0 )
            {
                $H_6_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_6_SharedVoicemailTranscription = $true
            }
            if ( $H_6_Redirect.Substring(18,1) -eq 0 )
            {
                $H_6_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_6_SharedVoicemailSuppress = $true
            }

            $H_6_Redirect = $H_6_Redirect.Substring(0,15)
        }

		if ( $H_6_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_6_Redirect.length -gt 19 )
			{
				$H_6_Redirect_Call_Priority = $H_6_Redirect.Substring(20,1)
				$H_6_Redirect = $H_6_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_6_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_6_Redirect.length -gt 21 )
			{
				$H_6_Redirect_Call_Priority = $H_6_Redirect.Substring(22,1)
				$H_6_Redirect = $H_6_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_6_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_6_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_6_Redirect_Comment = "ERROR: $H_6_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_6_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_6_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_6_RedirectTarget -eq "ERROR" )
		{
			$H_6_Redirect_Comment = "ERROR: Check H_6_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 7
        $H_7_VoiceCommand = $PSAAConfig.H_7_VoiceCommand[$i]
        $H_7_Redirect = $PSAAConfig.H_7_Redirect[$i]
        $H_7_RedirectTarget = $PSAAConfig.H_7_RedirectTarget[$i]
		
        if ( $H_7_Redirect -match "SharedVoicemail" )
        {
            if ( $H_7_Redirect.Substring(16,1) -eq 0 )
            {
                $H_7_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_7_SharedVoicemailTranscription = $true
            }
            if ( $H_7_Redirect.Substring(18,1) -eq 0 )
            {
                $H_7_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_7_SharedVoicemailSuppress = $true
            }

            $H_7_Redirect = $H_7_Redirect.Substring(0,15)
        }

		if ( $H_7_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_7_Redirect.length -gt 19 )
			{
				$H_7_Redirect_Call_Priority = $H_7_Redirect.Substring(20,1)
				$H_7_Redirect = $H_7_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_7_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_7_Redirect.length -gt 21 )
			{
				$H_7_Redirect_Call_Priority = $H_7_Redirect.Substring(22,1)
				$H_7_Redirect = $H_7_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_7_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_7_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_7_Redirect_Comment = "ERROR: $H_7_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_7_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_7_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_7_RedirectTarget -eq "ERROR" )
		{
			$H_7_Redirect_Comment = "ERROR: Check H_7_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 8
        $H_8_VoiceCommand = $PSAAConfig.H_8_VoiceCommand[$i]
        $H_8_Redirect = $PSAAConfig.H_8_Redirect[$i]
        $H_8_RedirectTarget = $PSAAConfig.H_8_RedirectTarget[$i]
		
        if ( $H_8_Redirect -match "SharedVoicemail" )
        {
            if ( $H_8_Redirect.Substring(16,1) -eq 0 )
            {
                $H_8_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_8_SharedVoicemailTranscription = $true
            }
            if ( $H_8_Redirect.Substring(18,1) -eq 0 )
            {
                $H_8_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_8_SharedVoicemailSuppress = $true
            }

           $H_8_Redirect = $H_8_Redirect.Substring(0,15)
        }

		if ( $H_8_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_8_Redirect.length -gt 19 )
			{
				$H_8_Redirect_Call_Priority = $H_8_Redirect.Substring(20,1)
				$H_8_Redirect = $H_8_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_8_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_8_Redirect.length -gt 21 )
			{
				$H_8_Redirect_Call_Priority = $H_8_Redirect.Substring(22,1)
				$H_8_Redirect = $H_8_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_8_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_8_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_8_Redirect_Comment = "ERROR: $H_8_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_8_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_8_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_8_RedirectTarget -eq "ERROR" )
		{
			$H_8_Redirect_Comment = "ERROR: Check H_8_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option 9
        $H_9_VoiceCommand = $PSAAConfig.H_9_VoiceCommand[$i]
        $H_9_Redirect = $PSAAConfig.H_9_Redirect[$i]
        $H_9_RedirectTarget = $PSAAConfig.H_9_RedirectTarget[$i]
		
        if ( $H_9_Redirect -match "SharedVoicemail" )
        {
            if ( $H_9_Redirect.Substring(16,1) -eq 0 )
            {
                $H_9_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_9_SharedVoicemailTranscription = $true
            }
            if ( $H_9_Redirect.Substring(18,1) -eq 0 )
            {
                $H_9_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_9_SharedVoicemailSuppress = $true
            }

            $H_9_Redirect = $H_9_Redirect.Substring(0,15)
        }

		if ( $H_9_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_9_Redirect.length -gt 19 )
			{
				$H_9_Redirect_Call_Priority = $H_9_Redirect.Substring(20,1)
				$H_9_Redirect = $H_9_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_9_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_9_Redirect.length -gt 21 )
			{
				$H_9_Redirect_Call_Priority = $H_9_Redirect.Substring(22,1)
				$H_9_Redirect = $H_9_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_9_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_9_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_9_Redirect_Comment = "ERROR: $H_9_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $9 -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_9_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_9_RedirectTarget -eq "ERROR" )
		{
			$H_9_Redirect_Comment = "ERROR: Check H_9_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option *
        $H_Star_VoiceCommand = $PSAAConfig.H_Star_VoiceCommand[$i]
        $H_Star_Redirect = $PSAAConfig.H_Star_Redirect[$i]
        $H_Star_RedirectTarget = $PSAAConfig.H_Star_RedirectTarget[$i]
		
        if ( $H_Star_Redirect -match "SharedVoicemail" )
        {
            if ( $H_Star_Redirect.Substring(16,1) -eq 0 )
            {
                $H_Star_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_Star_SharedVoicemailTranscription = $true
            }
            if ( $H_Star_Redirect.Substring(18,1) -eq 0 )
            {
                $H_Star_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_Star_SharedVoicemailSuppress = $true
            }

            $H_Star_Redirect = $H_Star_Redirect.Substring(0,15)
        }

		if ( $H_Star_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_Star_Redirect.length -gt 19 )
			{
				$H_Star_Redirect_Call_Priority = $H_Star_Redirect.Substring(20,1)
				$H_Star_Redirect = $H_Star_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_Star_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_Star_Redirect.length -gt 21 )
			{
				$H_Star_Redirect_Call_Priority = $H_Star_Redirect.Substring(22,1)
				$H_Star_Redirect = $H_Star_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_Star_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_Star_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_Star_Redirect_Comment = "ERROR: $H_Star_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_Star_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_Star_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_Star_RedirectTarget -eq "ERROR" )
		{
			$H_Star_Redirect_Comment = "ERROR: Check H_Star_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}


		# Holiday - Option #
        $H_Pound_VoiceCommand = $PSAAConfig.H_Pound_VoiceCommand[$i]
        $H_Pound_Redirect = $PSAAConfig.H_Pound_Redirect[$i]
        $H_Pound_RedirectTarget = $PSAAConfig.H_Pound_RedirectTarget[$i]
		
        if ( $H_Pound_Redirect -match "SharedVoicemail" )
        {
            if ( $H_Pound_Redirect.Substring(16,1) -eq 0 )
            {
                $H_Pound_SharedVoicemailTranscription = $false
            }
            else
            {
                $H_Pound_SharedVoicemailTranscription = $true
            }
            if ( $H_Pound_Redirect.Substring(18,1) -eq 0 )
            {
                $H_Pound_SharedVoicemailSuppress = $false
            }
            else
            {
                $H_Pound_SharedVoicemailSuppress = $true
            }

            $H_Pound_Redirect = $H_Pound_Redirect.Substring(0,15)
        }

		if ( $H_Pound_Redirect -match "ApplicationEndpoint" )
        {
			if ( $H_Pound_Redirect.length -gt 19 )
			{
				$H_Pound_Redirect_Call_Priority = $H_Pound_Redirect.Substring(20,1)
				$H_Pound_Redirect = $H_Pound_Redirect.Substring(0,19)
				
				$CallPriority = $true
			}
		}

		if ( $H_Pound_Redirect -match "ConfigurationEndpoint")
        {
			if ( $H_Pound_Redirect.length -gt 21 )
			{
				$H_Pound_Redirect_Call_Priority = $H_Pound_Redirect.Substring(22,1)
				$H_Pound_Redirect = $H_Pound_Redirect.Substring(0,21)
				
				$CallPriority = $true
			}
		}

		if ( $H_Pound_Redirect -EQ "FILE" )
		{
			$FileExists = CheckFileExists $H_Pound_RedirectTarget
			
			if ( ! $FileExists )
			{
				$H_Pound_Redirect_Comment = "ERROR: $H_Pound_RedirectTarget does not exist"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
		}

		if ( $H_Pound_Redirect -eq "Operator" )
		{
			if ( $H_Operator_Redirect )
			{
				$H_Pound_Redirect_Comment = "ERROR: Only one operator transfer allowed"
				$StopProcessing = $true
				$VerboseStopProcessing = $true
			}
			else
			{
				$H_Operator_Redirect = $true
			}
		}

		if ( $H_Pound_RedirectTarget -eq "ERROR" )
		{
			$H_Pound_Redirect_Comment = "ERROR: Check H_Pound_RedirectTarget on Config-HolidayMenu"
			$StopProcessing = $true
			$VerboseStopProcessing = $true
		}

		NewAutoAttendant
	}
}

Write-Host "Completed Auto Attendant Configuration."

if ( Test-Path -path ".\PS-AA.csv" )
{
    Remove-Item -Path ".\PS-AA.csv" | Out-Null
}

exit
