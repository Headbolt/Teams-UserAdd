###############################################################################################################
#
# ABOUT THIS PROGRAM
#
#   Teams-UserAdd.ps1
#   https://github.com/Headbolt/Teams-UserAdd
#
#   This script was designed to Create a Teams AutoAttendant and corresponding Call Queues
#
###############################################################################################################
#
# HISTORY
#
#   Version: 1.1 - 03/01/2023
#
#   - 12/07/2022 - V1.0 - Created by Headbolt
#
#   - 03/01/2023 - V1.1 - Updated by Headbolt
#							Now Allows for Emergency Location by finding ID
#				
###############################################################################################################
#
#   CUSTOMISABLE VARIABLES FUNCTION
#
function CustomisableVariables
{
	$global:LoggingEnabled="YES" # Enable Logging if needed
	$global:LogFileLocation="C:\Scripts\Teams\AddTeamsNumberToUser-Log.log" # Location Of LogFile if Enabled
	$global:PhoneNumberType="DirectRouting" # Type Of Number, Options are DirectRouting, CallingPlan and OperatorConnect
	$global:Language = "en-GB" # Sets the Laguage for users Voice Services, such as VoiceMail
	$global:OofGreetingFollowAutomaticRepliesEnabled = $true
	$global:CsOnlineVoiceRoutingPolicy = "HSO-RED UK Routing Policy" # Set as required
	$global:CsTenantDialPlan = "UK Dial Plan 1" # Set as required
	$global:CsOnlineVoicemailPolicy = "TranscriptionProfanityMaskingEnabled" # Set as required
	$global:CsDialoutPolicy = "DialoutCPCDisabledPSTNInternational" # Set as required
#
}
#
###############################################################################################################
#
#   START FUNCTION
#
function Logging
{
	if ( $global:LoggingEnabled -eq "YES" )
	{
		Start-Transcript $global:LogFileLocation # Start the logging
		Clear-Host #Clear Screen
		Write-Output "Logging to $global:LogFileLocation"
	}     
}
#
###############################################################################################################
#
#   END FUNCTION
#
function End
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	if ( $global:LoggingEnabled -eq "YES" )
	{
		Stop-Transcript # Stop Logging
	}
	Write-Host "END !!"
	Exit
}
#
###############################################################################################################
#
#   CONNECTIONS FUNCTION
#
function Connections
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host "Connecting To Teams Admin"
	# Connect to Teams
	Connect-MicrosoftTeams # Needs Teams Module https://www.powershellgallery.com/packages/MicrosoftTeams/4.5.0

	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   USER INPUT FUNCTION
#
function UserInput 
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	$global:Input = '' # Ensure Input Variable is Blank
	$global:Input=Read-Host -Prompt "Input the $InputVarible . eg. $InputVaribleExplanation" # Grab the Variable
	if ( "" -ne $global:Input )
	{
#		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host $InputVarible Value gathered is "'$global:Input'"
	}
	else
	{
		Write-Host Input was Blank, Ending Script
		End
	}
}
#
###############################################################################################################
#
#   COLLECT VARIABLES FUNCTION
#
function CollectVariables 
{
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Gathering Required Data'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="UserUPN"
	$InputVaribleExplanation="User.Name@domain.com"
	UserInput
	$global:UserUPN=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="OnpremPhoneNumber"
	$InputVaribleExplanation="+441112223333"
	UserInput
	$global:OnpremPhoneNumber=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
}
#
###############################################################################################################
#
#   EMERCENCY LOCATION FUNCTION
#
function EmergencyLocation 
{
	$InputVarible="City"
	$InputVaribleExplanation="London"
	UserInput
	$global:City=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$Locations = Get-CsOnlineLisLocation | 
    select City, StreetName, CompanyName, LocationId | 
    sort City, StreetName, CompanyName
	#
	$GridArguments = @{
		OutputMode = 'Single'
		Title      = 'Please select a location and click OK'
	}
	#
	$Location = $Locations | Out-GridView @GridArguments | foreach {
		$_.LocationId
	}
	#
	if ($Location) {
		$Arguments = @{
			LocationID      = $Location
		}
	}
}
#
###############################################################################################################
#
#   PROCESS USER FUNCTION
#
function ProcessUser 
{
#
Write-Host '' # Output To Make Screen Easier for User to read.
#
Write-Host 'Running Command "'Set-CsPhoneNumberAssignment -Identity "$global:UserUPN" -PhoneNumber "$global:OnpremPhoneNumber" -PhoneNumberType "$global:PhoneNumberType" -LocationID $Location'"'
Set-CsPhoneNumberAssignment -Identity "$global:UserUPN" -PhoneNumber "$global:OnpremPhoneNumber" -PhoneNumberType "$global:PhoneNumberType" -LocationID $Location
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Set-CsPhoneNumberAssignment -Identity "$UserUPN" -EnterpriseVoiceEnabled $true'"'
Set-CsPhoneNumberAssignment -Identity "$UserUPN" -EnterpriseVoiceEnabled $true
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Set-CsOnlineVoicemailUserSettings -Identity sip:"$UserUPN" -VoicemailEnabled $true -PromptLanguage $global:Language -OofGreetingFollowAutomaticRepliesEnabled $global:OofGreetingFollowAutomaticRepliesEnabled'"'
Set-CsOnlineVoicemailUserSettings -Identity sip:"$UserUPN" -VoicemailEnabled $true -PromptLanguage $global:Language -OofGreetingFollowAutomaticRepliesEnabled $global:OofGreetingFollowAutomaticRepliesEnabled
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Grant-CsOnlineVoiceRoutingPolicy -Identity "$UserUPN" -PolicyName "$global:CsOnlineVoiceRoutingPolicy"'"'
Grant-CsOnlineVoiceRoutingPolicy -Identity "$UserUPN" -PolicyName "$global:CsOnlineVoiceRoutingPolicy"
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Grant-CsTenantDialPlan -Identity "$UserUPN" -PolicyName "$global:CsTenantDialPlan"'"'
Grant-CsTenantDialPlan -Identity "$UserUPN" -PolicyName "$global:CsTenantDialPlan"
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Grant-CsOnlineVoicemailPolicy -PolicyName "$global:CsOnlineVoicemailPolicy" -Identity "$UserUPN"'"'
Grant-CsOnlineVoicemailPolicy -PolicyName "$global:CsOnlineVoicemailPolicy" -Identity "$UserUPN"
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host 'Running Command "'Grant-CsDialoutPolicy -identity "$UserUPN" -PolicyName "$global:CsDialoutPolicy"'"'
Grant-CsDialoutPolicy -identity "$UserUPN" -PolicyName "$global:CsDialoutPolicy"
#
Write-Host '' # Output To Make Screen Easier for User to read.
Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
#
}
#
###############################################################################################################
#
#   END OF FUNCTION DEFENITION
#
###############################################################################################################
#
#   BEGIN PROCESSING
#
###############################################################################################################
#
Write-Host '' # Output To Make Screen Easier for User to read.
#
CustomisableVariables
#
Logging
#
CollectVariables
#
#Connections
#
EmergencyLocation 
#
ProcessUser
#
End
