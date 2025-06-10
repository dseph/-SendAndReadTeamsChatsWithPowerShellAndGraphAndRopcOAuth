# SendAndReadTeamsChatsWithPowerShellAndGraphAndRopcOAuth.ps1 
# 
# Send and Read Teams Chats with PowerShell and Graph and ROPC.
#  
# ROPC will only work under some conditions, so read the documentaiton carefully. For example, MFA needs to be off. This auth flow should be avoided unless there is some dire need for it.
# Per Azure:
#    99.9% of account compromise could be stopped by using multifactor authentication, which is a feature that security defaults provides.
#    Microsoft's security teams see a drop of 80% in compromise rate when security defaults are enabled.
#
# Reading material and permissions ------------------------------------------

# https://learn.microsoft.com/en-us/powershell/module/teams/?view=teams-ps
# https://learn.microsoft.com/en-us/powershell/module/teams/connect-microsoftteams?view=teams-ps    - See Example 7
# https://learn.microsoft.com/en-us/microsoft-365/admin/security-and-compliance/set-up-multi-factor-authentication?view=o365-worldwide
# https://learn.microsoft.com/en-us/graph/api/channel-list-messages?view=graph-rest-1.0&tabs=http
# https://learn.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0&tabs=http
# https://learn.microsoft.com/en-us/graph/api/channel-list-messages?view=graph-rest-1.0&tabs=http
# https://learn.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-1.0&tabs=http
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/troubleshooting?view=graph-powershell-1.0
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
 
 

# Some points: 
#    A lot of samples on the web don't work or are not complete.
#    Per this article, no client secret is required for ROPC: https://docs.azure.cn/en-us/entra/identity-platform/v2-oauth-ropc
#    Be sure to add each permission carefully per each graph call needed and to keep a record of them in a document - you will need to specify them in code.  
#    Be sure to do an admin grant for all permissions in Azure.  
#    You can check permissions via documenation for the Graph call and also in Graph Explorer (there is a button which will show the permissions needed for the call based-upon the URL.
#    A client secret is not needed since this is a public flow.
#    A redirect is not needed.
#    Here is where to get the Graph PowerShell SDK: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
#    Permissions needed fro the Graph call need to have been Admin granted and must be be in the Scope statement forcode requesting the permissions - if they don't then you will get an admin grant error or a missing permission error.
#    An incorrectly formed scope string may cause an error which does not seem to represent it not being correct - so go back and check the Scope (and other setting) to be sure they are correct.
#    If you get an error saying you need to do admin consent and you have all permissions consented then recheck your Scope string against the permissions in the application and which were granted. Also check the string for type-os.


# MFA Disablement: 
# Its best not to disable it; below here is where to look to turn it off and on:
#     Disable MFA for Tenat:
#         portal.azure.com -> Entra ID -> (will be on the Overview page) -> Properies (second level tab) - > Manage security defaults (link on bottom
#
#     Disable MFA for a user:
#         portal.azure.com -> Entra ID -> Users  -> Per-user MFA (gear icon top right) -> select your tenant

# https://learn.microsoft.com/en-us/graph/api/teams-list?view=graph-rest-1.0&tabs=http   - all teams in org
#    Teams.ReadBasic.All   No
#
#    Teams.Read.All        Yes  higher
#    Teams.ReadWrite.All   Yes  higher
#
# https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams?view=graph-rest-1.0&tabs=http   - List my teams
#   Teams.ReadBasic.All        No 
#
#   Directory.Read.All         Higher
#   Directory.ReadWrite.All    Higher
#   TeamSettings.Read.All      Higher
#   TeamSettings.ReadWrite.All Higher
#   User.Read.All              Higher
#   User.ReadWrite.All         Higher
#
#   GET /me/joinedTeams
#   GET /users/{id | user-principal-name}/joinedTeams
#    https://graph.microsoft.com/v1.0/me/JoinedTeams
#  
#  https://learn.microsoft.com/en-us/graph/api/channel-list?view=graph-rest-1.0&tabs=http  - List channels for a team
#
#  https://learn.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-1.0&tabs=http    - List chats
#     Chat.ReadBasic, Chat.Read, Chat.ReadWrite
#  
#
#  https://learn.microsoft.com/en-us/graph/api/chat-post-messages?view=graph-rest-1.0&tabs=http  - Send Teams chat
#      ChatMessage.Send   
#      Chat.ReadWrite           Higher
#      Group.ReadWrite.All      Hihger
#
#      /chats/{chat-id}/messages
#
# Samples::
#    009c5afa-e3a7-4f95-b4f4-f4db2efd75d3    - Sample Team
#    https://graph.microsoft.com/v1.0/teams/009c5afa-e3a7-4f95-b4f4-f4db2efd75d3/channels
#    https://graph.microsoft.com/v1.0/me/chats
#       "id": "19:cdc596fd-008d-47f1-b042-ae66c018ae4d_d237c4a5-df47-4d2d-8d64-fdcafa3f4154@unq.gbl.spaces"
 
# Important note:  Preview/Beta APIs shoudl not be used in production since they may not work or work correctly and could cause issuess like concorrect updates, incorrect data returned and other things. They may also not be released.
# Also, they do not get mainline support, so you shoudl look to the forrums for assistance for Preview/Beta APIs if you need help trying them.
 
Import-Module Microsoft.Graph.Teams

 
# GraphDelegateNoRedirect
$ClientID = "aaaaaaaa-bbbb-cccc-ddddddddddd"        # The Application ID/Client ID from the application registration in Azure.
$TenantID = "xxxxxxxx-yyyy-zzzz-mmmmmmmmmmmmmmmmm"  # # The tenant ID from the application registration in Azure.
$Username = "somecooluser@contoso.onmicrosoft.com"  #User's ID/SMTP
$Password = "PoiQwe!14159"                          # User's password
$GrantType = "password"

$URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"

$Scope = "openid offline_access user.read ChatMessage.Send Chat.Read Chat.ReadWrite"  # Modify this as per the needs of your Graph calls

# The below line is just antoher sample.
#$Scope = "openid offline_access user.read ChannelMessage.Send ChatMessage.Send ChannelMessage.Read.All ChannelSettings.Read.All Chat.Read Chat.ReadWrite Chat.Create ChatMessage.Read"
 

$Body = @{
   client_id     = $ClientID
   scope         = $Scope
   username      = $Username
   password      = $Password
   grant_type    = $GrantType
}
 
 
Write-Output "Body Start ---------------------------------------"
$Body
Write-Output "Body End ---------------------------------------"
Write-Output ""

$RequestParameters = @{
  URI = $URI
  Method = "POST"
  ContentType = "application/x-www-form-urlencoded"
}


$oAuthToken = (Invoke-RestMethod @RequestParameters -Body $Body).access_token
Write-Output "GraphToken Start ---------------------------------------"
$oAuthToken
Write-Output "GraphToken End ---------------------------------------"
Write-Output ""

Connect-MgGraph -AccessToken ($oAuthToken |ConvertTo-SecureString -AsPlainText -Force)  # Need to convert for the Graph API - older samples may not show this

$DateTime = Get-Date  # This datetime is added to body only for testing purposes.
 
$params = @{
	body = @{
		content = "Hello World - $DateTime" 
	}
}

#$TeamId = "009c5afa-e3a7-4f95-b4f4-f4db2efd75d3"
#$ChannelId = "xxxxxxxxxxxxx"
#$Results = New-MgTeamChannelMessage -TeamId $TeamId -ChannelId $ChannelId -BodyParameter $params

$ChatId = "19:cdc596fd-008d-47f1-b042-ae66c018ae4d_d237c4a5-df47-4d2d-8d64-fdcafa3f4154@unq.gbl.spaces"   # This is a chat ID I got using Graph Explorer (/me/chats)
$Results = New-MgChatMessage -ChatId $ChatId -BodyParameter $params

$Results = Get-MgChatMessage -ChatId $ChatId -Top 3 -Sort "createdDateTime desc" 
#$Results.Body
 
 
#The below works:

#$Results = Invoke-MGGraphRequest -Method get -Uri 'https://graph.microsoft.com/v1.0/me/JoinedTeams' -OutputType PSObject -Headers $headers 

#$Results = Invoke-MGGraphRequest -Method get -Uri 'https://graph.microsoft.com/v1.0/teams/009c5afa-e3a7-4f95-b4f4-f4db2efd75d3/channels' -OutputType PSObject -Headers $headers 

#$Results = Invoke-MGGraphRequest -Method get -Uri 'https://graph.microsoft.com/v1.0/chats/19:cdc596fd-008d-47f1-b042-ae66c018ae4d_d237c4a5-df47-4d2d-8d64-fdcafa3f4154@unq.gbl.spaces/messages?$top=2' -OutputType PSObject -Headers $headers 

#$ChatId = "19:cdc596fd-008d-47f1-b042-ae66c018ae4d_d237c4a5-df47-4d2d-8d64-fdcafa3f4154@unq.gbl.spaces"
#$Results = New-MgChatMessage -ChatId $ChatId -BodyParameter $params
#$Results = Get-MgChatMessage -ChatId $ChatId -Top 2 -Sort "createdDateTime desc" 

 
#Write-Output "Results Start ---------------------------------------"
$Results.Body 
#$Results 
#Write-Output "ResultsEnd ---------------------------------------"
 