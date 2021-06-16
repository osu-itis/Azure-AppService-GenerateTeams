#Importing code for envs and graph token
############################################################################################################
import-module ".\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1"
. C:\Users\carrk\GitHub\CodeSnippet-Azure-AutoLoadENVs\AutoLoadENVs.ps1
AutoLoadENVs
$token = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$token.NewOAuthRequest()
############################################################################################################

# Getting the Group.Unified.Guest Template (Note we need this when posting a request)
$GroupUnifiedGuestTemplate = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/directorySettingTemplates/" -Method Get -Headers $token.Headers).value | Where-Object {$_.DisplayName -eq 'Group.Unified.Guest'}
$GroupUnifiedGuestTemplate

# The name of the group to find
$TeamName = "Keenan Testing Teams 4"

# Finding the group that matches the team name, filtering to ensure that we only have one match and then getting the group id, wich will be used for all future graph calls
$Results = Invoke-RestMethod -Headers $token.Headers -Method Get -Uri $('https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,' + "'"+$TeamName+"')")
$GroupID = $Results.value |Where-Object {$_.displayName -eq $TeamName}|Select-Object -ExpandProperty ID 

# Getting the group
$Group = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID" -Method "get" -Headers $token.Headers
$Group

# Getting the group settings (note the settings ID. we'll use that later)
$GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $token.Headers
$GroupInfo.value
       
# Creating the body of the patch request
$Post = [PSCustomObject]@{
    displayName = "GroupSettings"
    templateId = $GroupUnifiedGuestTemplate.id
    values = @(
        @{
            name = 'AllowToAddGuests'
            value = "true"
        }
    )
} 

# Converting the body to json
$PostJson = $Post | ConvertTo-Json
$PostJson

# If settings have been applied before, we need to Patch
if (($GroupInfo.value.displayname| Where-Object {$_ -eq "Group.Unified.Guest"}).count -eq 1) {
    # Making patch request following the example provided at "https://docs.microsoft.com/en-us/graph/api/groupsetting-update?view=graph-rest-1.0&tabs=http"
    # Patching the current settings using the group id and the group's settings id
    Write-Warning -Message "Patching settings"
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings/$($GroupInfo.value.id)" -Method Patch -Headers $token.Headers -Body $PostJson -ContentType 'application/json'
}
# Else we can Post
else {
    Write-Warning -Message "Posting settings"
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method Post -Headers $token.Headers -Body $PostJson -ContentType 'application/json'
}

# Getting the group settings
$GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $token.Headers
$GroupInfo.value

########################################################################
# Delete the settings
#Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings/$($GroupInfo.value.id)" -Method Delete -Headers $token.Headers
########################################################################