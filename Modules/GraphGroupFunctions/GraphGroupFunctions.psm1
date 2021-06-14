function Get-GraphGroupGuestTemplate {
    <#
    .SYNOPSIS
    Get the Group.Unified.Guest settings template
    
    .DESCRIPTION
    Get the Group.Unified.Guest settings template
    
    .PARAMETER Headers
    The headers to use when making Graph API calls
    
    .EXAMPLE
    Get-GraphGroupGuestTemplate -Headers $token.Headers
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$Headers
        )
        
    # Getting the Group.Unified.Guest Template (Note we need this when posting a request)
    $GroupUnifiedGuestTemplate = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/directorySettingTemplates/" -Method Get -Headers $Headers).value | Where-Object {$_.DisplayName -eq 'Group.Unified.Guest'}
    return $GroupUnifiedGuestTemplate
}

function Get-GraphGroup {
    <#
    .SYNOPSIS
    Get group information for a given team using the displayname of the team
    
    .DESCRIPTION
    Get group information for a given team using the displayname of the team
    
    .PARAMETER Headers
    The token headers to use when making a Graph API call
    
    .PARAMETER TeamName
    The displayname of the team to search
    
    .EXAMPLE
    $Group = Get-GraphGroup -Headers $token.Headers -TeamName <TEAMNAME>
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$Headers,
        [parameter(Mandatory=$true)]$TeamName
        )
        
    # Finding the group that matches the team name, filtering to ensure that we only have one match and then getting the group id, wich will be used for all future graph calls
    $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $('https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,' + "'"+$TeamName+"')")
    $GroupID = $Results.value |Where-Object {$_.displayName -eq $TeamName}|Select-Object -ExpandProperty ID 
    
    # Getting the group
    $Group = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID" -Method "get" -Headers $Headers
    return $Group
}

function Get-GraphGroupGuestSettings {
    <#
    .SYNOPSIS
    Get the Guest settings for the given team
    
    .DESCRIPTION
    Get the Guest settings for the given team
    
    .PARAMETER Headers
    The token headers to use when making a Graph API call
    
    .PARAMETER GroupID
    The group ID of the team
    
    .EXAMPLE
    Get-GraphGroupGuestSettings -Headers $token.Headers -GroupID $group.id
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$Headers,
        [parameter(Mandatory=$true)]$GroupID
    )
    
    # Getting the group settings (note the settings ID. we'll use that later)
    $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $Headers
    Return $GroupInfo.value
}

function Set-GraphGroupGuestSettings {
    <#
    .SYNOPSIS
    Set the group guest settings for a given team
    
    .DESCRIPTION
    Set the group guest settings for a given team
    
    .PARAMETER Headers
    The token headers to use when making a Graph API call
    
    .PARAMETER GroupUnifiedGuestTemplateID
    The Group Unified Guest Template ID
    
    .PARAMETER GroupID
    The ID of the group for a team
    
    .PARAMETER AllowToAddGuests
    True/FaLse in string format, allows or disallows adding guests to a team
    
    .EXAMPLE
    Set-GraphGroupGuestSettings -Headers $token.Headers -GroupUnifiedGuestTemplateID $GuestTemplate.id -GroupID $Group.id -AllowToAddGuests "True"
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$Headers,
        [parameter(Mandatory=$true)]$GroupUnifiedGuestTemplateID,
        [parameter(Mandatory=$true)]$GroupID,
        [parameter(Mandatory=$true)][ValidateSet("True","False")]$AllowToAddGuests
    )
    
    # Creating the body of the patch request
    $Post = [PSCustomObject]@{
        displayName = "GroupSettings"
        templateId = $GroupUnifiedGuestTemplateID
        values = @(
            @{
                name = 'AllowToAddGuests'
                value = $AllowToAddGuests
            }
        )
    } 

    # Converting the body to json
    $PostJson = $Post | ConvertTo-Json

    # Getting the group settings (note the settings ID. we'll use that later)
    $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $Headers

    # If settings have been applied before, we need to Patch
    if (($GroupInfo.value.displayname| Where-Object {$_ -eq "Group.Unified.Guest"}).count -eq 1) {
        # Making patch request following the example provided at "https://docs.microsoft.com/en-us/graph/api/groupsetting-update?view=graph-rest-1.0&tabs=http"
        # Patching the current settings using the group id and the group's settings id
        Write-Verbose -Message "Patching settings" -Verbose
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings/$($GroupInfo.value.id)" -Method Patch -Headers $Headers -Body $PostJson -ContentType 'application/json'
    }
    # Else we can Post
    else {
        Write-Verbose -Message "Posting settings" -Verbose
        $Null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method Post -Headers $Headers -Body $PostJson -ContentType 'application/json'
    }

    # Getting the group settings
    $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $Headers
    return $GroupInfo.value
}

function Remove-GraphGroupGuestSettings {
    <#
    .SYNOPSIS
    Remove the guest settings for a given team
    
    .DESCRIPTION
    Remove the guest settings for a given team
    
    .PARAMETER Headers
    The token headers to use when making a Graph API call
    
    .PARAMETER GroupID
    The group ID of the team
    
    .EXAMPLE
    Remove-GraphGroupGuestSettings -Headers $token.Headers -GroupID $group.id
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]$Headers,
        [parameter(Mandatory=$true)]$GroupID    
    )
    # Getting the group settings (note the settings ID. we'll use that later)
    $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings" -Method "get" -Headers $Headers

    # Delete the settings
    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupID/settings/$($GroupInfo.value.id)" -Method Delete -Headers $Headers
}

############################################################################################################
# Example using the Graph Group Functions together:
#
# import-module ".\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1"
# $token = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
#
# The name of the group to find
# $TeamName = "Testing Teams"
#
# Getting the template
# $GuestTemplate = Get-GraphGroupGuestTemplate -Headers $token.Headers
#
# Getting the graph group
# $Group = Get-GraphGroup -Headers $token.Headers -TeamName $TeamName
#
# Getting the graph group settings
# Get-GraphGroupGuestSettings -Headers $token.Headers -GroupID $group.id
#
# Setting the graph group settings
# Set-GraphGroupGuestSettings -Headers $token.Headers -GroupUnifiedGuestTemplateID $GuestTemplate.id -GroupID $Group.id -AllowToAddGuests True
#
# Removing the setting (for testing purposes)
# Remove-GraphGroupGuestSettings -Headers $token.Headers -GroupID $group.id
#
############################################################################################################
