#$RawData = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/862ae4fb-232d-49eb-a75c-5d9b17111ca6/settings" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"
$BADData = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/862ae4fb-232d-49eb-a75c-5d9b17111ca6/settings" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"


# Getting the Global Group Settings
$GGS = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groupsettings" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"


##Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/862ae4fb-232d-49eb-a75c-5d9b17111ca6/settings/b93e1646-bb70-4c67-9185-c9aade438055" -Authentication Bearer -Token $obj.GraphTokenString -Method "delete"






$GuestSetting = [PSCustomObject]@{
    '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#groupSettings'
    value            = @([pscustomobject]@{
            displayName = [string]'Group.Unified.Guest'
            templateID  = [string]$GGS.value.templateID
            values      = @([PSCustomObject]@{
                    Name  = [string]'AllowToAddGuests'
                    value = [string]'true'
                })
        })
}

Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/862ae4fb-232d-49eb-a75c-5d9b17111ca6/settings" -Authentication Bearer -Token $obj.GraphTokenString -Method "post" -Body $($GuestSetting|ConvertTo-Json -Depth 5) -ContentType 'application/json'


