using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Attempting to import the needed Modules
try
{
    Write-Output "Importing the MSAL.PS module"
    Import-Module .\Modules\MSAL.PS\4.21.0.1\MSAL.PS.psd1 -Force -ErrorAction stop
} catch
{
    Throw 'Failed to import the MSAL.PS Module'
}

# Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage))
{ Throw 'Could not find $env:AzureWebJobsStorage'
}
if ([string]::IsNullOrEmpty($env:ClientID))
{ Throw 'Could not find $env:ClientID'
}
if ([string]::IsNullOrEmpty($env:ClientSecret))
{ Throw 'Could not find $env:ClientSecret'
}
if ([string]::IsNullOrEmpty($env:TenantId))
{ Throw 'Could not find $env:TenantId'
}

# Using MSAL to generate and manage a (JWT) token
Write-Output "Generating token"
$token = Get-MsalToken -ClientId $env:ClientID -ClientSecret $(ConvertTo-SecureString $env:ClientSecret -AsPlainText -Force) -TenantId $env:TenantID

# Setting the headers as a variable for convienience
$Headers = @{Authorization = "Bearer $($token.AccessToken)" }

# Combinding into a custom object to be reused later
$MSAL = [PSCustomObject]@{
    Token = $token
    Headers = $Headers
}

# Defining functions
function Get-TeamInfo {
    <#
    .SYNOPSIS
    Gathers information about the guest access of a team

    .DESCRIPTION
    Gathers information about the guest access of a team

    .PARAMETER TeamName
    The display name of the team to inspect

    .EXAMPLE
    $GuestSettings = get-teamInfo -teamName "some team name"
    #>
    param ($TeamName)

    # Get team general info
    $URI = 'https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,' + "'"+$TeamName+"')"
    $Results = Invoke-RestMethod -Headers $MSAL.Headers -Method Get -Uri $URI
    $TeamInfo = $Results.value|Select-Object ID,Mail,displayname

    # Get current team guest access setting
    $results = Invoke-RestMethod -headers $MSAL.Headers -Method get -uri $(
        "https://graph.microsoft.com/v1.0/groups/"+$TeamInfo.id+"/settings"
    )

    Add-Member -InputObject $TeamInfo -NotePropertyName "AllowToAddGuests" -NotePropertyValue $Results.value.values.value
    Add-Member -InputObject $TeamInfo -NotePropertyName "GuestSettingsID" -NotePropertyValue $Results.value.id
    Return $TeamInfo
}

function Set-TeamInfo {
    <#
    .SYNOPSIS
    Sets the AllowToAddGuests setting

    .DESCRIPTION
    Sets the AllowToAddGuests setting

    .PARAMETER Headers
    The headers to make the post/patch request to the graph api

    .PARAMETER GroupID
    The ID of the group (team) that is to be modified

    .PARAMETER value
    The true/false value in string format

    .EXAMPLE
    Set-TeamInfo -Headers $Headers -GroupID <GUID> -value true
    #>
    [CmdletBinding()]
    param (
        $Headers,
        $GroupID,
        $value
    )

    $Func = [PSCustomObject]@{
        Headers = $Headers
        GroupID = $GroupID
        body = "{
            `n    `"displayName`": `"GroupSettings`",
            `n    `"templateId`": `"08d542b9-071f-4e16-94b0-74abb372e3d9`",
            `n    `"values`": [
            `n        {
            `n            `"name`": `"AllowToAddGuests`",
            `n            `"value`": `"$value`"
            `n        }
            `n    ]
            `n}"
    }

        $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($Func.GroupID)" -Method "get" -Headers $func.Headers
        $GroupSettingsInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($Func.GroupID)/settings" -Method "get" -Headers $func.Headers

    try {
        #Try to post the information first. If the setting has never been applied this will add it
        Write-Output "Attempting post to settings"
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($Func.GroupID)/settings" -Method "post" -Headers $func.Headers -Body $func.body -ContentType 'application/json'
    }
    catch {
        try {
            #If posting fails, try patching the information since the setting already exists on the object
            Write-Output "Attempting patch to settings"
            Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($func.GroupID)/settings/$($GroupSettingsInfo.value.id)" -Method 'patch' -Body $func.body -ContentType 'application/json'  -Headers $msal.Headers
        }
        catch {
            #Sometimes the resource is not available and needs to be tried again in a few moments
            start-sleep -Seconds 5
            Write-Output "Attempting patch to settings"
            Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($func.GroupID)/settings/$($GroupSettingsInfo.value.id)" -Method 'patch' -Body $func.body -ContentType 'application/json'  -Headers $msal.Headers
        }
    }
}

# Determine if this is a get or a post, and gather the request to process.
switch ($request.Method)
{
    'Get' {
        $RequestToProcess = $Request.Query
        if (-not $RequestToProcess.TeamName) {
            Write-Output "Sending bad request response"
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body = $("Request not formatted properly")
            })
            break
        }
        Write-output "Gathering team info"
        $body = $(get-teamInfo -teamName $RequestToProcess.TeamName)

        # Associate values to output bindings by calling 'Push-OutputBinding'.
        Write-Output "Sending response"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body = $($body|ConvertTo-Json)
        })
    }
    'Post' {
        $RequestToProcess = $Request.Body
        if (-not $RequestToProcess.TeamName) {
            Write-Output "Sending bad request response"
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body = $("Request not formatted properly")
            })
            break
        }
        Write-Output "Request: $($RequestToProcess|ConvertTo-Json)"
        Write-Output "Processing request for ticket $($RequestToProcess.TicketID)"
        Write-output "Gathering team info"
        $GuestSettings = get-teamInfo -teamName $RequestToProcess.TeamName
        Write-output "Setting team info"
        Set-TeamInfo -Headers $MSAL.Headers -GroupID $GuestSettings.id -value $RequestToProcess.GuestSettingsEnabled
        Write-output "Gathering team info"
        $Body = get-teamInfo -teamName $RequestToProcess.TeamName

        # Associate values to output bindings by calling 'Push-OutputBinding'.
        Write-Output "Sending response"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body = $($body|ConvertTo-Json)
        })
    }
    Default {
        Write-Error "Could not determine request type."
    }
}
