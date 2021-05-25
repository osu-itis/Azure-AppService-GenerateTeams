using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Output "PowerShell HTTP trigger function processed a request."

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

# make a graph api call to find the UserPrincipalName from the email address.
$URI = 'https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq '+"'"+"Team"+"')"

$MSAL = [PSCustomObject]@{
    Token = $token
    Headers = $Headers
}

#$MSAL |Export-Clixml -Path .\MSAL.cli.xml

Write-Output "Gathering raw results"
$Results = Invoke-RestMethod -Headers $MSAL.Headers -Method Get -Uri $URI

#$Results | Export-Clixml -path .\testing.cli.xml

$Collection = @()

do {
    foreach ($item in $Results.value) {
        $Collection += $item
    }
    
    if (($Results.'@odata.nextLink' -eq $null) -eq $false) {
        $URI = $Results.'@odata.nextLink'
        $Results = Invoke-RestMethod -Headers $MSAL.Headers -Method Get -Uri $URI
        Write-Output "Gathering raw results..."
    }
} until (($Results.'@odata.nextLink' -eq $null))

$Collection = $Collection|Select-Object ID, Displayname, mail

Write-Output "Reviewing settings"
$Teams = foreach ($item in $collection) {
    $GuestSettings = Invoke-RestMethod -Headers $MSAL.Headers -Method get -Uri $("https://graph.microsoft.com/v1.0/groups/"+$item.ID+"/settings")
    [PSCustomObject]@{
        Id = $item.id
        DisplayName = $item.displayName
        Mail = $item.mail
        AllowToAddGuests = $GuestSettings.value.values.value
    }
}

Write-Output "Filtering final output"
$FilteredTeams = $teams|Where-Object {$_.AllowToAddGuests -eq $true}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $( $FilteredTeams|ConvertTo-Json )
})
