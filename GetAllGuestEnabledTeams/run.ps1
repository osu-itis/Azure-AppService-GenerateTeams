using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Output "PowerShell HTTP trigger function processed a request."

# Attempting to import the needed Modules
import-module .\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1

# Checking if the needed ENVs exist
. .\Shared\Check-ENVs.ps1

# Gathering a token and setting the headers
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$Headers = $GraphAPIToken.Headers

# make a graph api call to find the UserPrincipalName from the email address.
$URI = 'https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq '+"'"+"Team"+"')"

Write-Output "Gathering raw results"
$Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI

#$Results | Export-Clixml -path .\testing.cli.xml

$Collection = @()

do {
    foreach ($item in $Results.value) {
        $Collection += $item
    }
    
    if (($Results.'@odata.nextLink' -eq $null) -eq $false) {
        $URI = $Results.'@odata.nextLink'
        $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI
        Write-Output "Gathering raw results..."
    }
} until (($Results.'@odata.nextLink' -eq $null))

$Collection = $Collection|Select-Object ID, Displayname, mail

Write-Output "Reviewing settings"
$Teams = foreach ($item in $collection) {
    $GuestSettings = Invoke-RestMethod -Headers $Headers -Method get -Uri $("https://graph.microsoft.com/v1.0/groups/"+$item.ID+"/settings")
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
