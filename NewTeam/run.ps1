using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) { Throw 'Could not find $env:AzureWebJobsStorage' }
if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
if ([string]::IsNullOrEmpty($env:CertificateThumbprint)) { Throw 'Could not find $env:CertificateThumbprint' }

#Importing all of the needed files:
Write-Host "Importing the Custom Graph API Token Class"
Import-Module .\Modules\New-GraphAPIToken -Force

Write-Host "Importing the Custom Team Object Class"
. .\NewTeam\CustomTeamObject.ps1

Write-Host "Gathering a new Graph Token"
#Generate the Graph Token info so we can make graph API calls
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID

#Grabbing the body of the request and setting it to a new hashtable
[hashtable]$Queue = $Request.body

#Generate the temporary object that will contain all of our temp variables
$TempObject = [CustomTeamObject]@{
    #Setting all of the items that are inputed from the Azure Queue
    TeamName         = $queue.TeamName
    TeamDescription  = $queue.TeamDescription
    TeamType         = $queue.TeamType
    TicketID         = $queue.TicketID
    Requestor        = $queue.Requestor
    #Using the Graph token info
    GraphTokenString = $GraphAPIToken.TokenString
}

# Write out the queue message and insertion time to the information log
Write-Host "PowerShell queue trigger function processed work item: $($Queue |convertto-json )"

#Generating the new team using the custom class that contains the needed methods
$TempObject.AutoCreateTeam()

#Converting to Json and pushing out to the host for humans to read
Write-Host -message "Final Output:"
Write-Host -message $($TempObject.Results | convertto-json)

#Determine if the the visibility in powershell (This can take a very long time as it requires that Exchange has replicated, so the request is processed by a seperate function)
switch ($TempObject.TeamType) {
    { $_ -eq "Public" } {
        #Do not make any changes to the visibility in the GAL
        Write-Host "No changes made to the visibility in the GAL"
    }
    { $_ -eq "Private" } {
        #Writing the Group ID to the timed-gal-changes storage queue
        Write-Host "Writing to the Azure Storage Queue for followup changes on the GAL"
        Push-OutputBinding -name "TimedGALChanges" -value $TempObject.GroupResults.id
    }
    Default {
        #Writing the Group ID to the timed-gal-changes storage queue
        Write-Host "Writing to the Azure Storage Queue for followup changes on the GAL"
        Push-OutputBinding -name "TimedGALChanges" -value $TempObject.GroupResults.id
    }
}

if ($TempObject.Results.status -eq "SUCCESS") {
    #Writing to the table (for logging purposes)
    write-host "Writing to the Azure Storage Table"
    Push-OutputBinding -Name LoggedTeamsRequests -Value $TempObject.Results

    # Sending the information back in the response
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = $($TempObject.Results | ConvertTo-Json)
        }
    )
}
else {
    # Sending the information back in the response
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = $($TempObject.Results | ConvertTo-Json)
        }
    )

    #If the script failed to create a new team, we want this to throw an error
    Throw "Script failed to generate a new Microsoft Teams team."
}
