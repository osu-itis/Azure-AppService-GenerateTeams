# Input bindings are passed in via param block.
param($Queue, $TriggerMetadata)

$StartErrorCount = [int]$(
    $error.Count
)

#Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) { Throw 'Could not find $env:AzureWebJobsStorage' }
if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
if ([string]::IsNullOrEmpty($env:CertificateThumbprint)) { Throw 'Could not find $env:CertificateThumbprint' }
if ([string]::IsNullOrEmpty($env:ServiceAccountUsername)) { Throw 'Could not find $env:ServiceAccountUsername' }
if ([string]::IsNullOrEmpty($env:ServiceAccountPassword)) { Throw 'Could not find $env:ServiceAccountPassword' }

#Importing all of the needed files:
. .\QueueTrigger\ClientInfo.ps1
. .\QueueTrigger\CustomTeamObject.ps1

#Generate the client info so we can make graph API calls
$ClientInfo = [GraphAPIToken]::new($env:ClientID, $env:ClientSecret, $env:TenantID)

#Generate the temporary object that will contain all of our temp variables
$TempObject = [CustomTeamObject]@{
    #Setting all of the items that are inputed from the Azure Queue
    TeamName                 = $queue.TeamName
    TeamDescription          = $queue.TeamDescription
    TeamType                 = $queue.TeamType
    TicketID                 = $queue.TicketID
    Requestor                = $queue.Requestor
    CallbackID               = $queue.CallbackID
    
    #Using the Graph token info
    GraphTokenString         = $ClientInfo.TokenString
    
    #Generating the service credential that is needed to make powershell calls
    ServiceAccountCredential = $(
        #Using the env variables set in the azure function app, generate a credential object
        $userName = $env:ServiceAccountUsername
        $userPassword = $env:ServiceAccountPassword
        [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
        [pscredential]$o365cred = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
        $o365cred
    )
}

# Write out the queue message and insertion time to the information log.
Write-Host "PowerShell queue trigger function processed work item: $($Queue |convertto-json )"
Write-Host "Queue item insertion time: $($TriggerMetadata.InsertionTime)"

$TempObject.ExportLastObject()

#Generate the new group
write-host "Generating a new group request via graph api"
$TempObject.NewGraphGroupRequest()

$TempObject.ExportLastObject()

#Wait for a few moments
Start-Sleep -Seconds 15

#Generate the new team (from the existing group)
write-host "Generating a new teams request via graph api"
$TempObject.NewGraphTeamRequest()

$TempObject.ExportLastObject()

#Wait for a few moments
Start-Sleep -Seconds 15

#Set the visibility in powershell (if needed) This can take a very long time as it requires that Exchange has replicated
switch ($TempObject.TeamType) {
    { $_ -eq "Public+Team" } {
        #Do not make any changes to the visibility in the GAL
        write-host "No changes made to the visibility in the GAL"
    }
    { $_ -eq "Private+Team" } {
        write-host "Using Powershell to hide visibility in the GAL"
        $TempObject.SetVisibilityInPowershell()
    }
    Default {
        write-host "Unable to determine team type, attempting to hide visibility in the GAL"
        $TempObject.SetVisibilityInPowershell()
    }
}

#Generate the results
write-host "Gathering a report of the results"
$TempObject.GenerateResults()

#Adding any logging or error information needed
write-host "Gathering logging and error information"
$TempObject.Results | Add-Member -NotePropertyMembers @{
    ErrorCount = $error.Count
    Errors     = $(
        if ($StartErrorCount -ne $error.Count) {
            $Error.Exception.Message
        }
    )
}

#Converting to Json and pushing out to the host for humans to read
Write-Host -message "Final Output:"
Write-Host -message $($TempObject.Results | convertto-json)

$TempObject.ExportLastObject()

#Writing to the table (for logging purposes)
write-host "Writing to the Azure Storage Table"
Push-OutputBinding -Name LoggedTeamsRequests -Value $TempObject.Results

#If the script failed to create a new team, we want this to throw an error
if ($TempObject.Results.status -eq "FAILED") {
    Throw "Script failed to generate a new Microsoft Teams team."
}
