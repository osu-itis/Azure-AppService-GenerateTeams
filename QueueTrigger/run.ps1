# Input bindings are passed in via param block.
param($QueueItem, $TriggerMetadata)

#Creating a new variable as our temp object (as a hash table)
[hashtable]$TempObject = $QueueItem

#Write-Host "PowerShell queue trigger function processed work item: $($QueueItem.Keys|convertto-json)"
#Write-Host "Queue item insertion time: $($TriggerMetadata.InsertionTime)"

#Fixing any email addresses that improperly formatted in the hashtable
$TempObject.TeamOwner = $TempObject.TeamOwner.tostring().replace("%40", "@")
$TempObject.Requestor = $TempObject.Requestor.tostring().replace("%40", "@")

#Creating the needed credentials to access Teams:
$ServiceAccountCredentials = New-Object System.Management.Automation.PSCredential (
    "$ENV:ServiceAccountUsername",
    $(ConvertTo-SecureString "$($ENV:ServiceAccountPassword)" -AsPlainText -Force)
)

#LOOKS LIKE STUFF IS BROKEN IN POWERSHELL CORE FOR AZURE FUNCTIONS ^^^^^^^^ https://github.com/MicrosoftDocs/office-docs-powershell/issues/5950


# $ServiceAccountCredentials = New-Object System.Management.Automation.PSCredential (
#    "$ENV:ServiceAccountUsername",
#    $(ConvertTo-SecureString "$((Get-AzKeyVaultSecret -vaultName "TeamsAutomationKeyVault" -name "TeamsAutomationSecret").SecretValueText)" -AsPlainText -Force)
# )

#function to format and create the new MS team
function New-MSTeams {
    [CmdletBinding()]
    PARAM (
        [parameter(Mandatory = $true)][string]$TeamName,
        [parameter(Mandatory = $true)][validateset("Public", "Private")][string]$TeamType,
        [parameter(Mandatory = $true)][string]$TeamOwner,
        [parameter(Mandatory = $false)][string]$TeamDescription
    )
    begin {
        #Connecting to Teams:
        Connect-MicrosoftTeams -Credential $ServiceAccountCredentials #-AccountID $env:ServiceAccountUsername -Verbose
    }
    Process {
        #Generating the new team with some basic consistent settings
        $TeamInfo = New-Team -DisplayName $TeamName -MailNickName $($TeamName.replace(" ", "")) -Owner $TeamOwner -Visibility $TeamType -Description $TeamDescription
        #Outputting the results
        $TeamInfo
    }
    End {
        Disconnect-MicrosoftTeams
    }
}

#Try to make the new team and report back the details and the status
try {
    $Results3 = New-MSTeams -TeamName $TempObject.TeamName -TeamType $TempObject.TeamType -TeamOwner $TempObject.TeamOwner
    $Status = "Passed"
}
catch {
    #This catch is activated if the function fails to run due to a bad or missing parameter, or a failed request to generate the new team
    $results3 = $null
    $Status = "Failed"
}

#Creating the table logging (needed table attributes)
$TabbleLogging = [hashtable]@{  
    partitionKey = 'TeamsLog'  
    rowKey       = (new-guid).guid  
}

#Adding our temp attributes, the table attributes & the status of the group
$Output = $TabbleLogging + $TempObject + @{
    Status  = $Status
    Results = $results3
}

#Converting to Json and pushing out to the host for humans to read
Write-Host -message $($Output | convertto-json)

#Writing to the table (for logging purposes)
Push-OutputBinding -Name LoggedTeamInstalls -Value $Output
