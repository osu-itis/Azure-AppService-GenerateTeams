# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

# Attempting to import the needed Modules
import-module .\Modules\New-GraphAPIToken\New-GraphAPIToken.psm1

# Checking if the needed ENVs exist
. .\Shared\Check-ENVs.ps1

# Gathering a token and setting the headers
$GraphAPIToken = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
$Headers = $GraphAPIToken.Headers

# Setting the Graph URI, We want to filter all unifed groups and then ensure that the output has both the ID and the ResourceProvisioningOptions
$URI = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(a:a eq 'unified')&$select=id,resourceProvisioningOptions"

# New empty array that we'll use to store our list of results
$CollectedResults = [array]@()

# Making a rest call to get our initial results
Write-Output "Gathering results..."
$Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI

# Collecting the first batch of results
$CollectedResults += $Results.value

# If there is a 'nextLink' then there are additional results to be gathered, collect the next batch of results
do {
    $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $Results.'@odata.nextLink'
    $CollectedResults += $Results.value
} until ($null -eq $Results.'@odata.nextLink')

# Now that we have all of the results, we can filter based on team (Right now this is our best option, there are some beta api calls to do this in a more simple way)
$FilteredResults = $CollectedResults.Where( { $_.resourceProvisioningOptions -like "*Team*" })
Write-Output "Found $($FilteredResults.count) Teams"

# Generating the service credential that is needed to make powershell calls
$ServiceAccountCredential = $(
    # Using the env variables set in the azure function app, generate a credential object
    $userName = $env:ClientID
    $userPassword = $env:ClientSecret
    [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
    [pscredential]$cred = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
    $cred
)

# Connecting to Azure using our service credentials
Write-Output "Connecting to Azure with App Registration"
$AZAccountInformation = Connect-AzAccount -Credential $ServiceAccountCredential -Tenant $env:TenantId -ServicePrincipal

# Creating the storage context needed to query Azure tables
$CTX = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage | Select-Object -ExpandProperty Context

# Getting the TableName from the function's parameters
$tableName = ( (Get-Content -Path .\TimedGroupCheck\function.json | ConvertFrom-Json).bindings.tablename | Out-String ).trim()

# Gathering the cloud table information
$cloudTable = (Get-AzStorageTable –Name $tableName –Context $ctx).CloudTable

# Gathering the known and exempt teams, starting with empty arrays
$KnownTeams = [array]@()
$ExemptTeams = [array]@()
$KnownTeams += Get-AzTableRow -table $cloudTable -partitionKey 'KnownTeams'
$ExemptTeams += Get-AzTableRow -Table $cloudTable -CustomFilter "(Exempt eq true)"
Write-Output "Pulled previous information, $($KnownTeams.count) TOTAL KnownTeams, $($ExemptTeams.count) ExemptTeams"

# Determining the new teams to be added to the list
$NewTeams = ($FilteredResults | Where-Object { ($KnownTeams.ID -notcontains $_.ID) })

if ($NewTeams.count -gt 0) {
    Write-Output "Found $($NewTeams.count) new teams:"
    foreach ($item in $NewTeams) {
        Write-Output "`t $($item.DisplayName) ($($item.ID))"
    }

    # Adding the partition and rowkey information, additionally only selecting some of the properties
    $output = foreach ($item in $NewTeams) {
        $item | Select-Object @{Name = 'PartitionKey'; Expression = { "KnownTeams" } }, @{Name = 'RowKey'; Expression = { $_.id } }, @{Name = 'Exempt'; Expression = { $false } }, id, displayName, description, mail, mailEnabled, mailNickname, createdDateTime, renewedDateTime, expirationDateTime, securityIdentifier, visibility
    }

    try {
        # Pushing our new teams to the table
        $output | Push-OutputBinding -Name outputTable
    }
    catch {
        throw "Failed to update the Teams to the table"
    }
}
