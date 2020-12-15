# Loading the expected ENVs that would exist when running in production
. C:\users\carrk\GitHub\CodeSnippet-Azure-AutoLoadENVs\AutoLoadENVs.ps1
AutoLoadENVs

# Using MSAL to generate and manage a (JWT) token
$token = Get-MsalToken -ClientId $env:ClientID -ClientSecret $(ConvertTo-SecureString $env:ClientSecret -AsPlainText -Force) -TenantId $env:TenantID

# Setting the headers as a variable for convienience
$Headers = @{Authorization = "Bearer $($token.AccessToken)" }

# Setting the Graph URI, We want to filter all unifed groups and then ensure that the output has both the ID and the ResourceProvisioningOptions
$URI = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(a:a eq 'unified')&$select=id,resourceProvisioningOptions"

# New empty array that we'll use to store our list of results
$CollectedResults = [array]@()

# Making a rest call to get our initial results
$Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI

# Collecting the first batch of results
$CollectedResults += $Results.value

# If there is a 'nextLink' then there are additional results to be gathered, collect the next batch of results
do {
    $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $Results.'@odata.nextLink'
    $CollectedResults += $Results.value
} until ($null -eq $Results.'@odata.nextLink')

# Now that we have all of the results, we can filter based on team (Right now this is our best option, there are some beta api calls to do this in a more simple way)
$FilteredResults = $CollectedResults.Where({$_.resourceProvisioningOptions -like "*Team*"})
