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

# Setting the Names and IDs of our groups that we use for the management of Teams.
$Management = @{
    # This is our group for Exceptions
    ExceptionsGroup = @{
        Name = 'm365_exceptionteams'
        ID = '15161ee5-a1a4-454b-a552-19ede6d66c13'
    }
    # This is our group of teams that we know of
    KnownGroup = @{
        Name = 'm365_knownteams'
        ID = 'f514152c-b08d-4c32-97e9-21c7941dd699'
    }
}

function ChunkSplitter {
    <#
    .SYNOPSIS
    Split an array into chunks based on size

    .DESCRIPTION
    Split an array into chunks based on size

    .PARAMETER chunkSize
    Int of the size the array should be split into

    .PARAMETER InputArray
    The array to be split into chunks

    .EXAMPLE
    ChunkSplitter -ChunkSize 20 -InputArray $MyLongArray

    .NOTES
    This was created to split out long arrays into a short enough array that it can be used within Microsoft Graph API Calls
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]
        $chunkSize,
        [Parameter(Mandatory = $true)]
        [array]
        $InputArray
    )

    # Setting a generic outarray
    $outArray = @()

    # Determing the chunks needed based on the InputArray array
    $parts = [math]::Ceiling($InputArray.Count / $chunkSize)

    # Iterating through the array and splitting whenever we hit the max of the chunk
    for ($i = 0; $i -le $parts; $i++) {
        $start = $i * $chunkSize
        $end = (($i + 1) * $chunkSize) - 1
        $outArray += , @($InputArray[$start..$end])
    }

    # Outputting the results into the calculated chunks
    $Output = $outArray | ForEach-Object { "$_" }

    # Returning the Output of the function
    return $Output
}

# Splitting the results into chunks of 20 to fit within Microsoft Graph's maximum number of IDs that can be added to a group at a time.
$ChunkedResults = ChunkSplitter -chunkSize 20 -InputArray $FilteredResults.ID

$GraphMembershipPatches = foreach ($Chunk in $ChunkedResults) {
    # Splitting the chunk by spaces to have a clean array of the IDs
    $Chunk = $Chunk.trim().Split(' ')

    # Formatting the IDs into the needed directory object format
    $DirectoryObjects = foreach ($ID in $Chunk) {
        # If the ID is NOT empty, put it into Graph's directory object format
        if (! [string]::IsNullOrEmpty($ID)) {
            "https://graph.microsoft.com/v1.0/directoryObjects/$ID"
        }
    }

    # A PSCustom Object containing the directory objects
    [PSCustomObject]@{
        "members@odata.bind" = $DirectoryObjects
    }
}

foreach ($item in $GraphMembershipPatches) {
    # Making a rest patch to to add the IDs as members of the group
    $call = @{
        Headers     = $Headers
        Method      = 'Patch'
        # Patching to our group of known Teams
        URI         = "https://graph.microsoft.com/v1.0/groups/$($Management.KnownGroup.ID)"
        ContentType = 'application/json'
        # Converting our graph membership patch to json for the REST call
        Body        = $($item | ConvertTo-Json)
    }
    Invoke-RestMethod @call
}