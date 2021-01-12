using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Writing out to a file for testing
$Request.body | Export-Clixml -Path .\GARBAGE.CLI.XML

# Setting this to a new variable
$RequestBody = $Request.body

# Gathering the three values needed from the request body
$FilterParameter = $($RequestBody.Keys -notlike 'exempt').tostring()
$FilterValue = $($RequestBody."$($FilterParameter)").tostring()
$Exempt = $(
    switch ($RequestBody.Exempt) {
        'true' { $true }
        'false' { $false }
    }
)

# Attempting to import the needed Modules
try {
    Write-Output "Importing the MSAL.PS module"
    Import-Module .\modules\MSAL.PS\4.21.0.1\MSAL.PS.psd1 -Force
    Write-Output "Importing the 'CustomToolkit' module"
    Import-Module .\Modules\CustomToolkit\CustomToolkit.psm1 -Force
}
catch {
    Throw 'Failed to import the required modules'
}

# Gathering the needed credentials
$ServiceAccountCredential = New-ServiceAccountCredential -ClientID $env:ClientID -ClientSecret $env:ClientSecret

# Determining the table name, NOTICE THAT THIS IS PULLING THE NAME OF WHATEVER TABLE THE TIMEDGROUPCHECK FUNCTION USES
$tableName = Get-FunctionTableName -Path '.\TimedGroupCheck\function.json'

# Gathering the context of the cloud table
$cloudTable = Connect-CloudTable -ServiceAccountCredential $ServiceAccountCredential -TenantId $env:TenantID -AzureWebJobsStorage $env:AzureWebJobsStorage -tableName $tableName

# Setting the exemption status
$Result = Set-ExemptStatus -cloudTable $cloudTable -FilterParameter $FilterParameter -FilterValue $FilterValue -Exempt $Exempt

# Associate values to output bindings by calling 'Push-OutputBinding'.
# Checking if any result was found, if not, send a badrequest
if ( [string]::IsNullOrWhiteSpace($Result) ) {
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = $("Could not find an entry with a property of '$FilterParameter' and a value of '$FilterValue'.")
        }
    )
}
else {
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = $Result
        }
    )
}
