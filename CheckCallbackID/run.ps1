using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) { Throw 'Could not find $env:AzureWebJobsStorage' }
if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }

$Settings = [pscustomobject]@{

    #Getting the Tenant ID
    TenantID              = $env:TenantId

    #The partition key of the azure storage table thats in use
    PartitionKey          = "TeamsLog"

    #Azure Storage Table
    AzureStorageTableName = "LoggedTeamsRequests"

    #StorageAccount Credentials
    StorageAccount        = $(
        #Building an empty hash
        $StorageAccount = @{}
        #Using a switch to pull the name and key out of the azurewebjobsstorage env
        switch (
            $($env:AzureWebJobsStorage.split(";"))
        ) {
            { $_ -like "AccountName=*" } {
                $StorageAccount += @{
                    Name = $($_.replace("AccountName=", ""))

                }
            }
            { $_ -like "AccountKey=*" } {
                $StorageAccount += @{
                    Key = $($_.replace("AccountKey=", ""))

                }
            }
        }
        #Returning the hash as an object
        [pscustomobject]$StorageAccount
    )

    #Service Principal
    ServicePrincipalID    = $env:ClientID
    ServicePrincipalKey   = $env:ClientSecret
}

#Gathering the Client ID from the query of the request, converted from Json
$CallbackID = $Request.Query.CallbackID | ConvertFrom-Json | Select-Object -ExpandProperty CallbackID

#Gather the AZ Storage Context which provides information about the account to be used
$StorageAccount = [pscustomobject]@{
    CTX = New-AzStorageContext -StorageAccountName $Settings.StorageAccount.Name -StorageAccountKey $Settings.StorageAccount.Key
}

#Using the context, get the storage table and gather the "CloudTable" properties
$AzureStorageTable = [PSCustomObject]@{
    CloudTable = (Get-AzStorageTable -Name $Settings.AzureStorageTableName -Context $StorageAccount.CTX.Context).CloudTable
}
$ServicePrincipalAccount = New-Object System.Management.Automation.PSCredential ($Settings.ServicePrincipalID, $(ConvertTo-SecureString $Settings.ServicePrincipalKey -AsPlainText -Force))

#Connecting to an AD service account (which auto-loads the "AzStorageTable" cmdlets, this is required to use this commands)
Connect-AzAccount -Tenant $Settings.TenantID -Credential $ServicePrincipalAccount -ServicePrincipal

#Generating the table entry to query
$TableEntryQuery = @{
    RowKey       = $($CallbackID.trim())
    PartitionKey = $Settings.PartitionKey
    Table        = $AzureStorageTable.CloudTable
}

#Gathering the AZ table row info (using information about the table to gather the correct entry)
$Results = Get-AzTableRow @TableEntryQuery

#Testing for the client id
switch ($Results) {
    #If the result status is success:
    {$_.status -eq "SUCCESS"} {
        $status = [HttpStatusCode]::OK
        $body = $Results
    }
    #If the result status is failed:
    {$_.status -eq "FAILED"} {
        $status = [HttpStatusCode]::BadRequest
        #Note that we are still returning the results, but it also returns a bad request status (which can be used to automate workflows based on the response)
        $body = $Results
    }
    #If there is no status or the $Results attribute does not even exist:
    Default {
        $status = [HttpStatusCode]::BadRequest
        $body = "$CallbackID was not found"
    }
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value (
    [HttpResponseContext]@{
        StatusCode = $status
        Body       = $body
    }
)
