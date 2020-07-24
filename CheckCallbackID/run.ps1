using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) { Throw 'Could not find $env:AzureWebJobsStorage' }
if ([string]::IsNullOrEmpty($env:AppRegistrationID)) { Throw 'Could not find $env:AppRegistrationID' }
if ([string]::IsNullOrEmpty($env:AppRegistrationSecret)) { Throw 'Could not find $env:AppRegistrationSecret' }

$Settings = [pscustomobject]@{

    #Getting the Tenant ID
    TenantID =  $env:TenantId

    #The partition key of the azure storage table thats in use
    PartitionKey = "TeamsLog"

    #Azure Storage Table
    AzureStorageTableName = "LoggedTeamInstalls"

    #StorageAccount Credentials
    StorageAccount = $(
        #Building an empty hash
        $StorageAccount = @{}
        #Using a switch to pull the name and key out of the azurewebjobsstorage env
        switch (
            $($env:AzureWebJobsStorage.split(";"))
        ) {
            {$_ -like "AccountName=*"}{
                $StorageAccount += @{
                    Name = $($_.replace("AccountName=",""))
        
                }
            }
            {$_ -like "AccountKey=*"}{
                $StorageAccount += @{
                    Key = $($_.replace("AccountKey=",""))
        
                }
            }
        }
        #Returning the hash as an object
        [pscustomobject]$StorageAccount
    )

    #Service Principal
    ServicePrincipalID = $env:AppRegistrationID
    ServicePrincipalKey = $env:AppRegistrationSecret
}

#Gathering the Client ID from the query of the request, converted from Json
$CallbackID = $Request.Query.CallbackID|ConvertFrom-Json|Select-Object -ExpandProperty CallbackID

#Generating the table entry to query
$TableEntry = @{
    RowKey = $($CallbackID)
    PartitionKey = $Settings.PartitionKey
}

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

#Gathering the AZ table row info (using information about the table to gather the correct entry)
$Results = Get-AzTableRow @TableEntry -Table $AzureStorageTable.CloudTable

#Testing for the client id
if ($Results) {
    $status = [HttpStatusCode]::OK
    $body = $Results
}
else {
    $status = [HttpStatusCode]::BadRequest
    $body = "$CallbackID was not found"
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $status
    Body = $body
})
