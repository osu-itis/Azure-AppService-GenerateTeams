using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Grabbing the body of the request and setting it to a new hashtable
[hashtable]$Temp = $Request.body

#Adding the callback id (which will be needed later)
$Temp = $Temp + @{CallbackID = $(new-guid | Select-Object -ExpandProperty GUID) }

# Associate values to output bindings by calling 'Push-OutputBinding':

## Exporting the information to the storage queue
Push-OutputBinding -Name "outputQueueItem" -Value $($temp | convertto-Json)

## Sending the information back in the response 
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $(@{CallbackID = $Temp.CallbackID } | ConvertTo-Json)
    })
