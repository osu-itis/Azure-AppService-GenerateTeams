using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Attempting to import the needed Modules
try
{
    Write-Output "Importing the MSAL.PS module"
    Import-Module .\Modules\MSAL.PS\4.21.0.1\MSAL.PS.psd1 -Force -ErrorAction stop
} catch
{
    Throw 'Failed to import the MSAL.PS Module'
}

# Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage))
{ Throw 'Could not find $env:AzureWebJobsStorage' 
}
if ([string]::IsNullOrEmpty($env:ClientID))
{ Throw 'Could not find $env:ClientID' 
}
if ([string]::IsNullOrEmpty($env:ClientSecret))
{ Throw 'Could not find $env:ClientSecret' 
}
if ([string]::IsNullOrEmpty($env:TenantId))
{ Throw 'Could not find $env:TenantId' 
}

# Using MSAL to generate and manage a (JWT) token
Write-Output "Generating token"
$token = Get-MsalToken -ClientId $env:ClientID -ClientSecret $(ConvertTo-SecureString $env:ClientSecret -AsPlainText -Force) -TenantId $env:TenantID

# Setting the headers as a variable for convienience
$Headers = @{Authorization = "Bearer $($token.AccessToken)" }

if ($Request.Body.user)
{
    
    # Setting the graph URI with the user to look up. 
    $URI = "https://graph.microsoft.com/v1.0/users/$($request.Body.user)/licenseDetails"
    
    
    try {
        $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI
    }
    catch {
        $Results = $null
        
        write-warning "Could not find a user with the value $($request.body.user)"
    }
    
    if ($results.value.serviceplans -eq $null) {
        Push-OutputBinding -Name Response -Value (
            [HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body = "Could not find a user with the value $($request.body.user)."
            })
    }
    else {
        $tempBody = [PSCustomObject]@{
            User = $Request.Body.user
            TeamsEnabled = $($results.value.servicePlans|Where-Object {$_.serviceplanname -like "*TEAMS*"}[0]|Select-Object -ExpandProperty provisioningStatus)
        }
        
        $body = $tempBody|ConvertTo-Json
        
        if ($tempBody.TeamsEnabled -eq "Success") {
            # Associate values to output bindings by calling 'Push-OutputBinding'.
            Push-OutputBinding -Name Response -Value (
                [HttpResponseContext]@{
                    StatusCode = [HttpStatusCode]::OK
                    Body = $body
                })
        }
        else {
            # Associate values to output bindings by calling 'Push-OutputBinding'.
            Push-OutputBinding -Name Response -Value (
                [HttpResponseContext]@{
                    StatusCode = [HttpStatusCode]::BadRequest
                    Body = "User with the value $($request.body.user) is not licenced for teams."
                })
        }
    }
    
} else
{
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = "Request was not correctly formatted."
        })
}






