using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

Write-Output "Ticket: $($Request.query.ticket)"
Write-Output "User: $($Request.query.user)"

# Attempting to import the needed Modules
. .\Shared\Import-MSAL.ps1

# Checking if the needed ENVs exist
. .\Shared\Check-ENVs.ps1

# Using MSAL to generate and manage a (JWT) token
Write-Output "Generating token"
$token = Get-MsalToken -ClientId $env:ClientID -ClientSecret $(ConvertTo-SecureString $env:ClientSecret -AsPlainText -Force) -TenantId $env:TenantID

# Setting the headers as a variable for convienience
$Headers = @{Authorization = "Bearer $($token.AccessToken)" }

# make a graph api call to find the UserPrincipalName from the email address.
$URI = $('https://graph.microsoft.com/v1.0/users?$filter=mail eq '+"'"+$Request.Query.user+"'")

$Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI

if ($Results.value.userprincipalname) {
    # If a upn was found, use that instead
    Write-Output "$($Request.query.user) resolved to $($Results.value.userprincipalname)"
    $Request.query.user = $Results.value.userprincipalname
}

if ($Request.query.user)
{
    
    # Setting the graph URI with the user to look up. 
    $URI = "https://graph.microsoft.com/v1.0/users/$($Request.query.user)/licenseDetails"
    
    
    try {
        $Results = Invoke-RestMethod -Headers $Headers -Method Get -Uri $URI
    }
    catch {
        $Results = $null
        
        write-warning "Could not find a user with the value $($Request.query.user)"
    }
    
    if ($results.value.serviceplans -eq $null) {
        Write-Output "Responding with bad request, user could not be found"
        Push-OutputBinding -Name Response -Value (
            [HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body = "Could not find a user with the value $($Request.query.user)."
            })
    }
    else {
        $tempBody = [PSCustomObject]@{
            User = $Request.query.user
            TeamsEnabled = $($results.value.servicePlans|Where-Object {$_.serviceplanname -like "*TEAMS*"}[0]|Select-Object -ExpandProperty provisioningStatus)
        }
        
        $body = $tempBody|ConvertTo-Json
        
        if ($tempBody.TeamsEnabled -eq "Success") {
            Write-Output "Responding with good request, user is licenced for teams"
            Push-OutputBinding -Name Response -Value (
                [HttpResponseContext]@{
                    StatusCode = [HttpStatusCode]::OK
                    Body = $body
                })
        }
        else {
            Write-Output "Responding with bad request, user does not have a licence for teams"
            Push-OutputBinding -Name Response -Value (
                [HttpResponseContext]@{
                    StatusCode = [HttpStatusCode]::BadRequest
                    Body = "User with the value $($Request.query.user) is not licenced for teams."
                })
        }
    }
    
} else
{
    Write-Output "Responding with bad request, request was not correctly formatted"
    Push-OutputBinding -Name Response -Value (
        [HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = "Request was not correctly formatted."
        })
}
