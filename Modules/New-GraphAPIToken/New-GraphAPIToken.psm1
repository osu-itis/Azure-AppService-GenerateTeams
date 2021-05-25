<# The entire custom GraphAPIToken class is wrapped within a function, this is for two reasons:
        1. The function both defines and instanciates a new instance of the class
        2. Classes are not imported from modules, but functions are. This benefits us by allowing us to import this module in other scripts
#>

function New-GraphAPIToken {
    <#
    .SYNOPSIS
    Generate a Powershell object that contains an API token that can be used for graph calls
    
    .DESCRIPTION
    Custom Class that contains both headers and a token string which can be used with Invoke-RestMethod
    
    .PARAMETER ClientID
    The ID of the service (application/app registration)
    
    .PARAMETER ClientSecret
    The password or secret
    
    .PARAMETER TenantID
    The Tenant ID (ID number only, dont use fqdn)
    
    .EXAMPLE
    $temp = New-GraphAPIToken -ClientID $env:ClientID -ClientSecret $env:ClientSecret -TenantID $env:TenantID
    
    Invoke-RestMethod -Method get -Uri https://graph.microsoft.com/v1.0/ -Authentication Bearer -Token $temp.TokenString
    
    .NOTES
    The output object contains a method (NewOAuthRequest) to generate a new token once the initial token expires
    #>
    param (
        [parameter(Mandatory = $true)]$ClientID,
        [parameter(Mandatory = $true)]$ClientSecret,
        [parameter(Mandatory = $true)]$TenantID
    )

    # Using a class for its strong typing and methods
    Class GraphAPIToken {
        [string]$ClientID
        [string]$ClientSecret
        [string]$TenantID

        [PSCustomObject]$OAuthReq
        [PSCustomObject]$Headers
        [SecureString]$TokenString

        [void]NewOAuthRequest() {
            try {
                $body = [hashtable]@{
                    client_id     = [string]$this.ClientID
                    client_secret = [string]$this.ClientSecret
                    grant_type    = [string]"client_credentials"
                    scope         = [uri]"https://graph.microsoft.com/.default"
                }
                $this.OAuthReq = $(Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($this.TenantId)/oauth2/v2.0/token" -Body $Body)
            }
            # If this fails out, stop everything, nothing will work without that token
            catch {
                write-error -message $Error[0].Exception -ErrorAction Stop
            }
        }
        [void]GenerateHeaders() {
            # Creating the headers in for format of 'Bearer <TOKEN>', this will be needed for all future requests
            $this.Headers = @{
                Authorization = "$($this.OAuthReq.token_type) $($this.OAuthReq.access_token)"
            }
        }
        [void]GenerateTokenString() {
            $this.TokenString = $(ConvertTo-SecureString -String ($this.OAuthReq.access_token) -AsPlainText -Force)
        }


        # Custom Constructor for generating the class:
        GraphAPIToken (
            [string]$ClientID,
            [string]$ClientSecret,
            [string]$TenantID
        ) {
            # Setting variables
            $this.ClientID = $ClientID
            $this.ClientSecret = $ClientSecret
            $this.TenantID = $TenantID

            # Running needed methods
            $this.NewOAuthRequest()
            $this.GenerateHeaders()
            $this.GenerateTokenString()
        }
    }

    # Outputting the new graph token object
    [GraphAPIToken]::new($ClientID, $ClientSecret, $TenantID)
}
