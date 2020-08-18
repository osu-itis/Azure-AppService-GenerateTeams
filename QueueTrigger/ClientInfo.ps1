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
        #If this fails out, stop everything, nothing will work without that token
        catch {
            write-error -message $Error[0].Exception -ErrorAction Stop
        }
    }
    [void]GenerateHeaders() {
        #Creating the headers in for format of 'Bearer <TOKEN>', this will be needed for all future requests
        $this.Headers = @{
            Authorization = "$($this.OAuthReq.token_type) $($this.OAuthReq.access_token))"
        }
    }
    [void]GenerateTokenString() {
        $this.TokenString = $(ConvertTo-SecureString -String ($this.OAuthReq.access_token) -AsPlainText -Force)
    }


    # Custom Constructor:
    GraphAPIToken (
        [string]$ClientID,
        [string]$ClientSecret,
        [string]$TenantID
    ) {
        #Setting variables
        $this.ClientID = $ClientID
        $this.ClientSecret = $ClientSecret
        $this.TenantID = $TenantID

        # #Running needed methods
        $this.NewOAuthRequest()
        $this.GenerateHeaders()
        $this.GenerateTokenString()
    }
}
