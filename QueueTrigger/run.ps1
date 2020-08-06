# Input bindings are passed in via param block.
param($Queue, $TriggerMetadata)

$StartErrorCount = [int]$(
    $error.Count
)

# Write out the queue message and insertion time to the information log.
Write-Host "PowerShell queue trigger function processed work item: $($Queue |convertto-json )"
Write-Host "Queue item insertion time: $($TriggerMetadata.InsertionTime)"

#Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }

$ClientInfo = [PSCustomObject]@{
    #This is the ClientID (Application ID) of registered AzureAD App
    ClientID     = $env:ClientID
    #This is the key of the registered AzureAD app
    ClientSecret = $env:ClientSecret
    #This is your our Tenant ID
    TenantId     = $env:TenantId
    #Leaving the headers blank for now, they'll be generated via a scriptmethod below
    Headers      = $null
    #Adding the token as a standalone variable (This can be used with invoke-restmethod after PS version 6.*)
    TokenString  = $null
}

#Adding the "NewOAuthRequest" script method to generate a token and the needed Headers (added to the ClientInfo object)
$ClientInfo | Add-Member -MemberType ScriptMethod -Name NewOAuthRequest -Value {
    $body = [hashtable]@{
        client_id     = [string]$this.ClientID
        client_secret = [string]$this.ClientSecret
        grant_type    = [string]"client_credentials"
        scope         = [uri]"https://graph.microsoft.com/.default"
    }

    try {
        $OAuthReq = $(Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($this.TenantId)/oauth2/v2.0/token" -Body $Body)
    }
    #If this fails out, stop everything, nothing will work without that token
    catch {
        write-error -message $Error[0].Exception -ErrorAction Stop
    }
    #Creating the headers in for format of 'Bearer <TOKEN>', this will be needed for all future requests
    $this.Headers = @{
        Authorization = "$($OAuthReq.token_type) $($OAuthReq.access_token))"
    }
    $this.TokenString = $(ConvertTo-SecureString -String ($OAuthReq.access_token) -AsPlainText -Force)
}

#Generate the needed auth request and token
$ClientInfo.NewOAuthRequest()

#Creating a new variable as our temp object (as a hash table)
[pscustomobject]$TempObject = [pscustomobject]$Queue

$TempObject | Add-Member -NotePropertyMembers @{
    GroupResults = $null
    TeamResults  = $null
    Results      = $null
}

#Adding the script method for the a group request
$TempObject | Add-Member -Force -MemberType ScriptMethod -Name NewGraphGroupRequest -Value {
        
    #Get the Owner ID based on the email address that was provided
    #Generating the params that are needed for the query
    $params = @{
        #This formatting is intentional, the $filter needs to be single quoted due to the dollarsign, the single quotes need to be double quoted and the variables should not be single quoted so they are evaluated properly
        #Example of the output: https://graph.microsoft.com/v1.0/users/?$filter=mail eq 'email.address@oregonstate.edu' or userprincipalname eq 'email.address@oregonstate.edu'
        Uri            = "https://graph.microsoft.com/v1.0/users/" + '?$filter=mail eq' + " '" + $($this.Requestor.replace("%40", "@")) + "' " + 'or userprincipalname eq' + " '" + $($this.Requestor.replace("%40", "@")) + "' "
        Authentication = "Bearer"
        Token          = $ClientInfo.TokenString
        Method         = "Get"
    }
    #Making the graph query and setting a variable with the ID that was returned in the response
    $Owner = (Invoke-RestMethod @params).value.id
    
    #setting the needed body parameters & converting to JSON format
    $Body = @{
        DisplayName          = $(
            try {
                #If the HTML URL encoding is using '+' replace with the space character, if the character is '%2B' Replace with '+' (so the displayname is properly set with the correct characters)
                [string]$this.TeamName.replace("+", " ").replace("%2B", "+").trim()
            }
            catch {
                Write-Error -Message "Failed to identity the display name" -ErrorAction Stop
            }
        )
        Description          = $(
            try {
                #Replace all + signs with spaces, fix slashes, and fix carage returns/new lines
                [string]$this.TeamDescription.replace("+", " ").replace("%2", "/").replace("%0D%0A", "`r`n").trim()
            }
            catch {
                Write-Error -Message "Failed to identity the description" -ErrorAction Stop
            }
        )
        groupTypes           = @([string]"Unified")
        MailEnabled          = [bool]$true
        MailNickname         = $(
            #Randomly generate a unique MailNickname based off of the TeamName and a randomized string
            try {
                [string]$this.TeamName.Replace(" ", "") + [string](Get-Random)
            }
            catch {
                Write-Error -Message "Failed to identity the Mail Nickname" -ErrorAction Stop
            }
        )
        securityEnabled      = [bool]$false
        Visibility           = $(
            try {
                #$this.TeamType
                switch ($this.TeamType) {
                    { $_ -like "Private+Team" } { "Private" }
                    { $_ -like "Public+Team" } { "Public" }
                    Default { "Private" }
                }
            }
            catch {
                Write-Error -Message "Failed to identity the Team type" -ErrorAction Stop
            }
        )
        "owners@odata.bind"  = [array]@(
            $(
                try {
                    [string]"https://graph.microsoft.com/v1.0/users/$Owner"
                }
                catch {
                    Write-Error -Message "Failed to identity the Owner" -ErrorAction Stop
                }
            )
        )
        "Members@odata.bind" = [array]@(
            $(
                try {
                    [string]"https://graph.microsoft.com/v1.0/users/$Owner"
                }
                catch {
                    Write-Error -Message "Failed to identity the Owner" -ErrorAction Stop
                }
            )
        )
    } | ConvertTo-Json

    try {
        $this.GroupResults = $(
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Authentication Bearer -Token $ClientInfo.TokenString -Method "Post" -ContentType "application/json" -Body $Body
        )
    }
    catch {
        write-error $(($error[0].ErrorDetails.Message | ConvertFrom-Json).error | Select-Object code, message)
    }
}

#Adding the script method for the teams request
$TempObject | Add-Member -Force -MemberType ScriptMethod -Name NewGraphTeamRequest -Value {
    #Setting the needed settings for the team and converting the data to Json for the API call
    $Body = @{
        MemberSettings    = @{
            allowCreatePrivateChannels = $true
            allowCreateUpdateChannels  = $true
        }
        MessagingSettings = @{
            allowUserEditMessages   = $true
            allowUserDeleteMessages = $true
        }
        FunSettings       = @{
            allowGiphy         = $true
            giphyContentRating = [string]"Moderate"
        }
        Visibility        = $(
            try {
                switch ($this.TeamType) {
                    { $_ -like "Private+Team" } { "Private" }
                    { $_ -like "Public+Team" } { "Public" }
                    Default { "Private" }
                }
            }
            catch {
                Write-Error -Message "Failed to identity the Team type" -ErrorAction Stop
            }
        )
    } | ConvertTo-Json

    try {
        $this.TeamResults = $(
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($this.GroupResults.ID)/team"  -Authentication Bearer -Token $ClientInfo.TokenString  -Method "Put" -ContentType "application/json" -Body $Body
        )
    }
    catch {
        write-error $(($error[0].ErrorDetails.Message | ConvertFrom-Json).error | Select-Object code, message)
    }
}

#Adding the script method to gather the results
$TempObject | Add-Member -Force -MemberType ScriptMethod -Name GenerateResults -Value {
    $this.Results = [hashtable]@{
        ID          = [string]$this.TeamResults.id
        DisplayName = [string]$this.TeamResults.displayName
        Description = [string]$this.TeamResults.description
        Mail        = [string]$this.GroupResults.mail
        Visibility  = [string]$this.GroupResults.visibility
    }
}

#Generate the new group
$TempObject.NewGraphGroupRequest()

#wait for a few moments
Start-Sleep -Seconds 15

#generate the new team (from the existing group)
$TempObject.NewGraphTeamRequest()

#wait for a few moments
Start-Sleep -Seconds 15

#Generate the results
$TempObject.GenerateResults()

#Creating the table logging (needed table attributes)
$TabbleLogging = [hashtable]@{
    partitionKey = 'TeamsLog'
    rowKey       = $($TempObject.CallbackID)
    TicketID     = $($TempObject.TicketID)
    Status       = $(
        #Any needed tests to confirm that the team was successfully created
        switch ($TempObject) {
            { [string]::isnullorempty($_.TeamResults) } { [string]"FAILED" }
            { -not [string]::isnullorempty($_.TeamResults.ID) } { [string]"SUCCESS" }
            Default { "UNKNOWN" }
        }
    )
}

#Adding our temp attributes, the table attributes & the status of the group
$Output = $TabbleLogging + $TempObject.Results

#Adding any logging or error information needed
$Output | Add-Member -NotePropertyMembers @{
    ErrorCount = $error.Count
    Errors     = $(
        if ($StartErrorCount -ne $error.Count) {
            $Error.Exception.Message
        }
    )
}

#Converting to Json and pushing out to the host for humans to read
Write-Host -message "Final Output:"
Write-Host -message $($Output | convertto-json)

#Writing to the table (for logging purposes)
Push-OutputBinding -Name LoggedTeamsRequests -Value $Output

#If the script failed to create a new team, we want this to throw an error
if ($Output.status -eq "FAILED") {
    Throw "Script failed to generate a new Microsoft Teams team."
}
