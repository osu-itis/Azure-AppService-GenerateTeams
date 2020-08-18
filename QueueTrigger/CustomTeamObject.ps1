Class CustomTeamObject {
    [String]$TeamDescription
    [String]$TeamType
    [String]$TeamName
    [String]$TicketID
    [String]$Requestor
    [String]$CallbackID
    [securestring]$GraphTokenString
    [pscredential]$ServiceAccountCredential
    [PSCustomObject]$GroupResults
    [PSCustomObject]$TeamResults
    [PSCustomObject]$Results
    [void]NewGraphGroupRequest() {


        Function convertformat {
            <#
    .SYNOPSIS
    Converts HTML encoding to standard formatting and removes any leading or trailing whitespace

    .PARAMETER InputText
    The input text to convert

    .EXAMPLE
    PS>$temp = "This+is+a+test%2fexample%0D%0A%0D%0AAnd+it+rocks"
    PS>convertformat -InputText $temp

    This is a test/example

    And it rocks
    "
    #>
            PARAM (
                [parameter(Mandatory = $true)][string]$InputText
            )

            Add-Type -AssemblyName System.Web

            $OutputText = [string]$(
                [System.Web.HttpUtility]::UrlDecode(
                    $InputText
                )
            ).Trim()

            Return $OutputText
        }


        #Get the Owner ID based on the email address that was provided
        #Generating the params that are needed for the query
        $params = @{
            #This formatting is intentional, the $filter needs to be single quoted due to the dollarsign, the single quotes need to be double quoted and the variables should not be single quoted so they are evaluated properly
            #Example of the output: https://graph.microsoft.com/v1.0/users/?$filter=mail eq 'email.address@oregonstate.edu' or userprincipalname eq 'email.address@oregonstate.edu'
            Uri            = "https://graph.microsoft.com/v1.0/users/" + '?$filter=mail eq' + " '" + $($this.Requestor.replace("%40", "@")) + "' " + 'or userprincipalname eq' + " '" + $($this.Requestor.replace("%40", "@")) + "' "
            Authentication = "Bearer"
            Token          = $this.GraphTokenString
            Method         = "Get"
        }
        #Making the graph query and setting a variable with the ID that was returned in the response
        $Owner = (Invoke-RestMethod @params).value.id

        #setting the needed body parameters & converting to JSON format
        $Body = @{
            DisplayName          = $(
                try {
                    #Convert any character encoding to plain text
                    convertformat -InputText $( $this.TeamName )
                }
                catch {
                    Write-Error -Message "Failed to identity the display name" -ErrorAction Stop
                }
            )
            Description          = $(
                try {
                    #Convert any character encoding to plain text
                    convertformat -InputText $( $this.TeamDescription )
                }
                catch {
                    Write-Error -Message "Failed to identity the description" -ErrorAction Stop
                }
            )
            groupTypes           = @([string]"Unified")
            MailEnabled          = [bool]$true
            MailNickname         = $(
                #Remove any spaces, remove slashes, and append a unique string to ensure that the MailNickname is unique and convert any character encoding to plain text
                try {
                    $(
                        convertformat -InputText $( $this.TeamName.tostring().replace(" ", "").replace("/", "").replace("\", "") + [string](Get-Random) )
                    )
                }
                catch {
                    Write-Error -Message "Failed to identity the Mail Nickname" -ErrorAction Stop
                }
            )
            securityEnabled      = [bool]$false
            Visibility           = $(
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
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups" -Authentication Bearer -Token $this.GraphTokenString -Method "Post" -ContentType "application/json" -Body $Body
            )
        }
        catch {
            write-error $(($error[0].ErrorDetails.Message | ConvertFrom-Json).error | Select-Object code, message)
        }

    }
    [void]NewGraphTeamRequest() {
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
            #Make the PUT request to create the team based on the existing group and do not output results to console
            $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString -Method "Put" -ContentType "application/json" -Body $Body

        }
        catch {
            Throw "Failed to generate Team using the existing O365 Unified Group "
        }

        try {
            #Setting the body of the request to show the team in the search or suggestions based off of the team type
            $body = $(
                @{
                    ShowInTeamsSearchAndSuggestions = $(
                        switch ($this.TeamType) {
                            { $_ -eq "Private+Team" } { $false }
                            { $_ -eq "Public+Team" } { $true }
                            Default { $false }
                        }
                    )
                }
            ) | ConvertTo-Json

            #PATCH that to the group team via the Graph API and do not output results to console
            $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString  -Method "Patch" -ContentType "application/json" -Body $Body
        }
        catch {
            throw "Failed to Patch the existing Team"
        }

        try {
            #Gathering the current settings of the Team and setting those to our TeamsResults attribute for later
            $this.TeamResults = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString -Method "Get"
        }
        catch {
            Throw "Failed to gather the team information"
        }

    }
    [void]SetVisibilityInPowershell() {
        #Create the new PSSession and then import it
        $o365session = New-PSSession -configurationname Microsoft.Exchange -connectionuri https://outlook.office365.com/powershell-liveid/ -credential $this.ServiceAccountCredential -authentication basic -allowredirection
        $null = Import-PSSession $o365session -allowclobber -disablenamechecking

        if (! [string]::IsNullOrWhiteSpace($(Get-PSSession | Where-Object { $_.configurationname -eq "Microsoft.Exchange" }))) {
            write-host "`tLoaded Exchange Online"
        }
        else {
            ThrowError "Failed to load Exchange Online"
        }

        #It can take up to 15 min to replicate, loop checks for the existance of the object and wait until its ready before proceeding
        Write-Host "`tWaiting for replication before proceeding"
        $attempt = $null
        $LoopCount = 0
        do {
            try {
                $attempt = Get-UnifiedGroup $this.GroupResults.id -ErrorAction Stop
            }
            catch {
                $LoopCount = $LoopCount + 1
                Start-Sleep -Seconds 60
            }
        } until ($null -ne $attempt)

        Write-Host "`tWaited $LoopCount Minutes for group replication"

        #After replication, we want to set the unified group
        Set-UnifiedGroup $this.GroupResults.id -HiddenFromAddressListsEnabled $true

        #Remove the session now that we no longer need the exchange module and powershell commands
        Remove-PSSession -Session $o365session

    }
    [void]GenerateResults() {
        $this.Results = [hashtable]@{
            ID           = [string]$this.TeamResults.id
            DisplayName  = [string]$this.TeamResults.displayName
            Description  = [string]$this.TeamResults.description
            Mail         = [string]$this.GroupResults.mail
            partitionKey = 'TeamsLog'
            rowKey       = $($this.CallbackID)
            TicketID     = $($this.TicketID)
            Status       = $(
                #Any needed tests to confirm that the team was successfully created
                switch ($this) {
                    { [string]::isnullorempty($_.TeamResults) } { [string]"FAILED" }
                    { -not [string]::isnullorempty($_.TeamResults.ID) } { [string]"SUCCESS" }
                    Default { "UNKNOWN" }
                }
            )
        }

    }
    [void]ExportLastObject() {
        #Can be used for debugging or manually ran instances of this class
        write-host "Exporting a copy of the last run object to an xml"
        Export-Clixml -InputObject $this .\TempObject.cli.xml
    }
    [void]AutoCreateTeam() {
        #Generate the new group
        write-host "Generating a new group request via graph api"
        $this.NewGraphGroupRequest()

        #Wait for a few moments
        Start-Sleep -Seconds 5

        #Generate the new team (from the existing group)
        write-host "Generating a new teams request via graph api"
        $this.NewGraphTeamRequest()

        #Set the visibility in powershell (if needed) This can take a very long time as it requires that Exchange has replicated
        switch ($this.TeamType) {
            { $_ -eq "Public+Team" } {
                #Do not make any changes to the visibility in the GAL
                write-host "No changes made to the visibility in the GAL"
            }
            { $_ -eq "Private+Team" } {
                write-host "Using Powershell to hide visibility in the GAL"
                $this.SetVisibilityInPowershell()
            }
            Default {
                write-host "Unable to determine team type, attempting to hide visibility in the GAL"
                $this.SetVisibilityInPowershell()
            }
        }

        #Generate the results
        write-host "Gathering a report of the results"
        $this.GenerateResults()

    }
}
