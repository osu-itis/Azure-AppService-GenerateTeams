Class CustomTeamObject {
    #Basic attributes
    [String]$TeamDescription
    [String]$TeamType
    [String]$TeamName
    [String]$TicketID
    [String]$Requestor
    [String]$CallbackID

    #Credentials that needed to run methods
    [securestring]$GraphTokenString
    [pscredential]$ServiceAccountCredential

    #Attribues that will store calculated values
    [string]$TeamOwner
    [string]$MailNickname

    #Attributes that contain the results of GRAPH API calls
    [PSCustomObject]$GroupResults
    [PSCustomObject]$TeamResults
    [PSCustomObject]$Results

    #Custom Methods
    [void]CleanAttributes() {
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

        #Processing the attributes
        $this.TeamDescription = convertformat -InputText $( $this.TeamDescription )
        $this.TeamType = $(
            switch ($this.TeamType) {
                { $_ -eq "Private+Team" } { "Private" }
                { $_ -eq "Public+Team" } { "Public" }
            }
        )
        $this.TeamName = convertformat -InputText $( $this.TeamName )
        $this.Requestor = convertformat -InputText $( $this.Requestor )

    }
    [void]ResolveTeamOwner() {
        #Get the Owner ID based on the email address that was provided
        #Generating the params that are needed for the query
        $params = @{
            #This formatting is intentional, the $filter needs to be single quoted due to the dollarsign, the single quotes need to be double quoted and the variables should not be single quoted so they are evaluated properly
            #Example of the output: https://graph.microsoft.com/v1.0/users/?$filter=mail eq 'email.address@oregonstate.edu' or userprincipalname eq 'email.address@oregonstate.edu'
            Uri            = "https://graph.microsoft.com/v1.0/users/" + '?$filter=mail eq' + " '" + $($this.Requestor) + "' " + 'or userprincipalname eq' + " '" + $($this.Requestor) + "' "
            Authentication = "Bearer"
            Token          = $this.GraphTokenString
            Method         = "Get"
        }
        #Making the graph query and setting a variable with the ID that was returned in the response
        try {
            $this.TeamOwner = (Invoke-RestMethod @params).value.id
        }
        catch {
            Write-Error -Message "Failed to identity the Owner" -ErrorAction Stop
        }
    }
    [void]GenerateMailNickname() {
        $this.MailNickname = $(
            #Remove any spaces, remove slashes, and append a unique string to ensure that the MailNickname is unique and convert any character encoding to plain text
            #Add a randomized string to make it unique, convert any character encoding to plain text and remove any slashes or spaces
            try {
                Add-Type -AssemblyName System.Web
                $( [System.Web.HttpUtility]::UrlDecode( $this.TeamName.tostring() + [string](Get-Random) ) ).replace(" ", "").replace("/", "").replace("\", "")
            }
            catch {
                Write-Error -Message "Failed to generate the Mail Nickname" -ErrorAction Stop
            }
        )
    }
    [void]NewGraphGroupRequest() {
        #Setting the needed body parameters & converting to JSON format
        $Body = @{
            DisplayName          = $this.TeamName
            Description          = $this.TeamDescription
            groupTypes           = @([string]"Unified")
            MailEnabled          = [bool]$true
            MailNickname         = $this.MailNickname
            securityEnabled      = [bool]$false
            Visibility           = $this.TeamType
            "owners@odata.bind"  = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($this.teamOwner)" ) )
            "Members@odata.bind" = [array]@( $( [string]"https://graph.microsoft.com/v1.0/users/$($this.TeamOwner)" ) )
        } | ConvertTo-Json

        #Making a graph request for a new group
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
            Visibility        = $this.TeamType
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
                            { $_ -eq "Private" } { $false }
                            { $_ -eq "Public" } { $true }
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

    #This particular method uses all of the other methods in order to generate a new Microsoft Team automatically
    [void]AutoCreateTeam() {
        #Start by cleaning attributes as needed
        Write-Host "Cleaning attributes"
        $this.CleanAttributes()

        #Resolve the Team Owner (based on the ID that will be used in graph requests)
        Write-Host "Resolving Team Owner"
        $this.ResolveTeamOwner()

        #Generate the new mail nickname
        Write-Host "Generating a unique mail nick name"
        $this.GenerateMailNickname()

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
            { $_ -eq "Public" } {
                #Do not make any changes to the visibility in the GAL
                write-host "No changes made to the visibility in the GAL"
            }
            { $_ -eq "Private" } {
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
