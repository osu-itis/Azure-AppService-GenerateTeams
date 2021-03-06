Class CustomTeamObject {
    # Basic attributes
    [String]$TeamDescription
    [String]$TeamType
    [String]$TeamName
    [String]$TicketID
    [String]$Requestor
    [String]$CallbackID

    # Credentials that are needed to run methods
    [securestring]$GraphTokenString
    [pscredential]$ServiceAccountCredential

    # Attribues that will store calculated values
    [string]$TeamOwner
    [string]$MailNickname

    # Attributes that contain the results of GRAPH API calls
    [PSCustomObject]$GroupResults
    [PSCustomObject]$TeamResults
    [PSCustomObject]$Results

    # Custom Methods
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

        # Processing the attributes
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
        # Get the Owner ID based on the email address that was provided
        # Generating the params that are needed for the query
        $params = @{
            # This formatting is intentional, the $filter needs to be single quoted due to the dollarsign, the single quotes need to be double quoted and the variables should not be single quoted so they are evaluated properly
            # Example of the output: https://graph.microsoft.com/v1.0/users/?$filter=mail eq 'email.address@oregonstate.edu' or userprincipalname eq 'email.address@oregonstate.edu'
            Uri            = "https://graph.microsoft.com/v1.0/users/" + '?$filter=mail eq' + " '" + $($this.Requestor) + "' " + 'or userprincipalname eq' + " '" + $($this.Requestor) + "' "
            Authentication = "Bearer"
            Token          = $this.GraphTokenString
            Method         = "Get"
        }
        # Making the graph query and setting a variable with the ID that was returned in the response
        try {
            $this.TeamOwner = (Invoke-RestMethod @params).value.id
        }
        catch {
            Write-Error -Message "Failed to identity the Owner" -ErrorAction Stop
        }
    }
    [void]GenerateMailNickname() {
        $this.MailNickname = $(
            # Convert any character encoding to plain text and remove any special characters, truncate and then append a randomized string to the end to ensure that the value is unique
            try {
                Add-Type -AssemblyName System.Web
                $TruncateTest = $( [System.Web.HttpUtility]::UrlDecode( $this.TeamName.tostring() ) ).replace(" ", "").replace("/", "").replace("\", "") -replace '[^\p{L}\p{Nd}]',''
                $RandString = [string](Get-Random)

                # Mail Nickname has a maximum length of 64 characters, we'll start shortening the nickname until its within the limit (without modifying our random string, to ensure that it's still a unique value)
                if (($TruncateTest+$RandString|Measure-Object -Character).characters -ge 64){
                    do {
                        $TruncateTest = $TruncateTest.Substring(0,$TruncateTest.Length-1)
                    } until (($TruncateTest+$RandString|Measure-Object -Character).characters -lt 64)
                }

                # Output the result
                $TruncateTest+$RandString
            }
            catch {
                Write-Error -Message "Failed to generate the Mail Nickname" -ErrorAction Stop
            }
        )
    }
    [void]NewGraphGroupRequest() {
        # Setting the needed body parameters & converting to JSON format
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

        # Making a graph request for a new group
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
        $Increment = 0
        $FailTest = $false
        do {
            try {
                # Setting the needed settings for the team and converting the data to Json for the API call
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

                # Waiting the recommended amount of time before attempting to use a unified group to create a new team
                Start-Sleep -Seconds 10
                $FailTest = $false
                # Make the PUT request to create the team based on the existing group and do not output results to console
                $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString -Method "Put" -ContentType "application/json" -Body $Body
    
            }
            catch {
                write-host "Failed to create a new team with the existing group, will try again in a moment..."
                $FailTest = $true
                $Increment ++
            }
        } until (($FailTest -eq $false) -or ($Increment -eq 6))

        if (($FailTest -eq $true) -and ($Increment -eq 6)) {
            throw "Failed to generate the team off of the existing unified group"
        }

        do {
            $Increment = 0
            $FailTest = $false
            try {
                $FailTest = $false
    
                # Setting the body of the request to show the team in the search or suggestions based off of the team type
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
    
                # PATCH that to the group team via the Graph API and do not output results to console
                $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString  -Method "Patch" -ContentType "application/json" -Body $Body
            }
            catch {
                Write-Host "Failed to Patch the existing Team, will try again in a moment..."
                $FailTest = $true
                $Increment ++
            }
        } until (($FailTest -eq $false) -or ($Increment -eq 6))

        if (($FailTest -eq $true) -and ($Increment -eq 6)) {
            throw "Failed to Patch the existing Team"
        }

        $Increment = 0
        $FailTest = $false
        do {
            $FailTest = $false
            try {
                # Gathering the current settings of the Team and setting those to our TeamsResults attribute for later
                $this.TeamResults = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($this.GroupResults.ID)/team" -Authentication Bearer -Token $this.GraphTokenString -Method "Get"
            }
            catch {
                Write-Host "Failed to gather the team information, will try again in a moment..."
                $FailTest = $true
                $Increment ++
            }
        } until (($FailTest -eq $false) -or ($Increment -eq 6))

        if (($FailTest -eq $true) -and ($Increment -eq 6)) {
            throw "Failed to gather the team information"
        }
    }
    [void]DisableAddingGuests() {
        # Setting a fail test status to true, If the code executes properly this will be toggled to false
        $FailTest = $true

        # Setting the starting increment for our loop
        $Increment = 0

        # The Group ID to be tested
        $GroupID = $this.TeamResults.id

        # Setting the body of the request to be made, notice that the template ID is a static value that should not change
        $body = "{
        `n    `"displayName`": `"GroupSettings`",
        `n    `"templateId`": `"08d542b9-071f-4e16-94b0-74abb372e3d9`",
        `n    `"values`": [
        `n        {
        `n            `"name`": `"AllowToAddGuests`",
        `n            `"value`": `"false`"
        `n        }
        `n    ]
        `n}"

        do {
            $Increment ++

            # Waiting for 1 second
            Start-Sleep -Seconds 1

            # Check and see if the Object returns (If not, we're probably querying the team too soon after it's creation)
            $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)" -Authentication Bearer -Token $this.GraphTokenString -Method "get"

            # If the Group info was returned...
            if ( ! [string]::IsNullOrEmpty($GroupInfo) ) {
                # Gather if it already has a "Group.Unified.Guest", if so, we need to patch. Otherwise we need to post.
                $PreChange = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Authentication Bearer -Token $this.GraphTokenString -Method "get"
                if ( $($PreChange.value | Where-Object { $_.displayname -Like "Group.Unified.Guest" }).count -eq 1 ) {
                    # Patching the setting on an object that already had the setting configured
                    $null = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings/$($PreChange.value.id)" -Method 'patch' -Body $body -ContentType 'application/json' -Authentication Bearer -Token $this.GraphTokenString
                }
                else {
                    # Posting the setting for the first time on the object
                    $null = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Method 'POST' -Body $body -ContentType 'application/json' -Authentication Bearer -Token $this.GraphTokenString
                }
                # Gathering the results of the change to be sure that the proper settings have been applied
                $PostChange = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Authentication Bearer -Token $this.GraphTokenString -Method "get"
    
                if (
                    # Allow to Add Guests is set to false
                    ($PostChange.value.values | Where-Object { $_.name -eq "AllowToAddGuests" }).value -eq $false
                ) {
                    $FailTest = $false
                }
            }
            # If the group was not returned...
            else {
                $FailTest = $true
            }
        } until (($FailTest -eq $false) -or ($Increment -eq 10))

        # Switch to output the final outcome
        switch ($FailTest) {
            { $_ -eq $true } { throw "Failed to set the 'AllowToAddGuests' attribute to False." }
            { $_ -eq $false } { Write-Host "Set the 'AllowToAddGuests' attribute to False." }
        }

    }
    [void]GenerateResults() {
        $this.Results = [hashtable]@{
            ID           = [string]$this.TeamResults.id
            DisplayName  = [string]$this.TeamResults.displayName
            Description  = [string]$this.TeamResults.description
            Visibility   = [string]$this.TeamType
            Mail         = [string]$this.GroupResults.mail
            partitionKey = 'TeamsLog'
            rowKey       = $([string](new-guid).guid)
            TicketID     = $($this.TicketID)
            Status       = $(
                # Any needed tests to confirm that the team was successfully created
                switch ($this) {
                    { [string]::isnullorempty($_.TeamResults) } { [string]"FAILED" }
                    { -not [string]::isnullorempty($_.TeamResults.ID) } { [string]"SUCCESS" }
                    Default { "UNKNOWN" }
                }
            )
        }
    }
    [void]ExportLastObject() {
        # Can be used for debugging or manually ran instances of this class
        write-host "Exporting a copy of the last run object to an xml"
        Export-Clixml -InputObject $this .\TempObject.cli.xml
    }

    # This particular method uses all of the other methods in order to generate a new Microsoft Team automatically
    [void]AutoCreateTeam() {
        # Start by cleaning attributes as needed
        Write-Host "Cleaning attributes"
        $this.CleanAttributes()

        # Resolve the Team Owner (based on the ID that will be used in graph requests)
        Write-Host "Resolving Team Owner"
        $this.ResolveTeamOwner()

        # Generate the new mail nickname
        Write-Host "Generating a unique mail nick name"
        $this.GenerateMailNickname()

        # Generate the new group
        write-host "Generating a new group request via graph api"
        $this.NewGraphGroupRequest()

        # Wait for a few moments
        Start-Sleep -Seconds 10

        # Generate the new team (from the existing group)
        write-host "Generating a new teams request via graph api"
        $this.NewGraphTeamRequest()

        # Disable Adding Guests to the team
        Write-Host "Disabling adding Guests to the Team by default"
        $this.DisableAddingGuests()

        # Generate the results
        write-host "Gathering a report of the results"
        $this.GenerateResults()
    }
}
