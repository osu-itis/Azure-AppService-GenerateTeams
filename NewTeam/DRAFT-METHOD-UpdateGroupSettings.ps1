# Setting a fail test status to true, If the code executes properly this will be toggled to false
$FailTest = $true

# Setting the starting increment for our loop
$Increment = 0

# The Group ID to be tested
$GroupID = ""

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
    $GroupInfo = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"

    # If the Group info was returned...
    if ( ! [string]::IsNullOrEmpty($GroupInfo) ) {
        # Gather if it already has a "Group.Unified.Guest", if so, we need to patch. Otherwise we need to post.
        $PreChange = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"
        if ( $($PreChange.value | Where-Object { $_.displayname -Like "Group.Unified.Guest" }).count -eq 1 ) {
            # Patching the setting on an object that already had the setting configured
            $null = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings/$($PreChange.value.id)" -Method 'patch' -Body $body -ContentType 'application/json' -Authentication Bearer -Token $obj.GraphTokenString
        }
        else {
            # Posting the setting for the first time on the object
            $null = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Method 'POST' -Body $body -ContentType 'application/json' -Authentication Bearer -Token $obj.GraphTokenString
        }
        # Gathering the results of the change to be sure that the proper settings have been applied
        $PostChange = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($GroupID)/settings" -Authentication Bearer -Token $obj.GraphTokenString -Method "get"
    
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
    {$_ -eq $true} {throw "Failed to set the 'AllowToAddGuests' attribute to False."}
    {$_ -eq $false} {Write-Host "Set the 'AllowToAddGuests' attribute to False."}
}
