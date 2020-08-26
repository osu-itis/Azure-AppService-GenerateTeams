# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function started TIME: $currentUTCtime"

# Checking if the needed ENVs exist:
if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) { Throw 'Could not find $env:AzureWebJobsStorage' }
if ([string]::IsNullOrEmpty($env:ClientID)) { Throw 'Could not find $env:ClientID' }
if ([string]::IsNullOrEmpty($env:ClientSecret)) { Throw 'Could not find $env:ClientSecret' }
if ([string]::IsNullOrEmpty($env:TenantId)) { Throw 'Could not find $env:TenantId' }
if ([string]::IsNullOrEmpty($env:ServiceAccountUsername)) { Throw 'Could not find $env:ServiceAccountUsername' }
if ([string]::IsNullOrEmpty($env:ServiceAccountPassword)) { Throw 'Could not find $env:ServiceAccountPassword' }

# Custom object that contains most of the important attributes used by the script
$Object = [pscustomobject]@{

    # Generating the service credential that is needed to make powershell calls
    ServiceAccountCredential   = $(
        # Using the env variables set in the azure function app, generate a credential object
        $userName = $env:ServiceAccountUsername
        $userPassword = $env:ServiceAccountPassword
        [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
        [pscredential]$cred = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
        $cred
    )
    # Setting the timeout for the queue message
    invisibleTimeout           = [System.TimeSpan]::FromSeconds(10)

    QueueName                  = "timed-gal-changes"
    TenantID                   = $env:TenantID
    AzureWebJobsStorage        = $env:AzureWebJobsStorage
    # Generating the needed credential object to connect to the storage queue
    ServicePrincipalCredential = $(
        $userName = $env:ClientID
        $userPassword = $env:ClientSecret
        [securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
        [pscredential]$Cred = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
        $Cred
    )
    # Gathering the context of the storage account, which is needed later
    StorageAccountName         = $(
        $env:AzureWebJobsStorage.split(";")[1].replace("AccountName=", "")
    )
    # Getting the needed resource group
    ResourceGroup              = "Infra-TeamsAutomation"
}

# Connecting to AZ Account
$null = Connect-AzAccount -Credential $Object.ServicePrincipalCredential -Tenant $Object.TenantID -ServicePrincipal

# Gathering the storage context
$CTX = New-AzStorageContext -ConnectionString $Object.AzureWebJobsStorage | Select-Object -ExpandProperty Context

# Checking the queue
$queue = Get-AzStorageQueue -Context $CTX -Name $Object.QueueName

# If there are some items in the queue, process them
if ($queue.CloudQueue.ApproximateMessageCount -gt 0) {
    # Create the new PSSession and then import it
    $o365session = New-PSSession -configurationname Microsoft.Exchange -connectionuri https://outlook.office365.com/powershell-liveid/ -credential $Object.ServiceAccountCredential -authentication basic -allowredirection
    $null = Import-PSSession $o365session -allowclobber -disablenamechecking
    if (! [string]::IsNullOrWhiteSpace($(Get-PSSession | Where-Object { $_.configurationname -eq "Microsoft.Exchange" }))) {
        Write-Host "Loaded Exchange Online"
    }
    else {
        ThrowError "Failed to load Exchange Online"
    }

    # Loop to run through all queue messages
    do {
        # Get the next queue message in the list
        $queueMessage = $queue.CloudQueue.GetMessageAsync($invisibleTimeout, $null, $null)

        try {
            # Attempt to set the group's visbility in the GAL
            Set-UnifiedGroup $queueMessage.Result.AsString -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            Write-Host "Setting $($queuemessage.Result.AsString) as hidden from the Global Address List"
            # Remove the Queue Message as it is no longer needed, dont output to console
            $null = $queue.CloudQueue.DeleteMessageAsync($queueMessage.Result.Id, $queueMessage.Result.popReceipt)
        }
        catch {
            "Could not find $($queuemessage.Result.AsString) in exchange, will try again later..."
        }

    } until ($null -eq $QueueMessage.Result)
    # Remove the session now that we no longer need the exchange module and powershell commands
    Remove-PSSession -Session $o365session
}
else {
    Write-Host "No items found pending in the queue"
}