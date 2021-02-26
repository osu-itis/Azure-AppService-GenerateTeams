function New-ServiceAccountCredential {
    <#
    .EXAMPLE
    $ServiceAccountCredential = New-ServiceAccountCredential -ClientID $env:ClientID -ClientSecret $env:ClientSecret
    #>
    [CmdletBinding()]
    param (
        [string]$ClientID,
        [string]$ClientSecret
    )

    # Converting client secret to a secure string password
    [securestring]$secStringPassword = ConvertTo-SecureString $ClientSecret -AsPlainText -Force

    # Generating the PSCredential Object
    [pscredential]$OutputCred = New-Object System.Management.Automation.PSCredential ($ClientID, $secStringPassword)

    # Returning the credential as the output
    return $OutputCred
}

function Connect-CloudTable {
    <#
    .EXAMPLE
    $cloudTable = Connect-CloudTable -ServiceAccountCredential $ServiceAccountCredential -TenantId $env:TenantID -AzureWebJobsStorage $env:AzureWebJobsStorage -tableName $tableName
    #>
    [CmdletBinding()]
    param (
        $ServiceAccountCredential,
        $TenantId,
        $AzureWebJobsStorage,
        $tableName
    )

    # Connecting to Azure using credentials
    $null = Connect-AzAccount -Credential $ServiceAccountCredential -Tenant $TenantId -ServicePrincipal

    # Creating the storage context needed to query Azure tables
    $CTX = (New-AzStorageContext -ConnectionString $AzureWebJobsStorage).context

    # Gathering the cloud table information
    $cloudTable = $(Get-AzStorageTable -Name $tableName -Context $CTX) | Select-Object -ExpandProperty "CloudTable"

    # Returning the cloud table information
    return $cloudTable
}

function Get-FunctionTableName {
    <#
    .EXAMPLE
    $tableName = Get-FunctionTableName -Path '.\FunctionFolder\function.json'
    #>
    [CmdletBinding()]
    param (
        $Path
    )

    # Getting the TableName from the function's parameters
    $tableName = ( (Get-Content -Path $Path | ConvertFrom-Json).bindings.tablename | Out-String ).trim()

    # Returning the table name
    return $tableName
}

function Set-ExemptStatus {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]$cloudTable,
        [parameter(Mandatory = $true)][string]$FilterParameter,
        [parameter(Mandatory = $true)][string]$FilterValue,
        [parameter(Mandatory = $true)][bool]$Exempt
    )

    try {
        $ROW = Get-AzTableRow -Table $cloudTable -ColumnName $FilterParameter -Value $FilterValue -Operator Equal
    }
    catch {
        Write-Error "Failed to find a matching result"
    }

    if ([string]::IsNullOrWhiteSpace($ROW)) {
        write-warning -Message "Could not find an entry in the table for $FilterValue"
    }
    else {
        $ROW | Add-Member -NotePropertyName 'Exempt' -NotePropertyValue $Exempt -Force
    
        $( $ROW | Update-AzTableRow -Table $cloudTable )
    
        return  $ROW
    }
}
