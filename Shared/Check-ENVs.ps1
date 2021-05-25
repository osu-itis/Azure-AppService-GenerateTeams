if ([string]::IsNullOrEmpty($env:AzureWebJobsStorage)) {
    Throw 'Could not find $env:AzureWebJobsStorage' 
}
if ([string]::IsNullOrEmpty($env:ClientID)) {
    Throw 'Could not find $env:ClientID' 
}
if ([string]::IsNullOrEmpty($env:ClientSecret)){
    Throw 'Could not find $env:ClientSecret' 
}
if ([string]::IsNullOrEmpty($env:TenantId)) {
    Throw 'Could not find $env:TenantId' 
}
