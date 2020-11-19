# Import the required AzureAD Module: https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0
Import-Module AzureAD

# Get all of the Unified Groups
$groupID = Get-UnifiedGroup -ResultSize Unlimited | Select-Object -ExpandProperty ExternalDirectoryObjectId

# Iterate through the groups and set the "Group.Unified.Guest" to "False"
Foreach ($Groups in $GroupID) {
    $template = Get-AzureADDirectorySettingTemplate | Where-Object { ​​​​​​​$_.displayname -eq "group.unified.guest" }​​​​​​​
    $settingsCopy = $template.CreateDirectorySetting()
    $settingsCopy["AllowToAddGuests"] = $False
    New-AzureADObjectSetting -TargetType Groups -TargetObjectId $groups -DirectorySetting $settingsCopy
}​​​​​​​
