## https://github.sig.oregonstate.edu/ExchangeOnline/allowguestsinteams



Import-module AzureADPreview
$GroupName = "Gem"
Connect-AzureAD
$template = Get-AzureADDirectorySettingTemplate | ? {$_.displayname -eq "group.unified.guest"}
$settingsCopy = $template.CreateDirectorySetting()
$settingsCopy["AllowToAddGuests"]=$True
$groupID= (Get-AzureADGroup -SearchString $GroupName).ObjectId
$id = (Get-AzureADObjectSetting -TargetObjectId $groupID -TargetType Groups).id
set-AzureADObjectSetting -TargetType Groups -TargetObjectId $groupID -Id $id -DirectorySetting $settingsCopy
Get-AzureADObjectSetting -TargetObjectId $groupID -TargetType Groups | fl Value