function Get-AzureADDirectorySettingTemplateConfiguration {
    <#
    .SYNOPSIS
    Get the curent Azure directory settings template configuration
    
    .DESCRIPTION
    Get the curent Azure directory settings template configuration for o365 groups (teams)
    
    .EXAMPLE
    Get-AzureADDirectorySettingTemplateConfiguration

    .NOTES
    This command requires the AzureADPreview module.
    #>
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        
    }
    
    process {
        try {
            $Template = Get-AzureADDirectorySettingTemplate | Where-Object -Property Id -Value "62375ab9-6b52-47ed-826b-58e47e0e304b" -EQ
            $Setting = $template.CreateDirectorySetting()
            $Setting.Values
        }
        catch {
            throw
        }
    }
    
    end {
    }
}

function Set-AzureADDirectorySettingTemplateConfiguration {
    <#
    .SYNOPSIS
    Sets the Azure active directory settings template configuration.
    
    .DESCRIPTION
    Sets the Azure active directory settings template configuration for O365 groups (Teams).
    
    .PARAMETER AllowToAddGuests
    True false boolean, allows the team to add external guests to teams.
    
    .EXAMPLE
    Set-AzureADDirectorySettingTemplateConfiguration -AllowToAddGuests $true
    
    .EXAMPLE
    Set-AzureADDirectorySettingTemplateConfiguration -AllowToAddGuests $false

    .NOTES
    This command requires the AzureADPreview module.
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)][bool]$AllowToAddGuests
    )
    
    begin {
        
    }
    
    process {
        try {
            $Template = Get-AzureADDirectorySettingTemplate | Where-Object -Property Id -Value "62375ab9-6b52-47ed-826b-58e47e0e304b" -EQ
            $Setting = $template.CreateDirectorySetting()
            $Setting["AllowToAddGuests"] = $False
        }
        catch {
            throw
        }

        try {
            Set-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ).id -DirectorySetting $Setting
            $Setting.Values
        }
        catch {
            throw
        }
    }
    
    end {
    }
}