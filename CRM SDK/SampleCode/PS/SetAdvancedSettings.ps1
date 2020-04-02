param
(
    #optional params
    [string]$ConfigurationEntityName,
    [string]$SettingName,
    [object]$SettingValue,
    [Guid]$Id
)

$RemoveSnapInWhenDone = $False

if (-not (Get-PSSnapin -Name Microsoft.Crm.PowerShell -ErrorAction SilentlyContinue))
{
    Add-PSSnapin Microsoft.Crm.PowerShell
    $RemoveSnapInWhenDone = $True
}

$setting = New-Object "Microsoft.Xrm.Sdk.Deployment.ConfigurationEntity"
$setting.LogicalName = $ConfigurationEntityName
if($Id) { $setting.Id = $Id }

$setting.Attributes = New-Object "Microsoft.Xrm.Sdk.Deployment.AttributeCollection"
$keypair = New-Object "System.Collections.Generic.KeyValuePair[String, Object]" ($SettingName, $SettingValue)
$setting.Attributes.Add($keypair)

Set-CrmAdvancedSetting -Entity $setting

if($RemoveSnapInWhenDone)
{
    Remove-PSSnapin Microsoft.Crm.PowerShell
}


