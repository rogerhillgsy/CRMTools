param
(
    #required params
    [string]$Name = $(throw "Name parameter not specified")
)

$RemoveSnapInWhenDone = $False

if (-not (Get-PSSnapin -Name Microsoft.Crm.PowerShell -ErrorAction SilentlyContinue))
{
    Add-PSSnapin Microsoft.Crm.PowerShell
    $RemoveSnapInWhenDone = $True
}

Enable-CrmServer -Name $Name

if($RemoveSnapInWhenDone)
{
    Remove-PSSnapin Microsoft.Crm.PowerShell
}
