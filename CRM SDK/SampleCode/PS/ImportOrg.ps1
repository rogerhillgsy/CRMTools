param
(
    #required params
    [string]$DatabaseName = $(throw "DatabaseName parameter not specified"),
    [string]$SqlServerName = $env:COMPUTERNAME,
    [string]$SrsUrl = "http://$SqlServerName/reportserver",
    [string]$UserMappingMethod = "ByAccount",
    
    #optional params (can accept nulls)
    [string]$DisplayName,
    [string]$Name,
    [switch]$WaitForJob
)

$RemoveSnapInWhenDone = $False

if (-not (Get-PSSnapin -Name Microsoft.Crm.PowerShell -ErrorAction SilentlyContinue))
{
    Add-PSSnapin Microsoft.Crm.PowerShell
    $RemoveSnapInWhenDone = $True
}

$opId = Import-CrmOrganization -SqlServerName $SqlServerName -DatabaseName $DatabaseName -SrsUrl $SrsUrl -DisplayName $DisplayName -Name $Name -UserMappingMethod $UserMappingMethod
$opId

if($WaitForJob)
{
    $opstatus = Get-CrmOperationStatus -OperationId $opid
    while($opstatus.State -eq "Processing")
    {
        Write-Host [(Get-Date)] Processing...
        Start-Sleep -s 30
        $opstatus = Get-CrmOperationStatus -OperationId $opid
    }

    if($opstatus.State -eq "Failed")
    {
        Throw ($opstatus.ProcessingError.Message)
    }

    Write-Host [(Get-Date)] Import org completed successfully.
}
else
{
    Write-Host [(Get-Date)] Import org async job requested.
}

if($RemoveSnapInWhenDone)
{
    Remove-PSSnapin Microsoft.Crm.PowerShell
}
