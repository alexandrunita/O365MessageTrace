<#
.SYNOPSIS
This script checks all valid recipients based on Summary Trace information exported by O365MessageTrace.ps1 script.
#>

using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$SummaryReportPath
)


################################################################################################
#region Main Script
$Summary = Import-CSV $SummaryReportPath

$ValidRecipientList = [List[PSObject]]::new()
$InvalidRecipientList = [List[string]]::new()
$Recipients = ($Summary|Group-Object RecipientAddress).Name
$TotalRecipientCount = $Recipients.Count

$ValidRecipientCount = 0
$i = 0
foreach($Recipient in $Recipients) {
    Write-Host -NoNewline "`rChecking Recipient $i from $TotalRecipientCount"
    $i++
    $CurrentRecipient = Get-EXORecipient $Recipient -ErrorAction SilentlyContinue | Select-Object PrimarySMTPAddress,RecipientTypeDetails
    if ($null -ne $CurrentRecipient) {
        $ValidRecipientList.Add($CurrentRecipient)
        $ValidRecipientCount++
    }
    else {$InvalidRecipientList.Add($Recipient)}
}
Write-Host -ForegroundColor Green -Message "#########################`r`nCount of valid Recipients"
$ValidRecipientCount

return $ValidRecipientList
#endregion