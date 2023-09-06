<#
.SYNOPSIS
This script generates statistics based on Summary Trace information exported by O365MessageTrace.ps1 script
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$SummaryReportPath
)


################################################################################################
#region Main Script
$Summary = Import-CSV $SummaryReportPath

Write-Host -ForegroundColor Green -Message "#########################`r`nCount of unique Emails"
($summary|Group-Object MessageID).Name.Count

Write-Host -ForegroundColor Green -Message "#########################`r`nSummary Report Delivery Status Statistics"
$Summary|Group-Object status|Sort-Object Count -Descending|ft Name,Count

Write-Host -ForegroundColor Green -Message "#########################`r`nCount of unique Recipients"
$Recipients = ($Summary|Group-Object RecipientAddress).Name
$Recipients.Count

Write-Host -ForegroundColor Green -Message "#########################`r`nCount of Emails for top 10 Recipients"
$Summary|Group-Object RecipientAddress|Sort-Object Count -Descending|Select-Object -first 10|ft Name,Count

Write-Host -ForegroundColor Green -Message "#########################`r`nCount of Emails for top 10 Senders"
$Summary|Group-Object SenderAddress|Sort-Object Count -Descending|Select-Object -first 10|ft Name,Count
#endregion