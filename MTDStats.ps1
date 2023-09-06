<#
.SYNOPSIS
This script generates statistics based on MTD information exported by O365MessageTrace.ps1 script
#>

using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$MTDReportPath
)


################################################################################################
#region Main Script
$MTDReport = Import-CSV $MTDReportPath

Write-Host -ForegroundColor Green -Message "#########################`r`nMTD Report Event Statistics"
$MTDReport|Group-Object Event|Sort-Object Count -descending|ft Name,Count

#Initializing FailDSNList
$FailDSNList = [List[string]]::new()
#TODO - Fail DSN stats
$FailMTD = $MTDReport|where{$_.Event -eq "Fail"}
if($FailMTD.count -ne 0) {
    foreach($Fail in $FailMTD) {
        if($Matches.count -ne 0) {$Matches.Clear()}
        $Fail.Detail -match "Reason: \[?{?(?'DSN'.*)}?;?" | Out-Null
        if($Matches['DSN'].Contains(";")) {
            $CurrentMatch = $Matches['DSN'].split(";")[0]
        }
        else {$CurrentMatch = $Matches['DSN']}
        $FailDSNList.Add($CurrentMatch)
    }
}
$FailDSNList|Group-Object|Sort-Object Count -descending|ft Name,Count

#TODO - Defer DSN stats
#endregion