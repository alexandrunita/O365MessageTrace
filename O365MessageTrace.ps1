<#
.SYNOPSIS
This script provides options for collecting Summary/Detailed Message Trace information from EXO.

By default, it will retrieve logs for the past 2 days, to customize within the maximum last 10 day interval you must provide StartDate & EndDate parameters.

By default, the script will export all traffic, you may filter the query based on SenderAddress.

By default, the script will only collect and export Summary Report data, you must request Extended Summary via dedicated switch to also get that output.
Please be aware Extended Summary collection will be a long running task especially on production tenants with large email volume.
#>
using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][string]$SenderAddress = $null,
    [Parameter(Mandatory = $false)][datetime]$StartDate,
    [Parameter(Mandatory = $false)][datetime]$EndDate,
    [Parameter(Mandatory = $false)][Int16]$NumberOfDaysInPast,
    [Parameter(Mandatory = $false)][bool]$IncludeExtendedSummary = $false
)

function Get-SummaryReport
{
    param (
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate,
        [Parameter(Mandatory = $false)][string]$SenderAddress
    )
    # initialize Empty generic list
    $SummaryReport = [List[PSObject]]::new()
    $i = 1
    while($true)
    {
        Write-Host "Collecting information from MessageTrace Trace Page $i"

        # check if SenderAddress provided
        if($SenderAddress.Length -ne 0){
            [PSObject[]]$CurrentMessageTrace = Get-MessageTrace -SenderAddress $SenderAddress -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page $i
        }
        else {
            [PSObject[]]$CurrentMessageTrace = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page $i
        }

        # if message trace results not null, add them to List&Export to local drive, else break out of the loop
        if($null -ne $CurrentMessageTrace) {
            [void]$SummaryReport.AddRange($CurrentMessageTrace)
            $CurrentMessageTrace | Export-Csv "$LogPath\SummaryReport.csv" -NoTypeInformation -Append
        }
        else {break}
        $i++

        Start-Sleep -m 500
    }

    # provide feedback to console if no emails found
    if($null -eq $SummaryReport[0]) {
        Write-Warning "No Emails found for this query"
    }    
    else {
        Write-Host -ForegroundColor Green "Exported Get-MessageTrace output to 
        $LogPath\SummaryReport.csv"
    }
    

    return $SummaryReport
}

function Get-ExtendedSummaryReport
{
    param (
        [Parameter(Mandatory = $true)][psobject[]]$SummaryReport,
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate
    )

    #initialize empty generic list
    $MTDReport = [List[PSObject]]::new()
    $i = 1

    #iterate through each Summary Report Entry and retrieve MTD data
    foreach($Report in $SummaryReport) {

        Write-Host "Collecting information from MessageTraceDetail Page $i"

        [PSObject[]]$CurrentMTD = Get-MessageTraceDetail -StartDate $StartDate -EndDate $EndDate -MessageTraceId $Report.MessageTraceId -RecipientAddress $Report.RecipientAddress | Select-Object Date, MessageId, MessageTraceId, @{Name="SenderAddress";expression={$Report.SenderAddress}}, @{Name="RecipientAddress";expression={$Report.RecipientAddress}}, Event, Action, Detail, Data

        # Check if MTD entry null, do not export if null
        if($null -ne $CurrentMTD) {
            [void]$MTDReport.AddRange($CurrentMTD)
            $CurrentMTD | Export-Csv "$LogPath\MTDReport.csv" -NoTypeInformation -Append
        }
        $i++

        Start-Sleep -m 500
    }

    #provide feedback to console if no emails found
    if($null -eq $MTDReport[0]) {
        Write-Warning "Extended Summary Report found no data for emails form Summary Report"
    }
    else {
        Write-Host -ForegroundColor Green "Exported Get-MessageTraceDetail output to
        $LogPath\MTDReport.csv"
    }

    return $MTDReport
}

################################################################################################
# Main Script Section

#Recommended Exchange Online Module Version
[string]$RecEXOModuleVersion = "3.2.0"
#Check if EXO Module imported
if($null -eq (Get-Module "ExchangeOnlineManagement")) {
    Write-Error -ErrorAction Continue "ExchangeOnlineManagement Module not imported, exiting script.
    Check Microsoft article on how to install and connect : 
    https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"
    Exit
}
#Provide warning if version lower than recommended version, recommending update
elseif ([System.Version]$RecEXOModuleVersion -gt (Get-Module "ExchangeOnlineManagement").Version) {
    Write-Warning -Message "Running older version than recommended Exchange Online Module Version $RecEXOModuleVersion, script may fail. Consider updating:
    https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#update-the-exchange-online-powershell-module"
}
#Check if Get-MessageTrace cmdlet is recognized
try {
    Get-MessageTrace -MessageId "BogusMessageID"
}
catch {
    Write-Error -ErrorAction Continue -Message "Failed to run test Get-MessageTrace cmdlet
    Check if Admin account has proper access rights, if the Connection to Exchange Online Powershell is successful and try again:
    https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
    https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/message-trace-scc?view=o365-worldwide#:~:text=You%20need%20to%20be,features%20in%20Microsoft%20365.
    
    Exiting the script"
    Exit
}
#Create Log File Directory on Desktop
$ts = Get-Date -Format yyyyMMdd_HHmm
$LogPath=[Environment]::GetFolderPath("Desktop")+"\MessageTraceScript\$($ts)_MessageTrace"
Write-Host "Created Directory on Desktop:"
mkdir "$LogPath"

#If no Start/End Dates provided by end user, default to StartDate 2 days ago and EndDate now
if($StartDate -eq $null) {
    $StartDate = ([DateTime]::UtcNow.AddDays(-2))
}
if($EndDate -eq $null) {
    $EndDate = ([DateTime]::UtcNow)
}

$SummaryReport = $null
$MTDReport = $null

#Collect Summary Report
$SummaryReport = Get-SummaryReport -SenderAddress $SenderAddress -StartDate $StartDate -EndDate $EndDate
#Check if ExtendedSummary requested and SummaryReport not empty before attempting to collect Extended Summary
if($IncludeExtendedSummary -and ($null -ne $SummaryReport))
{
    $MTDReport = Get-ExtendedSummaryReport -StartDate $StartDate -EndDate $EndDate -SummaryReport $SummaryReport
}