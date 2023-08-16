<#
.SYNOPSIS
This script provides options for collecting Summary/Detailed Message Trace information from EXO.
Before running the script, you need to be connected to Exchange Online Powershell:
https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
#>
using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][string]$SenderAddress,
    [Parameter(Mandatory = $false)][datetime]$StartDate,
    [Parameter(Mandatory = $false)][datetime]$EndDate,
    [Parameter(Mandatory = $false)][Int16]$NumberOfDaysInPast,
    [Parameter(Mandatory = $false)][bool]$DetailedReport = $false
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

        $CurrentMessageTrace = Get-MessageTrace -SenderAddress $SenderAddress -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page $i

        # if message trace results not null, add them to List, else break out of the loop
        if($null -ne $CurrentMessageTrace) {
            [void]$SummaryReport.Add($CurrentMessageTrace)
        }
        else {break}
        $i++

        Start-Sleep -m 500
    }

    # The result is an List of Arrays, we will need to export all of these to local file
    foreach($Report in $SummaryReport) {
        $Report| export-csv "$LogPath\SummaryReport.csv" -NoTypeInformation -Append
    }
    
    Write-Host "Exported Get-MessageTrace output to $LogPath\SummaryReport.csv"

    return $SummaryReport
}

function Get-SummaryReportDetail($SummaryReport)
{
    Write-Host -ForegroundColor Yellow "We have $SummaryReport.Count Get-MessageTraceDetail operations to run
    Do you want to run Get-MessageTraceDetail and export CSV output?
    Warning : This is a long running diagnostic
    A: Yes
    B: No"

    $MessageTraceDetailRequested = Read-Host

    switch ($MessageTraceDetailRequested) 
    {
        'A' 
        {
            $TotalNumberOfEmails = $SummaryReport.Count
            $i = 1
            foreach($MessageTrace in $SummaryReport)
            {
                Clear-Host
                Write-Host "Processing Email #$i from $TotalNumberOfEmails" 
                $MessageTraceDetail = Get-MessageTraceDetail -StartDate $StartDate -EndDate $EndDate -MessageTraceId $MessageTrace.MessageTraceId -RecipientAddress $MessageTrace.RecipientAddress
                $MTDReport += $MessageTraceDetail
                
                $i++
                Start-Sleep -m 500
            }
        
            $MTDReport | Select-Object -Property Date, MessageId, MessageTraceId, Event, Action, Detail, Data, FromIP, ToIP | Export-Csv "$LogPath\MessageTraceDetail.csv" -NoTypeInformation -Append
            Write-Host -ForegroundColor Green "Exported Get-MessageTraceDetail output to $LogPath\MessageTraceDetail.csv"
        }

        'B'
        {   
            Write-Host -ForegroundColor Green "Cancelled Get-MessageTraceDetail log collection"
            break
        }

        Default 
        {
            Get-SummaryReportDetail($SummaryReport)
        }
    }
}

################################################################################################
# Main Script Section

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
#Check if DetailedReport requested, and collect