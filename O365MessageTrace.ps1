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
    [Parameter(Mandatory = $false)][string]$SenderAddress,
    [Parameter(Mandatory = $false)][string]$RecipientAddress,
    [Parameter(Mandatory = $false)][string]$MessageId,
    [Parameter(Mandatory = $false)][string[]]$DeliveryStatuses = @("GettingStatus", "Failed", "Pending", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam"),
    [Parameter(Mandatory = $false)][datetime]$StartDate = [DateTime]::UtcNow.AddDays(-2),
    [Parameter(Mandatory = $false)][datetime]$EndDate = [DateTime]::UtcNow,
    [Parameter(Mandatory = $false)][bool]$IncludeExtendedSummary = $false
)

function Get-SummaryReport {
    param (
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate,
        [Parameter(Mandatory = $true)][string[]]$DeliveryStatuses,
        [Parameter(Mandatory = $false)][string]$SenderAddress,
        [Parameter(Mandatory = $false)][string]$RecipientAddress,
        [Parameter(Mandatory = $false)][string]$MessageId
    )
    # initialize Empty generic list
    $SummaryReport = [List[PSObject]]::new()
    $i = 1

    #Start building Get-MessageTrace cmdlet expression based on input params
    [string]$GetMessageTraceExpression = "Get-MessageTrace -Status $($DeliveryStatuses -join ',') -StartDate $($StartDate.Ticks) -EndDate $($EndDate.Ticks)"

    # check if MessageId provided
    if($MessageId.Length -ne 0) {
        $GetMessageTraceExpression += " -MessageId '$MessageId'"
    }

    # check if SenderAddress provided
    if($SenderAddress.Length -ne 0) {
        $GetMessageTraceExpression += " -SenderAddress $SenderAddress"
    }

    # check if RecipientAddress provided
    if($RecipientAddress.Length -ne 0) {
        $GetMessageTraceExpression += " -RecipientAddress $RecipientAddress"
    }

    # add pagination values
    $GetMessageTraceExpression +=  " -PageSize 5000 -Page $i"

    while($true) {
        Write-Host -NoNewline "`rCollecting Summary Trace Page $i"

        #Update expression Page number starting with second loop
        if($i -gt 1) {
            $GetMessageTraceExpression = $GetMessageTraceExpression.Substring(0, $GetMessageTraceExpression.Length - "-Page $($i -1)".Length) + "-Page $i"
        }

        #Invoke Expression we built to query EXO for required message trace information
        [PSObject[]]$CurrentMessageTrace = Invoke-Expression $GetMessageTraceExpression
        
        # if message trace results not null, add them to List&Export to local drive, else break out of the loop
        if($null -ne $CurrentMessageTrace) {
            [void]$SummaryReport.AddRange($CurrentMessageTrace)
            $CurrentMessageTrace | Export-Csv "$LogPath\SummaryReport.csv" -NoTypeInformation -Append
        }
        else {break}
        $i++

        Start-Sleep -m 200
    }

    # line feed after loop ended
    Write-Host
    
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

function Get-ExtendedSummaryReport {
    param (
        [Parameter(Mandatory = $true)][psobject[]]$SummaryReport,
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate
    )

    #initialize empty generic list
    $MTDReport = [List[PSObject]]::new()
    $MTDEmpty = $true
    
    #Counters for loop
    $i = 1
    $NumberOfLoops = $SummaryReport.Count
    $ResetAfter1000Loops = 0

    #iterate through each Summary Report Entry and retrieve MTD data
    foreach($Report in $SummaryReport) {

        Write-Host -NoNewline "`rCollecting MessageTraceDetail Entry $i out of $NumberOfLoops"

        [PSObject[]]$CurrentMTD = Get-MessageTraceDetail -StartDate $StartDate -EndDate $EndDate -MessageTraceId $Report.MessageTraceId -RecipientAddress $Report.RecipientAddress | Select-Object Date, MessageId, MessageTraceId, @{Name="SenderAddress";expression={$Report.SenderAddress}}, @{Name="RecipientAddress";expression={$Report.RecipientAddress}}, Event, Action, Detail, Data

        # Check if MTD entry null, do not export if null
        if($null -ne $CurrentMTD) {
            [void]$MTDReport.AddRange($CurrentMTD)
            #make sure MTDEmpty is false, data was retrieved
            $MTDEmpty = $false
            #Skip for now, too many writes to disk, too slow
            #$CurrentMTD | Export-Csv "$LogPath\MTDReport.csv" -NoTypeInformation -Append

        }

        $i++
        $ResetAfter1000Loops++
        # Check if 1000 loops performed, export and clear list, then reset counter
        if($ResetAfter1000Loops -eq 1000) {
            $MTDReport | Export-Csv "$LogPath\MTDReport.csv" -NoTypeInformation -Append
            $MTDReport.Clear()
            $ResetAfter1000Loops = 0
        }

        Start-Sleep -m 200
    }

    # line feed after loop ended
    Write-Host

    #provide feedback to console if no emails found
    if($MTDEmpty) {
        Write-Warning "Extended Summary Report found no data for emails form Summary Report"
    }
    else {
        # see if it makes sense, will be a lot of memory used though
        $MTDReport | Export-Csv "$LogPath\MTDReport.csv" -NoTypeInformation -Append
        Write-Host -ForegroundColor Green "Exported Get-MessageTraceDetail output to
        $LogPath\MTDReport.csv"
    }

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

#Check if Delivery Status input is valid
foreach($Status in $DeliveryStatuses) {
    if ($Status -notin @("GettingStatus", "Failed", "Pending", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam")) {
        Write-Error -ErrorAction Continue "Invalid Delivery Status list provided, valid values are:
        'GettingStatus', 'Failed', 'Pending', 'Delivered', 'Expanded', 'Quarantined', 'FilteredAsSpam'
        The script will now exit, retry with valid input."
        Exit
    }
}

#Create Log File Directory on Desktop
$ts = Get-Date -Format yyyyMMdd_HHmm_ff
$LogPath=[Environment]::GetFolderPath("Desktop")+"\MessageTraceScript\$($ts)_MessageTrace"
Write-Host "Created Directory on Desktop:"
mkdir "$LogPath"

#For long running task Get-ExtendedSummaryReport we need to keep Windows Alive
#Creating object to call to hit key
#$KeepPCAlive = New-Object -ComObject WScript.Shell #Credit for this option goes to : https://gist.github.com/jamesfreeman959/231b068c3d1ed6557675f21c0e346a9c

#If no Start/End Dates provided by end user, default to StartDate 2 days ago and EndDate now
<#if($StartDate -eq $null) {
    $StartDate = ([DateTime]::UtcNow.AddDays(-2))
}
if($EndDate -eq $null) {
    $EndDate = ([DateTime]::UtcNow)
}#>

$SummaryReport = $null
$MTDReport = $null

#Collect Summary Report
$SummaryReport = Get-SummaryReport -StartDate $StartDate -EndDate $EndDate -DeliveryStatuses $DeliveryStatuses -SenderAddress $SenderAddress -RecipientAddress $RecipientAddress -MessageId $MessageId
#Check if ExtendedSummary requested and SummaryReport not empty before attempting to collect Extended Summary
if($IncludeExtendedSummary -and ($null -ne $SummaryReport)) {
    Get-ExtendedSummaryReport -StartDate $StartDate -EndDate $EndDate -SummaryReport $SummaryReport
}