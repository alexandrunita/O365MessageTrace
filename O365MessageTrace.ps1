<#
.SYNOPSIS
This script provides options for collecting Summary/Detailed Message Trace information from EXO.

.DESCRIPTION
By default, it will retrieve logs for the past 2 days, to customize within the maximum last 10 day interval you must provide StartDate & EndDate parameters.

By default, the script will export all traffic, you may filter the query based on SenderAddress.

By default, the script will only collect and export Summary Report data, you must request Extended Summary via dedicated switch to also get that output.
Please be aware Extended Summary collection will be a long running task especially on production tenants with large email volume.

.PARAMETER StartDate
DateTime .Net object must be passed, taking into account that Exchange Online Message Trace data always reflects UTC time.
If omitted, script will default to value : [DateTime]::UtcNow.AddDays(-2) .
For more information: https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=net-7.0

.PARAMETER EndDate
DateTime .Net object must be passed, taking into account that Exchange Online Message Trace data always reflects UTC time.
If omitted, script will default to value : [DateTime]::UtcNow .
For more information: https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=net-7.0

.PARAMETER SenderAddress
Single Sender Email Address used to query Exchange Online Message Trace data.
Accepts wildcard value such as "*@contoso.com".

.PARAMETER RecipientAddress
Single Recipient Email Address used to query Exchange Online Message Trace data.
Accepts wildcard value such as "*@contoso.com".

.PARAMETER MessageId
MessageId header value used to query Exchange Online Message Trace data.
For more information: https://learn.microsoft.com/en-us/powershell/module/exchange/get-messagetrace?view=exchange-ps#-messageid 

.PARAMETER DeliveryStatuses
Email status MultiValuedProperty used to query Exchange Online Message Trace Data.
Valid values are : GettingStatus, Failed, Pending, Delivered, Expanded, Quarantined, FilteredAsSpam
For more information: https://learn.microsoft.com/en-us/powershell/module/exchange/get-messagetrace?view=exchange-ps#-status

.PARAMETER IncludeExtendedSummary
Will default to $false, must be explicitly provided with $true value to instruct the script to contain Get-MessageTraceDetail logs from Exchange Online.

.EXAMPLE
.\O365MessageTrace.ps1 -DeliveryStatuses Failed, Pending
Retrieves all emails where Delivery Status is Failed or Pending in the past 48 hours.

.EXAMPLE
.\O365MessageTrace.ps1 -SenderAddress user@contoso.com -StartDate ([datetime]::UtcNow).AddDays(-10)
Retrieves all email sent by user@contoso.com starting with 10 days ago, EndDate will default to [datetime]::UtcNow as it was not specified.

.EXAMPLE
.\O365MessageTrace.ps1 -SenderAddress user@contoso.com -RecipientAddress user@fabrikam.com -MessageId "<1241241@contoso.com>"
Retrieves events for the email with the specified MessageId string where SenderAddress and RecipientAddress match provided inputs.

.EXAMPLE
.\O365MessageTrace.ps1 -RecipientAddress user@contoso.com -IncludeExtendedSummary $true
Retrieves all emails sent to user@contoso.com in the past 48 hours and exports both Summary Report and Enhanced Summary Report that contains extra routing information which may assist in advanced troubleshooting.

#>
using namespace System.Collections.Generic

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][string]$SenderAddress,
    [Parameter(Mandatory = $false)][string]$RecipientAddress,
    [Parameter(Mandatory = $false)][string]$MessageId,
    [Parameter(Mandatory = $false)][string]$FromIP,
    [Parameter(Mandatory = $false)][string]$ToIP,
    [Parameter(Mandatory = $false)]
        [ValidateSet("GettingStatus", "Failed", "Pending", "Delivered", "Expanded", "Quarantined", "FilteredAsSpam")]
        [string[]]$DeliveryStatuses,
    [Parameter(Mandatory = $false)][datetime]$StartDate = [DateTime]::UtcNow.AddDays(-2),
    [Parameter(Mandatory = $false)][datetime]$EndDate = [DateTime]::UtcNow,
    [Parameter(Mandatory = $false)][bool]$IncludeExtendedSummary = $false
)

function Get-SummaryReport {
    param (
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate,
        [Parameter(Mandatory = $false)][string[]]$DeliveryStatuses,
        [Parameter(Mandatory = $false)][string]$SenderAddress,
        [Parameter(Mandatory = $false)][string]$RecipientAddress,
        [Parameter(Mandatory = $false)][string]$MessageId,
        [Parameter(Mandatory = $false)][string]$FromIP,
        [Parameter(Mandatory = $false)][string]$ToIP
    )
    # initialize Empty generic list
    $SummaryReport = [List[PSObject]]::new()
    $i = 1

    #Start building Get-MessageTrace cmdlet expression based on input params
    [string]$GetMessageTraceExpression = "Get-MessageTrace -StartDate $($StartDate.Ticks) -EndDate $($EndDate.Ticks)"

    # check if DeliveryStatus provided
    if($DeliveryStatuses.Length -ne 0) {
        $GetMessageTraceExpression += " -Status $($DeliveryStatuses -join ',')"
    }

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

    if($FromIP.Length -ne 0) {
        $GetMessageTraceExpression += " -FromIP $FromIP"
    }

    if($ToIP.Length -ne 0) {
        $GetMessageTraceExpression += " -ToIP $ToIP"
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

        [PSObject[]]$CurrentMTD = Get-MessageTraceDetail -StartDate $StartDate -EndDate $EndDate -MessageTraceId $Report.MessageTraceId -RecipientAddress $Report.RecipientAddress | Select-Object Date, MessageId, MessageTraceId, @{Name="SenderAddress";expression={$Report.SenderAddress}}, @{Name="RecipientAddress";expression={$Report.RecipientAddress}}, @{Name="Subject";expression={$Report.Subject}}, Event, Action, Detail, Data

        # Check if MTD entry null, do not export if null
        if($null -ne $CurrentMTD) {
            #must truncate any events after resolve, none will be relevant for current recipient email address, they will just duplicate events for the new recipient email address
            if($CurrentMTD.Event -contains "Resolve") {
                # make sure logs are in chronological order
                $CurrentMTD = $CurrentMTD|Sort-Object Date
                $CurrentMTD = $CurrentMTD[0..$CurrentMTD.Event.IndexOf("Resolve")]
            }
            #must check if Drop Event occurred for original recipient
            #if Drop event occurred, we must truncate MTDReport array to avoid confusing output
            #by default Get-MessageTraceDetail will provide output for the email entity
            #it will not take into account recipient change in transit
            #For example : we send email to originalRecipient@contoso.com who redirects email to newRecipient@contoso.com
            #originalRecipient@contoso.com has DeliverToMailboxAndForward : False
            #Get-MessageTraceDetail -MessageID "xyz" -RecipientAddress originalRecipient@contoso.com will display output tracking email entity to delivery to newRecipient@contoso.com.
            #It does not store recipient address, this can confuse someone who looks at the report that an email was Delivered/SentExternal/etc.. to originalRecipient@contoso.com 
            #although the email was actually dropped for that recipient and the events we see afterwards actually belong to email entity and describe handling for newRecipient@contoso.com
            if($CurrentMTD.Event -contains "Drop") {
                # make sure logs are in chronological order
                $CurrentMTD = $CurrentMTD|Sort-Object Date
                $CurrentMTD = $CurrentMTD[0..$CurrentMTD.Event.IndexOf("Drop")]
            }
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
        Write-Warning "Extended Summary Report found no data for emails from Summary Report"
    }
    else {
        # see if it makes sense, will be a lot of memory used though
        $MTDReport | Export-Csv "$LogPath\MTDReport.csv" -NoTypeInformation -Append
        Write-Host -ForegroundColor Green "Exported Get-MessageTraceDetail output to
        $LogPath\MTDReport.csv"
    }

}

################################################################################################
#region Main Script

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

#If Delivery Status input provided, perform Delivery Status de-duplication for user input
#this prevents errors due to duplicate statuse when running Get-MessageTrace later on during logic
if($DeliveryStatuses.Length -ne 0){
    $DeliveryStatuses = $DeliveryStatuses|Select-Object -Unique
}

#Create Log File Directory on Desktop
$ts = Get-Date -Format yyyyMMdd_HHmm_ss_ff
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
$SummaryReport = Get-SummaryReport -StartDate $StartDate -EndDate $EndDate -DeliveryStatuses $DeliveryStatuses -SenderAddress $SenderAddress -RecipientAddress $RecipientAddress -MessageId $MessageId -FromIP $FromIP -ToIP $ToIP
#Check if ExtendedSummary requested and SummaryReport not empty before attempting to collect Extended Summary
if($IncludeExtendedSummary -and ($null -ne $SummaryReport)) {
    Get-ExtendedSummaryReport -StartDate $StartDate -EndDate $EndDate -SummaryReport $SummaryReport
}

#endregion