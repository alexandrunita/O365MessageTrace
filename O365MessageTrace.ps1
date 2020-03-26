function Connect-O365EXOPS
{
    $Error.Clear()

    Set-ExecutionPolicy RemoteSigned

    $O365GlobalAdminCredentials = Get-Credential

    $O365EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365GlobalAdminCredentials -Authentication Basic -AllowRedirection -ErrorAction Continue

    Import-PSSession $O365EXOSession -AllowClobber
}

function Get-O365MessageTrace
{
    $i = 1
    do
    {
        Write-Host "Collecting information from MessageTrace Trace Page $i"

        $CurrentMessageTrace = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize 5000 -Page $i

        $SummaryReport += $CurrentMessageTrace
        $i++

        Start-Sleep -m 500
    }
    while($null -ne  $CurrentMessageTrace)

    $SummaryReport | Select-Object -Property Received, SenderAddress, RecipientAddress, Subject, MessageId, MessageTraceId, Status, ToIP, FromIP, Size | export-csv "$LogPath\SummaryReport.csv" -NoTypeInformation -Append

    return $SummaryReport

    Write-Host "Exported Get-MessageTrace output to $LogPath\SummaryReport.csv"
}

function Get-O365MessageTraceDetail($SummaryReport)
{
    Write-Host -ForegroundColor Yellow "Do you want to run Get-MessageTraceDetail and export CSV output?
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
            Get-O365MessageTraceDetail($SummaryReport)
        }
    }
}


# Main Script Section
###Calls EXO Connection function
###Collects Get-MessageTrace & Get-MessageTraceDetail output
###Writes Log files on Desktop

$ts = Get-Date -Format yyyyMMdd_HHmm
$LogPath=[Environment]::GetFolderPath("Desktop")+"\$($ts)_MessageTrace"
Write-Host "Created Directory on Desktop:"
mkdir "$LogPath"

$StartDate = ([DateTime]::Now.AddDays(-10))
$EndDate = ([DateTime]::Now)

$SummaryReport = $null
$MTDReport = $null

Connect-O365EXOPS
$SummaryReport = Get-O365MessageTrace
Get-O365MessageTraceDetail($SummaryReport)