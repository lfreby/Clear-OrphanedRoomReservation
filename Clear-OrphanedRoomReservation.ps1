<#
	.SYNOPSIS
		Removes meeting from conference rooms where the organizer is a deleted mailbox.

	.DESCRIPTION
		The script will lookup the list of disconnected mailboxes on Exchange, then build a query for search-mailbox to find and remove all meetings from organizers
        matching the disconnected mailbox name.

	.PARAMETER SearchDays
        Filter the input list to mailboxes that were disconnected in the last <SearchDays>

	.EXAMPLE
		.\Clear-OrphanedRoomReservation -SearchDays 2

	.NOTES

	.LINK
		https://technet.microsoft.com/en-us/library/dd298173(v=exchg.141).aspx

#>
Param
(
    [parameter(Mandatory = $false)]
    [int]
    $SearchDays

)

#Requires -Version 3.0

#region begin VARIABLES
# Exchange 2010 FQDN. The script will create a remote PowerShell session on this server
$serverFQDN = 'Exchange2010.fqdn'

# Set variables for sending e-mail
$sendMailMessageParams = @{
    To         = 'recipient@domain' # Put the recipients here (you can use an coma separated array)
    From       = 'sender@domain' # Put the sender address here (optimally you use an authenticated service account with mailbox)
    SmtpServer = 'Exchange2010.fqdn' # Put the smtp server FQDN here. It can be an Exchange server, a shared connector, or any opened SMTP relay in your org.
    Subject    = 'Orphaned Meetings Cleanup'
    BodyAsHtml = $true
    Priority   = 'normal'
    UseSsl     = $true
    Port       = 25
}

#endregion
#region begin INIT

$Global:ErrorActionPreference = 'Stop'
$TimeStamp = get-date -Format yyyy.MM.dd-HH.mm.ss
$scriptName = ($MyInvocation.MyCommand.Name) -replace ".ps1", "_LOG"
$currentPath = Split-Path $MyInvocation.MyCommand.Path

$logPath = Join-Path -Path $currentPath -ChildPath $scriptName
$transcriptFilePath = Join-Path -Path $logPath -ChildPath "_$($TimeStamp)_Transcript.log"
$reportFilePath = Join-Path -Path $logPath -ChildPath "_$($TimeStamp)_Report.csv"

$lastrunFilePath = Join-Path -Path $logPath -ChildPath "_timeStamp.xml"

$jobResults = @()

# Stop existing transcript, if already running
Try
{
    Stop-Transcript
}
Catch
{
    Write-Verbose "No existing transcript to stop"
}

# Create Log Subfolder
New-Item $logPath -ItemType Directory -Force

# Start new transcript logging
Try
{
    Start-Transcript -Path $transcriptFilePath
}
Catch
{
    Write-Warning "Could not start a new transcript. Error: $PSItem"
}

#endregion
#region begin FUNCTIONS

Function Write-LogResult
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Activity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Item,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Result,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ErrMessage
    )

    $currentTime = Get-Date -Format yyyy.MM.dd-HH.mm.ss
    $psObjectLogAction = [pscustomobject]@{
        StartTime  = $currentTime
        Activity   = $Activity
        Item       = $Item
        Result     = $Result
        ErrMessage = $ErrMessage
    }
    $psObjectLogAction | Export-Csv -Path $reportFilePath -Append -NoTypeInformation
}

function Send-Report
{
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Subject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TargetServer,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $JobStatus,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [Array]
        $Facts,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ThemeColor,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [Array]
        $EmailDetails
    )

    # Define htmlReport
    $htmlReport = @()
    $htmlReport += "<head>"
    $htmlReport += "<style>"
    $htmlReport += "body {text-align: left;font-size: 12px;font-family:'Segoe UI Web Regular','Segoe UI','Segoe WP',Tahoma,Arial,sans-serif;color:#696969;}"
    $htmlReport += "h1 {font-family: 'Segoe UI Web Light','Segoe UI Light','Segoe WP Light','Segoe UI','Segoe WP',Tahoma,Arial,sans-serif;font-size: 21px;color:black;}"
    $htmlReport += "h2 {font-family: 'Segoe UI Web Light','Segoe UI Light','Segoe WP Light','Segoe UI','Segoe WP',Tahoma,Arial,sans-serif;font-size: 17px;color:black;}"
    $htmlReport += "h3 {font-family: 'Segoe UI Web Regular','Segoe UI','Segoe WP',Tahoma,Arial,sans-serif;font-size: 14px;color:black;font-weight:normal;}"
    $htmlReport += "table {table-layout: fixed;border-collapse: collapse;border-spacing: 0;}"
    $htmlReport += "table tr {border: solid;border-width: 1px 0;border-color:black;}"
    $htmlReport += "table td {border-bottom:solid black 1.0pt;padding:10px;}"
    $htmlReport += "table th {border: none;padding:10px;text-align:left;height:30px;color:white;background:#2376BC;}"
    $htmlReport += "table.notification tr {border:none;padding:1px;}"
    $htmlReport += "table.notification td {border:none;padding:1px;}"
    $htmlReport += "ul {list-style-type:square}"
    $htmlReport += "</style>"
    $htmlReport += "</head>"

    # Define Message Body
    $htmlReport += "<body>"
    $htmlReport += "<table class='notification'>"
    $htmlReport += "<tr>"
    $htmlReport += "<td style=width:5px;background:" + $notification.themeColor + "></td>"
    $htmlReport += "<td>"
    $htmlReport += "<table class='notification'>"
    $htmlReport += "<tr>"
    $htmlReport += "<td><h1>" + $notification.title + "</h1></td>"
    $htmlReport += "</tr>"
    $htmlReport += "<tr>"
    $htmlReport += "<td>" + $notification.text + "</td>"
    $htmlReport += "</tr>"
    $htmlReport += "</table>"
    $htmlReport += "</td>"
    $htmlReport += "</tr>"
    Foreach ($section in $notification.sections)
    {
        $htmlReport += "<tr>"
        $htmlReport += "<td style=width:5px></td>"
        $htmlReport += "<td>"
        $htmlReport += "<table class='notification'>"
        $htmlReport += "<tr>"
        $htmlReport += "<td></td>"
        $htmlReport += "<td>"
        $htmlReport += "<table class='notification'>"
        $htmlReport += "<tr>"
        $htmlReport += "<td><h2>" + $section.activityTitle + "</h2></td>"
        $htmlReport += "</tr>"
        $htmlReport += "<tr>"
        $htmlReport += "<td><h3>" + $section.activitySubTitle + "</h3></td>"
        $htmlReport += "</tr>"
        $htmlReport += "</table>"
        $htmlReport += "</td>"
        $htmlReport += "</tr>"
        $htmlReport += "</table>"
        $htmlReport += "</td>"
        $htmlReport += "</tr>"
        $htmlReport += "<tr>"
        $htmlReport += "<td style=width:5px></td>"
        $htmlReport += "<td>"
        $htmlReport += "<table class='notification'>"
        foreach ($fact in $section.facts)
        {
            $htmlReport += "<tr>"
            $htmlReport += "<td style=color:black;width:140px>" + $fact.name + "</td>"
            $htmlReport += "<td>" + $fact.value + "</td>"
            $htmlReport += "</tr>"
        }

        $htmlReport += "</table>"
        $htmlReport += "</td>"
        $htmlReport += "</tr>"
    }
    $htmlReport += "</table><br />"
    $htmlReport += $EmailDetails
    $htmlReport += "</p>"
    $htmlReport += "</body>"

    # Send email notification
    $sendMailMessageParams.Body = $htmlReport | Out-String
    Send-MailMessage @sendMailMessageParams
}


#endregion
#region begin PROCESS

#connect to to Exchange Remote Shell
try
{
    Write-Verbose "Connecting to $serverFQDN via Remote Powershell"
    $exSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$serverFQDN/PowerShell/ -Authentication Kerberos
    Import-PSSession $exSession -DisableNameChecking -AllowClobber | out-null
}
catch
{
    Write-LogResult -Activity "Connecting to Exchange via Remote Powershell" -Item $serverFQDN -Result "Failure" -ErrMessage $PSItem
    Write-Error $PSItem
}

if ($SearchDays)
{
    $searchDate = (Get-Date).AddDays(-$SearchDays)
}
else
{
    $searchDate = Import-Clixml $lastrunFilePath
}

# Retrieve list of all mailbox DBs
$mailboxDatabaseList = Get-MailboxDatabase

# Retrieve list of disconnected mailboxes
$disconnectedMailboxList = @()
foreach ($mailboxDatabase in $mailboxDatabaseList)
{
    Write-Verbose "Parsing Database $mailboxDatabase for disconnected mailboxes"
    $disconnectedMailboxList += Get-MailboxStatistics -Database $mailboxDatabase.Name | `
        Where-Object {$_.DisconnectReason -eq "Disabled" -and $_.DisconnectDate -ge $searchDate} | `
        Select-Object -ExpandProperty DisplayName
}

# Retrieve list of recently disabled user mailboxes
$sessionDomainController = Get-ADDomainController | Select-Object -ExpandProperty HostName
$disabledMailboxList = Get-Mailbox -Filter "ExchangeUserAccountControl -eq 'AccountDisabled' -and RecipientTypeDetails -eq 'UserMailbox' -and WhenChanged -ge '$searchDate'"
foreach ($disabledMailbox in $disabledMailboxList)
{
    $userMetaData = Get-ADReplicationAttributeMetadata $disabledMailbox.Distinguishedname -Server $sessionDomainController -Properties userAccountControl
    if ($userMetaData.LastOriginatingChangeTime -ge $searchDate)
    {
        $disconnectedMailboxList += $disabledMailbox.DisplayName
    }
}

# Build search query for Seach-Mailbox
$searchQuery = "kind:meetings and ("
foreach ($disconnectedMailbox in $disconnectedMailboxList)
{
    $searchQuery += "From:`"$disconnectedMailbox`" or "
}
$searchQuery = $searchQuery -Replace " or $", ")"
Write-Verbose "Search query: $searchQuery"

# Run search query and delete against each room mailbox
Try
{
    Write-Verbose "Running orphaned meeting search and destroy against room mailboxes"
    $SearchResultList = @(Get-mailbox -Filter {recipientTypeDetails -eq "roomMailbox"} | Search-Mailbox -SearchQuery $searchQuery -DeleteContent -Force)
}
Catch
{
    Write-Warning "Unable to run Search-Mailbox. Error: $PSItem"
}

$actionResultList = $SearchResultList | Where-Object {$_.ResultItemsCount -ne 0}
Foreach ($actionResult in $actionResultList)
{
    $action = "Removing $ResultItemsCount meeting(s)"
    Write-Action $action $actionResult.DisplayName
    if ($actionResult.Success)
    {
        Write-LogResult -Item $actionResult.Identity -Activity "Removing meeting from retired organizers" -Result 'Success'
    }
    else
    {
        Write-LogResult -Item $actionResult.Identity -Activity "Removing meeting from retired organizers" -Result 'Error'
    }
}
# TimeStamp run
$date.AddMinutes(-5) | Export-Clixml -Path $lastrunFilePath

# Send Email
Send-EmailReport

# Remove PSEssion
Remove-PSSession $exSession | Write-Verbose
#endregion
