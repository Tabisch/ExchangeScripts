Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

$logFile = ".\Mailbox_Counters.csv"

$header = '"Mailbox","SMTPAdress","ReceiveCount","SendCount"'

$mailboxes = Get-Mailbox -Filter *

$trackingLog = Get-MessageTrackingLog -Start (Get-Date).AddDays(-90) -ResultSize unlimited

$header | Out-File $logFile -Encoding utf8

foreach($mailbox in $mailboxes)
{
    Write-Host "Group: $($mailbox.Name)"

    foreach($emailAddress in $mailbox.EmailAddresses)
    {
        $smtpAdress = $emailAddress.SmtpAddress

        if($emailAddress.Prefix.PrimaryPrefix -eq "SMTP")
        {
            Write-Host "Emailadress: $($smtpAdress)"
            $receiveMessageCount = ($trackingLog | Where-Object{ $_.Recipients -contains $smtpAdress } | Measure-Object).Count
            $sendMessageCount = ($trackingLog | Where-Object{ $_.Sender -eq $smtpAdress } | Measure-Object).Count

            "`"$($mailbox.Name)`",`"$($smtpAdress)`",`"$($receiveMessageCount)`",`"$($sendMessageCount)`""  | Out-File $logFile -Append -Encoding utf8
        }
        else
        {
            Write-Host "Not Smtp: $($emailAddress)"
        }
    }
}