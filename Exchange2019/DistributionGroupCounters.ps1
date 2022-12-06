Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

$logFile = ".\DistributionGroup_Counters.csv"

$header = '"DistributionGroup","SMTPAdress","ReceiveCount","SendCount"'

$distributionGroups = Get-DistributionGroup

$trackingLog = Get-MessageTrackingLog -Start (Get-Date).AddDays(-90) -ResultSize unlimited

$header | Out-File $logFile -Encoding utf8

foreach($distributionGroup in $distributionGroups)
{
    Write-Host "Group: $($distributionGroup.Name)"

    foreach($emailAddress in $distributionGroup.EmailAddresses)
    {
        $smtpAdress = $emailAddress.SmtpAddress

        if($smtpAdress)
        {
            Write-Host "Emailadress: $($smtpAdress)"
            $receiveMessageCount = ($trackingLog | Where-Object{ $_.Recipients -contains $smtpAdress } | Measure-Object).Count
            $sendMessageCount = ($trackingLog | Where-Object{ $_.Sender -eq $smtpAdress } | Measure-Object).Count

            "`"$($distributionGroup.Name)`",`"$($smtpAdress)`",`"$($receiveMessageCount)`",`"$($sendMessageCount)`""  | Out-File $logFile -Append -Encoding utf8
        }
        else
        {
            Write-Host "Not Smtp: $($emailAddress)"
        }
    }
}