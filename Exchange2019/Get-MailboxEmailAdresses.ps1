$path = ".\mailboxes_smtpadresses.csv"

Remove-Item -Path $path -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

#Head
$output = "`"DisplayName`",`"PrimarySmtpAddress`",`"SmtpAddress`""
$output | Out-File -FilePath $path -Encoding utf8 -Append

$mailboxes = Get-Mailbox 

$mailboxes | ForEach-Object{

    Write-Host $_.DisplayName

    $addresses = $_.EmailAddresses | Where-Object{ $_.Prefix.PrimaryPrefix -contains "SMTP" } | Where-Object{ $_.SmtpAddress -notmatch ".onmicrosoft.com" } | Where-Object{ $_.SmtpAddress -notmatch "DiscoverySearchMailbox" }

    foreach($address in $addresses)
    {
        $output = "`"$($_.DisplayName)`",`"$($_.PrimarySmtpAddress)`",`"$($address.SmtpAddress)`""
        #Write-Host $output
        $output | Out-File -FilePath $path -Encoding utf8 -Append
    }
}