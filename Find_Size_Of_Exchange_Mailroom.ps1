
Connect-ExchangeOnline

Get-Mailbox -ResultSize 10 | ForEach-Object {Get-MailboxStatistics -Identity $_.UserPrincipalName | Select Identity, DisplayName,TotalItemSize,ItemCount, proxyAddresses} | Format-Table -AutoSize

Get-Mailbox | ForEach-Object {Get-MailboxStatistics -Identity $_.UserPrincipalName | Select DisplayName,TotalItemSize,ItemCount} | Export-Csv -Path "C:\MailboxStatistics.csv" -NoTypeInformation -Encoding UTF8

Get-Mailbox -ResultSize 10 | ForEach-Object {
    $mailbox = $_
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName
    [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimarySMTPAddress = $mailbox.PrimarySmtpAddress
        TotalSize = $stats.TotalItemSize
        ItemCount = $stats.ItemCount
        ProxyAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like "SMTP:*" }) -join "; "
    }
} | Format-Table -AutoSize


Get-Mailbox | ForEach-Object {
    $mailbox = $_
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName
    [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimarySMTPAddress = $mailbox.PrimarySmtpAddress
        TotalSize = $stats.TotalItemSize
        ItemCount = $stats.ItemCount
        ProxyAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like "SMTP:*" }) -join "; "
    }
} | Export-Csv -Path "C:\MailboxStatistics_WithProxyAddresses.csv" -NoTypeInformation -Encoding UTF8


