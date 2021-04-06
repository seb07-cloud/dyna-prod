Start-Transcript -Path .\Remove-SMTP-Address.log -Append
 
# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited
 
# Loop through each mailbox
foreach ($Mailbox in $Mailboxes) {
 
    # Change @contoso.com to the domain that you want to remove
    $Mailbox.EmailAddresses | Where-Object { $_.AddressString -like "*@domain.at" } | ForEach-Object {
 
        # Remove the -WhatIf parameter after you tested and are sure to remove the secondary email addresses
        Set-Mailbox $Mailbox -EmailAddresses @{remove = $_ } -WhatIf
 
        # Write output
        Write-Host "Removing $_ from $Mailbox Mailbox" -ForegroundColor Green
    }
}
 
Stop-Transcript