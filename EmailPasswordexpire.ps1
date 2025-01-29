# TODO: log it

# Import Active Directory Module
Import-Module ActiveDirectory

# Set the number of days before password expiry for notification
$daysToExpire = 5
$expireThreshold = (Get-Date).AddDays($daysToExpire)

# SMTP Server settings
$smtpServer = "smtp.yourdomain.com"  # Change this to your SMTP server
$smtpFrom = "admin@yourdomain.com"    # Change this to the sender's email
$smtpSubject = "Password Expiration Notice"

try {
    # Get all users in AD whose passwords will expire in the next 5 days
$users = Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Property DisplayName, EmailAddress, msDS-UserPasswordExpiryTimeComputed | 
Select-Object -Property DisplayName, EmailAddress, 
    @{Name = "PasswordExpiresSoon"; Expression = {
        [datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed') -lt $expireThreshold
    }}

}
catch {
    <#Do this if a terminating exception happens#>
}

try {
    # Loop through each user and send email
foreach ($user in $users) {
    $emailAddress = $user.EmailAddress

    # Check if the user has an email address
    if ($emailAddress) {
        # Calculate days until password expires
        $expiryDate = [datetime]::FromFileTime($user.'msDS-UserPasswordExpiryTimeComputed')
        $daysLeft = ($expiryDate - (Get-Date)).Days

        # Email body content
        $smtpBody = @"
                        Beste $($user.DisplayName),

                        Hierbij is een herinnering dat u uw wachtwoord moet veranderen binnen $($daysleft) dagen. Als u hierbij hulp nodig hebt, kunt u contact opnemen met de helpdesk.

                        Met vriendelijke groeten, 
                        De ICT-team
"@

        # Send email using SMTP
        Send-MailMessage -SmtpServer $smtpServer 
                        -From $smtpFrom 
                        -To $emailAddress 
                        -Subject $smtpSubject 
                        -Body $smtpBody 
                        -Priority High
                        
        Write-Host "Password expiration notice sent to $($user.DisplayName) <$emailAddress>."
    } else {
        Write-Host "No email address found for user $($user.DisplayName)."
    }
}
}
catch {
    <#Do this if a terminating exception happens#>
}


