# Define log path
$LogFile = "C:\Logs\emailnotificationlog.txt" 

# Redirect output streams to the log file
Start-Transcript -Path $LogFile -Append

# Ensure the log directory exists
$LogDir = Split-Path $LogFile
if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Import Active Directory Module
Import-Module ActiveDirectory

# Set the number of days before password expiry for notification
$daysToExpire = 5
$expireThreshold = (Get-Date).AddDays($daysToExpire)

# SMTP Server settings
$smtpServer = "smtp.yourdomain.com"  # Change this to your SMTP server
$smtpFrom = "admin@yourdomain.com"    # Change this to the sender's email
$smtpSubject = "Password Expiration Notice"

# Logging starts

write-host (Get-Date) "script started" -ForegroundColor Cyan

try {
    # Get all users in AD whose passwords will expire in the next 5 days
$users = Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Property DisplayName, EmailAddress, msDS-UserPasswordExpiryTimeComputed | 
Select-Object -Property DisplayName, EmailAddress, 
    @{Name = "PasswordExpiresSoon"; Expression = {
        [datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed') -lt $expireThreshold
    }}
    write-host (Get-Date) "users are searched" -ForegroundColor Cyan

}
catch {
    Write-Host (get-date) "An error occurred: $_" -ForegroundColor Red
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
                        
        Write-Host (Get-Date) "Password expiration notice sent to $($user.DisplayName) <$emailAddress>." -ForegroundColor Green
    } else {
        Write-Host (Get-Date )"No email address found for user $($user.DisplayName)." -ForegroundColor Black # Leave open?
    }
}
}
catch {
    Write-Host (get-date) "An error occurred: $_" -ForegroundColor Red
}
finally {
     # Completion message
     Write-Host (Get-Date)"Operation complete. Check the log at $LogFile for details." -ForegroundColor Cyan
     # Always close the transcript
     Stop-Transcript
}

