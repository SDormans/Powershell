<h1>Powershell scripts</h1>

<h2>Description</h2>
Multiple scripts that automate some tasks. 
Every script I will explain below:

The automatic email-notification about password expire
This will send a email if the password will expire in 5 days or less. 

<br />


<h2>Languages and Utilities Used</h2>

- <b>PowerShell</b> 

<h2>Environments Used </h2>

- <b>Windows 10</b> (21H2)

<h2>Program walk-through:</h2>
The automatic email-notification about password expire
Before you try the script make sure the Remote Server Administration Tools (RSAT) is installed
if not follow the next steps: 

1. Open Settings > Apps > Optional Features.
2. Scroll down and look for RSAT: Active Directory and LDAP Tools.
3. If not present, click Add a feature, search for RSAT: Active Directory and LDAP Tools, and install it.
4. Restart your computer if prompted.
5. open terminal and type 'Get-Module -ListAvailable' to see if ActiveDirectory is in the list. 
