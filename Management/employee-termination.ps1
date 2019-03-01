# Tell script to stop on error
$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['*:ErrorAction']='Stop'

# Get Credentials
$credential = Get-Credential -Message "* Please enter domain administrator credentials"
$techname = Read-Host -Prompt "Please enter your name"

# Set company variables
. ./company-variables.ps1

$alias = Read-Host -Prompt "* Enter the account of the user to be offboarded"
Write-Host ("* The following actions will be taken:")
Write-Host ("*     Remove the user from all Active Directory security groups")
Write-Host ("*     Reset the account password to a random 24-character string")
Write-Host ("*     Hide the mailbox from Exchange address lists")
Write-Host ("*     Remove the email from all internal distribution lists")
Write-Host ("*     Set an out-of-office auto-reply for the mailbox")
Write-Host ("*     Send confirmation email upon completion")
Read-Host -Prompt "* If this is okay to do for the user " + $alias + ", press enter to continue..."

### DOMAIN CONTROLLER ###
# Connect to Domain Controller
Write-Host ("* Attempting to Connect to Domain Controller: " + $DomainController)
$session = New-PSSession -ComputerName $DomainController -Credential $credential
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -DisableNameChecking -module ActiveDirectory
Write-Host ("* Connected to Domain Controller: " + $DomainController)

# Do Active Directory Stuff
$ADAccount = Get-ADUser -identity $alias
$FullName = $ADAccount.Name
$DistinguishedName = $ADAccount.DistinguishedName
$date = [datetime]::Today.ToString('MM-dd-yyyy')
$Description = ("Disabled on $date by $techname `r`n`r`n")

# Remove account from AD Groups
$Description = $Description + "Removed from the security groups:`r`n"
Get-ADUser $alias -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {
    Remove-ADGroupMember $_ -member $alias
    $Description = $Description + "$_`r`n"
}
Write-Host ("* " + $Fullname + "'s account has been removed from all security groups")

# Reset User Password
$NewPassword = [system.web.security.membership]::GeneratePassword(24, 0)
$NewPassword = ConvertTo-SecureString $NewPassword -AsPlainText -Force
Set-ADAccountPassword $alias -NewPassword $NewPassword
Write-Host ("* " + $Fullname + "'s Active Directory password has been changed.")

# Set account description
Set-ADUser $alias -Description $Description
Write-Host ("* " + $Fullname + "'s account description has been set.")

# Disable the account
Disable-ADAccount $ADAccount
Write-Host ("* " + $Fullname + "'s Active Directory account has been disabled.")

# Disconnect
Write-Host ("* Disconnecting from Domain Controller")
Get-PSSession | Remove-PSSession

### EXCHANGE ###
# Connect to Exchange Server
Write-Host ("* Attempting to Connect to Exchange Server: " + $ExchangeServer)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $credential
Import-PSSession -Session $session -DisableNameChecking
Write-Host ("* Connected to Exchange Server: " + $ExchangeServer)

$Mailbox = Get-Mailbox -identity $alias

# Hide email from exchange address lists
Set-Mailbox -identity $alias -HiddenFromAddressListsEnabled $true
Write-Host ("* " + $Fullname + "'s email account has been hidden from address lists.")

# Remove from all distribution groups
$AllDistGroups = Get-DistributionGroup | Where-Object {(Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySMTPAddress}) -contains "$alias.domain.tld"}
$AllDistGroups | ForEach-Object {
    $members = Get-DistributionGroupMember -Identity $_
    if ($alias -in $members) {
        Remove-DistributionGroupMember $_ -Member $Mailbox -Confirm:$False
    }
}
Write-Host ("* " + $Fullname + " has been removed from email distribution groups.")

# Set out of office auto-reply
$OOOReply = "$FullName is no longer with $CompanyName  Please note that this email address is not actively monitored.  You may direct any general inquiries to info@domain.tld and someone will respond as soon as possible."
Set-MailboxAutoReplyConfiguration -identity $alias -AutoReplyState enabled -InternalMessage $OOOReply -ExternalMessage $OOOReply

# Allow access to mailbox
# NOTE: The automapping parameter prevents the exchange autodiscover from mapping the account into the user's local outlook client!
Add-MailboxPermission -Identity $DistinguishedName -User (Get-ADUser $CIOAlias).DistinguishedName -AccessRights "FullAccess" -Automapping $false
Write-Host ("* Permission to access the mailbox has been granted to $CIOAlias.")

# Wait for user input before sending email
Read-Host -Prompt "Please review any errors, then press Enter to continue"

# Email Notification
# Reference: https://interworks.com/blog/trhymer/2013/07/08/powershell-how-encrypt-and-store-credentials-securely-use-automation-scripts/
$MailParams = @{

    SmtpServer 	= $SMTPServer
    From       	= "notifications@domain.tld"
    To         	= $CIOEmail
    cc         	= $AdminEmail
    Subject    	= "User Account for $FullName has been Offboarded"
    UseSsl     	= $true
    BodyAsHtml 	= $true
    Credential 	= $credential
    Body       	= "Hello,<br><br>

    This is an automated message to let you know that the terminated user, <b>$FullName</b>, has been offboarded from the company network.<br><br>

    The following actions were taken:<br>
    Display Name = $FullName<br><br>
    <u>Network</u><br>
    <ul>
        <li>The user's Active Directory account was disabled on $date.<br>
        <li>The user's Active Directory account's password was changed to a random 24-character string.</li>
        <li>The user's Active Directory security permissions were stripped and stored in the account description for future reference.</li>
    </ul>
    <u>Email</u><br>
    <ul>
        <li>An out-of-office reply message has been set directing inquiries to the generic information address.</li>
        <li>The mailbox has been removed from all internal distribution lists.</li>
        <li>The mailbox has been hidden from internal address lists.</li>
        <li>Full-Access permissions were set for $CIOAlias allowing access through the Outlook Web App.</li>
    </ul>
    <br>-- $techname<br><br>

    <H6>Automated user off-boarding script written for $CompanyName using Microsoft PowerShell.</H6><br>"

}
Send-MailMessage @MailParams
Write-Host ("* Email notification has been sent.")

# Disconnect
Write-Host ("* Disconnecting from Exchange Server")
Get-PSSession | Remove-PSSession