# Set company variables
. ./company-variables.ps1

# Get variables
$credential = Get-Credential -Message "* Please enter domain administrator credentials"

$date = [datetime]::Today.ToString('MM-dd-yyyy')
$ToBeDisabled = @()
$ToBeDisabledInput = Read-Host -Prompt "* Please enter the name of an account to be disabled"
while ("exit" -ne $ToBeDisabledInput) {
    if (Get-ADUser -Identity $ToBeDisabledInput) {
        $ToBeDisabled += (Get-ADUser -Identity $ToBeDisabledInput).DistinguishedName
        Write-Host ("* Added $ToBeDisabledInput to the list of accounts to be disabled")
    }
    else {
        Write-Host ("* Could not find the user account: $ToBeDisabledInput")
    }
    $ToBeDisabledInput = Read-Host -Prompt "* Please enter the name of an account to be disabled ['exit' to continue]"
}

$Custodian = (Get-ADUser -Identity (Read-Host -Prompt "* Please enter the alias of the email custodian")).DistinguishedName
if ((Read-Host -Prompt "*Set new password? [y/n]") -eq "y") {
    $NewPasswordFlag = $true
    if ((Read-Host -Prompt "* Generate random password? [y/n]") -eq "y") {
        $RandomPasswordFlag = $true
    }
    else {
        $RandomPasswordFlag = $false
        $NewPassword = Read-Host -AsSecureString -Prompt "* Please enter a new password for the user accounts"
    }
}
else {
    $NewPasswordFlag = $false
}


### DOMAIN CONTROLLER ###
# Connect to Domain Controller
Write-Host ("* Attempting to Connect to Domain Controller: " + $DomainController)
$session = New-PSSession -ComputerName $DomainController -Credential $credential
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -DisableNameChecking -module ActiveDirectory -AllowClobber
Write-Host ("* Connected to Domain Controller: " + $DomainController)

# Do Active Directory Stuff
foreach ($DN in $ToBeDisabled) {
    $FullName = (Get-ADUser -identity $DN).Name
    $Description = ("Disabled on $date`r`n`r`n")

    # Remove account from AD Groups
    $Description = $Description + "Removed from the security groups:`r`n"
    Get-ADUser $DN -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object {
        Remove-ADGroupMember $_ -member $DN
        $Description = $Description + "$_`r`n"
    }
    Write-Host ("* " + $Fullname + "has been removed from Active Directory groups.")

    # Reset User Password
    if ($NewPasswordFlag) {
        if ($RandomPasswordFlag) {
            $NewPassword = [system.web.security.membership]::GeneratePassword(24, 0)
            $NewPassword = ConvertTo-SecureString $NewPassword -AsPlainText -Force
        }
    Set-ADAccountPassword $DN -NewPassword $NewPassword
    Write-Host ("* " + $Fullname + "'s Active Directory password has been changed.")
    }

    # Set account description
    Set-ADUser $DN -Description $Description

    # Disable the account
    Disable-ADAccount -Identity $DN
    Write-Host ("* " + $Fullname + "'s Active Directory account has been disabled.")

}

# Disconnect
Write-Host ("* Disconnecting from Domain Controller")
Get-PSSession | Remove-PSSession

### EXCHANGE ###
# Connect to Exchange Server
Write-Host ("* Attempting to Connect to Exchange Server: " + $ExchangeServer)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $credential
Import-PSSession -Session $session -DisableNameChecking -AllowClobber
Write-Host ("* Connected to Exchange Server: " + $ExchangeServer)

# Do Exchange Stuff
foreach ($DN in $ToBeDisabled) {
    $FullName = (Get-ADUser -identity $DN).Name

    # Allow access to mailbox
    # NOTE: The automapping parameter prevents the exchange autodiscover from mapping the account into the user's local outlook client!
    Add-MailboxPermission -Identity $DN -User $Custodian -AccessRights "FullAccess" -Automapping $false | Out-Null
    Write-Host ("* Permission to access $FullName's mailbox has been granted to $Custodian.")
}

# Disconnect
Write-Host ("* Disconnecting from Exchange Server")
Get-PSSession | Remove-PSSession