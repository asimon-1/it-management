# Get Credentials
$credential = Get-Credential -Message "* Please enter domain administrator credentials"

# Set company variables
. ./Management/company-variables.ps1

### EXCHANGE ###
# Connect to Exchange Server
Write-Host ("* Attempting to Connect to Exchange Server: " + $ExchangeServer)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $credential
Import-PSSession -Session $session -DisableNameChecking
Write-Host ("* Connected to Exchange Server: " + $ExchangeServer)

# Do Exchange Stuff
$alias = "testuser"
Get-Mailbox -identity $alias

# Disconnect
Write-Host ("* Disconnecting from Exchange Server")
Get-PSSession | Remove-PSSession

### DOMAIN CONTROLLER ###
# Connect to Domain Controller
Write-Host ("* Attempting to Connect to Domain Controller: " + $DomainController)
$session = New-PSSession -ComputerName $DomainController -Credential $credential
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -DisableNameChecking -module ActiveDirectory
Write-Host ("* Connected to Domain Controller: " + $DomainController)

# Do Active Directory Stuff
Get-ADUser $alias

# Disconnect
Write-Host ("* Disconnecting from Domain Controller")
Get-PSSession | Remove-PSSession
