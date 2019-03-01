# Tell script to stop on error
$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['*:ErrorAction']='Stop'

# Get Credentials
$credential = Get-Credential -Message "* Please enter domain administrator credentials"

# Set company variables
. ./company-variables.ps1

# Get User Information
Write-Host ("* Getting user information")
$FirstName = Read-Host -Prompt "Please enter the new employee's first name"
$MiddleInitial = Read-Host -Prompt "Please enter the new employee's middle initial"
$LastName = Read-Host -Prompt "Please enter the new employee's last name"
$alias = ($FirstName[0] + $LastName).ToLower()
$ComputerModel = Read-Host -Prompt "Please enter the computer model number (Numbers only)"
$ComputerName = $FirstName[0] + $MiddleInitial + $LastName + "-" + $ComputerModel
# Ensure Computer and User account names are within Active Directory limits
if ($ComputerName.Length -gt 15) {
    $ComputerName = $FirstName[0] + $LastName + "-" + $ComputerModel  # Try to leave out the middle initial to fit 15 char limit
}
if ($ComputerName.Length -gt 15) {
    throw "Computer name is too long! The Active Directory limit is 15 characters."
}
if ($alias.Length -gt 20) {
    throw "User alias is too long! The Active Directory limit is 20 characters."
}

$Extension = Read-Host -Prompt "Please enter the new employee's office phone extension"
$PhoneNumber = $Extension

### DOMAIN CONTROLLER ###
# Connect to Domain Controller
Write-Host ("* Attempting to Connect to Domain Controller: $DomainController")
$session = New-PSSession -ComputerName $DomainController -Credential $credential
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -DisableNameChecking -module ActiveDirectory
Write-Host ("* Connected to Domain Controller: $DomainController")

# Create User Object
$UserProperties = @{
    Name = "$FirstName $LastName"
    SamAccountName = "$alias"
    UserPrincipalname = "$alias@$UPNSuffix"
    GivenName = $FirstName
    Initials = $MiddleInitial
    Surname = $LastName
    OfficePhone =$PhoneNumber
    Path = $ManagedUsers
    AccountPassword = (Read-Host -AsSecureString -Prompt "Enter the desired password")
    Enabled = $True
}
New-ADUser @UserProperties
Write-Host ("* Created Active Directory User Object: $alias")

# Create Computer Object
$ComputerProperties = @{
    Name = $ComputerName
    SamAccountName = $ComputerName
    Path = $ManagedComputers
    Enabled = $True
}
New-ADComputer @ComputerProperties
Write-Host ("* Created Active Directory Computer Object: $ComputerName")

# Add Computer and User to Security Groups
$User = Get-ADUser -Identity $alias
$Computer = (Get-ADComputer -Identity $ComputerName).DistinguishedName
$EmployeeComputers = "Employee Computers"
$Employees = "Employees"
Add-ADGroupMember -Identity $Employees -Members $alias
Write-Host ("* Added $alias to the 'Employees' security group")
Add-ADGroupMember -Identity $EmployeeComputers -Members $Computer
Write-Host ("* Added $ComputerName to the 'Employee Computers' security group")
$SecurityGroup = "Office 365 Users"
while ("exit" -ne $SecurityGroup) {
    if (Get-ADGroup -Identity $SecurityGroup) {
        Add-ADGroupMember -Identity $SecurityGroup -Members $alias
        Write-Host ("* Added $alias to the Security Group: $SecurityGroup")
    }
    else {
        Write-Host ("* Could not find the Security Group: $SecurityGroup")
    }
    $SecurityGroup = Read-Host -Prompt "Enter any other security groups that $alias should be in ['exit' to continue]"
}


# Allow Dial-In for VPN
Set-ADUser $alias -replace @{msNPAllowDialIn = $True}
Write-Host ("* Allowed dial-in network access")

# Disconnect
Write-Host ("* Disconnecting from Domain Controller")
Get-PSSession | Remove-PSSession


### EXCHANGE ###
# Connect to Exchange Server
Write-Host ("* Attempting to Connect to Exchange Server: " + $ExchangeServer)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $credential
Import-PSSession -Session $session -DisableNameChecking
Write-Host ("* Connected to Exchange Server: " + $ExchangeServer)

# Enable Mailbox
Enable-Mailbox -Identity $alias | Out-Null
$Mailbox = Get-Mailbox -Identity $alias
$Email = $Mailbox.PrimarySmtpAddress
Write-Host ("* Enabled Mailbox for $alias : $Email")

# Add to distribution groups
$DistributionGroup = "All Employees"
while ("exit" -ne $DistributionGroup) {
    if (Get-DistributionGroup -Identity $DistributionGroup) {
        Add-DistributionGroupmember -Identity $DistributionGroup -Member $Email
        Write-Host ("* Added $Email to the Distribution Group: $DistributionGroup")
    }
    else {
        Write-Host ("* Could not find the distribution group: $DistributionGroup")
    }
    $DistributionGroup = Read-Host -Prompt "Enter any other distribution groups that $alias should be in ['exit' to continue]"
}

# Disconnect
Write-Host ("* Disconnecting from Exchange Server")
Get-PSSession | Remove-PSSession


# Create User Filedrop Folder
$Path = Join-Path $FileShareBase "Filedrop"
New-Item -Path $Path -Name "$LastName Filedrop" -ItemType "directory"
Write-Host ("* Created Filedrop Folder")