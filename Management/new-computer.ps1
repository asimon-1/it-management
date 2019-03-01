# Adds a new computer to the domain. Optionally adds the computer to the Employee Computers security group

# Tell script to stop on error
$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['*:ErrorAction']='Stop'

# Get Credentials
$credential = Get-Credential -Message "* Please enter domain administrator credentials"

# Set company variables
. ./company-variables.ps1

### DOMAIN CONTROLLER ###
# Connect to Domain Controller
Write-Host ("* Attempting to Connect to Domain Controller: " + $DomainController)
$session = New-PSSession -ComputerName $DomainController -Credential $credential
Invoke-Command -Session $session -ScriptBlock { Import-Module ActiveDirectory }
Import-PSSession -Session $session -DisableNameChecking -module ActiveDirectory -AllowClobber
Write-Host ("* Connected to Domain Controller: " + $DomainController)

# Do Active Directory Stuff
$ComputerName = Read-Host -Prompt "Please enter the name of the new computer"

$Path = $ManagedComputers
New-ADComputer -Name $ComputerName -sAMAccountName $ComputerName -Path $Path -Enabled $True
Write-Host ("* Added New Computer Object: $ComputerName")

$EmployeeComputerFlag = Read-Host -Prompt "Is this an Employee Computer? [y/n]"
while ("y","n" -notcontains $EmployeeComputerFlag) {
    $EmployeeComputerFlag = Read-Host -Prompt "Is this an Employee Computer? [y/n]"
}
if ("y" -contains $EmployeeComputerFlag) {
    $EmployeeComputers = "Employee Computers"
    $ComputerDN = Get-ADComputer $ComputerName | Select-Object -Expand Distinguishedname
    Add-ADGroupMember -Identity $EmployeeComputers -Members $ComputerDN
    Write-Host ("* Added $ComputerName to the group $EmployeeComputers")
}

# Disconnect
Write-Host ("* Disconnecting from Domain Controller")
Get-PSSession | Remove-PSSession
