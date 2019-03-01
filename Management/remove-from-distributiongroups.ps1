# Tell script to stop on error
$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['*:ErrorAction']='Stop'

# Get Credentials
$credential = Get-Credential -Message "* Please enter domain administrator credentials"

# Set company variables
. ./company-variables.ps1

### EXCHANGE ###
# Connect to Exchange Server
Write-Host ("* Attempting to Connect to Exchange Server: " + $ExchangeServer)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $credential
Import-PSSession -Session $session -DisableNameChecking
Write-Host ("* Connected to Exchange Server: " + $ExchangeServer)

$alias = Read-Host -Prompt "** Please enter alias of user to be removed from all distribution groups"
$user = Get-ADUser $alias
$name = $user.Name

# Get distribution groups which the user is a part of
$distgroups = @()
ForEach ($dg in Get-DistributionGroup) {
    ForEach ($member in Get-DistributionGroupMember -identity $dg.name | Where-Object {$_.Name -eq $name}) {
        $distgroups += $dg.Name
    }
}

# Require confirmation
Write-Host ("* $name is about to be removed from the following distribution groups:")
ForEach ($dg in $distgroups) {Write-Host ("*    $dg")}
$continue = Read-Host "** Is this okay? [y/n]"
if ($continue -eq "y") {
    ForEach ($dg in $distgroups) {
        Remove-DistributionGroupMember $dg -Member $alias -Confirm:$False
        Write-Host ("*    $dg : Removed")
    }
}
else {
    Write-Host ("* Aborting operation, no changes made.")
}

# Disconnect
Write-Host ("* Disconnecting from Exchange Server")
Get-PSSession | Remove-PSSession