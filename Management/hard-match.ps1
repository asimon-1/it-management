# Performs hard matching for all users within the "Office 365 Users" security group.
$credential = Get-Credential  -Message "Please enter credentials for an Azure Active Directory Global Admin"
Connect-MsolService -Credential $credential

. ./company-variables.ps1

# Only match the users which are being synced with O365
$users = Get-ADGroupMember -Identity "Office 365 Users"
ForEach ($user in $users) {
    $ADUser = $user.SamAccountName
    $365User = "$ADUser@$UPNSuffix"  # Note that UPN Suffixes must match!
    $guid =(Get-ADUser $ADUser).Objectguid
    $immutableID=[system.convert]::ToBase64String($guid.tobytearray())
    try {
        $MsolUser = Get-MsolUser -UserPrincipalName "$365User"
    }
    catch {
        Write-Host ("* Could not find the user $ADUser in Azure Active Directory! Check the recent sync results.")
        continue
    }
    if ($MsolUser.ImmutableId -eq $immutableID) {
        Write-Host ("* ID already matches for user: $ADUser. Skipping...")
    }
    else {
        Write-Host ("* ID does not match for user: $ADUser. Changing AAD to match on-prem AD.")
        Set-MsolUser -UserPrincipalName "$365User" -ImmutableId $immutableID
    }
}