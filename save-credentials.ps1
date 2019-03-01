# Save encrypted credentials to file. Uses Windows profile information as encryption key!
# Note that this means the resulting file may only be used by the same user on the same computer as it was generated.
# Discards the username field, only keeps the password.
# From https://interworks.com/blog/trhymer/2013/07/08/powershell-how-encrypt-and-store-credentials-securely-use-automation-scripts/
$credential = Get-Credential
$credential.Password | ConvertFrom-SecureString | Set-Content encrypted_password.txt