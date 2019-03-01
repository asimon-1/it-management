$credential = Get-Credential -Message "* Please enter domain admin credentials"
$DateCutoff = (Get-Date).AddDays(-120)
# Get-ADComputer -Properties LastLogonDate -Filter {LastLogonDate -lt $DateCutoff} | Sort LastLogonDate | Format-Table Name,LastLogonDate -Autosize
Get-ADComputer -Properties LastLogonDate -Filter {LastLogonDate -lt $DateCutoff} | Set-ADComputer -Enabled $false -Credential $credential