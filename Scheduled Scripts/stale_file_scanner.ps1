# Run this script as a scheduled task every hour. It will scan $ScanPath and generate a list of filenames.
# If any files have been in the directory for longer than $TimeThreshold, it will notify you.

# results.csv Format
# 1     YYYY-mm-dd HHMM # Datetime of last notification
# 2     YYYY-mm-dd HHMM,file_a,file_b,file_c # Oldest directory snapshot
# ...   ...
# n     YYYY-mm-dd HHMM,file_c,file_d,file_e,file_f # Newest directory snapshot

# Globals
. ../Management/company-variables.ps1
$ScanPath = "$FileShareBase\File Inbox"
$OutputPath = "$FileShareBase\Stale File Scanner\results.csv"
$EmailThreshold = 12 # Hours
$TimeThreshold = 72 # Hours

# Email Credentials https://interworks.com/blog/trhymer/2013/07/08/powershell-how-encrypt-and-store-credentials-securely-use-automation-scripts/
$SendEmail = $true
$From = $NotificationsMail
$To = $AdminEmail
$Subject = "Stale Files in $ScanPath"
$BodyList = ""
$emailusername = "Administrator"
$encrypted = Get-Content "$FileShareBase\Stale File Scanner\notifications_encrypted_password.txt" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PsCredential($emailusername, $encrypted)

function Send-Output
{
    param([bool]$MailFlag, [string]$Output)
    # Optionally send an email or write to console depending on $MailFlag
    if ($MailFlag)
    {
            Send-MailMessage -From $From -To $ToAddress -Subject $Subject -Body $Output -SmtpServer $SMTPServer -UseSsl -Credential $credential -BodyAsHtml
    }
    else
    {
        Write-Output $Output
    }
}

function Format-RotatingFile
{
    param([string]$OutputPath, [datetime]$LastEmail, [int]$Skip)
    # Rotates the file in $OutputPath by $Skip lines, and setting the first line in the file to $LastEmail
    $Content = (Get-Content $OutputPath | Select-Object -Skip $Skip)
    $Content[0] = $LastEmail.ToString("yyyy-MM-dd HHmm")
    $Content | Set-Content $OutputPath
}

$CurrentTime = Get-Date
$CurrentTimeString = Get-Date -uformat "%Y-%m-%d %H%M"

# Check if output file exists, and create it if it doesnt
if (-Not (Test-Path $OutputPath -PathType Leaf))
{
    $CurrentTimeString | Set-Content $OutputPath
}
else
{
    # Read current files and store at end of file
    $CurrentFiles = Get-ChildItem -Path $ScanPath -Name -Recurse -Filter '*.*' -ErrorAction 'SilentlyContinue'
    $NewLine =,$CurrentTimeString + $CurrentFiles -join ","
    Add-Content -Path $OutputPath $NewLine

    # Read the file
    $LastEmail = [datetime]::ParseExact((Get-Content $OutputPath -Head 1),"yyyy-MM-dd HHmm", $null) # The last time an email notification was set out.
    $Head = (Get-Content $OutputPath -Head 2 | Select-Object -skip 1).split(",") # Oldest Entry

    # Check if any of files are in previous entry
    $OldTime = [datetime]::ParseExact($Head[0],"yyyy-MM-dd HHmm", $null)
    $OldFiles = $Head[1..$Head.Length]
    $match = $CurrentFiles | Where-Object { $_ -in $OldFiles } # Get any current files that are also in the oldest entry

    if ($match.Length -gt 0 -and $OldTime.AddHours($TimeThreshold) -lt $CurrentTime)
    {
        # There are files that have been in the folder longer than is allowed!
        ForEach ($file in $match) {
            $BodyList = $BodyList + "<li>$file</li>"
        }
        if ($LastEmail.AddHours($EmailThreshold) -lt $CurrentTime)
        {
            $LastEmail = $CurrentTime
            $Body = "<p>This is an automated email to alert you that the following files have been in the '$ScanPath' folder longer than $TimeThreshold hours:</p><ul>$BodyList</ul>"
            Send-Output $SendEmail $Body
            Format-RotatingFile $OutputPath $LastEmail 0
        }
    }
    else
    {
        # No files have been in the folder beyond the threshold.
        Write-Output "No stale files exist in $ScanPath"
    }

    # Delete oldest entries if there are too many.
    while ((Get-Content $OutputPath | Measure-Object).count -gt $TimeThreshold + 1)
    {
        Format-RotatingFile $OutputPath $LastEmail 1
    }
}
