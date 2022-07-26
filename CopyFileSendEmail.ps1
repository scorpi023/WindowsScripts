
param(
    $source = $null,
    $destination = $null,
	$email = $null
)

If ( ($PSBoundParameters.Values.Count -eq 0)  -or ($source -eq $null -or $destination -eq $null))
{ write-host "##############################################################"
  write-host "Usage: .\CopyFileSendEmail.ps1 <Source file/Directory> <Destination>"
  write-host "                                                 <Email/PDL to be notified>"
  write-host "##############################################################"
  exit 1
}

If(!(Test-Path -Path $source))
{
	write-host "Source location is not accessible by this machine or the file is not available"
	exit 1
}

$obj = New-Item -Path $destination -Name "testfile1.txt" -ItemType "file" -Value "This is a test file."
if ($obj -eq $null)
{
	write-host "Check if the machine has write access to destination"
	exit 1
}
Remove-Item $obj
Copy-Item $source -Destination $destination
write-host "Copy Completed"
write-host "File/Directory $source copied to $destination"

if ($email -ne $null)
{
$Message = new-object Net.Mail.MailMessage 
$smtp = new-object Net.Mail.SmtpClient("smtpauth.test.com", 587) 
$smtp.Credentials = New-Object System.Net.NetworkCredential("User_email@domain.com", "PasswordForEmailAccount"); 
$smtp.EnableSsl = $true 
$smtp.Timeout = 400000  
$Message.From = "sender_email@domain.com" 
$Message.To.Add($email) 
$Message.Subject = "Copy Completed"
$Message.Body = "Copy of $source to $destination was successful at $(Get-Date)"
$smtp.Send($Message)
write-host "Email receipt sent to $email"
}