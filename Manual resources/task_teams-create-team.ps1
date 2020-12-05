$connected = $false
try {
	Import-Module MicrosoftTeams
	$pwd = ConvertTo-SecureString -string $TeamsAdminPWD -AsPlainText –Force
	$cred = New-Object System.Management.Automation.PSCredential $TeamsAdminUser, $pwd
	Connect-MicrosoftTeams -Credential $cred
    HID-Write-Status -Message "Connected to Microsoft Teams" -Event Information
    HID-Write-Summary -Message "Connected to Microsoft Teams" -Event Information
	$connected = $true
}
catch
{	
    HID-Write-Status -Message "Could not connect to Microsoft Teams. Error: $($_.Exception.Message)" -Event Error
    HID-Write-Summary -Message "Failed to connect to Microsoft Teams" -Event Failed
}

if ($connected)
{
	try {
		New-Team -displayName $displayName -MailNickName $MailNickName -Visibility $Visibility
		HID-Write-Status -Message "Created Team [$displayName] Mailnickname [$MailNickName] Visibility [$Visibility]" -Event Success
		HID-Write-Summary -Message "Successfully created Team [$displayName] Mailnickname [$MailNickName] Visibility [$Visibility]" -Event Success
	}
	catch
	{
		HID-Write-Status -Message "Could not create Team [$displayName]. Error: $($_.Exception.Message)" -Event Error
		HID-Write-Summary -Message "Failed to create Team [$displayName]" -Event Failed
	}
}