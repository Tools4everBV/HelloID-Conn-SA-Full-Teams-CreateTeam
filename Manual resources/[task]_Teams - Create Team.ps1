# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$displayName = $form.DisplayName
$mailNickName = $form.Mailnickname
$Visibility = $form.Visibility

$connected = $false
try {
	$module = Import-Module MicrosoftTeams

	$pwd = ConvertTo-SecureString -string $TeamsAdminPWD -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential $TeamsAdminUser, $pwd
	$connectTeams = Connect-MicrosoftTeams -Credential $cred
    Write-Information "Connected to Microsoft Teams"
	$connected = $true
}
catch
{	
    Write-Error "Could not connect to Microsoft Teams. Error: $($_.Exception.Message)"
}

if ($connected)
{
	try {
		$team = New-Team -displayName $displayName -MailNickName $MailNickName -Visibility $Visibility
		Write-Information "Created Team [$displayName] Mailnickname [$MailNickName] Visibility [$Visibility]"
	}
	catch
	{
		Write-Error "Could not create Team [$displayName]. Error: $($_.Exception.Message)"
	}
}
