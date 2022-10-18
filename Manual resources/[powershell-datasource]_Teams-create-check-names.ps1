# AzureAD Application Parameters #
$Mailsuffix = $AzureMailSuffix
$Name = $datasource.displayName
$Description = $datasource.description

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#region Supporting Functions
function Get-ADSanitizeGroupName
{
    param(
        [parameter(Mandatory = $true)][String]$Name
    )
    $newName = $name.trim();
    $newName = $newName -replace ' - ','_'
    $newName = $newName -replace '[`,~,!,#,$,%,^,&,*,(,),+,=,<,>,?,/,'',",;,:,\,|,},{,.]',''
    $newName = $newName -replace '\[','';
    $newName = $newName -replace ']','';
    $newName = $newName -replace ' ','_';
    $newName = $newName -replace '\.\.\.\.\.','.';
    $newName = $newName -replace '\.\.\.\.','.';
    $newName = $newName -replace '\.\.\.','.';
    $newName = $newName -replace '\.\.','.';
    return $newName;
}
#endregion Supporting Functions

try {
    $iterationMax = 10
    $iterationStart = 1;

    for($i = $iterationStart; $i -lt $iterationMax; $i++) {
        $tempName = Get-ADSanitizeGroupName -Name $Name
        
        if($i -eq $iterationStart) {
            $tempName = $tempName
        } else {
            $tempName = $tempName + "$i"
        }

        #Shorten Name to max. 20 characters
        #$Name = $Name.substring(0, [System.Math]::Min(20, $Name.Length)) 
        
        #$DisplayName    = $tempName
        $DisplayName    = $Name
        #Shorten DisplayName to max. 20 characters
        #$DisplayName = $DisplayName.substring(0, [System.Math]::Min(20, $DisplayName.Length)) 
        $Description    = $Description
        $Mail           = $tempName.Replace(" ","") + "@" + $Mailsuffix 
        $MailNickname   = $tempName.Replace(" ","")

        Write-Information "Generating Microsoft Graph API Access Token.."

        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$AADTenantID/oauth2/token"

        $body = @{
            grant_type      = "client_credentials"
            client_id       = "$AADAppId"
            client_secret   = "$AADAppSecret"
            resource        = "https://graph.microsoft.com"
        }
    
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $accessToken = $Response.access_token;

        Write-Information "Searching for AzureAD group.."

        #Add the authorization header to the request
        $authorization = @{
            Authorization       = "Bearer $accesstoken";
            'Content-Type'      = "application/json";
            Accept              = "application/json";
            ConsistencyLevel    = "eventual";
        }

        Write-Verbose -Verbose "Searching for Group displayName=$DisplayName or mail=$Mail or mailNickname=$MailNickname"
        $baseSearchUri = "https://graph.microsoft.com/"
        $searchUri = $baseSearchUri + 'v1.0/groups?$filter=displayName+eq+' + "'$DisplayName'" + ' OR mail+eq+' + "'$Mail'" + ' OR mailNickname+eq+' + "'$MailNickname'" + '&$count=true'

        $azureADGroupResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
        $azureADGroup = $azureADGroupResponse.value

        if(@($azureADGroup).count -eq 0) {
            Write-Information "Group displayName=$DisplayName or mail=$Mail or mailNickname=$MailNickname not found"

            $returnObject = @{
                displayName=$DisplayName; 
                description=$Description; 
                mail=$Mail; 
                mailNickname=$MailNickname
            }
            
            Write-Output $returnObject
            break;
        } else {
            Write-Warning "Group displayName=$DisplayName or mail=$PrimarySmtpAddress or mailNickname=$MailNickname found"
        }
    }
} catch {
    if($_.ErrorDetails.Message) { $errorDetailsMessage = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message } 
    Write-Verbose -Verbose ("Error generating names. Error: $_" + $errorDetailsMessage)
}
