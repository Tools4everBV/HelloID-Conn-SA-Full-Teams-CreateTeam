# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$baseGraphUri = "https://graph.microsoft.com/"

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$displayName = $form.generatedNames.displayName
$description = $form.generatedNames.description
$visibility = $form.visibility
$owner = $form.owner.id

# Create authorization token and add to headers
try{
    Write-Information "Generating Microsoft Graph API Access Token"

    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AADAppId"
        client_secret = "$AADAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }
}
catch{
    throw "Could not generate Microsoft Graph API Access Token. Error: $($_.Exception.Message)"    
}

try {
    Write-Information "Creating Team [$displayName] with description [$description]."

    $createTeamUri = $baseGraphUri + "v1.0/teams"
    #Write-Information $createTeamUri

    #$URLOwner = "https://graph.microsoft.com/v1.0/users/$($owner.user)"
    #Write-Information $URLOwnwer

    #$ResultOwner = Invoke-RestMethod -Headers $authorization -Uri $URLOwner -Method Get
    #Write-Information ($ResultOwner | ConvertTo-Json -Depth 10)

    $bodyJson = @"
    {
        "Template@odata.bind":"https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        "DisplayName":"$displayName",
        "Description":"$description",
        "visibility":"$visibility",
        "Members":[
            {
                "`@odata.type":"#microsoft.graph.aadUserConversationMember",
                "Roles":[
                    "owner"
                ],
                "User`@odata.bind":"https://graph.microsoft.com/v1.0/users/$owner"
            }
        ]
    }
"@

    Write-Information $bodyJson 

    $team = Invoke-RestMethod -Method POST -Uri $createTeamUri -Body $bodyJson -Headers $authorization -Verbose:$false
    
    Write-Information "Successfully created Team [$displayName] with description [$description]."
    $Log = @{
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = "Successfully created team [$displayName] with description [$description]." # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text)
        TargetIdentifier  = $($team.id) # optional (free format text)
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch
{
    Write-Error "Failed to create Team [$displayName]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = "Failed to create team [$displayName] with description [$description]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text)
        TargetIdentifier  = $($team.id) # optional (free format text)
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

