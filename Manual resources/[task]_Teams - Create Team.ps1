#######################################################################
# Template: HelloID SA Delegated form task
# Name: Teams - Create Team
# Date: 01-04-2026
#######################################################################

# For basic information about delegated form tasks see:
# https://docs.helloid.com/en/service-automation/delegated-forms/delegated-form-tasks.html

# Service automation variables:
# https://docs.helloid.com/en/service-automation/service-automation-variables.html

#region init

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# global variables (Automation --> Variable libary):
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# variables configured in form:
$displayName = $form.DisplayName
$description = $form.Description
$visibility = $form.visibility
$owners = $form.owners

#endregion init

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)
            if ($errorDetailsObject.error_description) {
                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description
            }
            elseif ($errorDetailsObject.error.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)"
            }
            elseif ($errorDetailsObject.error.details.message) {
                $httpErrorObj.FriendlyMessage = "$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)"
            }
            else {
                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Get-MSEntraAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Certificate
    )
    try {
        # Get the DER encoded bytes of the certificate
        $derBytes = $Certificate.RawData

        # Compute the SHA-256 hash of the DER encoded bytes
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha256.ComputeHash($derBytes)
        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Create a JWT (JSON Web Token) header
        $header = @{
            'alg'      = 'RS256'
            'typ'      = 'JWT'
            'x5t#S256' = $base64Thumbprint
        } | ConvertTo-Json
        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))

        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for 'exp', 'nbf' and 'iat'
        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]'1970-01-01T00:00:00Z').ToUniversalTime()).TotalSeconds)

        # Create a JWT payload
        $payload = [Ordered]@{
            'iss' = "$entraidappid"
            'sub' = "$entraidappid"
            'aud' = "https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token"
            'exp' = ($currentUnixTimestamp + 3600) # Expires in 1 hour
            'nbf' = ($currentUnixTimestamp - 300) # Not before 5 minutes ago
            'iat' = $currentUnixTimestamp
            'jti' = [Guid]::NewGuid().ToString()
        } | ConvertTo-Json
        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace('+', '-').Replace('/', '_').Replace('=', '')

        # Extract the private key from the certificate
        $rsaPrivate = $Certificate.PrivateKey
        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))

        # Sign the JWT
        $signatureInput = "$base64Header.$base64Payload"
        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), 'SHA256')
        $base64Signature = [System.Convert]::ToBase64String($signature).Replace('+', '-').Replace('/', '_').Replace('=', '')
	
        # Extract the private key from the certificate
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $entraidappid
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token"
            Body        = $createEntraAccessTokenBody
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded'
            Verbose     = $false
            ErrorAction = 'Stop'
        }

        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams
        Write-Output $createEntraAccessTokenResponse.access_token
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-MSEntraCertificate {
    [CmdletBinding()]
    param()
    try {
        $rawCertificate = [system.convert]::FromBase64String($EntraIdCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $EntraIdCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

#endregion functions

try {
    # Setup Connection with Entra/Exo
    Write-Verbose 'connecting to MS-Entra'
    $certificate = Get-MSEntraCertificate
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate
      
    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $entraToken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    } 
    
    $actionMessage = "creating team"

    $body = @{
        "Template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        "DisplayName"         = $displayName
        "Description"         = $description
        "visibility"          = $visibility
        "Members"             = @()
    }

    # API only supports a single member during team creation.
    # Add the first owner at creation time and add remaining owners afterwards.

    $body.Members += @{
        "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
        roles             = @("owner")
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$($owners[0].id)"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    $createTeamSplatParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/teams"
        Body        = ($body | ConvertTo-Json -Depth 10)
        Headers     = $authorization
        Method      = 'POST'
        ContentType = 'application/json'
        Verbose     = $false
        ErrorAction = 'Stop'
    }
    $createTeamResponse = Invoke-WebRequest @createTeamSplatParams

    Write-Information "Successfully created team [$displayName] with description [$description] and owner [$($owners[0].displayName)]."

    $teamId = $null
    $locationHeader = [string]$createTeamResponse.Headers['Location']

    # Team creation is asynchronous and typically returns 202 Accepted.
    if ($createTeamResponse.StatusCode -eq 202) {

        $actionMessage = "adding additional owners"
        Start-Sleep -Seconds 5

        if ([string]::IsNullOrEmpty($locationHeader)) {
            throw "Team creation returned 202 Accepted but did not include a Location header for operation status."
        }

        $locationMatch = [regex]::Match($locationHeader, "^/teams\('([^']+)'\)/operations\('([^']+)'\)$")
        if (-not $locationMatch.Success) {
            throw "Location header format is unexpected: [$locationHeader]"
        }
        $teamId = $locationMatch.Groups[1].Value

        if ($locationHeader.StartsWith("http")) {
            $operationUri = $locationHeader
        }
        else {
            $operationUri = "https://graph.microsoft.com/v1.0$locationHeader"
        }
        $maxPollAttempts = 5
        $pollDelaySeconds = 30

        for ($attempt = 1; $attempt -le $maxPollAttempts; $attempt++) {
            $operation = Invoke-RestMethod -Uri $operationUri -Headers $authorization -Method 'GET' -ErrorAction Stop

            if ($operation.status -eq 'succeeded') {
                break
            }

            if ($operation.status -eq 'failed') {
                throw "Team creation operation failed. Status: [$($operation.status)]. Error: [$($operation.error.message)]"
            }

            Start-Sleep -Seconds $pollDelaySeconds
        }

        if ($operation.status -ne 'succeeded') {
            throw "Timed out waiting for team creation operation to complete. Last status: [$($operation.status)]"
        }
        
        if ([string]::IsNullOrEmpty($teamId)) {
            throw "Unable to determine Team ID from Graph create team response headers."
            
        }

        if ($owners.Count -gt 1) {
            $additionalMembers = @()
            foreach ($owner in ($owners | Select-Object -Skip 1)) {
                $additionalMembers += @{
                    "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                    roles             = @("owner")
                    "user@odata.bind" = "https://graph.microsoft.com/v1.0/users/$($owner.id)"
                }
            }

            $addMembersBody = @{
                values = $additionalMembers
            }

            $addMembersSplatParams = @{
                Uri         = "https://graph.microsoft.com/v1.0/teams/$($teamId)/members/add"
                Body        = ($addMembersBody | ConvertTo-Json -Depth 10)
                Headers     = $authorization
                Method      = 'POST'
                ContentType = 'application/json'
                Verbose     = $false
                ErrorAction = 'Stop'
            }
            $null = Invoke-RestMethod @addMembersSplatParams

            Write-Information "Successfully added [$(($owners | Select-Object -Skip 1 | ForEach-Object { $_.displayName }) -join ', ')] as additional owners to team [$displayName]"
        }
    }

    $Log = @{
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = "Successfully created team [$displayName] with description [$description] and owners [$(($owners | ForEach-Object { $_.displayName }) -join ', ')]" # required (free format text)
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text) 
        TargetIdentifier  = $([string]$teamId) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log    
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    $Log = @{
        Action            = "CreateResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = $auditMessage # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $displayName # optional (free format text) 
        TargetIdentifier  = $([string]$teamId) # optional (free format text) 
    }
    Write-Information -Tags "Audit" -MessageData $log
    Write-Warning $warningMessage
    Write-Error $auditMessage
}

