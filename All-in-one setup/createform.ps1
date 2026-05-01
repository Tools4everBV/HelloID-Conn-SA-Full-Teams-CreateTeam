# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Teams") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> EntraIdCertificatePassword
$tmpName = @'
EntraIdCertificatePassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #2 >> EntraIdAppId
$tmpName = @'
EntraIdAppId
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> EntraIdCertificateBase64String
$tmpName = @'
EntraIdCertificateBase64String
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #4 >> EntraIdTenantId
$tmpName = @'
EntraIdTenantId
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #5 >> TeamsMailsuffix
$tmpName = @'
TeamsMailsuffix
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #6 >> companyName
$tmpName = @'
companyName
'@ 
$tmpValue = @'
{{company.name}}
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});


#make sure write-information logging is visual
$InformationPreference = "continue"

# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}

# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}

# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  

# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}

function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )

    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false

        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid

            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}

function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}

        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task

            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid

            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }

    $returnObject.Value = $taskGuid
}

function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter()][String][AllowEmptyString()]$DatasourceRunInCloud,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }

    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
                runInCloud         = $DatasourceRunInCloud;
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
        Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }

    $returnObject.Value = $datasourceGuid
}

function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )

    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }

        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }

    $returnObject.Value = $formGuid
}


function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })

    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }

        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100

            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body

            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true

            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }

    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}

<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "teams-create-team | Validation" #>
$tmpPsScript = @'
#######################################################################
# Template: HelloID SA Powershell data source
# Name:     teams-create-team | Validation
# Date:     01-04-2026
#######################################################################

# For basic information about powershell data sources see:
# https://docs.helloid.com/en/service-automation/dynamic-forms/data-sources/powershell-data-sources/add,-edit,-or-remove-a-powershell-data-source.html#add-a-powershell-data-source

# Service automation variables:
# https://docs.helloid.com/en/service-automation/service-automation-variables/service-automation-variable-reference.html

#region init

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

$outputText = [System.Collections.Generic.List[PSCustomObject]]::new()

# global variables (Automation --> Variable library):
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""
# $TeamsMailsuffix = ""

# variables configured in form:
$displayName = $dataSource.displayName
$mail = $displayName.Replace(" ", "") + "@" + $TeamsMailsuffix 
$mailNickname = $displayName.Replace(" ", "")
#endregion init

#region functions
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
#endregion functions

#region lookup
try {

    $actionMessage = "checking Entra ID for uniqueness"

    # Setup Connection with Entra/Exo
    Write-Verbose 'connecting to MS-Entra'
    $certificate = Get-MSEntraCertificate
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate
    
    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $entraToken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    } 

    $graphApiUrl = "https://graph.microsoft.com/v1.0/groups"
    $select = '&$select=id,displayName,mail,mailNickname' + '&$top=999 + &$count=true'
    $filter = "?`$filter=displayName eq '$displayName' or mail eq '$mail' or mailNickname eq '$mailNickname'"
    $searchUri = $graphApiUrl + $filter + $select

    $entraIDGroupsParams = @{
        Uri     = $searchUri
        Method  = 'Get'
        Headers = $authorization
        Verbose = $false
    }

    $entraIDGroupsResponse = Invoke-RestMethod @entraIDGroupsParams

    $entraIDGroups = $entraIDGroupsResponse.value
    while (![string]::IsNullOrEmpty($entraIDGroupsResponse.'@odata.nextLink')) {
        $entraIDGroupsResponse = Invoke-RestMethod -Uri $entraIDGroupsResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $entraIDGroups += $entraIDGroupsResponse.value
    }  

    Write-Information "Found [$($entraIDGroups.Count)] groups"

    foreach ($record in $entraIDGroups) {
        if ($record.displayName -eq $displayName) {
            $outputText.Add([PSCustomObject]@{
                    Message  = "Display name [$displayName] not unique, found on [$($record.displayName)]"
                    IsError  = $true
                    Property = "displayName"
                })
        }
        if ($record.mail -eq $mail) {
            $outputText.Add([PSCustomObject]@{
                    Message  = "Mail [$mail] not unique, found on [$($record.displayName)]"
                    IsError  = $true
                    Property = "mail"
                })
        }
        if ($record.mailNickname -eq $mailNickname) {
            $outputText.Add([PSCustomObject]@{
                    Message  = "Mail nickname [$mailNickname] not unique, found on [$($record.displayName)]"
                    IsError  = $true
                    Property = "mailNickname"
                })
        }
    }

    if ($outputText.isError -contains $true) {
        $outputMessage = "Invalid:"
    }
    elseif (-not($outputText.isError -contains $false)) {
        $outputMessage = "Valid:"
        $outputText.Add([PSCustomObject]@{
                Message  = "Team with displayName [$displayName] is unique"
                IsError  = $false
                Property = "displayName"
            })
    }
    else {
        $outputMessage = "Valid:"
    }

    foreach ($text in $outputText) {
        $outputMessage += "`n" + $($text.Message)
    }

    Write-Output $outputMessage
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
    Write-Warning $warningMessage
    Write-Error $auditMessage
}
#endregion lookup

'@ 
$tmpModel = @'
[{"key":"output","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"displayName","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
teams-create-team | Validation
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "teams-create-team | Validation" #>

<# Begin: DataSource "teams-create-team | Teams-Get-All-Entra-Id-Users" #>
$tmpPsScript = @'
#######################################################################
# Template: HelloID SA Powershell data source
# Name:     teams-create-team | Teams-Get-All-Entra-Id-Users
# Date:     01-04-2026
#######################################################################

# For basic information about powershell data sources see:
# https://docs.helloid.com/en/service-automation/dynamic-forms/data-sources/powershell-data-sources/add,-edit,-or-remove-a-powershell-data-source.html#add-a-powershell-data-source

# Service automation variables:
# https://docs.helloid.com/en/service-automation/service-automation-variables/service-automation-variable-reference.html

#region init

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Fixed values
$filter = "`$filter=userType eq 'Member' and accountEnabled eq true"

$propertiesToSelect = @(
    "id",
    "userPrincipalName",
    "displayName",
    "mail",
    "description",
    "department",
    "jobTitle",
    "companyName",
    "accountEnabled"
) # Properties to select from Microsoft Graph API, comma separated

# global variables (Automation --> Variable library):
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""
# $TeamsMailsuffix = ""

# variables configured in form:
# $formValue1 = $datasource.<formElementKey>.<value>
# $formValue2 = $datasource.<formElementKey>

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
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    } 

    # Get Microsoft Entra ID Users
    # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching search value [$filter]"
    $microsoftEntraIDUsers = [System.Collections.ArrayList]@()
    do {
        $getMicrosoftEntraIDUsersSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
            Headers     = $authorization
            Method      = "GET"
            Verbose     = $false
            ErrorAction = "Stop"
        }
        if (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink')) {
            $getMicrosoftEntraIDUsersSplatParams["Uri"] = $getMicrosoftEntraIDUsersResponse.'@odata.nextLink'
        }
        
        $getMicrosoftEntraIDUsersResponse = $null
        $getMicrosoftEntraIDUsersResponse = Invoke-RestMethod @getMicrosoftEntraIDUsersSplatParams
    
        # Select only specified properties to limit memory usage
        $getMicrosoftEntraIDUsersResponse.Value = $getMicrosoftEntraIDUsersResponse.Value | Select-Object $propertiesToSelect

        if ($getMicrosoftEntraIDUsersResponse.Value -is [array]) {
            [void]$microsoftEntraIDUsers.AddRange($getMicrosoftEntraIDUsersResponse.Value)
        }
        else {
            [void]$microsoftEntraIDUsers.Add($getMicrosoftEntraIDUsersResponse.Value)
        }
    } while (-not[string]::IsNullOrEmpty($getMicrosoftEntraIDUsersResponse.'@odata.nextLink'))
    Write-Information "Queried Microsoft Entra ID Users matching search value [$filter]. Result count: $(@($microsoftEntraIDUsers).Count)"

    # Send results to HelloID
    $actionMessage = "sending results to HelloID"
    $microsoftEntraIDUsers | ForEach-Object {
        Write-Output $_
    }
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
    Write-Warning $warningMessage
    Write-Error $auditMessage
}

'@ 
$tmpModel = @'
[{"key":"id","type":0},{"key":"userPrincipalName","type":0},{"key":"displayName","type":0},{"key":"mail","type":0},{"key":"description","type":0},{"key":"department","type":0},{"key":"jobTitle","type":0},{"key":"companyName","type":0},{"key":"accountEnabled","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
teams-create-team | Teams-Get-All-Entra-Id-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -DataSourceRunInCloud "True" -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "teams-create-team | Teams-Get-All-Entra-Id-Users" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Teams - Create Team" #>
$tmpSchema = @"
[{"key":"DisplayName","templateOptions":{"label":"Displayname","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"Description","templateOptions":{"label":"Description","required":true},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"visibility","templateOptions":{"label":"Security","useObjects":false,"options":["Private","Public"],"required":true},"type":"radio","defaultValue":"Private","summaryVisibility":"Show","textOrLabel":"label","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"owners","templateOptions":{"label":"Select owner(s)","required":true,"grid":{"columns":[{"headerName":"User Principal Name","field":"userPrincipalName"},{"headerName":"Display Name","field":"displayName"}],"height":300,"rowSelection":"multiple"},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[]}},"useFilter":true,"useDefault":false,"allowCsvDownload":true},"type":"multiselectgrid","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":true},{"key":"tiValidation","templateOptions":{"label":"Validation","pattern":"^Valid:\\s*.*","minLength":1,"required":true,"useDataSource":true,"readonly":true,"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"displayName","otherFieldValue":{"otherFieldKey":"DisplayName"}}]}},"displayField":"output"},"validation":{"messages":{"pattern":"No valid teams"}},"type":"input","summaryVisibility":"Show","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Teams - Create Team
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
        
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}

$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $response = $response | Where-Object {$_.name.en -eq $category}
    
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
    
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100

        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid

        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Teams - Create Team
'@
$tmpTask = @'
{"name":"Teams - Create Team","script":"#######################################################################\n# Template: HelloID SA Delegated form task\n# Name: Teams - Create Team\n# Date: 01-04-2026\n#######################################################################\n\n# For basic information about delegated form tasks see:\n# https://docs.helloid.com/en/service-automation/delegated-forms/delegated-form-tasks.html\n\n# Service automation variables:\n# https://docs.helloid.com/en/service-automation/service-automation-variables.html\n\n#region init\n\n$VerbosePreference = \"SilentlyContinue\"\n$InformationPreference = \"Continue\"\n$WarningPreference = \"Continue\"\n\n# global variables (Automation --\u003e Variable libary):\n# Outcommented as these are set from Global Variables\n# $EntraIdTenantId = \"\"\n# $EntraIdAppId = \"\"\n# $EntraIdCertificateBase64String = \"\"\n# $EntraIdCertificatePassword = \"\"\n\n# variables configured in form:\n$displayName = $form.DisplayName\n$description = $form.Description\n$visibility = $form.visibility\n$owners = $form.owners\n\n#endregion init\n\n#region functions\nfunction Resolve-MicrosoftGraphAPIError {\n    [CmdletBinding()]\n    param (\n        [Parameter(Mandatory)]\n        [object]\n        $ErrorObject\n    )\n    process {\n        $httpErrorObj = [PSCustomObject]@{\n            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber\n            Line             = $ErrorObject.InvocationInfo.Line\n            ErrorDetails     = $ErrorObject.Exception.Message\n            FriendlyMessage  = $ErrorObject.Exception.Message\n        }\n        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {\n            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message\n        }\n        elseif ($ErrorObject.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027) {\n            if ($null -ne $ErrorObject.Exception.Response) {\n                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()\n                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {\n                    $httpErrorObj.ErrorDetails = $streamReaderResponse\n                }\n            }\n        }\n        try {\n            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)\n            if ($errorDetailsObject.error_description) {\n                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description\n            }\n            elseif ($errorDetailsObject.error.message) {\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)\"\n            }\n            elseif ($errorDetailsObject.error.details.message) {\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)\"\n            }\n            else {\n                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\n            }\n        }\n        catch {\n            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\n        }\n        Write-Output $httpErrorObj\n    }\n}\n\nfunction Get-MSEntraAccessToken {\n    [CmdletBinding()]\n    param(\n        [Parameter(Mandatory)]\n        $Certificate\n    )\n    try {\n        # Get the DER encoded bytes of the certificate\n        $derBytes = $Certificate.RawData\n\n        # Compute the SHA-256 hash of the DER encoded bytes\n        $sha256 = [System.Security.Cryptography.SHA256]::Create()\n        $hashBytes = $sha256.ComputeHash($derBytes)\n        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\n        # Create a JWT (JSON Web Token) header\n        $header = @{\n            \u0027alg\u0027      = \u0027RS256\u0027\n            \u0027typ\u0027      = \u0027JWT\u0027\n            \u0027x5t#S256\u0027 = $base64Thumbprint\n        } | ConvertTo-Json\n        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))\n\n        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for \u0027exp\u0027, \u0027nbf\u0027 and \u0027iat\u0027\n        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]\u00271970-01-01T00:00:00Z\u0027).ToUniversalTime()).TotalSeconds)\n\n        # Create a JWT payload\n        $payload = [Ordered]@{\n            \u0027iss\u0027 = \"$entraidappid\"\n            \u0027sub\u0027 = \"$entraidappid\"\n            \u0027aud\u0027 = \"https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token\"\n            \u0027exp\u0027 = ($currentUnixTimestamp + 3600) # Expires in 1 hour\n            \u0027nbf\u0027 = ($currentUnixTimestamp - 300) # Not before 5 minutes ago\n            \u0027iat\u0027 = $currentUnixTimestamp\n            \u0027jti\u0027 = [Guid]::NewGuid().ToString()\n        } | ConvertTo-Json\n        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\n        # Extract the private key from the certificate\n        $rsaPrivate = $Certificate.PrivateKey\n        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()\n        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))\n\n        # Sign the JWT\n        $signatureInput = \"$base64Header.$base64Payload\"\n        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), \u0027SHA256\u0027)\n        $base64Signature = [System.Convert]::ToBase64String($signature).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\n\t\n        # Extract the private key from the certificate\n        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {\n            throw \"The certificate does not have a private key.\"\n        }\n\n        # Create the JWT token\n        $jwtToken = \"$($base64Header).$($base64Payload).$($base64Signature)\"\n\n        $createEntraAccessTokenBody = @{\n            grant_type            = \u0027client_credentials\u0027\n            client_id             = $entraidappid\n            client_assertion_type = \u0027urn:ietf:params:oauth:client-assertion-type:jwt-bearer\u0027\n            client_assertion      = $jwtToken\n            resource              = \u0027https://graph.microsoft.com\u0027\n        }\n\n        $createEntraAccessTokenSplatParams = @{\n            Uri         = \"https://login.microsoftonline.com/$EntraIdTenantId/oauth2/token\"\n            Body        = $createEntraAccessTokenBody\n            Method      = \u0027POST\u0027\n            ContentType = \u0027application/x-www-form-urlencoded\u0027\n            Verbose     = $false\n            ErrorAction = \u0027Stop\u0027\n        }\n\n        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams\n        Write-Output $createEntraAccessTokenResponse.access_token\n    }\n    catch {\n        $PSCmdlet.ThrowTerminatingError($_)\n    }\n}\n\nfunction Get-MSEntraCertificate {\n    [CmdletBinding()]\n    param()\n    try {\n        $rawCertificate = [system.convert]::FromBase64String($EntraIdCertificateBase64String)\n        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $EntraIdCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)\n        Write-Output $certificate\n    }\n    catch {\n        $PSCmdlet.ThrowTerminatingError($_)\n    }\n}\n\n#endregion functions\n\ntry {\n    # Setup Connection with Entra/Exo\n    Write-Verbose \u0027connecting to MS-Entra\u0027\n    $certificate = Get-MSEntraCertificate\n    $entraToken = Get-MSEntraAccessToken -Certificate $certificate\n      \n    #Add the authorization header to the request\n    $authorization = @{\n        Authorization  = \"Bearer $entraToken\";\n        \u0027Content-Type\u0027 = \"application/json\";\n        Accept         = \"application/json\";\n    } \n    \n    $actionMessage = \"creating team\"\n\n    $body = @{\n        \"Template@odata.bind\" = \"https://graph.microsoft.com/v1.0/teamsTemplates(\u0027standard\u0027)\"\n        \"DisplayName\"         = $displayName\n        \"Description\"         = $description\n        \"visibility\"          = $visibility\n        \"Members\"             = @()\n    }\n\n    # API only supports a single member during team creation.\n    # Add the first owner at creation time and add remaining owners afterwards.\n\n    $body.Members += @{\n        \"@odata.type\"     = \"#microsoft.graph.aadUserConversationMember\"\n        roles             = @(\"owner\")\n        \"user@odata.bind\" = \"https://graph.microsoft.com/v1.0/users/$($owners[0].id)\"\n        \"ConsistencyLevel\" = \"eventual\" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)\n    }\n\n    $createTeamSplatParams = @{\n        Uri         = \"https://graph.microsoft.com/v1.0/teams\"\n        Body        = ($body | ConvertTo-Json -Depth 10)\n        Headers     = $authorization\n        Method      = \u0027POST\u0027\n        ContentType = \u0027application/json\u0027\n        Verbose     = $false\n        ErrorAction = \u0027Stop\u0027\n    }\n    $createTeamResponse = Invoke-WebRequest @createTeamSplatParams\n\n    Write-Information \"Successfully created team [$displayName] with description [$description] and owner [$($owners[0].displayName)].\"\n\n    $teamId = $null\n    $locationHeader = [string]$createTeamResponse.Headers[\u0027Location\u0027]\n\n    # Team creation is asynchronous and typically returns 202 Accepted.\n    if ($createTeamResponse.StatusCode -eq 202) {\n\n        $actionMessage = \"adding additional owners\"\n        Start-Sleep -Seconds 5\n\n        if ([string]::IsNullOrEmpty($locationHeader)) {\n            throw \"Team creation returned 202 Accepted but did not include a Location header for operation status.\"\n        }\n\n        $locationMatch = [regex]::Match($locationHeader, \"^/teams\\(\u0027([^\u0027]+)\u0027\\)/operations\\(\u0027([^\u0027]+)\u0027\\)$\")\n        if (-not $locationMatch.Success) {\n            throw \"Location header format is unexpected: [$locationHeader]\"\n        }\n        $teamId = $locationMatch.Groups[1].Value\n\n        if ($locationHeader.StartsWith(\"http\")) {\n            $operationUri = $locationHeader\n        }\n        else {\n            $operationUri = \"https://graph.microsoft.com/v1.0$locationHeader\"\n        }\n        $maxPollAttempts = 5\n        $pollDelaySeconds = 30\n\n        for ($attempt = 1; $attempt -le $maxPollAttempts; $attempt++) {\n            $operation = Invoke-RestMethod -Uri $operationUri -Headers $authorization -Method \u0027GET\u0027 -ErrorAction Stop\n\n            if ($operation.status -eq \u0027succeeded\u0027) {\n                break\n            }\n\n            if ($operation.status -eq \u0027failed\u0027) {\n                throw \"Team creation operation failed. Status: [$($operation.status)]. Error: [$($operation.error.message)]\"\n            }\n\n            Start-Sleep -Seconds $pollDelaySeconds\n        }\n\n        if ($operation.status -ne \u0027succeeded\u0027) {\n            throw \"Timed out waiting for team creation operation to complete. Last status: [$($operation.status)]\"\n        }\n        \n        if ([string]::IsNullOrEmpty($teamId)) {\n            throw \"Unable to determine Team ID from Graph create team response headers.\"\n            \n        }\n\n        if ($owners.Count -gt 1) {\n            $additionalMembers = @()\n            foreach ($owner in ($owners | Select-Object -Skip 1)) {\n                $additionalMembers += @{\n                    \"@odata.type\"     = \"#microsoft.graph.aadUserConversationMember\"\n                    roles             = @(\"owner\")\n                    \"user@odata.bind\" = \"https://graph.microsoft.com/v1.0/users/$($owner.id)\"\n                }\n            }\n\n            $addMembersBody = @{\n                values = $additionalMembers\n            }\n\n            $addMembersSplatParams = @{\n                Uri         = \"https://graph.microsoft.com/v1.0/teams/$($teamId)/members/add\"\n                Body        = ($addMembersBody | ConvertTo-Json -Depth 10)\n                Headers     = $authorization\n                Method      = \u0027POST\u0027\n                ContentType = \u0027application/json\u0027\n                Verbose     = $false\n                ErrorAction = \u0027Stop\u0027\n            }\n            $null = Invoke-RestMethod @addMembersSplatParams\n\n            Write-Information \"Successfully added [$(($owners | Select-Object -Skip 1 | ForEach-Object { $_.displayName }) -join \u0027, \u0027)] as additional owners to team [$displayName]\"\n        }\n    }\n\n    $Log = @{\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \n        System            = \"MicrosoftTeams\" # optional (free format text) \n        Message           = \"Successfully created team [$displayName] with description [$description] and owners [$(($owners | ForEach-Object { $_.displayName }) -join \u0027, \u0027)]\" # required (free format text)\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $displayName # optional (free format text) \n        TargetIdentifier  = $([string]$teamId) # optional (free format text) \n    }\n    #send result back  \n    Write-Information -Tags \"Audit\" -MessageData $log    \n}\ncatch {\n    $ex = $PSItem\n    if ($($ex.Exception.GetType().FullName -eq \u0027Microsoft.PowerShell.Commands.HttpResponseException\u0027) -or\n        $($ex.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027)) {\n        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex\n        $auditMessage = \"Error $($actionMessage). Error: $($errorObj.FriendlyMessage)\"\n        $warningMessage = \"Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)\"\n    }\n    else {\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\n    }\n    $Log = @{\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \n        System            = \"MicrosoftTeams\" # optional (free format text) \n        Message           = $auditMessage # required (free format text) \n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \n        TargetDisplayName = $displayName # optional (free format text) \n        TargetIdentifier  = $([string]$teamId) # optional (free format text) \n    }\n    Write-Information -Tags \"Audit\" -MessageData $log\n    Write-Warning $warningMessage\n    Write-Error $auditMessage\n}\n","runInCloud":true}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-plus-square" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

