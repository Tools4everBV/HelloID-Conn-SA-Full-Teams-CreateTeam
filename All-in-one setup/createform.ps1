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

#Global variable #1 >> EntraIdCertificateBase64String
$tmpName = @'
EntraIdCertificateBase64String
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #2 >> EntraIdAppId
$tmpName = @'
EntraIdAppId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> EntraIdTenantId
$tmpName = @'
EntraIdTenantId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #4 >> EntraIdCertificatePassword
$tmpName = @'
EntraIdCertificatePassword
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});

#Global variable #5 >> TeamsMailsuffix
$tmpName = @'
TeamsMailsuffix
'@ 
$tmpValue = @'
'@ 
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
# Name:     Teams - Create Team | validation
# Date:     04-03-2026
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
            'iss' = "$EntraIdAppId"
            'sub' = "$EntraIdAppId"
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
            client_id             = $EntraIdAppId
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
    Write-Information 'Checking Entra ID for uniqueness, connecting to MS-Entra'
    $certificate = Get-MSEntraCertificate
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate
    
    #Add the authorization header to the request
    $authorization = @{
        Authorization    = "Bearer $entraToken"
        'Content-Type'   = "application/json"
        Accept           = "application/json"
        ConsistencyLevel = "eventual"
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

    Write-Warning "groups: [$($entraIDGroups | ConvertTo-Json)]"

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

# Global variables
# Outcommented as these are set from Global Variables
# $EntraIdTenantId = ""
# $EntraIdAppId = ""
# $EntraIdCertificateBase64String = ""
# $EntraIdCertificatePassword = ""

# Fixed values
$filter = "`$filter=accountEnabled eq true" # Get all enabled users

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

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

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
        [ValidateNotNull()]
        $Certificate,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AppId,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TenantId
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
            'iss' = "$($AppId)"
            'sub' = "$($AppId)"
            'aud' = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
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

        # Ensure the certificate has a private key
        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {
            throw "The certificate does not have a private key."
        }

        # Create the JWT token
        $jwtToken = "$($base64Header).$($base64Payload).$($base64Signature)"

        $createEntraAccessTokenBody = @{
            grant_type            = 'client_credentials'
            client_id             = $AppId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwtToken
            resource              = 'https://graph.microsoft.com'
        }

        $createEntraAccessTokenSplatParams = @{
            Uri         = "https://login.microsoftonline.com/$($TenantId)/oauth2/token"
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
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificateBase64String,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CertificatePassword
    )
    try {
        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        Write-Output $certificate
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
#endregion functions

try {
    # Convert base64 certificate string to certificate object
    $actionMessage = "converting base64 certificate string to certificate object"
    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword

    # Create access token
    $actionMessage = "creating access token"
    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId

    # Create headers
    $actionMessage = "creating headers"
    $headers = @{
        "Authorization"    = "Bearer $($entraToken)"
        "Accept"           = "application/json"
        "Content-Type"     = "application/json"
        "ConsistencyLevel" = "eventual" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
    }

    # Get Microsoft Entra ID Users
    # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying Microsoft Entra ID Users matching search value [$filter]"
    $microsoftEntraIDUsers = [System.Collections.ArrayList]@()
    do {
        $getMicrosoftEntraIDUsersSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/users?$filter&`$select=$($propertiesToSelect -join ',')&`$top=999&`$count=true"
            Headers     = $headers
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
{"name":"Teams - Create Team","script":"# variables configured in form:\r\n$displayName = $form.DisplayName\r\n$description = $form.Description\r\n$visibility = $form.visibility\r\n$owners = $form.owners\r\n\r\n# Global variables\r\n# Outcommented as these are set from Global Variables\r\n# $EntraIdTenantId = \"\"\r\n# $EntraIdAppId = \"\"\r\n# $EntraIdCertificateBase64String = \"\"\r\n# $EntraIdCertificatePassword = \"\"\r\n\r\n# Enable TLS1.2\r\n[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12\r\n\r\n# Set debug logging\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\n#region functions\r\nfunction Resolve-MicrosoftGraphAPIError {\r\n    [CmdletBinding()]\r\n    param (\r\n        [Parameter(Mandatory)]\r\n        [object]\r\n        $ErrorObject\r\n    )\r\n    process {\r\n        $httpErrorObj = [PSCustomObject]@{\r\n            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber\r\n            Line             = $ErrorObject.InvocationInfo.Line\r\n            ErrorDetails     = $ErrorObject.Exception.Message\r\n            FriendlyMessage  = $ErrorObject.Exception.Message\r\n        }\r\n        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {\r\n            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message\r\n        }\r\n        elseif ($ErrorObject.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027) {\r\n            if ($null -ne $ErrorObject.Exception.Response) {\r\n                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()\r\n                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {\r\n                    $httpErrorObj.ErrorDetails = $streamReaderResponse\r\n                }\r\n            }\r\n        }\r\n        try {\r\n            $errorDetailsObject = ($httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop)\r\n            if ($errorDetailsObject.error_description) {\r\n                $httpErrorObj.FriendlyMessage = $errorDetailsObject.error_description\r\n            }\r\n            elseif ($errorDetailsObject.error.message) {\r\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.code): $($errorDetailsObject.error.message)\"\r\n            }\r\n            elseif ($errorDetailsObject.error.details.message) {\r\n                $httpErrorObj.FriendlyMessage = \"$($errorDetailsObject.error.details.code): $($errorDetailsObject.error.details.message)\"\r\n            }\r\n            else {\r\n                $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\r\n            }\r\n        }\r\n        catch {\r\n            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails\r\n        }\r\n        Write-Output $httpErrorObj\r\n    }\r\n}\r\n\r\n\r\nfunction Get-MSEntraAccessToken {\r\n    [CmdletBinding()]\r\n    param(\r\n        [Parameter(Mandatory)]\r\n        [ValidateNotNull()]\r\n        $Certificate,\r\n        \r\n        [Parameter(Mandatory)]\r\n        [ValidateNotNullOrEmpty()]\r\n        [string]\r\n        $AppId,\r\n        \r\n        [Parameter(Mandatory)]\r\n        [ValidateNotNullOrEmpty()]\r\n        [string]\r\n        $TenantId\r\n    )\r\n    try {\r\n        # Get the DER encoded bytes of the certificate\r\n        $derBytes = $Certificate.RawData\r\n\r\n        # Compute the SHA-256 hash of the DER encoded bytes\r\n        $sha256 = [System.Security.Cryptography.SHA256]::Create()\r\n        $hashBytes = $sha256.ComputeHash($derBytes)\r\n        $base64Thumbprint = [System.Convert]::ToBase64String($hashBytes).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\r\n\r\n        # Create a JWT (JSON Web Token) header\r\n        $header = @{\r\n            \u0027alg\u0027      = \u0027RS256\u0027\r\n            \u0027typ\u0027      = \u0027JWT\u0027\r\n            \u0027x5t#S256\u0027 = $base64Thumbprint\r\n        } | ConvertTo-Json\r\n        $base64Header = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header))\r\n\r\n        # Calculate the Unix timestamp (seconds since 1970-01-01T00:00:00Z) for \u0027exp\u0027, \u0027nbf\u0027 and \u0027iat\u0027\r\n        $currentUnixTimestamp = [math]::Round(((Get-Date).ToUniversalTime() - ([datetime]\u00271970-01-01T00:00:00Z\u0027).ToUniversalTime()).TotalSeconds)\r\n\r\n        # Create a JWT payload\r\n        $payload = [Ordered]@{\r\n            \u0027iss\u0027 = \"$($AppId)\"\r\n            \u0027sub\u0027 = \"$($AppId)\"\r\n            \u0027aud\u0027 = \"https://login.microsoftonline.com/$($TenantId)/oauth2/token\"\r\n            \u0027exp\u0027 = ($currentUnixTimestamp + 3600) # Expires in 1 hour\r\n            \u0027nbf\u0027 = ($currentUnixTimestamp - 300) # Not before 5 minutes ago\r\n            \u0027iat\u0027 = $currentUnixTimestamp\r\n            \u0027jti\u0027 = [Guid]::NewGuid().ToString()\r\n        } | ConvertTo-Json\r\n        $base64Payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\r\n\r\n        # Extract the private key from the certificate\r\n        if (-not $Certificate.HasPrivateKey -or -not $Certificate.PrivateKey) {\r\n            throw \"The certificate does not have a private key.\"\r\n        }\r\n        $rsaPrivate = $Certificate.PrivateKey\r\n        $rsa = [System.Security.Cryptography.RSACryptoServiceProvider]::new()\r\n        $rsa.ImportParameters($rsaPrivate.ExportParameters($true))\r\n\r\n        # Sign the JWT\r\n        $signatureInput = \"$base64Header.$base64Payload\"\r\n        $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signatureInput), \u0027SHA256\u0027)\r\n        $base64Signature = [System.Convert]::ToBase64String($signature).Replace(\u0027+\u0027, \u0027-\u0027).Replace(\u0027/\u0027, \u0027_\u0027).Replace(\u0027=\u0027, \u0027\u0027)\r\n\r\n        # Create the JWT token\r\n        $jwtToken = \"$($base64Header).$($base64Payload).$($base64Signature)\"\r\n\r\n        $createEntraAccessTokenBody = @{\r\n            grant_type            = \u0027client_credentials\u0027\r\n            client_id             = $AppId\r\n            client_assertion_type = \u0027urn:ietf:params:oauth:client-assertion-type:jwt-bearer\u0027\r\n            client_assertion      = $jwtToken\r\n            resource              = \u0027https://graph.microsoft.com\u0027\r\n        }\r\n\r\n        $createEntraAccessTokenSplatParams = @{\r\n            Uri         = \"https://login.microsoftonline.com/$($TenantId)/oauth2/token\"\r\n            Body        = $createEntraAccessTokenBody\r\n            Method      = \u0027POST\u0027\r\n            ContentType = \u0027application/x-www-form-urlencoded\u0027\r\n            Verbose     = $false\r\n            ErrorAction = \u0027Stop\u0027\r\n        }\r\n\r\n        $createEntraAccessTokenResponse = Invoke-RestMethod @createEntraAccessTokenSplatParams\r\n        Write-Output $createEntraAccessTokenResponse.access_token\r\n    }\r\n    catch {\r\n        $PSCmdlet.ThrowTerminatingError($_)\r\n    }\r\n}\r\n\r\nfunction Get-MSEntraCertificate {\r\n    [CmdletBinding()]\r\n    param(\r\n        [Parameter(Mandatory)]\r\n        [ValidateNotNullOrEmpty()]\r\n        [string]\r\n        $CertificateBase64String,\r\n        \r\n        [Parameter(Mandatory)]\r\n        [ValidateNotNullOrEmpty()]\r\n        [string]\r\n        $CertificatePassword\r\n    )\r\n    try {\r\n        $rawCertificate = [system.convert]::FromBase64String($CertificateBase64String)\r\n        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)\r\n        Write-Output $certificate\r\n    }\r\n    catch {\r\n        $PSCmdlet.ThrowTerminatingError($_)\r\n    }\r\n}\r\n#endregion functions\r\n\r\ntry {\r\n    # Convert base64 certificate string to certificate object\r\n    $actionMessage = \"converting base64 certificate string to certificate object\"\r\n    $certificate = Get-MSEntraCertificate -CertificateBase64String $EntraIdCertificateBase64String -CertificatePassword $EntraIdCertificatePassword\r\n\r\n    # Create access token\r\n    $actionMessage = \"creating access token\"\r\n    $entraToken = Get-MSEntraAccessToken -Certificate $certificate -AppId $EntraIdAppId -TenantId $EntraIdTenantId\r\n\r\n    # Create headers\r\n    $actionMessage = \"creating headers\"\r\n    $headers = @{\r\n        \"Authorization\"    = \"Bearer $($entraToken)\"\r\n        \"Accept\"           = \"application/json\"\r\n        \"Content-Type\"     = \"application/json\"\r\n        \"ConsistencyLevel\" = \"eventual\" # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)\r\n    }\r\n    \r\n    $actionMessage = \"creating team\"\r\n\r\n    $body = @{\r\n        \"Template@odata.bind\" = \"https://graph.microsoft.com/v1.0/teamsTemplates(\u0027standard\u0027)\"\r\n        \"DisplayName\"         = $displayName\r\n        \"Description\"         = $description\r\n        \"visibility\"          = $visibility\r\n        \"Members\"             = @()\r\n    }\r\n\r\n    # API only supports a single member during team creation.\r\n    # Add the first owner at creation time and add remaining owners afterwards.\r\n\r\n    $body.Members += @{\r\n        \"@odata.type\"     = \"#microsoft.graph.aadUserConversationMember\"\r\n        roles             = @(\"owner\")\r\n        \"user@odata.bind\" = \"https://graph.microsoft.com/v1.0/users/$($owners[0].id)\"\r\n    }\r\n\r\n    $createTeamSplatParams = @{\r\n        Uri         = \"https://graph.microsoft.com/v1.0/teams\"\r\n        Body        = ($body | ConvertTo-Json -Depth 10)\r\n        Headers     = $headers\r\n        Method      = \u0027POST\u0027\r\n        ContentType = \u0027application/json\u0027\r\n        Verbose     = $false\r\n        ErrorAction = \u0027Stop\u0027\r\n    }\r\n    $createTeamResponse = Invoke-WebRequest @createTeamSplatParams\r\n\r\n    Write-Information \"Successfully created team [$displayName] with description [$description] and owner [$($owners[0].displayName)].\"\r\n\r\n    $teamId = $null\r\n    $locationHeader = [string]$createTeamResponse.Headers[\u0027Location\u0027]\r\n\r\n    # Team creation is asynchronous and typically returns 202 Accepted.\r\n    if ($createTeamResponse.StatusCode -eq 202) {\r\n\r\n        $actionMessage = \"adding additional owners\"\r\n        Start-Sleep -Seconds 5\r\n\r\n        if ([string]::IsNullOrEmpty($locationHeader)) {\r\n            throw \"Team creation returned 202 Accepted but did not include a Location header for operation status.\"\r\n        }\r\n\r\n        $locationMatch = [regex]::Match($locationHeader, \"^/teams\\(\u0027([^\u0027]+)\u0027\\)/operations\\(\u0027([^\u0027]+)\u0027\\)$\")\r\n        if (-not $locationMatch.Success) {\r\n            throw \"Location header format is unexpected: [$locationHeader]\"\r\n        }\r\n        $teamId = $locationMatch.Groups[1].Value\r\n\r\n        if ($locationHeader.StartsWith(\"http\")) {\r\n            $operationUri = $locationHeader\r\n        }\r\n        else {\r\n            $operationUri = \"https://graph.microsoft.com/v1.0$locationHeader\"\r\n        }\r\n        $maxPollAttempts = 5\r\n        $pollDelaySeconds = 30\r\n\r\n        for ($attempt = 1; $attempt -le $maxPollAttempts; $attempt++) {\r\n            $operation = Invoke-RestMethod -Uri $operationUri -Headers $headers -Method \u0027GET\u0027 -ErrorAction Stop\r\n\r\n            if ($operation.status -eq \u0027succeeded\u0027) {\r\n                break\r\n            }\r\n\r\n            if ($operation.status -eq \u0027failed\u0027) {\r\n                throw \"Team creation operation failed. Status: [$($operation.status)]. Error: [$($operation.error.message)]\"\r\n            }\r\n\r\n            Start-Sleep -Seconds $pollDelaySeconds\r\n        }\r\n\r\n        if ($operation.status -ne \u0027succeeded\u0027) {\r\n            throw \"Timed out waiting for team creation operation to complete. Last status: [$($operation.status)]\"\r\n        }\r\n        \r\n        if ([string]::IsNullOrEmpty($teamId)) {\r\n            throw \"Unable to determine Team ID from Graph create team response headers.\"\r\n            \r\n        }\r\n\r\n        if ($owners.Count -gt 1) {\r\n            $additionalMembers = @()\r\n            foreach ($owner in ($owners | Select-Object -Skip 1)) {\r\n                $additionalMembers += @{\r\n                    \"@odata.type\"     = \"#microsoft.graph.aadUserConversationMember\"\r\n                    roles             = @(\"owner\")\r\n                    \"user@odata.bind\" = \"https://graph.microsoft.com/v1.0/users/$($owner.id)\"\r\n                }\r\n            }\r\n\r\n            $addMembersBody = @{\r\n                values = $additionalMembers\r\n            }\r\n\r\n            $addMembersSplatParams = @{\r\n                Uri         = \"https://graph.microsoft.com/v1.0/teams/$($teamId)/members/add\"\r\n                Body        = ($addMembersBody | ConvertTo-Json -Depth 10)\r\n                Headers     = $headers\r\n                Method      = \u0027POST\u0027\r\n                ContentType = \u0027application/json\u0027\r\n                Verbose     = $false\r\n                ErrorAction = \u0027Stop\u0027\r\n            }\r\n            $null = Invoke-RestMethod @addMembersSplatParams\r\n\r\n            Write-Information \"Successfully added [$(($owners | Select-Object -Skip 1 | ForEach-Object { $_.displayName }) -join \u0027, \u0027)] as additional owners to team [$displayName]\"\r\n        }\r\n    }\r\n\r\n    $Log = @{\r\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \r\n        System            = \"MicrosoftTeams\" # optional (free format text) \r\n        Message           = \"Successfully created team [$displayName] with description [$description] and owners [$(($owners | ForEach-Object { $_.displayName }) -join \u0027, \u0027)]\" # required (free format text)\r\n        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $displayName # optional (free format text) \r\n        TargetIdentifier  = $([string]$teamId) # optional (free format text) \r\n    }\r\n    #send result back  \r\n    Write-Information -Tags \"Audit\" -MessageData $log    \r\n}\r\ncatch {\r\n    $ex = $PSItem\r\n    if ($($ex.Exception.GetType().FullName -eq \u0027Microsoft.PowerShell.Commands.HttpResponseException\u0027) -or\r\n        $($ex.Exception.GetType().FullName -eq \u0027System.Net.WebException\u0027)) {\r\n        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex\r\n        $auditMessage = \"Error $($actionMessage). Error: $($errorObj.FriendlyMessage)\"\r\n        $warningMessage = \"Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)\"\r\n    }\r\n    else {\r\n        $auditMessage = \"Error $($actionMessage). Error: $($ex.Exception.Message)\"\r\n        $warningMessage = \"Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)\"\r\n    }\r\n    $Log = @{\r\n        Action            = \"CreateResource\" # optional. ENUM (undefined = default) \r\n        System            = \"MicrosoftTeams\" # optional (free format text) \r\n        Message           = $auditMessage # required (free format text) \r\n        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n        TargetDisplayName = $displayName # optional (free format text) \r\n        TargetIdentifier  = $([string]$teamId) # optional (free format text) \r\n    }\r\n    Write-Information -Tags \"Audit\" -MessageData $log\r\n    Write-Warning $warningMessage\r\n    Write-Error $auditMessage\r\n}","runInCloud":true}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-plus-square" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

