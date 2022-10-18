# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
        Write-Information "Generating Microsoft Graph API Access Token user.."

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
        
        #Add the authorization header to the request
        $authorization = @{
            Authorization = "Bearer $accesstoken";
            'Content-Type' = "application/json";
            Accept = "application/json";
        }

        $propertiesToSelect = @(
            "UserPrincipalName",
            "GivenName",
            "Surname",
            "EmployeeId",
            "AccountEnabled",
            "DisplayName",
            "OfficeLocation",
            "Department",
            "JobTitle",
            "Mail",
            "MailNickName",
            "Id",
            "assignedLicenses",
            "assignedPlans"
        )
 
        $usersUri = "https://graph.microsoft.com/v1.0/users"                
        $usersUri = $usersUri + ('?$select=' + ($propertiesToSelect -join "," | Out-String))
        
        $data = @()
        $query = Invoke-RestMethod -Method Get -Uri $usersUri -Headers $authorization -ContentType 'application/x-www-form-urlencoded'
        $data += $query.value
        
        while($query.'@odata.nextLink' -ne $null){
            $query = Invoke-RestMethod -Method Get -Uri $query.'@odata.nextLink' -Headers $authorization -ContentType 'application/x-www-form-urlencoded'
            $data += $query.value 
        }
        
        $users = $data | Sort-Object -Property Name
        $resultCount = @($users).Count
        Write-Information "Result count: $resultCount"        
          
        if($resultCount -gt 0){
            foreach($user in $users){   
                if($user.assignedLicenses.count -gt 0) {
                    if($user.UserPrincipalName -ne $TeamsAdminUser)
                    {                        
                        $returnObject = @{Id=$user.id; User=$user.UserPrincipalName; DisplayName=$user.displayName }
                        Write-Output $returnObject
                    }
                }
            }
        } else {
            return
        }
    }
 catch {
    $errorDetailsMessage = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message
    Write-Error ("Error searching for AzureAD groups. Error: $($_.Exception.Message)" + $errorDetailsMessage)
     
    return
}
