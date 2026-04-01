# HelloID-Conn-SA-Full-Teams-CreateTeam

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description
_HelloID-Conn-SA-Full-Teams-CreateTeam_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

By using this delegated form , you can create Microsoft Teams teams through Microsoft Graph. The delegated form supports the following flow:
1. Enter team details (display name, description, and privacy)
2. Select one or more owners from Microsoft Entra ID
3. Validate uniqueness of display name, mail address, and mail nickname
4. Create the team in Microsoft 365
5. Add additional selected owners after team creation

## Getting started
### Requirements

- **Microsoft Entra application registration (certificate-based)**:
  The connector authenticates to Microsoft Graph using a certificate (client credentials flow).
- **Microsoft Graph application permissions**:
  Configure and grant admin consent for the following minimal application permissions:
  - `User.Read.All`
  - `Group.Read.All`
  - `Team.Create`
  - `TeamMember.ReadWrite.All`

### Connection settings

The following user-defined variables are used by the connector.

| Setting                        | Description                                                                | Mandatory |
|--------------------------------|----------------------------------------------------------------------------|-----------|
| EntraIdTenantId                | Microsoft Entra tenant ID                                                  | Yes       |
| EntraIdAppId                   | Application (client) ID of the app registration                            | Yes       |
| EntraIdCertificateBase64String | Base64 encoded certificate (including private key) used for authentication | Yes       |
| EntraIdCertificatePassword     | Password for the certificate                                               | Yes       |
| TeamsMailsuffix                | Mail suffix used when building mail address from display name              | Yes       |

## Remarks

### Microsoft Graph Query Behavior
- `ConsistencyLevel: eventual` is added on Graph requests where advanced query capabilities are used (for example filtering and counting large result sets). This allows Microsoft Graph to evaluate those queries correctly and consistently, especially when data has just changed.

### Team Provisioning Task
- Team creation starts with a standard template and includes the first selected owner in the initial create request.
- Provisioning is asynchronous; the task follows the Graph operation endpoint from the `Location` header and polls until completion.
- If multiple owners are selected, additional owners are added after successful provisioning through the team members add endpoint.

## Development resources

### API endpoints

The following endpoints are used by the connector.

| Endpoint                                                      | Description                                                             |
|---------------------------------------------------------------|-------------------------------------------------------------------------|
| `https://login.microsoftonline.com/{tenantId}/oauth2/token`   | Retrieve OAuth2 access token using certificate-based client credentials |
| `https://graph.microsoft.com/v1.0/users`                      | Retrieve enabled Entra ID users for owner selection                     |
| `https://graph.microsoft.com/v1.0/groups`                     | Validate uniqueness of display name, mail, and mail nickname            |
| `https://graph.microsoft.com/v1.0/teams`                      | Create Microsoft Teams team                                             |
| `https://graph.microsoft.com/v1.0/teams/{teamId}/members/add` | Add additional owners to the created team                               |

### API documentation

- https://learn.microsoft.com/graph/api/overview
- https://learn.microsoft.com/graph/api/team-post
- https://learn.microsoft.com/graph/api/user-list
- https://learn.microsoft.com/graph/api/group-list

## Getting help
> :bulb: **Tip:**  
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/