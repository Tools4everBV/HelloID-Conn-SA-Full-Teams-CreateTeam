<!-- Description -->
## Description
This HelloID Service Automation Delegated Form can create a Microsoft Teams Team

<!-- Requirements -->
## Requirements
The Powershell Module "MicrosoftTeams" is required on the HelloID Agent Server in order to use this Form
 
<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Description](#description)
* [Requirements](#requirements)
* [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  * [Getting started](#getting-started)
* [Post-setup configuration](#post-setup-configuration)
* [Manual resources](#manual-resources)


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

_Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_

### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.


## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)
<table>
  <tr><td><strong>Variable name</strong></td><td><strong>Example value</strong></td><td><strong>Description</strong></td></tr>
  <tr><td>TeamsAdminUser</td><td>admin@contoso.onmicrosoft.com</td><td>Username of a user with Microsoft Teams Administrator Rights</td></tr>
  <tr><td>TeamsAdminPWD</td><td>Password</td><td>Password of the Admin User</td></tr>
</table>

## Manual resources
This Delegated Form uses the following resources in order to run

### Delegated form task 'Teams-create-team'
