---
Graph API Used: https://graph.microsoft.com/v1.0
Total Rest Options: Post,Get
---

# Generate Teams

This repository contains the code used for azure functions allowing a single Rest Post request to generate a new Microsoft Teams team based off of the information provided. This was originally created for use with TDx Web Request features to automate our workflows when creating Microsoft Teams

## Table Of Contents

- [Generate Teams](#generate-teams)
  - [Table Of Contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Setup](#setup)
    - [Hardcoded Script Values](#hardcoded-script-values)
    - [TDx configuration](#tdx-configuration)
      - [TDx User account permission](#tdx-user-account-permission)
      - [Web Service Auth Account](#web-service-auth-account)
      - [Web Service Provider](#web-service-provider)
  - [Workflow and Usage](#workflow-and-usage)

## Requirements

- Azure App Registration
  - Application (client) ID
  - Client Secret
  - Graph Permissions:
    | Graph API Permission       | Type        | Description                         | Admin Consent Required |
    | -------------------------- | ----------- | ----------------------------------- | ---------------------- |
    | Directory.ReadWrite.All    | Application | Read and write directory data       | Yes                    |
    | Group.ReadWrite.All        | Application | Read and write all groups           | Yes                    |
    | Team.Create                | Application | Create teams                        | Yes                    |
    | Team.ReadBasic.All         | Application | Get a list of all teams             | Yes                    |
    | TeamSettings.ReadWrite.All | Application | Read and change all teams' settings | Yes                    |
- Azure Resource Group
  - Technically optional, used to store and organize all of these resources mentioned.
- Azure App Service
- Azure Storage Account
  - Azure Storage Queue
  - Azure Storage Table
  - >NOTE: Account name and key are not directly used in this script, instead the `function.json` files use the default "AzureWebJobsStorage" connection which provides this information to the functions.

## Setup

- Generate the Azure App Registration
  - Set the proper Graph permissions (listed in the requirement section)
  - Make note of the Application (client) ID
    - Generate a new client secret (description does not matter)
    - > NOTE: Make note of the new client secret, you will not be able to view it later, if lost, a new client secret needs to be generated
  - > NOTE: The name of the App Registration will be visible when the team is created and the requestor is invited, Teams uses the name of the app registration for the notification and invite displayed in Microsoft Teams
- Create a new Azure Resource Group, this will be used to "store" all of the additional components
- Create an azure storage account
  - Create a new Azure Storage Queue and Table
    - Make note of both the Queue and Table name, they will be needed later
- Create a new App Service
  - The code from this Repo can be cloned down from git or use an SFTP transfer to the app service.
  - >NOTE: Review the `local.settings.json.template` "values" section for a list of attributes that will need to exist in the "application settings and configuration" in the app service (set these using the Azure Portal GUI).
- Review the hardcoded values section below and ensure that those entries are updated to match the current Azure storage.

### Hardcoded Script Values

There are a few hardcoded values that need to be set that are based on the configuration of the Azure Storage:

- `CheckCallbackID\run.ps1`
  - PartitionKey
  - AzureStorageTableName
- `HttpTrigger\function.json`
  - QueueName
- `QueueTrigger\function.json`
  - TableName
  - QueueName

### TDx configuration

#### TDx User account permission

  | Application   | Security Role                                      |
  | ------------- | -------------------------------------------------- |
  | Chat          | ✔                                                  |
  | Client Portal | Client + Knowledge Base, Services, Ticket Requests |
  | Community     | ✔                                                  |
  | IT            | Technician                                         |
  | TDNext        | ✔                                                  |

#### Web Service Auth Account

  | Name                       | type                | active |
  | -------------------------- | ------------------- | ------ |
  | Existing user account Name | TeamDynamix Web API | ✔      |

#### Web Service Provider

  | Name                | Base Service Provider URL                   | Active |
  | ------------------- | ------------------------------------------- | ------ |
  | Azure Teams Creator | [https://AzureAppName.azurewebsites.net/api/](https://AzureAppName.azurewebsites.net/api/) | ✔      |

## Workflow and Usage

- >NOTE: Review the `sample.dat` files for a REST example of how to trigger the respective function.
- `HTTPTrigger` function is triggered via a post request.
  - This returns a "CallbackID" which can later be used to query the status of the team.
  - The request is then stored within an azure storage queue.
- The `QueueTrigger` function is triggered as soon as a new item is created in the queue, the function then processes the new request and makes a Graph request to create a group.
  - A new O365 Group is generated and populated with a single member (the owner of the group).
  - The group is then used to create a Microsoft Team via Microsoft Graph.
    - >NOTE: Currently this is best practice, Graph API calls newer 1.* may have a single Graph Request to create a team rather than a two part process.
  - The queue trigger takes a mixture of the Group and Team attributes and posts the results to an azure table (this is both used for checking the status of the new team and as long term storage logs for the requests).
A new Rest Get request is made (on demand by an application, like TDx) to the `CheckCallbackID` Function, this function takes the provided callback ID and checks the Azure storage table, Returning a 400 response if not found, and a 200 response with the team information if found. This status can be used in TDx to automatically move through a predefined workflow.
- Shortly after the new team is created, the owner (and one and only member) will be granted access to the team and receive a notification if the teams client is running.
