# CDE-Helper

## Description
This is application provided CDE search capabilities for the following applications: 
* Outlook add-ins
* MS Teams message extension
* M365 Copilot plugins. 

## side the project
* run the following cmd to side the outlook add-ins. this will output a title_id which will be used to unload add-ins later
  ```bash
  npm run outlook:localstart
  ```
* run the following cmd to unload the add-ins
  ```bash
  npm run outlook:localstop
  ```
* if for some reason, you mix the deployment with another add-ins development, you might need to run the below cmd to retrieve title_id and delete it separately: 
  ```bash
  # login to M365 using Teamsfx CLI
  teamsfx login
  # get title_id via using manifest-id which is same as "id" in manifest file or "TEAMS_APP_ID" in .env.local or .env.dev files. 
  teamsfx m365 launchinfo --manifest-id ab1c0b65-5df9-4df5-9721-2516daafc282
  # using the title_id get from previous cmd
  teamsfx m365 unacquire --title-id [title-id(usually start with "U_")]
  ```

## Customization from template: 
### Update to ngrok tunnel
* tasks.json: 
  * add the following code to start ngrok
  ```JSON
         {
            // Start ngrok local. NOTE: need to install ngrok NPM package first
            "label": "ngrok:start",
            "type": "shell",
            "command": "npx ngrok http 3978 --subdomain=ezcode",
            "isBackground": true,
            "problemMatcher": {
                "pattern": [
                    {
                        "regexp": "^.*$",
                        "file": 0,
                        "location": 1,
                        "message": 2
                    }
                ],
                "background": {
                    "activeOnStart": true,
                    "beginsPattern": "Take our ngrok in production survey!",
                    "endsPattern": "Connections"
                }
            },
        },
  ```
  * commented out "Start local tunnel" in task "Start Teams App Locally" because it doesn't need to start local dev tunnel. 
  ```JSON
  {
            "label": "Start Teams App Locally",
            "dependsOn": [
                "Validate prerequisites",
                // "Start local tunnel",
                "ngrok:start",
                "Provision",
                "Deploy",
                "Start application"
            ],
            "dependsOrder": "sequence"
        }
  ```
* .env.local: 
  * updated ```BOT_DOMAIN``` to use ngrok custom domain
  ```
  BOT_DOMAIN=ezcode.ngrok.io
  ``` 
* teamsapp.local.yml: 
  * update ```-uses botFramework/create``` to include
  ```
  messagingEndpoint: https://${{BOT_DOMAIN}}/api/messages
  ```
  * update ```-uses: file/createOrUpdateEnvironmentFile``` to include the following environment variables. 
  ```
        INITIATE_LOGIN_ENDPOINT: ${{BOT_ENDPOINT}}/auth-start.html
        AZUREDEVOPS_CLIENTID: DB4D0B62-824D-4879-B24B-11B9A94AC664
        AZUREDEVOPS_AUTHURL: https://app.vssps.visualstudio.com/oauth2/authorize
        AZUREDEVOPS_TOKENURL: https://app.vssps.visualstudio.com/oauth2/token
        AZUREDEVOPS_PROJECTURL: https://dev.azure.com/O365DSE/POD%208
        AZUREDEVOPS_REDIRECTURL: https://ezcode.ngrok.io/auth-end.html
        AZUREDEVOPS_CLIENTSECRET: [secret]

  ```
* you need to run ngrok separately, otherwise, you need to update tasks.json to automatically run it. ```ngrok http 3978 --subdomain=ezcode```

## Deployment
### Cosmos DB copy. 
run 
```
C:\Tools\dmt\windows-package\dmt.exe --settings settings-CompletedTasks.json
```

### Run from package deployment
MS Teams toolkit is using [Run your app in Azure App Service directly from a ZIP package](https://learn.microsoft.com/en-us/azure/app-service/deploy-run-package) which will provisioning a configuration ```WEBSITE_RUN_FROM_PACKAGE="1"```. After deployment, you won't be able to find the deployment from site/wwwroot folder. the Zip file is located under home\data\SitePackages and will be mount automatically. 
```
# generate zip file
#./createzip.ps1
# run 7zip to zip package.json, package-lock.json, dist and node_modules folders

# run below command to make sure you are using right subscription
az account show
az webapp deploy --resource-group CDEHelperRG --name cdehelper-web --src-path deployment\deployment_2024090601.zip
```

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 16, 18
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Set up your dev environment for extending Teams apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
>   Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
> - Join Microsoft 365 Copilot Plugin development [early access program](https://aka.ms/plugins-dev-waitlist).

> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
3. To directly trigger the Message Extension in Teams App Test Tool, you can:
   1. Press F5 to start debugging which launches your app in Teams App Test Tool using a web browser. Select `Debug in Test Tool`.
   2. When Test Tool launches in the browser, click the `+` in compose message area and select `Search command` to trigger the search commands.
3. To directly trigger the Message Extension in Teams, you can:
   1. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
   2. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
   3. `@mention` Your message extension from the `search box area`, `@mention` your message extension from the `compose message area` or click the `...` under compose message area to find your message extension.
4. To trigger the Message Extension through Copilot, you can:
   1. Select `Debug in Copilot (Edge)` or `Debug in Copilot (Chrome)` from the launch configuration dropdown.
   2. When Teams launches in the browser, click the `Apps` icon from Teams client left rail to open Teams app store and search for `Copilot`.
   3. Open the `Copilot` app, select `Plugins`, and from the list of plugins, turn on the toggle for your message extension. Now, you can send a prompt to trigger your plugin.
   4. Send a message to Copilot to find an NPM package information. For example: `Find the npm package info on teamsfx-react`.
      > Note: This prompt may not always make Copilot include a response from your message extension. If it happens, try some other prompts or leave a feedback to us by thumbing down the Copilot response and leave a message tagged with [MessageExtension].

**Congratulations**! You are running an application that can now search npm registries in Teams and Copilot.

![Search ME Copilot](https://github.com/OfficeDev/TeamsFx/assets/107838226/0beaa86e-d446-4ab3-a701-eec205d1b367)

## What's included in the template

| Folder        | Contents                                     |
| ------------- | -------------------------------------------- |
| `.vscode/`    | VSCode files for debugging                   |
| `appPackage/` | Templates for the Teams application manifest |
| `env/`        | Environment files                            |
| `infra/`      | Templates for provisioning Azure resources   |
| `src/`        | The source code for the search application   |

The following files can be customized and demonstrate an example implementation to get you started.

| File               | Contents                                                                                       |
| ------------------ | ---------------------------------------------------------------------------------------------- |
| `src/searchApp.ts` | Handles the business logic for this app template to query npm registry and return result list. |
| `src/index.ts`     | `index.ts` is used to setup and configure the Message Extension.                               |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                  |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions. |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                     |
| `teamsapp.testtool.yml`| This overrides `teamsapp.yml` with actions that enable local execution and debugging in Teams App Test Tool.                            |

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Teams Toolkit CLI](https://aka.ms/teams-toolkit-cli/debug)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
- [Extend Microsoft 365 Copilot](https://aka.ms/teamsfx-copilot-plugin)

## change logs: 
* 1.0.3: 
  * Moved to separate tenant for hosting
  * fixed office ows token retriving issues. introduced backend process to renew the token to avoid long time. 
  * enhanced the action dialog for task item