const config = {
    botId: process.env.BOT_ID,
    botPassword: process.env.BOT_PASSWORD,
    botEndpoint: process.env.BOT_ENDPOINT,
    azureDevOpsProviderConfig: {
        tenantId: process.env.TENANT_ID,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        projectUrl: process.env.PROJECT_URL,
        scopes: process.env.SCOPES,
        loginUrl: `${process.env.BOT_ENDPOINT}/web/auth-start.html`, //   process.env.INITIATE_LOGIN_ENDPOINT,
        redirectUrl: `${process.env.BOT_ENDPOINT}/web/auth-end.html` //process.env.REDIRECT_URL
    },
    cosmosDbConfig: {
        CosmosDbEndPoint: process.env.COSMOSDBENDPOINT,
        CosmosDbId: process.env.COSMOSDBID,
        CosmosDbKey: process.env.COSMOSDBKEY,
        CosmosDbContainerId_APIKeys: "APIKeys",
        CosmosDbContainerId_CompletedTasks: "CompletedTasks",
        CosmosDbContainerId_CustomerAreaMappings: "CustomerAreaMappings",
    }
};

export default config;
