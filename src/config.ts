const config = {
    botId: process.env.BOT_ID,
    botPassword: process.env.BOT_PASSWORD,
    botEndpoint: process.env.BOT_ENDPOINT,
    azureDevOpsProviderConfig: {
        tenantId: process.env.TENANT_ID,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        projectUrl: process.env.PROJECT_URL,
        adoScopes: process.env.ADO_SCOPES,
        apiScopes: process.env.API_SCOPES,
        loginUrl: `${process.env.BOT_ENDPOINT}/web/auth-start.html`, //   process.env.INITIATE_LOGIN_ENDPOINT,
        redirectUrl: `${process.env.BOT_ENDPOINT}/web/auth-end.html` //process.env.REDIRECT_URL
    },
    entraIdAppConfig: {
        // Requried
        identityMetadata: `https://login.microsoftonline.com/${process.env.TENANT_ID}/v2.0/.well-known/openid-configuration`,
        // or 'https://login.microsoftonline.com/<your_tenant_guid>/.well-known/openid-configuration'
        // or you can use the common endpoint
        // 'https://login.microsoftonline.com/common/.well-known/openid-configuration'

        // Required
        clientID: process.env.API_CLIENT_ID,

        // Required.
        // If you are using the common endpoint, you should either set `validateIssuer` to false, or provide a value for `issuer`.
        validateIssuer: false,

        // Required. 
        // Set to true if you use `function(req, token, done)` as the verify callback.
        // Set to false if you use `function(req, token)` as the verify callback.
        passReqToCallback: true,

        isB2C: false,

        policyName: "test_b2c",

        // Required if you are using common endpoint and setting `validateIssuer` to true.
        // For tenant-specific endpoint, this field is optional, we will use the issuer from the metadata by default.
        issuer: null,

        // Optional, default value is clientID
        audience: process.env.API_AUDIENCE,
        //audience: process.env.CLIENTID,

        // Optional. Default value is false.
        // Set to true if you accept access_token whose `aud` claim contains multiple values.
        allowMultiAudiencesInToken: false,

        // Optional. 'error', 'warn' or 'info'
        loggingLevel: 'info',
        loggingNoPII: false
    },
    cosmosDbConfig: {
        CosmosDbEndPoint: process.env.COSMOSDBENDPOINT,
        CosmosDbId: process.env.COSMOSDBID,
        CosmosDbKey: process.env.COSMOSDBKEY,
        CosmosDbContainerId_APIKeys: "APIKeys",
        CosmosDbContainerId_CompletedTasks: "CompletedTasks",
        CosmosDbContainerId_CustomerAreaMappings: "CustomerAreaMappings",
        CosmosDbContainerId_UserSettings: "UserSettings"
    }
};

export default config;
