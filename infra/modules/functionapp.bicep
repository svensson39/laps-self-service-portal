@description('Project name used as prefix for resource names')
param projectName string

@description('Azure region')
param location string

@description('Resource ID of the App Service Plan (Linux B1)')
param appServicePlanId string

@description('Name of the Storage Account used for AzureWebJobsStorage and audit logging')
param storageAccountName string

@description('Name of the Azure Table used for audit logging')
param auditTableName string

@description('Application Insights connection string')
param appInsightsConnectionString string

@description('Entra ID Tenant ID (written to TENANT_ID app setting)')
param tenantId string

@description('Client ID of the App Registration (for Easy Auth audience validation)')
param authClientId string

@description('Allowed CORS origins – the Static Web App URL and any custom domain')
param allowedOrigins array

@description('Resource tags')
param tags object

// ── Existing Storage Account reference ───────────────────────────────────────
// AzureWebJobsStorage still requires a shared-key connection string (Functions
// runtime limitation). The audit log does NOT – it authenticates with the
// Managed Identity so the app cannot be handed a key that also lets it rewrite
// its own audit trail. See modules/roleAssignments.bicep.

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-04-01' existing = {
  name: storageAccountName
}

var storageConnectionString = 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};AccountKey=${storageAccount.listKeys().keys[0].value};EndpointSuffix=${az.environment().suffixes.storage}'

// ── Function App ──────────────────────────────────────────────────────────────

resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: '${projectName}-func'
  location: location
  tags: tags
  kind: 'functionapp,linux'
  identity: {
    // System-assigned Managed Identity – used to call Microsoft Graph and Table
    // Storage without stored secrets
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: appServicePlanId
    reserved: true      // Required for Linux
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'Node|24'
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      cors: {
        allowedOrigins: allowedOrigins
        // The API authenticates with a Bearer header only – it never relies on
        // cookies, so credentialed cross-origin requests are not needed.
        supportCredentials: false
      }
      appSettings: [
        // ── Azure Functions runtime ─────────────────────────────────────────
        {
          name: 'AzureWebJobsStorage'
          value: storageConnectionString
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~24'
        }

        // ── Monitoring ──────────────────────────────────────────────────────
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsightsConnectionString
        }
        {
          name: 'ApplicationInsightsAgent_EXTENSION_VERSION'
          value: '~3'
        }

        // ── Azure SDK logging ───────────────────────────────────────────────
        // Suppress known deprecation warnings from auto-generated Azure SDK
        // clients (e.g. @azure/data-tables internally uses 'baseUri' instead
        // of 'endpoint'). Only errors are logged – warnings/info/debug are off.
        {
          name: 'AZURE_LOG_LEVEL'
          value: 'error'
        }

        // ── Application settings ────────────────────────────────────────────
        {
          name: 'TENANT_ID'
          value: tenantId
        }
        {
          // Used by lib/auth.js for JWT verification (audience + issuer check)
          name: 'AUTH_CLIENT_ID'
          value: authClientId
        }
        {
          // Declares that Easy Auth (authsettingsV2 below) sits in front of this
          // app and therefore strips any client-supplied X-MS-CLIENT-PRINCIPAL
          // header. lib/auth.js refuses to trust that unsigned header unless
          // this is 'true'. Never set it on a deployment without Easy Auth.
          name: 'EASY_AUTH_ENABLED'
          value: 'true'
        }
        {
          // OAuth2 scope the caller's token must carry
          name: 'REQUIRED_SCOPE'
          value: 'access_as_user'
        }
        {
          // Audit log authenticates with the Managed Identity – no account key
          name: 'AUDIT_STORAGE_ACCOUNT_NAME'
          value: storageAccountName
        }
        {
          name: 'AUDIT_TABLE_NAME'
          value: auditTableName
        }
        {
          name: 'GRAPH_API_ENDPOINT'
          value: 'https://graph.microsoft.com'
        }
        {
          name: 'JUSTIFICATION_MIN_LENGTH'
          value: '10'
        }
        {
          // Server-side cap. Without it an oversized justification exceeds the
          // Azure Table entity limit, the audit write fails, and the caller
          // still gets the password.
          name: 'JUSTIFICATION_MAX_LENGTH'
          value: '500'
        }
        {
          name: 'PASSWORD_DISPLAY_SECONDS'
          value: '60'
        }
        {
          // Maximum age of the interactive sign-in (Entra auth_time claim).
          // MSAL silent renewals keep the original auth_time, so this caps the
          // real session length, not the token lifetime. 0 = disabled.
          name: 'SESSION_MAX_AGE_MINUTES'
          value: '480'
        }
      ]
    }
  }
}

// ── Easy Auth (Entra ID built-in authentication) ──────────────────────────────
// Validates JWT tokens before requests reach the function code.
// Unauthenticated requests return HTTP 401 without reaching any function code.
// lib/auth.js re-verifies the forwarded Bearer token independently – this is
// defence in depth, not the only check.

resource authSettings 'Microsoft.Web/sites/config@2023-01-01' = {
  parent: functionApp
  name: 'authsettingsV2'
  properties: {
    globalValidation: {
      requireAuthentication: true
      unauthenticatedClientAction: 'Return401'
    }
    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        registration: {
          clientId: authClientId
          openIdIssuer: 'https://sts.windows.net/${tenantId}/v2.0'
        }
        validation: {
          // Only the API audience. The bare client ID is the audience of ID
          // tokens, which must not be accepted as bearer credentials here.
          allowedAudiences: [
            'api://${authClientId}'
          ]
        }
      }
    }
    login: {
      tokenStore: {
        // Do not persist tokens server-side – stateless auth
        enabled: false
      }
    }
  }
}

// ── Outputs ───────────────────────────────────────────────────────────────────

output functionAppName string = functionApp.name
output functionAppId string = functionApp.id

@description('HTTPS URL of the Function App (base URL for all API calls)')
output url string = 'https://${functionApp.properties.defaultHostName}'

@description('Object ID of the system-assigned Managed Identity – assign Graph permissions to this ID')
output principalId string = functionApp.identity.principalId
