// Data-plane RBAC for the Function App's system-assigned Managed Identity.
//
// Replaces two shared-key usages:
//   * the audit log previously authenticated with a storage account key that
//     was stored in plaintext app settings, which also let anyone able to read
//     those settings rewrite the audit trail;
//   * WEBSITE_RUN_FROM_PACKAGE previously used a two-year account-key SAS that
//     could only be revoked by rotating the account key.

@description('Name of the Storage Account holding the audit table and deployment packages')
param storageAccountName string

@description('Principal ID of the Function App system-assigned Managed Identity')
param principalId string

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-04-01' existing = {
  name: storageAccountName
}

// Built-in role definition IDs (constant across all clouds)
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var storageBlobDataReaderRoleId       = '2a2b9908-6ea1-4ae2-8e65-a410df84e7d1'

// Write audit records to the LapsAuditLog table
resource tableDataContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, principalId, storageTableDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageTableDataContributorRoleId)
    principalId: principalId
    principalType: 'ServicePrincipal'
  }
}

// Read the deployment package referenced by WEBSITE_RUN_FROM_PACKAGE
resource blobDataReader 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, principalId, storageBlobDataReaderRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageBlobDataReaderRoleId)
    principalId: principalId
    principalType: 'ServicePrincipal'
  }
}
