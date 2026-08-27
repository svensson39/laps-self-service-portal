# Architecture

## Overview

The LAPS Self-Service Portal is a two-tier web application hosted entirely in Azure:

| Tier | Technology | Azure Service |
|------|------------|---------------|
| Frontend | Static HTML + MSAL.js + Vanilla JS | Azure Static Web App |
| Backend | Node.js 24 Azure Functions v4 | Azure Function App (Dedicated Linux B1) |
| Identity | Entra ID | Built-in Easy Auth |
| Graph access | Managed Identity | System-assigned to Function App |
| Audit storage | Azure Table Storage | Included in Function App storage account |
| Monitoring | Application Insights | Linked to Log Analytics Workspace |

---

## Authentication Flow

```
Browser                    Entra ID                 Function App          Microsoft Graph
  │                            │                          │                      │
  │── Open portal ────────────▶│                          │                      │
  │                            │                          │                      │
  │◀── Redirect to login ──────│                          │                      │
  │                            │                          │                      │
  │── Credentials ────────────▶│                          │                      │
  │                            │                          │                      │
  │◀── ID token + access token─│                          │                      │
  │                            │                          │                      │
  │── GET /api/my-devices ───────────────────────────────▶│                      │
  │   Authorization: Bearer <token>                        │                      │
  │                            │                          │                      │
  │                            │◀── Easy Auth validates ──│                      │
  │                            │    token, injects header │                      │
  │                            │                          │                      │
  │                            │                          │── GET /users/{oid}/registeredDevices ────▶│
  │                            │                          │   (Managed Identity token)                │
  │                            │                          │◀── Device list (Windows + macOS only) ────│
  │                            │                          │                      │
  │◀── 200 { devices: [...] } ─────────────────────────────│                      │
  │                            │                          │                      │
  │── POST /api/laps-password ───────────────────────────▶│                      │
  │   { deviceId, justification }                         │                      │
  │                            │                          │                      │
  │                            │                          │── verify ownership ──▶│
  │                            │                          │── GET /v1.0/directory/deviceLocalCredentials/{deviceId} ─▶│
  │                            │                          │◀── password ──────────────────────────────│
  │                            │                          │                      │
  │                            │                          │── write audit log ──▶ Table Storage
  │                            │                          │                      │
  │◀── 200 { password, ... } ──────────────────────────────│                      │
```

---

## Security Model

### Token Validation

The Function App's built-in authentication (Easy Auth) validates the Bearer token on every
request before the function code runs. Unauthenticated requests receive HTTP 401.

Easy Auth is defence in depth, not the only check. `lib/auth.js` independently verifies the
forwarded Bearer token against Entra ID's JWKS (signature, RS256, audience `api://<clientId>`,
issuer, tenant and the `access_as_user` scope). The unsigned `X-MS-CLIENT-PRINCIPAL` header is
only trusted when no Bearer token is present *and* the app setting `EASY_AUTH_ENABLED` is
`true` - that flag is the deployment's explicit statement that the platform strips any
client-supplied copy of the header. Without it, a forged header is rejected rather than
believed.

### "Only My Device" Rule

Device ownership is enforced **in the backend** on every request, not just in the UI:

1. The backend reads the user's Object ID (OID) from the verified access token
2. It queries Graph for all registered devices of that user (`GET /users/{oid}/registeredDevices`)
3. The requested device object ID must appear in that list - otherwise HTTP 403 is returned
4. The LAPS credential is then fetched using that same device's Entra `deviceId`. The
   authorization decision and the data fetch are keyed on the same identifier; looking the
   credential up by display name would let a device with a colliding name resolve to somebody
   else's credential, since display names are not unique in Entra ID.

The frontend device list is a UX convenience only; it does not constitute an authorization boundary.

### Managed Identity

The Function App uses a system-assigned Managed Identity to authenticate against Microsoft Graph.
No client secrets, certificates, or connection strings are stored for Graph access.

Token acquisition flow:

```
Function App
  └── @azure/identity DefaultAzureCredential
        └── ManagedIdentityCredential
              └── Azure IMDS endpoint (http://169.254.169.254, internal)
                    └── Entra ID token endpoint
                          └── access_token for https://graph.microsoft.com
```

Required application permissions on the Managed Identity:

| Permission | Purpose |
|-----------|---------|
| `Device.Read.All` | Read device properties |
| `DeviceLocalCredential.Read.All` | Read LAPS passwords |

### Audit Trail

Every invocation of `POST /api/laps-password` writes a record to Azure Table Storage before
returning a response, regardless of outcome (success, denial, or error).

| Field | Value |
|-------|-------|
| PartitionKey | `YYYY-MM-DD` (UTC date) |
| RowKey | UUID v4 |
| UserId | Entra Object ID |
| UserPrincipalName | UPN |
| DeviceId | Entra Device Object ID |
| DeviceName | Display name |
| Justification | User-provided reason |
| Action | `SUCCESS` / `DENIED` / `ERROR` |
| DenialReason | e.g. `DEVICE_NOT_OWNED`, `NO_LAPS_CREDENTIAL` |
| ClientIp | Source IP address |
| UserAgent | Browser user agent |

---

## Resource Topology

```
Subscription
└── Resource Group: rg-<projectName>
    ├── Storage Account: <projectName-prefix><uniquehash>
    │   ├── Blob container: func-deployments  (WEBSITE_RUN_FROM_PACKAGE zip)
    │   ├── Blob containers: azure-webjobs-*  (Functions runtime internal)
    │   └── Table: LapsAuditLog
    │
    ├── App Service Plan: <projectName>-plan  (Linux B1 Dedicated)
    │
    ├── Function App: <projectName>-func
    │   ├── System-assigned Managed Identity
    │   └── Graph permissions: Device.Read.All, DeviceLocalCredential.Read.All
    │   ├── Easy Auth → Entra ID (validates JWT before code runs)
    │   ├── WEBSITE_RUN_FROM_PACKAGE → Blob Storage URL (read via Managed Identity)
    │   └── App Settings (TENANT_ID, AUTH_CLIENT_ID, AUDIT_*, GRAPH_API_ENDPOINT, …)
    │
    ├── Static Web App: <projectName>-swa  (Standard tier, westeurope by default)
    │   └── Custom Domain (optional)
    │
    ├── Log Analytics Workspace: <projectName>-law
    │
    └── Application Insights: <projectName>-ai
```

> **Note:** The Static Web App must be deployed to one of the five supported regions:
> `westus2`, `centralus`, `eastus2`, `westeurope`, `eastasia`. All other resources
> can use any Azure region (controlled by the `--location` parameter).

---

## Data Flow – LAPS Password Retrieval

```
POST /api/laps-password
{ "deviceId": "...", "justification": "..." }

Step 1 - Easy Auth validates the Bearer token (audience: api://<clientId>)
         and forwards it, plus X-MS-CLIENT-PRINCIPAL, to the function

Step 2 - getCallerIdentity() re-verifies the Bearer token against JWKS
         (RS256, aud, iss, tid, scp) and returns OID + UPN. If a principal
         header is also present its OID must match the verified token.

Step 3 - Input validation
         -> deviceId present?
         -> JUSTIFICATION_MIN_LENGTH <= justification <= JUSTIFICATION_MAX_LENGTH?

Step 4 – findOwnedDevice(deviceId, oid)
         → Graph: GET /v1.0/users/{oid}/registeredDevices
         → Is deviceId in the result set?
         → No → write DENIED audit log → return HTTP 403

Step 5 - getLapsPassword(device.deviceId, device.name)
         -> Graph: GET /v1.0/directory/deviceLocalCredentials/{deviceId}
                       ?$select=credentials,deviceName,refreshDateTime
         -> Returned deviceName must match the authorized device -> else HTTP 403
         -> No credential -> write DENIED audit log -> return HTTP 404

Step 6 - writeAuditLog(action: 'SUCCESS')
         -> Fail-closed: if the record cannot be persisted the password is
            withheld and HTTP 503 is returned. Otherwise a caller could
            suppress their own audit trail and still get the credential.

Step 7 – Return { deviceName, password, expiresAt, auditId }
```

---

## Backend Deployment

The Function App is deployed using the `WEBSITE_RUN_FROM_PACKAGE` pattern:

1. Backend source is zipped locally
2. Zip is uploaded to the project's Storage Account (`func-deployments` container)
3. The blob URL (no SAS) is written to `WEBSITE_RUN_FROM_PACKAGE` via the ARM REST API,
   together with `WEBSITE_RUN_FROM_PACKAGE_BLOB_MI_RESOURCE_ID=SystemAssigned`
4. The platform reads the package with the Function App's Managed Identity, which holds
   *Storage Blob Data Reader* on the account. No SAS token and no account key ever land in
   app settings, and access is revoked by removing the role assignment.
5. Azure Functions runtime mounts the zip read-only and runs from it

This avoids the Kudu SCM endpoint (which is unreliable for Linux Dedicated plans) and
gives deterministic, fast deployments.

---

## Scalability & Cost

The portal runs on a **Dedicated B1** App Service Plan to avoid cold-start delays that
are common with the Consumption plan. Expected cost breakdown for a typical 500-user organization:

| Resource | Estimated monthly cost |
|----------|----------------------|
| App Service Plan (Linux B1) | ~€12 |
| Static Web App (Standard) | ~€9 |
| Storage Account | < €1 |
| Application Insights | < €2 (first 5 GB/month free) |
| Log Analytics Workspace | < €1 |
| **Total** | **~€25/month** |

> To reduce costs, the App Service Plan can be changed to `B1` → `Y1` (Consumption) in
> `infra/modules/appServicePlan.bicep`. This trades cold-start latency (~3–5 s on first
> request after idle) for near-zero compute cost.
