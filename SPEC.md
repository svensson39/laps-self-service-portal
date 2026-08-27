# SPEC – LAPS Self-Service Portal

> Full Product Specification
> Version: 1.0
> Date: 2026-02-26
> Author: Daniel Fraubaum / base-IT GmbH

---

## 1. Objective

The LAPS Self-Service Portal allows end users to retrieve the Windows LAPS (Local Administrator Password Solution) password for **their own Intune-managed device** without assistance from IT. No helpdesk ticket is required.

### Business Requirements

| Requirement | Description |
|-------------|-------------|
| Self-service | User retrieves LAPS password independently |
| Security | Access limited to own devices only (backend-enforced) |
| Accountability | Every retrieval logged with reason and timestamp |
| Zero-secret backend | No stored credentials – Managed Identity only |
| Deployment automation | Fully deployable via Bicep |
| Customizability | CSS theming for corporate branding |

---

## 2. Architecture

### 2.1 Components

```
┌─────────────────────────────────────────────────────────────────┐
│                        Azure Tenant                              │
│                                                                  │
│  ┌─────────────────────┐     ┌───────────────────────────────┐  │
│  │  Azure Static        │     │  Azure Function App           │  │
│  │  Web App (Frontend)  │────▶│  (Backend, Node.js)           │  │
│  │                      │     │  + Easy Auth (Entra ID)       │  │
│  │  - index.html        │     │  + Managed Identity           │  │
│  │  - MSAL.js           │     │                               │  │
│  │  - CSS Theming       │     │  Functions:                   │  │
│  └─────────────────────┘     │  - GET /api/my-devices         │  │
│           │                  │  - POST /api/laps-password     │  │
│           │ MSAL OIDC        └───────────────┬───────────────┘  │
│           ▼                                  │                  │
│  ┌─────────────────────┐                     │ Managed Identity │
│  │  Entra ID           │                     │ (App Permission) │
│  │  (Azure AD)         │◀────────────────────┤                  │
│  │                     │                     ▼                  │
│  │  App Registration   │     ┌───────────────────────────────┐  │
│  │  (Frontend SPA)     │     │  Microsoft Graph API          │  │
│  └─────────────────────┘     │                               │  │
│                              │  - deviceManagement/          │  │
│  ┌─────────────────────┐     │    managedDevices             │  │
│  │  Azure Table Storage│◀────│  - deviceLocalCredentials     │  │
│  │  (Audit Log)        │     └───────────────────────────────┘  │
│  └─────────────────────┘                                        │
│  ┌─────────────────────┐                                        │
│  │  Application        │                                        │
│  │  Insights           │                                        │
│  └─────────────────────┘                                        │
└─────────────────────────────────────────────────────────────────┘
```

### 2.2 Data Flow

1. User opens the portal (Azure Static Web App)
2. MSAL.js redirects to the Entra ID sign-in page
3. After successful sign-in, the browser receives an ID token and access token
4. Frontend calls `GET /api/my-devices` with the Bearer token
5. Backend (Easy Auth) validates the token and extracts the user identity
6. Backend queries Microsoft Graph for the user's Intune-managed devices
7. Frontend displays the device list
8. User selects their device and enters a justification
9. Frontend calls `POST /api/laps-password`
10. Backend re-verifies device ownership (backend-enforced "only my device" rule)
11. Backend retrieves the LAPS password via Microsoft Graph
12. An audit log entry is written to Azure Table Storage
13. The password is returned to the browser (session only, never cached)

---

## 3. Authentication & Authorization

### 3.1 Frontend (MSAL.js)

- **Library**: `@azure/msal-browser` (CDN)
- **Flow**: Authorization Code Flow with PKCE
- **Scopes**: `openid`, `profile`, `User.Read`
- **Token storage**: `sessionStorage` (no persistent cache)
- **Silent refresh**: Automatic via MSAL

### 3.2 Backend (Easy Auth + Managed Identity)

- **Easy Auth**: Azure Function App is configured with Entra ID authentication
  - All unauthenticated requests are rejected with HTTP 401
  - User information is forwarded as HTTP headers (`X-MS-CLIENT-PRINCIPAL`)
  - The backend does **not** rely on this alone: `lib/auth.js` re-verifies the forwarded
    Bearer token against Entra ID's JWKS (RS256, audience `api://<clientId>`, issuer, tenant
    and the `access_as_user` scope). The unsigned principal header is only trusted when no
    Bearer token is present *and* `EASY_AUTH_ENABLED=true` declares that the platform strips
    client-supplied copies of it.
- **Managed Identity**: System-assigned Managed Identity of the Function App
  - Communicates with Microsoft Graph without stored credentials
  - Token acquisition via `@azure/identity` (ManagedIdentityCredential)

### 3.3 Graph API Permissions (Application Permissions)

| Permission | Purpose |
|------------|---------|
| `Device.Read.All` | Read device metadata and verify ownership (`/users/{oid}/registeredDevices`) |
| `DeviceLocalCredential.Read.All` | Retrieve LAPS passwords |

> **Note**: These Application Permissions must be assigned to the Managed Identity via the
> Azure CLI (`az rest`). The deployment scripts handle this automatically (Step 3 of the
> manual guide). Bicep does not support direct Graph permission assignment.

---

## 4. Only-My-Device Rule

The restriction that users may only retrieve the LAPS password for their own device is enforced **exclusively in the backend**.

**Flow:**

1. Backend extracts the Object ID (OID) from the cryptographically verified access token
2. Backend queries all Entra ID registered devices for that user via Microsoft Graph:
   ```
   GET https://graph.microsoft.com/v1.0/users/{oid}/registeredDevices
       ?$filter=operatingSystem eq 'Windows' or operatingSystem eq 'macOS'
       &$select=id,displayName,operatingSystem,isManaged,approximateLastSignInDateTime
   ```
3. For `POST /api/laps-password`: backend checks whether the requested device object ID is in the user's device list
4. If not → HTTP 403 Forbidden
5. If yes → the LAPS credential is fetched with that device's Entra `deviceId`
   (`GET /v1.0/directory/deviceLocalCredentials/{deviceId}`), i.e. the same identifier the
   authorization decision was made on. Display names are **not** unique in Entra ID, so a
   name-based lookup would let a colliding device name resolve to another user's credential.
   The `deviceName` returned by Graph is compared against the authorized device as a final check.

The frontend device list is a UX convenience only. The actual authorization check happens in the backend on every request.

---

## 5. Mandatory Justification

Every password retrieval requires a justification:

- **Required field**: Justification must be provided (minimum 10 characters)
- **Maximum length**: 500 characters, enforced by the backend (`JUSTIFICATION_MAX_LENGTH`).
  Without a server-side cap an oversized justification exceeds the Azure Table entity limit,
  the audit write fails, and the password would still be released.
- **Stored in audit log**: Yes, as free text
- **Validation**: Frontend (UX) + backend (authoritative, both bounds)

---

## 6. Audit Logging

Every successful and failed LAPS password retrieval attempt is logged.

### 6.1 Storage

Azure Table Storage – table `LapsAuditLog` in the dedicated storage account.

### 6.2 Log Schema

| Field | Type | Description |
|-------|------|-------------|
| `PartitionKey` | String | Date in `YYYY-MM-DD` format (UTC) |
| `RowKey` | String | UUID v4 (unique event ID) |
| `Timestamp` | DateTime | UTC timestamp |
| `UserPrincipalName` | String | UPN of the requesting user |
| `UserId` | String | Entra Object ID of the user |
| `DeviceId` | String | Intune Device ID |
| `DeviceName` | String | Device display name |
| `Justification` | String | User-provided reason |
| `Action` | String | `SUCCESS`, `DENIED`, or `ERROR` |
| `DenialReason` | String | Reason when Action != `SUCCESS` |
| `ClientIp` | String | Source IP address |
| `UserAgent` | String | Browser user agent string |

### 6.3 Retention

- Azure Table Storage: unlimited (low cost)
- Recommended: archive after 365 days via lifecycle policy

---

## 7. Frontend

### 7.1 Views

The frontend is a single-page application. Views are toggled via CSS `display`.

| View | Description |
|------|-------------|
| `#view-login` | Welcome screen with sign-in button |
| `#view-loading` | Loading spinner during API calls |
| `#view-devices` | User's managed device list |
| `#view-password` | LAPS password display (with countdown) |
| `#view-error` | Error message |

### 7.2 UX Requirements

- LAPS password is automatically hidden after **60 seconds** (countdown)
- **Copy-to-clipboard** button next to the password
- Password initially masked (`•••••••`), toggle button to reveal
- Device name and password expiry date are displayed
- Mobile-friendly layout (responsive, minimum 320px width)
- No JavaScript framework, no build pipeline required

### 7.3 CSS Theming

All visual properties are controlled via CSS Custom Properties in `theme.css`.

Customizable variables:

```css
:root {
  --color-primary: #0078d4;        /* Main action color (buttons, links) */
  --color-primary-hover: #106ebe;  /* Hover state                        */
  --color-danger: #a4262c;
  --color-success: #107c10;
  --color-bg: #f3f2f1;
  --color-surface: #ffffff;
  --color-text: #201f1e;
  --color-text-muted: #605e5c;
  --color-border: #e1dfdd;
  --font-family: 'Segoe UI', system-ui, sans-serif;
  --border-radius: 8px;
  --shadow: 0 2px 8px rgba(0,0,0,0.10);
  --logo-url: none;                /* Set to url('/assets/logo.svg') */
}
```

### 7.4 Custom Domain Support

- Azure Static Web Apps supports custom domains natively
- Create a CNAME record pointing to the SWA hostname
- TLS/SSL is provisioned automatically by Azure (Let's Encrypt)
- Configured in Bicep via the `customDomain` parameter

---

## 8. Backend

### 8.1 Azure Function App

- **Runtime**: Node.js 24 LTS
- **Hosting**: App Service Plan (Linux B1 Basic)
- **Auth**: Easy Auth (built-in authentication) with Entra ID provider
- **CORS**: Restricted to the Static Web App domain only

### 8.2 API Endpoints

#### `GET /api/my-devices`

Returns all Entra ID registered devices for the signed-in user.

**Request:**
```
GET /api/my-devices
Authorization: Bearer <msal-access-token>
```

**Response 200:**
```json
{
  "devices": [
    {
      "id": "device-id-guid",
      "name": "LAPTOP-ABC123",
      "operatingSystem": "Windows",
      "osVersion": "10.0.22631",
      "complianceState": "compliant",
      "lastSyncDateTime": "2026-02-25T14:30:00Z"
    }
  ]
}
```

**Response 401:** Token missing or invalid
**Response 500:** Graph API error

#### `POST /api/laps-password`

Retrieves the LAPS password for a device, after ownership verification.

**Request body:**
```json
{
  "deviceId": "device-id-guid",
  "justification": "Local software installation, ticket #1234"
}
```

**Response 200:**
```json
{
  "deviceName": "LAPTOP-ABC123",
  "password": "Xk9!mP2#nQ7$",
  "expiresAt": "2026-03-01T00:00:00Z",
  "auditId": "550e8400-e29b-41d4-a716-446655440000"
}
```

**Response 400:** Missing or invalid request body (e.g. justification too short)
**Response 401:** Not authenticated
**Response 403:** Device does not belong to the requesting user
**Response 404:** No LAPS password stored for this device
**Response 500:** Graph API error

---

## 9. Infrastructure (Bicep)

### 9.1 Resources

| Resource | Type | Description |
|----------|------|-------------|
| Resource Group | `Microsoft.Resources/resourceGroups` | Container for all resources |
| Storage Account | `Microsoft.Storage/storageAccounts` | Function App storage + audit log |
| App Service Plan | `Microsoft.Web/serverfarms` | Linux B1 Basic |
| Function App | `Microsoft.Web/sites` | Backend with Managed Identity |
| Static Web App | `Microsoft.Web/staticSites` | Frontend |
| Log Analytics Workspace | `Microsoft.OperationalInsights/workspaces` | Monitoring |
| Application Insights | `Microsoft.Insights/components` | Function App telemetry |

### 9.2 Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `location` | string | `germanywestcentral` | Azure region |
| `environmentName` | string | – | Short name for resource naming (e.g. `prod`, `dev`) |
| `entraClientId` | string | – | App Registration Client ID (for Easy Auth) |
| `entraTenantId` | string | – | Entra ID Tenant ID |
| `customDomain` | string | `''` | Optional custom domain for Static Web App |

### 9.3 Naming Convention

Resources follow the pattern: `{projectName}-{resourceType}`

Examples:
- `laps-prod-func` – Function App
- `laps-prod-swa` – Static Web App
- `laps-prod-plan` – App Service Plan
- `{prefix}{uniqueString}` – Storage Account (lowercase, alphanumeric, max 24 chars)
- `laps-prod-law` – Log Analytics Workspace
- `laps-prod-ai` – Application Insights

---

## 10. Security Requirements

| Requirement | Measure |
|-------------|---------|
| No stored secrets | Managed Identity for all Azure service access |
| Token validation | Easy Auth on the Function App (not manual code) |
| Device ownership | Backend check on every request |
| HTTPS only | HSTS via Static Web App, TLS 1.2+ enforced |
| CORS | Function App accepts requests only from the SWA origin |
| Audit trail | Full logging of all retrieval attempts |
| Password display | Temporary only (60 s countdown), no browser caching |
| Mandatory justification | Backend-side validation enforced |
| Least privilege | Only the required Graph API permissions assigned |

---

## 11. Out of Scope

- Password rotation (read-only)
- Admin dashboard for IT staff
- Email notifications on password retrieval
- Multi-language support (initial release: English)
- Offline operation
- Mobile app
