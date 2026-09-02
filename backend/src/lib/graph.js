/**
 * lib/graph.js – Microsoft Graph API client using DefaultAzureCredential.
 *
 * In production the Function App's system-assigned Managed Identity is used
 * automatically (no secrets required). Locally, DefaultAzureCredential falls
 * back to Azure CLI, VS Code, or environment credentials.
 *
 * Required application permissions on the Managed Identity:
 *   Device.Read.All                  – enumerate user's registered devices
 *   DeviceLocalCredential.Read.All   – read LAPS passwords
 *
 * Directory.Read.All is deliberately NOT requested: every call this module
 * makes is covered by the two roles above, and Directory.Read.All would grant
 * tenant-wide read access to the whole directory.
 *
 * See docs/deployment.md (Step 3) for the az CLI commands to assign these.
 */

'use strict';

const { DefaultAzureCredential }    = require('@azure/identity');
const { Client }                    = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require(
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials',
);

const GRAPH_ENDPOINT = process.env.GRAPH_API_ENDPOINT ?? 'https://graph.microsoft.com';

// Singleton credential – reused across warm invocations to avoid redundant
// IMDS/token-endpoint round-trips on every request
let _credential = null;

function getCredential() {
  if (!_credential) {_credential = new DefaultAzureCredential();}
  return _credential;
}

function getGraphClient() {
  const authProvider = new TokenCredentialAuthenticationProvider(getCredential(), {
    scopes: [`${GRAPH_ENDPOINT}/.default`],
  });
  return Client.initWithMiddleware({ authProvider, baseUrl: GRAPH_ENDPOINT });
}

// ---------------------------------------------------------------------------
// Device list
// ---------------------------------------------------------------------------

/**
 * @typedef {object} DeviceInfo
 * @property {string}  id              - Entra Device *Object* ID (used to match the client's request)
 * @property {string}  deviceId        - Entra *Device* ID (GUID) – the key for the LAPS lookup
 * @property {string}  name            - Device display name
 * @property {string}  operatingSystem - e.g. "Windows"
 * @property {boolean} isManaged       - Whether the device is Intune-managed
 * @property {string|null} lastSignIn  - ISO 8601 date of last sign-in activity
 */

/**
 * Return all Entra ID registered devices for the given user.
 * Uses GET /v1.0/users/{userId}/registeredDevices (requires Device.Read.All).
 *
 * Note: registeredDevices is a reference/navigation property – Graph does not
 * support $filter on it. We fetch all registered devices and filter client-side.
 *
 * @param {string} userId  Entra Object ID of the authenticated user
 * @returns {Promise<DeviceInfo[]>}
 */
async function getRegisteredDevices(userId) {
  const client = getGraphClient();
  const result = await client
    .api(`/users/${userId}/registeredDevices`)
    .select('id,deviceId,displayName,operatingSystem,isManaged,approximateLastSignInDateTime')
    .get();

  const SUPPORTED_OS = new Set(['windows', 'macos']);

  return (result?.value ?? [])
    .filter(d => SUPPORTED_OS.has((d.operatingSystem ?? '').toLowerCase()))
    .map(d => ({
      id:              d.id,
      deviceId:        d.deviceId ?? null,
      name:            d.displayName ?? '',
      operatingSystem: d.operatingSystem ?? '',
      isManaged:       d.isManaged ?? false,
      lastSignIn:      d.approximateLastSignInDateTime ?? null,
    }));
}

// ---------------------------------------------------------------------------
// Device ownership verification
// ---------------------------------------------------------------------------

/**
 * Verify that a device (by Entra Device Object ID) is in the user's registered
 * devices. Returns the matching device object if owned, or null if not.
 *
 * @param {string} deviceId  Entra Device *Object* ID as supplied by the client
 * @param {string} userId    Entra Object ID of the authenticated user
 * @returns {Promise<DeviceInfo|null>}
 */
async function findOwnedDevice(deviceId, userId) {
  const devices = await getRegisteredDevices(userId);
  return devices.find(d => d.id === deviceId) ?? null;
}

// ---------------------------------------------------------------------------
// LAPS password retrieval
// ---------------------------------------------------------------------------

/**
 * @typedef {object} LapsCredential
 * @property {string}      deviceName       - Display name of the device
 * @property {string}      accountName      - Local administrator account name
 * @property {string}      password         - Plaintext local administrator password
 * @property {string|null} passwordCreated  - ISO 8601 timestamp when this credential was backed up to Entra ID
 *                                            (sourced from credentials[].backupDateTime)
 * @property {string|null} nextRotation     - ISO 8601 timestamp when the password will next be refreshed and
 *                                            backed up to Entra ID (sourced from top-level refreshDateTime)
 */

/**
 * Retrieve the LAPS credential for the given device.
 *
 * The lookup key is the Entra **deviceId** – the same value the ownership check
 * is performed against. deviceLocalCredentialInfo.id *is* the device's deviceId,
 * so the credential can be fetched directly:
 *
 *   GET /v1.0/directory/deviceLocalCredentials/{deviceId}?$select=credentials,deviceName,refreshDateTime
 *
 * Do NOT look this up by device display name. Display names are not unique in
 * Entra ID (they are just the hostname), so a name-based lookup lets a user who
 * owns a device called "X" receive the LAPS credential of somebody else's
 * device also called "X" – the authorization decision and the data fetch would
 * be keyed on different identifiers.
 *
 * As belt and braces, the display name returned by Graph is compared with the
 * name of the device the caller was authorized for.
 *
 * We use native fetch (Node 18+) instead of the Graph SDK client to avoid
 * the SDK's version-override mechanism conflicting with a custom baseUrl.
 *
 * Requires: DeviceLocalCredential.Read.All
 *
 * @param {string} deviceId            Entra deviceId (GUID) of the owned device
 * @param {string} [expectedDeviceName] Display name the caller was authorized for
 * @returns {Promise<LapsCredential>}
 * @throws {Error} err.code === 'NOT_FOUND' if no credential is stored
 * @throws {Error} err.code === 'DEVICE_MISMATCH' if Graph returns a different device
 */
async function getLapsPassword(deviceId, expectedDeviceName) {
  if (!deviceId) {
    const err = new Error('No Entra deviceId available for this device.');
    err.code = 'NOT_FOUND';
    throw err;
  }

  const tokenResponse = await getCredential().getToken(`${GRAPH_ENDPOINT}/.default`);
  const authHeader = { Authorization: `Bearer ${tokenResponse.token}` };

  const url = `${GRAPH_ENDPOINT}/v1.0/directory/deviceLocalCredentials/${encodeURIComponent(deviceId)}`
            + '?$select=credentials,deviceName,refreshDateTime,expirationDateTime';
  const res = await fetch(url, { headers: authHeader });

  let result;
  try {
    result = await res.json();
  } catch {
    result = {};
  }

  if (res.status === 404) {
    const notFound = new Error('No LAPS credential found for this device.');
    notFound.code = 'NOT_FOUND';
    throw notFound;
  }

  if (!res.ok) {
    const msg = result?.error?.message ?? `HTTP ${res.status}`;
    const err = new Error(msg);
    err.statusCode = res.status;
    throw err;
  }

  // Graph should never return a different device for a deviceId lookup, but the
  // whole security model rests on this – verify rather than assume.
  if (expectedDeviceName && result.deviceName
      && result.deviceName.toLowerCase() !== expectedDeviceName.toLowerCase()) {
    const mismatch = new Error(
      `LAPS lookup returned "${result.deviceName}" but the caller was authorized for "${expectedDeviceName}".`);
    mismatch.code = 'DEVICE_MISMATCH';
    throw mismatch;
  }

  const credential = result?.credentials?.[0];
  if (!credential) {
    const notFound = new Error('No LAPS credential found for this device.');
    notFound.code = 'NOT_FOUND';
    throw notFound;
  }

  // The password is returned base64-encoded; decode to plaintext
  const password = credential.passwordBase64
    ? Buffer.from(credential.passwordBase64, 'base64').toString('utf8')
    : (credential.password ?? '');

  return {
    deviceName:      result.deviceName ?? expectedDeviceName ?? '',
    accountName:     credential.accountName ?? '',
    password,
    passwordCreated: credential.backupDateTime ?? null,
    nextRotation:    result.refreshDateTime ?? null,
    passwordExpires: result.expirationDateTime ?? credential.expirationDateTime ?? null,
  };
}

module.exports = { getRegisteredDevices, findOwnedDevice, getLapsPassword };
