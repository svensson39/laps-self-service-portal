/**
 * lib/audit.js – Persistent audit log in Azure Table Storage.
 *
 * Complements Application Insights (real-time monitoring) with a durable,
 * queryable record of every LAPS access attempt.
 *
 * Table : LapsAuditLog  (configurable via AUDIT_TABLE_NAME env var)
 * PartitionKey : YYYY-MM-DD UTC date of the event
 * RowKey       : UUID v4 (unique event ID, returned as auditId)
 *
 * Authentication
 * ──────────────
 * Preferred: AUDIT_STORAGE_ACCOUNT_NAME + Managed Identity holding the
 * "Storage Table Data Contributor" role. No account key is stored anywhere.
 *
 * Fallback: AUDIT_STORAGE_CONNECTION_STRING, for local development against
 * Azurite. Using a shared-key connection string in Azure means anyone who can
 * read the app settings can also rewrite the audit trail – avoid it there.
 *
 * Failure handling
 * ────────────────
 * writeAuditLog() throws when the record cannot be persisted. Callers must
 * decide: a successful password disclosure that could not be audited has to
 * fail closed, while a denial that could not be audited may still be returned.
 * See functions/lapsPassword.js.
 */

'use strict';

const { TableClient }            = require('@azure/data-tables');
const { DefaultAzureCredential } = require('@azure/identity');
const { v4: uuidv4 }             = require('uuid');

const ACCOUNT_NAME      = process.env.AUDIT_STORAGE_ACCOUNT_NAME;
const CONNECTION_STRING = process.env.AUDIT_STORAGE_CONNECTION_STRING;
const TABLE_NAME        = process.env.AUDIT_TABLE_NAME ?? 'LapsAuditLog';
const TABLE_SUFFIX      = process.env.AUDIT_TABLE_ENDPOINT_SUFFIX ?? 'table.core.windows.net';

// Cap free-text fields so a caller cannot exceed Azure Table's per-property
// (32 KB) or per-entity (1 MB) limits and thereby suppress their own audit record.
const MAX_FIELD_LENGTH = 512;

let _client = null;

function getTableClient() {
  if (_client) return _client;

  if (ACCOUNT_NAME) {
    _client = new TableClient(
      `https://${ACCOUNT_NAME}.${TABLE_SUFFIX}`,
      TABLE_NAME,
      new DefaultAzureCredential(),
    );
  } else if (CONNECTION_STRING) {
    _client = TableClient.fromConnectionString(CONNECTION_STRING, TABLE_NAME);
  } else {
    throw new Error('Audit storage is not configured: set AUDIT_STORAGE_ACCOUNT_NAME or AUDIT_STORAGE_CONNECTION_STRING.');
  }

  return _client;
}

/** Coerce to string and clamp to the maximum stored length. */
function clamp(value) {
  const str = String(value ?? '');
  return str.length > MAX_FIELD_LENGTH ? `${str.slice(0, MAX_FIELD_LENGTH)}…[truncated]` : str;
}

/**
 * @typedef {object} AuditEntry
 * @property {string}  oid           Entra Object ID of the requesting user
 * @property {string}  upn           User Principal Name
 * @property {string}  deviceId      Entra Device Object ID
 * @property {string}  [deviceName]  Device display name
 * @property {string}  justification User-provided reason
 * @property {'SUCCESS'|'DENIED'|'ERROR'} action
 * @property {string}  [denialReason]
 * @property {string}  [clientIp]
 * @property {string}  [userAgent]
 */

/**
 * Write an audit record to Azure Table Storage.
 * Returns the generated audit ID (UUID v4) for inclusion in API responses.
 *
 * @param {AuditEntry} entry
 * @returns {Promise<string>} auditId
 * @throws {Error} if the record could not be persisted
 */
async function writeAuditLog(entry) {
  const now          = new Date();
  const partitionKey = now.toISOString().slice(0, 10);  // YYYY-MM-DD
  const rowKey       = uuidv4();

  const entity = {
    partitionKey,
    rowKey,
    Timestamp:         now.toISOString(),
    UserId:            clamp(entry.oid),
    UserPrincipalName: clamp(entry.upn),
    DeviceId:          clamp(entry.deviceId),
    DeviceName:        clamp(entry.deviceName),
    Justification:     clamp(entry.justification),
    Action:            clamp(entry.action ?? 'UNKNOWN'),
    DenialReason:      clamp(entry.denialReason),
    ClientIp:          clamp(entry.clientIp),
    UserAgent:         clamp(entry.userAgent),
  };

  await getTableClient().createEntity(entity);
  return rowKey;
}

module.exports = { writeAuditLog };
