/**
 * POST /api/laps-password
 *
 * Retrieves the LAPS (local administrator) password for a device after:
 *   1. Authenticating the caller via verified JWT (or Easy Auth, when declared)
 *   2. Validating the request body (deviceId + justification)
 *   3. Verifying the device is owned by the authenticated user (Graph API)
 *   4. Fetching the LAPS credential from Microsoft Graph, keyed on the same
 *      device identity the ownership check was performed against
 *
 * Every attempt (success or failure) is:
 *   - Tracked as a custom event in Application Insights
 *   - Written to Azure Table Storage (persistent audit log)
 *
 * The audit log is fail-closed for successful retrievals: if the record cannot
 * be persisted, the password is not returned. Otherwise a caller could suppress
 * their own audit trail and still obtain the credential.
 *
 * The password is NEVER persisted anywhere – it only exists in the response body.
 *
 * Request body:
 *   { "deviceId": "<entra-device-object-id>", "justification": "..." }
 *
 * Response 200:
 *   { "deviceName": "LAPTOP-ABC", "password": "...", "passwordCreated": "...", "nextRotation": "...", "auditId": "..." }
 *
 * Response 400: Missing or invalid body
 * Response 401: Not authenticated
 * Response 403: Device not owned by the requesting user, or insufficient scope
 * Response 404: No LAPS password stored for this device
 * Response 500: Graph API error
 * Response 503: Audit log unavailable – password deliberately withheld
 */

'use strict';

const { app }                          = require('@azure/functions');
const { getCallerIdentity, AuthError } = require('../lib/auth');
const { findOwnedDevice, getLapsPassword } = require('../lib/graph');
const { trackPasswordAccess }          = require('../lib/telemetry');
const { writeAuditLog }                = require('../lib/audit');

const MIN_JUSTIFICATION = parseInt(process.env.JUSTIFICATION_MIN_LENGTH ?? '10', 10);
const MAX_JUSTIFICATION = parseInt(process.env.JUSTIFICATION_MAX_LENGTH ?? '500', 10);

// Client-controlled headers are recorded for forensics only – never trusted.
const MAX_HEADER_LENGTH = 256;

/** First hop of X-Forwarded-For, trimmed of any port and clamped. */
function firstForwardedIp(headerValue) {
  const first = String(headerValue ?? '').split(',')[0].trim();
  // Azure appends :port for IPv4; leave bracketed IPv6 literals intact
  const withoutPort = first.startsWith('[') ? first : first.replace(/:\d+$/, '');
  return withoutPort.slice(0, 64);
}

/** Strip control characters so a crafted header cannot forge log structure. */
function sanitizeHeader(value) {
  return Array.from(String(value ?? ''))
    .map((ch) => {
      const code = ch.charCodeAt(0);
      return (code < 0x20 || code === 0x7f) ? ' ' : ch;
    })
    .join('')
    .slice(0, MAX_HEADER_LENGTH);
}

app.http('laps-password', {
  methods:   ['POST'],
  route:     'laps-password',
  authLevel: 'anonymous',  // Token validation handled by lib/auth.js (verified JWT or declared Easy Auth)

  handler: async (request, context) => {
    context.log('POST /api/laps-password');

    const clientIp  = sanitizeHeader(firstForwardedIp(request.headers.get('x-forwarded-for')));
    const userAgent = sanitizeHeader(request.headers.get('user-agent'));

    // ── 1. Authenticate ────────────────────────────────────────────────────
    let caller;
    try {
      caller = await getCallerIdentity(request);
    } catch (err) {
      // Diagnostics stay server-side; the client gets a generic message
      context.log('Authentication rejected:', err.detail ?? err.message);
      const status = err instanceof AuthError ? err.status : 401;
      return {
        status,
        jsonBody: {
          error:   status === 403 ? 'FORBIDDEN' : 'UNAUTHORIZED',
          message: err instanceof AuthError ? err.message : 'Authentication required.',
        },
      };
    }

    // ── 2. Validate request body ────────────────────────────────────────────
    let body;
    try {
      body = await request.json();
    } catch {
      return {
        status:   400,
        jsonBody: { error: 'INVALID_BODY', message: 'Request body must be valid JSON.' },
      };
    }

    const { deviceId, justification } = body ?? {};

    if (!deviceId || typeof deviceId !== 'string') {
      return {
        status:   400,
        jsonBody: { error: 'MISSING_DEVICE_ID', message: 'deviceId is required.' },
      };
    }

    const trimmedJustification = typeof justification === 'string' ? justification.trim() : '';

    if (trimmedJustification.length < MIN_JUSTIFICATION) {
      return {
        status:   400,
        jsonBody: {
          error:   'JUSTIFICATION_TOO_SHORT',
          message: `Justification must be at least ${MIN_JUSTIFICATION} characters.`,
        },
      };
    }

    // Enforced server-side: the frontend's maxlength attribute is not a control.
    // An oversized justification would otherwise break the audit write while the
    // password was still returned.
    if (trimmedJustification.length > MAX_JUSTIFICATION) {
      return {
        status:   400,
        jsonBody: {
          error:   'JUSTIFICATION_TOO_LONG',
          message: `Justification must be at most ${MAX_JUSTIFICATION} characters.`,
        },
      };
    }

    context.log(`User ${caller.upn} requesting LAPS for device ${deviceId}`);

    // ── 3. Verify device ownership (backend-enforced "only my device" rule) ──
    let device;
    try {
      device = await findOwnedDevice(deviceId, caller.oid);
    } catch (err) {
      context.log('Graph API error during ownership check:', err.message);
      trackPasswordAccess({
        oid: caller.oid, upn: caller.upn, deviceId,
        justification: trimmedJustification, success: false, failReason: 'OWNERSHIP_CHECK_ERROR',
      });
      return {
        status:   500,
        jsonBody: { error: 'GRAPH_ERROR', message: 'Failed to verify device ownership.' },
      };
    }

    if (!device) {
      context.log(`Device ${deviceId} not found in registered devices of ${caller.upn}`);
      trackPasswordAccess({
        oid: caller.oid, upn: caller.upn, deviceId,
        justification: trimmedJustification, success: false, failReason: 'DEVICE_NOT_OWNED',
      });
      await auditDenial(context, {
        oid: caller.oid, upn: caller.upn, deviceId, deviceName: '',
        justification: trimmedJustification, denialReason: 'DEVICE_NOT_OWNED', clientIp, userAgent,
      });
      return {
        status:   403,
        jsonBody: { error: 'DEVICE_NOT_OWNED', message: 'This device is not registered to your account.' },
      };
    }

    // ── 4. Retrieve LAPS password from Microsoft Graph ─────────────────────
    // Keyed on device.deviceId – the same device identity that was just
    // authorized. Looking up by display name would allow a device with a
    // colliding name to resolve to somebody else's credential.
    let lapsResult;
    try {
      lapsResult = await getLapsPassword(device.deviceId, device.name);
    } catch (err) {
      if (err.code === 'NOT_FOUND') {
        trackPasswordAccess({
          oid: caller.oid, upn: caller.upn, deviceId, deviceName: device.name,
          justification: trimmedJustification, success: false, failReason: 'NO_LAPS_CREDENTIAL',
        });
        await auditDenial(context, {
          oid: caller.oid, upn: caller.upn, deviceId, deviceName: device.name,
          justification: trimmedJustification, denialReason: 'NO_LAPS_CREDENTIAL', clientIp, userAgent,
        });
        return {
          status:   404,
          jsonBody: { error: 'NO_LAPS_CREDENTIAL', message: 'No LAPS password is stored for this device.' },
        };
      }

      if (err.code === 'DEVICE_MISMATCH') {
        context.log('SECURITY: LAPS lookup returned a different device than authorized:', err.message);
        trackPasswordAccess({
          oid: caller.oid, upn: caller.upn, deviceId, deviceName: device.name,
          justification: trimmedJustification, success: false, failReason: 'DEVICE_MISMATCH',
        });
        await auditDenial(context, {
          oid: caller.oid, upn: caller.upn, deviceId, deviceName: device.name,
          justification: trimmedJustification, denialReason: 'DEVICE_MISMATCH', clientIp, userAgent,
        });
        return {
          status:   403,
          jsonBody: { error: 'DEVICE_MISMATCH', message: 'Could not confirm the identity of this device.' },
        };
      }

      context.log('Graph API error retrieving LAPS password:', err.message);
      trackPasswordAccess({
        oid: caller.oid, upn: caller.upn, deviceId, deviceName: device.name,
        justification: trimmedJustification, success: false, failReason: 'LAPS_API_ERROR',
      });
      return {
        status:   500,
        jsonBody: { error: 'GRAPH_ERROR', message: 'Failed to retrieve LAPS password.' },
      };
    }

    // ── 5. Persist the audit record BEFORE disclosing the password ──────────
    let auditId;
    try {
      auditId = await writeAuditLog({
        oid: caller.oid, upn: caller.upn, deviceId,
        deviceName:    lapsResult.deviceName,
        justification: trimmedJustification,
        action:        'SUCCESS',
        clientIp,
        userAgent,
      });
    } catch (err) {
      // Fail closed: an unauditable disclosure is worse than a failed request
      context.log('AUDIT FAILURE – withholding password:', err.message);
      trackPasswordAccess({
        oid: caller.oid, upn: caller.upn, deviceId, deviceName: lapsResult.deviceName,
        justification: trimmedJustification, success: false, failReason: 'AUDIT_WRITE_FAILED',
      });
      return {
        status:   503,
        jsonBody: {
          error:   'AUDIT_UNAVAILABLE',
          message: 'The audit log is unavailable, so the password cannot be released. Please try again shortly.',
        },
      };
    }

    trackPasswordAccess({
      oid: caller.oid, upn: caller.upn, deviceId,
      deviceName:    lapsResult.deviceName,
      justification: trimmedJustification,
      success:       true,
    });

    context.log(`Password delivered for ${lapsResult.deviceName} (audit: ${auditId})`);

    // ── 6. Return – password is never persisted beyond this response ────────
    return {
      status: 200,
      jsonBody: {
        deviceName:      lapsResult.deviceName,
        accountName:     lapsResult.accountName,     // local admin username
        password:        lapsResult.password,         // plaintext, never logged or stored
        passwordCreated: lapsResult.passwordCreated,  // when this credential was backed up to Entra ID
        nextRotation:    lapsResult.nextRotation,     // when the password will next be rotated
        auditId,
      },
    };
  },
});

/**
 * Audit a denial. Unlike a success, a failed write here is logged and swallowed:
 * nothing sensitive is being released, so blocking the 403/404 adds no safety.
 */
async function auditDenial(context, entry) {
  try {
    await writeAuditLog({ ...entry, action: 'DENIED' });
  } catch (err) {
    context.log('[audit] Failed to write denial record:', err.message);
  }
}
