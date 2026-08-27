/**
 * GET /api/my-devices
 *
 * Returns all Entra ID registered devices for the authenticated user.
 * Only device name and object ID are returned – no sensitive information.
 *
 * Response 200:
 *   {
 *     "devices": [
 *       { "id": "...", "name": "LAPTOP-ABC", "operatingSystem": "Windows",
 *         "isManaged": true, "lastSignIn": "2026-02-25T14:00:00Z" }
 *     ]
 *   }
 *
 * Response 401: Missing or invalid token
 * Response 403: Insufficient scope
 * Response 500: Graph API error
 */

'use strict';

const { app }                 = require('@azure/functions');
const { getCallerIdentity, AuthError } = require('../lib/auth');
const { getRegisteredDevices }         = require('../lib/graph');
const { trackDeviceListAccess }        = require('../lib/telemetry');

app.http('my-devices', {
  methods:    ['GET'],
  route:      'my-devices',
  authLevel:  'anonymous',  // Token validation handled by lib/auth.js (verified JWT or declared Easy Auth)

  handler: async (request, context) => {
    context.log('GET /api/my-devices');

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

    context.log(`Fetching devices for user ${caller.upn} (${caller.oid})`);

    // ── 2. Fetch registered devices from Microsoft Graph ───────────────────
    let devices;
    try {
      devices = await getRegisteredDevices(caller.oid);
    } catch (err) {
      context.log('Graph API error in getRegisteredDevices:', err.message, 'statusCode:', err.statusCode, 'code:', err.code);
      return {
        status:   500,
        jsonBody: { error: 'GRAPH_ERROR', message: 'Failed to retrieve device list.' },
      };
    }

    context.log(`Returning ${devices.length} device(s) for ${caller.upn}`);

    // ── 3. Track in Application Insights ──────────────────────────────────
    trackDeviceListAccess({ oid: caller.oid, upn: caller.upn, deviceCount: devices.length });

    // Expose only what the UI needs. The Entra deviceId stays server-side so a
    // client can never supply it as a lookup key.
    return {
      status:   200,
      jsonBody: {
        devices: devices.map(({ id, name, operatingSystem, isManaged, lastSignIn }) =>
          ({ id, name, operatingSystem, isManaged, lastSignIn })),
      },
    };
  },
});
