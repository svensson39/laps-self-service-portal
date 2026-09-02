/**
 * lib/auth.js – Caller identity extraction and JWT validation.
 *
 * Resolution order (fail-safe: cryptographic verification always wins)
 * ───────────────────────────────────────────────────────────────────
 *
 *  1. Authorization: Bearer <token>
 *     The token is verified against Entra ID's public JWKS – signature,
 *     algorithm, audience, issuer, tenant and scope are all checked.
 *     Easy Auth forwards the caller's original Authorization header, so this
 *     path is taken in production too. If the X-MS-CLIENT-PRINCIPAL header is
 *     also present, its OID must match the verified token's OID.
 *
 *  2. X-MS-CLIENT-PRINCIPAL (only when EASY_AUTH_ENABLED === 'true')
 *     Base64-encoded claims injected by the Function App's built-in
 *     authentication. This header carries no signature – it is only
 *     trustworthy because the platform strips any client-supplied copy before
 *     the request reaches function code. That guarantee holds only while Easy
 *     Auth is actually enabled, so the app setting must state so explicitly.
 *     Without the flag a forged header is rejected rather than trusted.
 *
 * Requires TENANT_ID and AUTH_CLIENT_ID.
 */

'use strict';

const jwt      = require('jsonwebtoken');
const jwksRsa  = require('jwks-rsa');

const TENANT_ID      = process.env.TENANT_ID;
const CLIENT_ID      = process.env.AUTH_CLIENT_ID;
const REQUIRED_SCOPE = process.env.REQUIRED_SCOPE ?? 'access_as_user';

// Maximum age of the *interactive sign-in* (Entra auth_time claim), not the
// token. MSAL silently renews access tokens, so token iat/exp never reflect
// how long the user has been on the page – auth_time does.
// 0 = disabled (no session age limit).
const SESSION_MAX_AGE_MINUTES = parseInt(process.env.SESSION_MAX_AGE_MINUTES ?? '60', 10);

// Easy Auth's unsigned principal header is only honoured when the deployment
// explicitly confirms that Easy Auth is in front of this app.
const EASY_AUTH_ENABLED = String(process.env.EASY_AUTH_ENABLED ?? '').toLowerCase() === 'true';

// Claim type URIs used by the Easy Auth principal header
const CLAIM_OID    = 'http://schemas.microsoft.com/identity/claims/objectidentifier';
const CLAIM_TID    = 'http://schemas.microsoft.com/identity/claims/tenantid';
const CLAIM_SCOPE  = 'http://schemas.microsoft.com/identity/claims/scope';
const CLAIM_UPN    = 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn';
const CLAIM_AUTH_TIME = 'http://schemas.microsoft.com/identity/claims/auth_time';

// JWKS client – caches public keys for 10 minutes to avoid repeated HTTP calls
const jwksClient = jwksRsa({
  jwksUri:       `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`,
  cache:         true,
  cacheMaxEntries: 5,
  cacheMaxAge:   10 * 60 * 1000,
  rateLimit:     true,
});

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * @typedef {object} CallerIdentity
 * @property {string} oid  - Entra ID Object ID (stable, use for Graph calls and audit)
 * @property {string} upn  - User Principal Name (human-readable, use for logging)
 */

/**
 * Extract the verified caller identity from the incoming request.
 *
 * @param {import('@azure/functions').HttpRequest} request
 * @returns {Promise<CallerIdentity>}
 * @throws {AuthError}
 */
async function getCallerIdentity(request) {
  const easyAuthHeader = request.headers.get('x-ms-client-principal');
  const authHeader     = request.headers.get('authorization');

  // Path 1: cryptographically verifiable Bearer token – always preferred
  if (authHeader?.startsWith('Bearer ')) {
    const identity = await verifyBearerToken(authHeader.slice(7));

    // Defence in depth: a forged principal header cannot contradict a verified token
    if (easyAuthHeader) {
      const injected = decodeEasyAuthPrincipal(easyAuthHeader);
      if (injected.oid !== identity.oid) {
        throw new AuthError(401, 'Authentication failed.',
          `X-MS-CLIENT-PRINCIPAL oid (${injected.oid}) does not match verified token oid (${identity.oid}).`);
      }
    }

    return identity;
  }

  // Path 2: Easy Auth principal header, only when the platform guarantee is declared
  if (easyAuthHeader) {
    if (!EASY_AUTH_ENABLED) {
      throw new AuthError(401, 'Authentication required.',
        'X-MS-CLIENT-PRINCIPAL received but EASY_AUTH_ENABLED is not "true" – refusing to trust an unsigned identity header.');
    }
    return decodeEasyAuthPrincipal(easyAuthHeader, { validateClaims: true });
  }

  throw new AuthError(401, 'Authentication required.', 'Missing or invalid Authorization header.');
}

// ---------------------------------------------------------------------------
// Easy Auth principal header
// ---------------------------------------------------------------------------

/**
 * Decode the base64 X-MS-CLIENT-PRINCIPAL header.
 *
 * @param {string} headerValue
 * @param {{validateClaims?: boolean}} [opts]  Also enforce tenant and scope claims
 * @returns {CallerIdentity}
 */
function decodeEasyAuthPrincipal(headerValue, opts = {}) {
  let principal;
  try {
    principal = JSON.parse(Buffer.from(headerValue, 'base64').toString('utf8'));
  } catch {
    throw new AuthError(401, 'Authentication failed.', 'Failed to decode X-MS-CLIENT-PRINCIPAL header.');
  }

  const claims = principal.claims ?? [];

  /** Find the first non-empty value among the given claim type URIs. */
  const getClaim = (...types) => {
    for (const type of types) {
      const match = claims.find(c => c.typ === type);
      if (match?.val) {return match.val;}
    }
    return null;
  };

  const oid = getClaim(CLAIM_OID, 'oid');
  const upn = getClaim(CLAIM_UPN, 'upn', 'preferred_username');

  if (!oid) {throw new AuthError(401, 'Authentication failed.', 'OID claim missing from principal.');}

  if (opts.validateClaims) {
    const tid = getClaim(CLAIM_TID, 'tid');
    if (tid !== TENANT_ID) {
      throw new AuthError(401, 'Authentication failed.',
        `Tenant claim mismatch (expected ${TENANT_ID}, got ${tid}).`);
    }

    // Only enforce the scope when Easy Auth surfaced one – some token shapes omit it.
    const scope = getClaim(CLAIM_SCOPE, 'scp');
    if (scope && !hasRequiredScope(scope)) {
      throw new AuthError(403, 'Insufficient scope.',
        `Scope claim "${scope}" does not include "${REQUIRED_SCOPE}".`);
    }

    // Session age limit for the Easy Auth path. Fail closed: if the platform
    // did not surface auth_time we cannot measure session age, so the request
    // is rejected rather than silently bypassing the limit.
    if (SESSION_MAX_AGE_MINUTES > 0) {
      const authTimeRaw = getClaim(CLAIM_AUTH_TIME, 'auth_time');
      const authTime    = authTimeRaw ? Math.floor(Number(authTimeRaw)) : NaN;

      if (!Number.isFinite(authTime) || authTime <= 0) {
        throw new AuthError(401, 'Session expired. Please sign in again.',
          'Easy Auth principal has no usable auth_time claim while SESSION_MAX_AGE_MINUTES is active.');
      }

      const sessionAgeMinutes = (Math.floor(Date.now() / 1000) - authTime) / 60;
      if (sessionAgeMinutes > SESSION_MAX_AGE_MINUTES) {
        throw new AuthError(401, 'Session expired. Please sign in again.',
          `Session age ${Math.round(sessionAgeMinutes)} min exceeds SESSION_MAX_AGE_MINUTES (${SESSION_MAX_AGE_MINUTES}).`);
      }
    }
  }

  return { oid, upn: upn ?? '' };
}

// ---------------------------------------------------------------------------
// Direct JWT verification
// ---------------------------------------------------------------------------

async function verifyBearerToken(token) {
  if (!TENANT_ID || !CLIENT_ID) {
    throw new AuthError(500, 'Server configuration error.',
      'TENANT_ID and AUTH_CLIENT_ID must be set for JWT validation.');
  }

  // Decode header to extract the key ID (kid) without verifying
  const unverified = jwt.decode(token, { complete: true });
  if (!unverified?.header?.kid) {
    throw new AuthError(401, 'Authentication failed.', 'Invalid JWT: missing kid header.');
  }

  // Fetch the matching public key from Entra ID's JWKS endpoint
  let signingKey;
  try {
    const key = await jwksClient.getSigningKey(unverified.header.kid);
    signingKey = key.getPublicKey();
  } catch (err) {
    throw new AuthError(401, 'Authentication failed.',
      `Failed to retrieve token signing key: ${err.message}`);
  }

  // Verify signature, algorithm, audience and issuer
  let payload;
  try {
    payload = jwt.verify(token, signingKey, {
      // Pin the algorithm – never let the token's own header choose it
      algorithms: ['RS256'],
      // Only the API audience. The bare client ID is the audience of ID tokens,
      // which are not bearer credentials for this API.
      audience:   `api://${CLIENT_ID}`,
      // Accept both Entra issuer formats for this tenant. Which one appears
      // depends on the app registration's requestedAccessTokenVersion; both are
      // signed by the same tenant-scoped JWKS, so allowing both is not a
      // weakening. NOTE: if you set requestedAccessTokenVersion to 2 the
      // audience becomes the bare client ID and must be added above.
      issuer: [
        `https://sts.windows.net/${TENANT_ID}/`,
        `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      ],
    });
  } catch (err) {
    throw new AuthError(401, 'Authentication failed.', `Token verification failed: ${err.message}`);
  }

  if (payload.tid !== TENANT_ID) {
    throw new AuthError(401, 'Authentication failed.',
      `Tenant claim mismatch (expected ${TENANT_ID}, got ${payload.tid}).`);
  }

  // Reject app-only tokens: this API is only reachable on behalf of a signed-in user
  if (!payload.scp) {
    throw new AuthError(403, 'Insufficient scope.',
      'Token has no scp claim – app-only tokens are not accepted.');
  }
  if (!hasRequiredScope(payload.scp)) {
    throw new AuthError(403, 'Insufficient scope.',
      `Scope claim "${payload.scp}" does not include "${REQUIRED_SCOPE}".`);
  }

  const oid = payload.oid;
  const upn = payload.preferred_username ?? payload.upn ?? '';

  if (!oid) {throw new AuthError(401, 'Authentication failed.', 'OID claim missing from token.');}

  // ── Session age limit ──────────────────────────────────────────────────────
  // Entra issues auth_time on interactive sign-ins; silent token renewals
  // carry the ORIGINAL auth_time forward, so this measures true session age.
  if (SESSION_MAX_AGE_MINUTES > 0) {
    if (!payload.auth_time || typeof payload.auth_time !== 'number') {
      throw new AuthError(401, 'Session expired. Please sign in again.',
        'Token has no numeric auth_time claim – cannot verify session age.');
    }

    const sessionAgeMinutes = (Math.floor(Date.now() / 1000) - payload.auth_time) / 60;

    if (sessionAgeMinutes > SESSION_MAX_AGE_MINUTES) {
      throw new AuthError(401, 'Session expired. Please sign in again.',
        `Session age ${Math.round(sessionAgeMinutes)} min exceeds SESSION_MAX_AGE_MINUTES (${SESSION_MAX_AGE_MINUTES}).`);
    }

    // A small negative drift (clock skew) is tolerated by comparing only the
    // upper bound; anything wildly negative would indicate a forged token,
    // which the signature verification above has already ruled out.
  }

  return { oid, upn };
}

/** Entra returns scopes as a space-separated list. */
function hasRequiredScope(scp) {
  return String(scp).split(' ').includes(REQUIRED_SCOPE);
}

// ---------------------------------------------------------------------------
// AuthError
// ---------------------------------------------------------------------------

class AuthError extends Error {
  /**
   * @param {number} status  HTTP status code (401, 403 or 500)
   * @param {string} message Generic text safe to return to the client
   * @param {string} [detail] Diagnostic text for server-side logs only – never returned
   */
  constructor(status, message, detail) {
    super(message);
    this.name = 'AuthError';
    this.status = status;
    this.detail = detail ?? message;
  }
}

module.exports = { getCallerIdentity, AuthError };
