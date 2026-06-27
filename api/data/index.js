/*
 * SFS SAF Compliance Tracker — server-side data API
 * ────────────────────────────────────────────────────────────────
 * THIS is the security boundary. Everything here runs on Azure, not
 * in the browser. The browser can lie about anything; this code does
 * not trust it. Identity comes ONLY from the x-ms-client-principal
 * header, which Azure Easy Auth injects and the client cannot forge.
 *
 * Flow per request:
 *   1. Parse verified identity (oid + roles) from the injected header.
 *   2. Derive role + which operators this user may touch — SERVER-SIDE.
 *   3. Reject out-of-scope reads/writes BEFORE any SharePoint call.
 *   4. Exchange the user's token (OBO) for a Graph token and read/write
 *      SharePoint AS THE USER, so SharePoint's own permissions are a
 *      second backstop.
 *
 * Data keys (in the SharePoint list, Title column):
 *   "operators"      → JSON array of operators incl. assignedOids[]
 *   "state_<opId>"   → JSON compliance state for one operator
 * ────────────────────────────────────────────────────────────────
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

// Single-client build: at most one operator/facility. Mirrors MAX_OPS in
// index.html, but THIS is the authoritative cap — the browser's is cosmetic.
const MAX_OPERATORS = 1;

// ── env (set via `az staticwebapp appsettings set`, NEVER in code/repo) ──
const TENANT_ID    = process.env.AAD_TENANT_ID;
const CLIENT_ID    = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET= process.env.AAD_CLIENT_SECRET;   // OBO requires this
const SITE_URL     = process.env.SP_SITE_URL;          // https://x.sharepoint.com/sites/y
const LIST_NAME    = process.env.SP_LIST_NAME || 'SFS_SAF_Tracker';

module.exports = async function (context, req) {
  try {
    // ── 1. VERIFIED IDENTITY (cannot be forged by the browser) ──────
    const principal = parsePrincipal(req);
    if (!principal) {
      return done(context, 401, { error: 'Not authenticated' });
    }
    const me = identityOf(principal);     // { oid, email, roles[] }
    const role = roleOf(me.roles);        // 'admin' | 'auditor' | 'client'

    // The user's Easy Auth access token (for OBO). Injected by SWA.
    const userToken = req.headers['x-ms-token-aad-access-token'];
    if (!userToken) {
      return done(context, 401, {
        error: 'No user token — Easy Auth not configured for token store. ' +
               'Enable token store / loginParameters offline_access + Sites.ReadWrite.All.'
      });
    }

    // OBO exchange → Graph token, then resolve site + list ids.
    const graphToken = await oboToken(userToken);
    const { siteId, listId } = await resolveList(graphToken);

    // Load the operator roster (server-side, authoritative scope source).
    const operators = JSON.parse(await spGet(graphToken, siteId, listId, 'operators') || '[]');

    // Which operator ids may THIS user touch? Decided here, not in browser.
    const allowedOpIds = scopeFor(role, me.oid, operators);

    // ── 2. ROUTE ────────────────────────────────────────────────────
    if (req.method === 'GET') {
      const key = (req.query.key || '').trim();
      if (!key) return done(context, 400, { error: 'Missing key' });

      if (key === 'operators') {
        // Return only operators this user is allowed to see.
        const visible = (role === 'admin' || role === 'auditor')
          ? operators
          : operators.filter(o => allowedOpIds.includes(o.id));
        return done(context, 200, { value: JSON.stringify(visible) });
      }

      if (key.startsWith('state_')) {
        const opId = key.slice('state_'.length);
        if (!allowedOpIds.includes(opId)) {
          return done(context, 403, { error: 'Not authorised for this operator' });
        }
        const v = await spGet(graphToken, siteId, listId, key);
        return done(context, 200, { value: v });
      }

      return done(context, 400, { error: 'Unknown key' });
    }

    if (req.method === 'POST') {
      // Auditors are read-only — enforced HERE, not just hidden in the UI.
      if (role === 'auditor') {
        return done(context, 403, { error: 'Auditor role is read-only' });
      }
      const body = req.body || {};
      const key = (body.key || '').trim();
      const value = body.value;
      if (!key || typeof value !== 'string') {
        return done(context, 400, { error: 'Missing key or value' });
      }

      if (key === 'operators') {
        // Only admins may rewrite the operator roster / assignments.
        if (role !== 'admin') {
          return done(context, 403, { error: 'Only SFS admin can modify operators' });
        }
        // Single-client build: refuse a roster with more than one operator,
        // even if the UI cap is bypassed by POSTing here directly.
        try {
          const parsed = JSON.parse(value);
          if (Array.isArray(parsed) && parsed.length > MAX_OPERATORS) {
            return done(context, 400, {
              error: 'Single-client build: at most ' + MAX_OPERATORS + ' operator allowed'
            });
          }
        } catch {
          return done(context, 400, { error: 'operators value must be a JSON array' });
        }
        await spSet(graphToken, siteId, listId, key, value);
        return done(context, 200, { ok: true });
      }

      if (key.startsWith('state_')) {
        const opId = key.slice('state_'.length);
        if (!allowedOpIds.includes(opId)) {
          return done(context, 403, { error: 'Not authorised for this operator' });
        }
        await spSet(graphToken, siteId, listId, key, value);
        return done(context, 200, { ok: true });
      }

      return done(context, 400, { error: 'Unknown key' });
    }

    return done(context, 405, { error: 'Method not allowed' });

  } catch (e) {
    context.log.error('data fn error', e);
    return done(context, 500, { error: 'Server error', detail: String(e.message || e) });
  }
};

// ── identity helpers ──────────────────────────────────────────────
function parsePrincipal(req) {
  const h = req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf8')); }
  catch { return null; }
}

function identityOf(p) {
  // claims is an array of { typ, val }. oid is the stable AAD object id.
  const claims = p.claims || [];
  const get = (t) => {
    const c = claims.find(c => c.typ === t || (c.typ || '').endsWith('/' + t));
    return c ? c.val : null;
  };
  const oid = get('http://schemas.microsoft.com/identity/claims/objectidentifier')
            || get('oid');
  const email = (get('preferred_username') || p.userDetails || '').toLowerCase();
  const roles = (p.userRoles || []).filter(r => r !== 'anonymous' && r !== 'authenticated');
  return { oid, email, roles };
}

function roleOf(roles) {
  if (roles.includes('sfs.admin')) return 'admin';
  if (roles.includes('sfs.supervisor')) return 'admin'; // supervisor = full admin privileges
  if (roles.includes('auditor')) return 'auditor';
  if (roles.includes('client.editor')) return 'client';
  return 'client'; // authenticated but unassigned → least privilege, never admin
}

// Server-side scope decision. For a client.editor, match the caller's
// verified oid against each operator's assignedOids[]. No email trust,
// no browser input.
function scopeFor(role, oid, operators) {
  if (role === 'admin' || role === 'auditor') return operators.map(o => o.id);
  return operators
    .filter(o => Array.isArray(o.assignedOids) && o.assignedOids.includes(oid))
    .map(o => o.id);
}

// ── OBO token exchange ────────────────────────────────────────────
async function oboToken(userAssertion) {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    assertion: userAssertion,
    scope: 'https://graph.microsoft.com/.default',
    requested_token_use: 'on_behalf_of'
  });
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  });
  if (!r.ok) throw new Error('OBO exchange failed: ' + (await r.text()));
  return (await r.json()).access_token;
}

// ── SharePoint via Graph (as the user) ────────────────────────────
let _siteCache = null;
async function resolveList(token) {
  if (_siteCache) return _siteCache;
  const host = SITE_URL.replace('https://', '').split('/sites/')[0];
  const path = '/sites/' + SITE_URL.split('/sites/')[1];
  const sr = await fetch(`${GRAPH}/sites/${host}:${path}`,
    { headers: { Authorization: 'Bearer ' + token } });
  if (!sr.ok) throw new Error('Site lookup failed: ' + (await sr.text()));
  const siteId = (await sr.json()).id;
  const lr = await fetch(
    `${GRAPH}/sites/${siteId}/lists?$filter=displayName eq '${LIST_NAME}'`,
    { headers: { Authorization: 'Bearer ' + token } });
  const listId = (await lr.json()).value[0]?.id;
  if (!listId) throw new Error('List not found: ' + LIST_NAME);
  _siteCache = { siteId, listId };
  return _siteCache;
}

async function spGet(token, siteId, listId, key) {
  const r = await fetch(
    `${GRAPH}/sites/${siteId}/lists/${listId}/items?$filter=fields/Title eq '${encodeURIComponent(key)}'&$expand=fields`,
    { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  const d = await r.json();
  if (d.value && d.value.length) return d.value[0].fields.DataValue || null;
  return null;
}

async function spSet(token, siteId, listId, key, value) {
  // find existing
  const r = await fetch(
    `${GRAPH}/sites/${siteId}/lists/${listId}/items?$filter=fields/Title eq '${encodeURIComponent(key)}'&$expand=fields`,
    { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } });
  const d = await r.json();
  const existing = d.value && d.value[0];
  if (existing) {
    await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${existing.id}/fields`, {
      method: 'PATCH',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ DataValue: value })
    });
  } else {
    await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: { Title: key, DataValue: value } })
    });
  }
}

function done(context, status, body) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json' },
    body
  };
}
