/*
 * SFS SAF Tracker — SharePoint data Function
 * Simple client credentials approach — app authenticates as itself.
 * No user token needed. Reads/writes SFS_SAF_Tracker SharePoint list.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';
let _token = null;
let _tokenExpiry = 0;
let _siteId = null;
let _listId = null;

async function getAppToken() {
  if (_token && Date.now() < _tokenExpiry) return _token;
  const url = `https://login.microsoftonline.com/${process.env.AAD_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: process.env.AAD_CLIENT_ID,
    client_secret: process.env.AAD_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  });
  const r = await fetch(url, { method: 'POST', body });
  if (!r.ok) throw new Error('Token failed: ' + await r.text());
  const d = await r.json();
  _token = d.access_token;
  _tokenExpiry = Date.now() + (d.expires_in - 60) * 1000;
  return _token;
}

async function resolveIds(token) {
  if (_siteId && _listId) return;
  const host = process.env.SP_SITE_URL.replace('https://','').split('/sites/')[0];
  const path = '/sites/' + process.env.SP_SITE_URL.split('/sites/')[1];
  const sr = await fetch(`${GRAPH}/sites/${host}:${path}`, { headers: { Authorization: 'Bearer ' + token } });
  if (!sr.ok) throw new Error('Site lookup failed: ' + sr.status);
  _siteId = (await sr.json()).id;
  const lr = await fetch(`${GRAPH}/sites/${_siteId}/lists?$filter=displayName eq '${process.env.SP_LIST_NAME || 'SFS_SAF_Tracker'}'`, { headers: { Authorization: 'Bearer ' + token } });
  const ld = await lr.json();
  _listId = ld.value?.[0]?.id;
  if (!_listId) throw new Error('List not found: ' + process.env.SP_LIST_NAME);
}

module.exports = async function(context, req) {
  try {
    const token = await getAppToken();
    await resolveIds(token);

    if (req.method === 'GET') {
      const key = (req.query.key || '').trim();
      if (!key) return respond(context, 400, { error: 'Missing key' });
      const r = await fetch(
        `${GRAPH}/sites/${_siteId}/lists/${_listId}/items?$filter=fields/Title eq '${encodeURIComponent(key)}'&$expand=fields`,
        { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } }
      );
      const d = await r.json();
      const val = d.value?.[0]?.fields?.DataValue || null;
      return respond(context, 200, { value: val });
    }

    if (req.method === 'POST') {
      const { key, value } = req.body || {};
      if (!key || typeof value !== 'string') return respond(context, 400, { error: 'Missing key or value' });
      // Check if exists
      const r = await fetch(
        `${GRAPH}/sites/${_siteId}/lists/${_listId}/items?$filter=fields/Title eq '${encodeURIComponent(key)}'&$expand=fields`,
        { headers: { Authorization: 'Bearer ' + token, Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' } }
      );
      const d = await r.json();
      const existing = d.value?.[0];
      if (existing) {
        await fetch(`${GRAPH}/sites/${_siteId}/lists/${_listId}/items/${existing.id}/fields`, {
          method: 'PATCH',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ DataValue: value })
        });
      } else {
        await fetch(`${GRAPH}/sites/${_siteId}/lists/${_listId}/items`, {
          method: 'POST',
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ fields: { Title: key, DataValue: value } })
        });
      }
      return respond(context, 200, { ok: true });
    }

    return respond(context, 405, { error: 'Method not allowed' });
  } catch(e) {
    context.log.error('data fn error', e);
    return respond(context, 500, { error: String(e.message) });
  }
};

function respond(context, status, body) {
  context.res = { status, headers: { 'Content-Type': 'application/json' }, body };
}
