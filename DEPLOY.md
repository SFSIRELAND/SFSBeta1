# SFS SAF Compliance Tracker — Beta: Clean Deploy Runbook

**Single-client build (one operator/facility).** `MAX_OPS = 1`, enforced
in the browser *and* server-side — the Function rejects any operator roster
with more than one entry, so the cap can't be bypassed by POSTing directly
to `/api/data`.

**Path B — server-side RBAC.** The browser no longer holds a Graph
token or talks to SharePoint directly. All data goes through `/api/data`,
a Function that verifies identity from the un-forgeable
`x-ms-client-principal` header and enforces operator scope **server-side**
before any SharePoint call.

Commands you run yourself. I can't touch your Azure tenant — provisioning,
deleting, AAD registration and consent grants are all credentialed actions
that are yours to execute deliberately. Where a real secret is involved, it
never appears in any file; you set it directly in Azure.

---

## 0. Prerequisites (one-time)

```bash
# Azure CLI + SWA CLI
az --version            # install: https://learn.microsoft.com/cli/azure/install-azure-cli
npm install -g @azure/static-web-apps-cli
az login                # sign in to the SFS tenant
```

You'll need: **Contributor** on the subscription/resource group, and rights
to create an **App registration** in Entra ID.

---

## 1. Azure AD App registration (the auth + OBO identity)

This single registration does double duty: Easy Auth login AND the
on-behalf-of exchange the Function uses to reach SharePoint.

```bash
# Create the app registration
az ad app create --display-name "SFS SAF Tracker Beta" \
  --sign-in-audience AzureADMyOrg

# Note the appId (this is AAD_CLIENT_ID) and the directory tenant id.
```

Then in the **Entra portal → App registrations → SFS SAF Tracker Beta**:

**a) Authentication**
- Add a **Web** platform.
- Redirect URI (fill in once you have the SWA hostname, step 3):
  `https://<your-swa-name>.azurestaticapps.net/.auth/login/aad/callback`
- Enable **ID tokens**.

**b) Certificates & secrets**
- New client secret. Copy the **value** immediately — you'll set it in step 4.
  This is `AAD_CLIENT_SECRET`. Do **not** put it in any file or in GitHub.

**c) API permissions (delegated — this is the secure model)**
- Microsoft Graph → **Delegated**:
  - `openid`, `profile`, `email`, `offline_access`
  - `Sites.ReadWrite.All`  *(delegated — acts AS the user; SharePoint
    permissions remain a second backstop)*
- Click **Grant admin consent** for the tenant.

**d) App roles** (Entra → App registrations → App roles → Create)
Create three, each *Allowed member type: Users/Groups*:
| Display name | Value          | 
|--------------|----------------|
| SFS Admin    | `sfs.admin`    |
| Client Editor| `client.editor`|
| Auditor      | `auditor`      |

**e) Assign roles to people**
Entra → **Enterprise applications** → SFS SAF Tracker Beta → Users and
groups → assign each user the right role. Unassigned authenticated users
get least privilege (client with no operators), never admin.

---

## 2. SharePoint list (the datastore)

In the target SharePoint site (`SP_SITE_URL`), create a list named
**`SFS_SAF_Tracker`** with:
- **Title** (default single line — used as the data key)
- **DataValue** — *Multiple lines of text*, plain text, no versioning limit
  issues. (Stores the JSON blob per key.)

**Lock the library/list down**: this is the real backstop. Set list
permissions so only SFS admins + the specific assigned client users have
access. The Function acts *as the user* (delegated), so SharePoint's own
permissions enforce a second layer — but only if you've set them. A
tenant-wide-shared list undoes the scoping. This step is yours and it matters.

---

## 3. Create the Static Web App (fresh — alongside the old one)

> Keep the old app running until this one is verified green. Don't delete
> first — there's no upside to a window where nothing's live. Teardown is
> step 7.

```bash
RG=sfs-rg                      # your resource group
NAME=sfs-tracker-beta
LOCATION=westeurope

az staticwebapp create \
  --name $NAME \
  --resource-group $RG \
  --location $LOCATION \
  --sku Standard               # Standard needed for AAD custom config + auth

# Get the default hostname — use it for the redirect URI in step 1a.
az staticwebapp show --name $NAME --resource-group $RG \
  --query "defaultHostname" -o tsv
```

Go back and add the redirect URI from step 1a now that you have the hostname.

---

## 4. Set application settings (secrets live here, NOT in code)

```bash
az staticwebapp appsettings set --name $NAME --resource-group $RG \
  --setting-names \
    AAD_TENANT_ID="<tenant-guid>" \
    AAD_CLIENT_ID="<app-appId>" \
    AAD_CLIENT_SECRET="<the-secret-value-from-step-1b>" \
    SP_SITE_URL="https://<yourco>.sharepoint.com/sites/<site>" \
    SP_LIST_NAME="SFS_SAF_Tracker"
```

The `staticwebapp.config.json` references `AAD_CLIENT_ID` /
`AAD_CLIENT_SECRET` by **name** — the values resolve from these settings at
runtime. Also replace `<AAD_TENANT_ID>` in `staticwebapp.config.json`'s
`openIdIssuer` with your real tenant guid before deploying (it's in the URL,
not a secret).

---

## 5. Deploy — SWA CLI, no GitHub Actions

```bash
cd sfs-tracker-beta

# Get a deployment token (treat as a secret; don't commit it)
az staticwebapp secrets list --name $NAME --resource-group $RG \
  --query "properties.apiKey" -o tsv

# Deploy app + api in one shot
swa deploy ./ \
  --api-location ./api \
  --deployment-token "<token-from-above>" \
  --env production
```

That's it — no Git, no Actions, no token-in-repo. `swa deploy` pushes the
static files and the Functions directly.

---

## 6. Test both modes

**Demo/local mode** (no Azure needed):
```bash
# Open index.html locally or `npx serve`. CONFIG still has placeholders →
# SP_MODE is false → localStorage, full demo access. Confirms UI works.
```

**Easy Auth mode** (on Azure):
1. Visit `https://<name>.azurestaticapps.net` → should redirect to AAD login.
2. Log in as an **admin** user → see all operators, can add/edit.
3. Add an operator, paste a **client user's Object ID** (Entra → Users →
   their Object ID) into the new "Client AAD Object ID(s)" field.
4. Log in (incognito) as that **client** user → see only their operator(s).
   Open dev tools, try `fetch('/api/data?key=state_<other-op-id>')` → must
   return **403**. That 403 is the whole point — the server refuses, not the UI.
5. Log in as an **auditor** → can read, but `POST /api/data` returns **403**.

If step 4's 403 doesn't happen, stop — the scope isn't enforcing; check the
operator's `assignedOids` and the user's app-role assignment.

---

## 7. Teardown the old app (only after the above is green)

Deliberate, irreversible. Run it yourself, eyes on the portal:
```bash
# Confirm you're naming the OLD app, not the new one:
az staticwebapp list --resource-group <old-rg> -o table

az staticwebapp delete --name <OLD-APP-NAME> --resource-group <old-rg>
```
Also remove the old app's redirect URI from the App registration if it was
a separate registration.

---

## What this build does and doesn't guarantee

**Does:** access control is genuinely server-enforced. A client.editor
cannot pull another operator's data — the server checks their verified oid
against the assignment list before Graph is touched. Auditors can't write.
That closes the original hole.

**Doesn't:** this is beta. It's sound architecture, not a security audit.
Three things sit in your tenant, not in this code: (1) SharePoint list
permissions — the real backstop, set them tight; (2) the client secret —
keep it only in app settings, never in a repo; (3) before *real* Bia Energy
evidence of record goes in, have someone who does security professionally
review it. For a regulated cert programme, that sign-off matters.
