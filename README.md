# GoalLens SharePoint Mock

A dependency-free local contract simulator for the Microsoft identity token
endpoint and the subset of Microsoft Graph/SharePoint used by GoalLens.

It is intentionally separate from the GoalLens repository. It does not require
a Microsoft account, Microsoft 365 license, SharePoint tenant, API key, or
client secret.

## Start it

Requires Python 3.10 or newer.

```bash
python3 mock_server.py
```

The service listens on `http://127.0.0.1:8765` by default.

```bash
curl http://127.0.0.1:8765/health
```

## GoalLens development configuration

The GoalLens Graph client should make its authority and Graph base URLs
configurable. For local development, use:

```env
MICROSOFT_AUTH_BASE_URL=http://127.0.0.1:8765
MICROSOFT_GRAPH_BASE_URL=http://127.0.0.1:8765/v1.0
MICROSOFT_TENANT_ID=mocktenant
MICROSOFT_CLIENT_ID=mock-client-id
MICROSOFT_CLIENT_SECRET=mock-client-secret
```

The mock accepts any client credentials and returns a local bearer token. Graph
routes require an `Authorization: Bearer ...` header unless the scenario itself
is testing an invalid token. `MICROSOFT_AUTH_BASE_URL` also covers the
admin-consent redirect (`GET /organizations/v2.0/adminconsent`) — GoalLens's
"Connect with Microsoft" button reads it from the same override, so pointing
it here redirects that button here too, instead of to real Microsoft.

Production configuration must continue to use Microsoft's real endpoints. Do
not allow a user-supplied Graph base URL.

## Endpoints

- `POST /{tenant}/oauth2/v2.0/token`
- `GET /organizations/v2.0/adminconsent` — simulates Microsoft's tenant-level
  admin-consent screen: redirects straight back to the given `redirect_uri`
  with `admin_consent=True&tenant=mocktenant&state=...` (or, under the
  `admin_consent_denied` scenario, `error=access_denied&...`). No sign-in UI,
  no `client_id` validation — it's a stand-in for an admin clicking
  Accept/Cancel, not a real consent screen.
- `GET /v1.0/sites/{hostname}:/sites/{path}`
- `GET /v1.0/sites/{site-id}/drives`
- `GET /v1.0/drives/{drive-id}/items/{folder-id}/children`
- `GET /v1.0/drives/{drive-id}/root/delta`
- `GET /v1.0/drives/{drive-id}/items/{item-id}/content`

Local administration:

- `GET /health`
- `GET /__admin/state`
- `POST /__admin/scenario`
- `POST /__admin/reset`

## Scenarios

Change the active scenario:

```bash
curl -X POST http://127.0.0.1:8765/__admin/scenario \
  -H 'Content-Type: application/json' \
  -d '{"scenario":"new_version"}'
```

Available scenarios:

| Scenario | Behavior |
|---|---|
| `happy_path` | Two-page initial delta and no subsequent changes |
| `permission_denied` | Graph returns `403 accessDenied` |
| `invalid_token` | Graph returns `401 InvalidAuthenticationToken` |
| `throttled_once` | Each Graph route returns `429` once, then succeeds |
| `expired_delta` | Requests carrying a delta token return `410 resyncRequired` |
| `new_version` | Incremental delta returns a changed eTag and updated file |
| `deleted` | Incremental delta returns a Graph deleted facet |
| `restored` | Incremental delta returns the original item and eTag |
| `download_failure` | File content requests return `500` |
| `admin_consent_denied` | The admin-consent redirect comes back with `error=access_denied` instead of approval |

## Example

```bash
TOKEN=$(curl -s -X POST \
  http://127.0.0.1:8765/mocktenant/oauth2/v2.0/token | \
  python3 -c 'import json,sys; print(json.load(sys.stdin)["access_token"])')

curl -H "Authorization: Bearer $TOKEN" \
  'http://127.0.0.1:8765/v1.0/sites/goallensdev.sharepoint.com:/sites/GoalLensSandbox'
```

## Tests

```bash
python3 -m unittest -v
```

The server uses only Python's standard library, so there is no package install
or virtual environment to maintain.

## Scope and limitations

This validates GoalLens connector behavior, including pagination, cursor
handling, file downloads, throttling, versioning inputs, errors, and the
admin-consent redirect's request/response shape (approve/deny, state
round-trip). It cannot prove that a real Entra tenant actually grants
`Sites.Selected` consent, enforces tenant boundaries, or matches Microsoft's
live delta behavior — the admin-consent endpoint here approves unconditionally
by default, it doesn't model real Entra authorization logic. Those require a
final test against a real Microsoft 365 tenant.
