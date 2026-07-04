# Beacon Backend

OAuth + mail-API proxy for the Beacon client. Handles sign-in and read/write
for **Google (Gmail)**, **Microsoft (Graph)**, and **Apple (iCloud over IMAP/SMTP)**,
exposing one uniform REST API the React frontend consumes.

## Why a backend is required

The browser can't do this itself: OAuth client secrets must stay server-side,
and Gmail/Graph APIs and iCloud IMAP all block direct cross-origin browser calls.
This server holds the secrets, runs the OAuth flows, stores tokens encrypted, and
proxies every mail request.

## Quick start

```bash
cd beacon-backend
npm install
cp .env.example .env
# generate the two secrets the .env asks for:
node -e "console.log(require('crypto').randomBytes(48).toString('hex'))"  # SESSION_SECRET
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"  # TOKEN_ENC_KEY
# fill in provider credentials (below), then:
npm run dev
```

Visit `http://localhost:4000/health` — it reports which providers are configured.
The server runs fine with only one provider set up; the others just stay disabled.

## Registering each provider

**Google** — [console.cloud.google.com](https://console.cloud.google.com): create a
project, enable the **Gmail API**, create an **OAuth client ID (Web application)**.
Add redirect URI `http://localhost:4000/auth/google/callback`. Put the client ID
and secret in `.env`. While the app is in "testing," add your account as a test user.

**Microsoft** — [entra.microsoft.com](https://entra.microsoft.com) → **App
registrations** → **New registration**. Add a **Web** redirect URI
`http://localhost:4000/auth/microsoft/callback`. Under **API permissions** add
delegated `Mail.ReadWrite`, `Mail.Send`, `offline_access`, `User.Read`. Under
**Certificates & secrets** create a client secret. `MS_TENANT=common` supports
both personal and work/school accounts.

**Apple / iCloud** — nothing to register. iCloud mail has no API. The user creates
an **app-specific password** at [appleid.apple.com](https://appleid.apple.com) →
Sign-In & Security, and the client posts it to `/auth/apple/password`. The server
connects over IMAP/SMTP. (This requires the Apple ID to have 2FA enabled.)

## Wiring the frontend

`client-mailService.js` is a drop-in replacement for the mock `mailService` in
`Beacon.jsx`. Replace the mock object with the contents of that file. Then:

- **Google / Microsoft** connect buttons call `mailService.connectRedirect("google")`
  — a full-page redirect to the provider, returning to your app with `?connect=ok`.
  Call `handleConnectReturn(refresh)` once on mount to finish.
- **Apple** connect collects an email + app-specific password in your modal and
  calls `mailService.connectApple(address, appPassword)`.
- The label editor calls `mailService.setLabel(accountId, label)`.

Set `BACKEND` at the top of `client-mailService.js` to your server URL.

## API surface

```
GET    /health
GET    /auth/:provider                       redirect → OAuth consent (google|microsoft)
GET    /auth/:provider/callback              OAuth return → bounces to client
POST   /auth/apple/password                  { address, appPassword }

GET    /api/accounts                         connected accounts (no secrets)
PATCH  /api/accounts/:id                     { label }
DELETE /api/accounts/:id
GET    /api/accounts/:id/folders             folders + unread counts
GET    /api/accounts/:id/folders/:folder/messages?limit=&cursor=
                                             → { messages, nextCursor }
GET    /api/accounts/:id/messages/:msgId?folder=inbox
PATCH  /api/accounts/:id/messages/:msgId     { read?, starred?, folder? }
POST   /api/accounts/:id/send                { to, subject, body }
POST   /api/accounts/:id/refresh             force a poll → { folders, inboxUnread }
GET    /api/events                           SSE stream of unread-count changes
```

## Pagination

`listMessages` returns `{ messages, nextCursor }`. Pass `nextCursor` back as the
`cursor` query param to load the next page; `nextCursor: null` means the end.
The cursor is provider-specific and opaque to the client — Gmail uses its
`pageToken`, Graph a `$skip` offset, IMAP a sequence offset. On the client:

```js
let cursor;
const first = await mailService.listMessages(accId, "inbox");
cursor = first.nextCursor;
// ...on "load more":
if (cursor) {
  const next = await mailService.listMessages(accId, "inbox", { cursor });
  cursor = next.nextCursor;
}
```

## Live sync

A background poller checks each account's folder unread counts every
`SYNC_INTERVAL_MS` (default 60s) and pushes changes to the browser over SSE.
Subscribe once on mount:

```js
import { subscribeToUpdates } from "./client-mailService.js";

useEffect(() => subscribeToUpdates(({ accountId, inboxUnread, folders }) => {
  // update that account's sidebar badge from inboxUnread
}), []);
```

`POST /api/accounts/:id/refresh` forces an immediate poll — wire it to the
refresh button in the message list header.

Interval polling is the portable default. For true push, upgrade `pollAccount()`
in `src/sync.js` to Gmail `watch` + Pub/Sub, Microsoft Graph webhook
subscriptions, and IMAP `IDLE` — all need a publicly reachable callback URL.

## Docker

```bash
cp .env.example .env   # fill in secrets + provider credentials
docker compose up --build
```

The encrypted token store persists in the named volume `beacon-data`. The
container runs as a non-root user and exposes a `/health` healthcheck.

## Production notes

This scaffold favors clarity. Before real use:

- **Sessions/users** — `userId` is derived per browser session. Wire your own user
  auth and key accounts to real user IDs.
- **Token store** — `store.js` persists to an AES-256-GCM encrypted file. Swap
  `load()/save()` for a database; keep `TOKEN_ENC_KEY` in a secret manager.
- **Cookies** — set `SERVER_ORIGIN` to `https://…` in production so the session
  cookie is marked `Secure`; consider `sameSite:"none"` if the client is on a
  different site.
- **Rate/error handling** — providers throttle; add retries/backoff and pagination
  (cursors are stubbed at `limit`).
- **Scopes** — drop `gmail.send` / `Mail.Send` if you only need read access.

## Layout

```
beacon-backend/
├── .env.example
├── Dockerfile
├── docker-compose.yml
├── package.json
├── client-mailService.js        ← paste into the React app
└── src/
    ├── server.js                ← Express entry
    ├── config.js
    ├── store.js                 ← encrypted account/token store
    ├── sync.js                  ← polling engine + SSE event bus
    ├── routes/
    │   ├── auth.js              ← OAuth + Apple password connect
    │   └── api.js              ← accounts, folders, messages, send, refresh, events
    └── providers/
        ├── index.js
        ├── google.js           ← Gmail API
        ├── microsoft.js        ← Microsoft Graph
        └── apple.js            ← iCloud IMAP/SMTP
```
