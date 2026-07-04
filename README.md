# Beacon

A unified workspace client: multiple email accounts (Google, Apple, Microsoft),
Slack workspaces, and Microsoft Teams orgs in one app, with a home dashboard
(clock, weather, calendar aggregated across accounts, Slack Lists).

## Structure

```
beacon/
├── frontend/   React app (Vite). UI with realistic mock data; service layer
│               marked with REPLACE comments for wiring real APIs.
└── backend/    Express server: OAuth flows (Google, Microsoft), iCloud IMAP,
                encrypted token store, mail API proxy, pagination, SSE sync.
```

## Quick start (npm workspaces)

One install from the repo root sets up both halves:

```bash
npm install        # installs frontend + backend dependencies
npm run dev        # starts both: frontend on :5173, backend on :4000
```

Or run either side alone:

```bash
npm run dev:frontend   # just the UI (mock data — no backend needed)
npm run dev:backend    # just the API server
```

The backend boots with safe dev defaults; to use real providers,
`cp backend/.env.example backend/.env` and fill in credentials
(see backend/README.md).

The frontend currently uses mock data. To connect it to the backend, swap the
mock `mailService` in `frontend/src/Beacon.jsx` for
`backend/client-mailService.js` (drop-in replacement) — details in backend/README.md.
