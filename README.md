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

## Run the frontend

```bash
cd frontend
npm install
npm run dev        # http://localhost:5173
```

## Run the backend

```bash
cd backend
npm install
cp .env.example .env   # fill in secrets + provider credentials (see backend/README.md)
npm run dev            # http://localhost:4000
```

The frontend currently uses mock data. To connect it to the backend, swap the
mock `mailService` in `frontend/src/Beacon.jsx` for
`backend/client-mailService.js` (drop-in replacement) — details in backend/README.md.
