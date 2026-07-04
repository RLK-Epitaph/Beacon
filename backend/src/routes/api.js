import express from "express";
import { getProvider } from "../providers/index.js";
import { store } from "../store.js";
import { syncEngine } from "../sync.js";

export const apiRouter = express.Router();

function userId(req) {
  return req.session.userId || null;
}

// Require a session for everything under /api.
apiRouter.use((req, res, next) => {
  if (!userId(req)) return res.status(401).json({ error: "Not signed in" });
  next();
});

// Resolve :accountId → raw account (with secrets) for this user.
function loadAccount(req, res) {
  const acc = store.getRaw(userId(req), req.params.accountId);
  if (!acc) {
    res.status(404).json({ error: "Account not found" });
    return null;
  }
  return acc;
}

/* ---- accounts ---- */

// GET /api/accounts → connected accounts (no secrets)
apiRouter.get("/accounts", (req, res) => {
  res.json({ accounts: store.listForUser(userId(req)) });
});

// PATCH /api/accounts/:accountId  { label }
apiRouter.patch("/accounts/:accountId", (req, res) => {
  const updated = store.setLabel(userId(req), req.params.accountId, req.body.label ?? "");
  if (!updated) return res.status(404).json({ error: "Account not found" });
  res.json({ account: updated });
});

// DELETE /api/accounts/:accountId
apiRouter.delete("/accounts/:accountId", (req, res) => {
  const ok = store.remove(userId(req), req.params.accountId);
  res.json({ removed: ok });
});

/* ---- folders + messages ---- */

// GET /api/accounts/:accountId/folders
apiRouter.get("/accounts/:accountId/folders", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  try {
    const folders = await getProvider(acc.provider).listFolders(acc);
    res.json({ folders });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// GET /api/accounts/:accountId/folders/:folder/messages?limit=&cursor=
apiRouter.get("/accounts/:accountId/folders/:folder/messages", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  const limit = Math.min(Number(req.query.limit) || 30, 100);
  const cursor = req.query.cursor || undefined;
  try {
    const result = await getProvider(acc.provider).listMessages(acc, req.params.folder, { limit, cursor });
    // Providers now return { messages, nextCursor }.
    res.json({ messages: result.messages, nextCursor: result.nextCursor ?? null });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// GET /api/accounts/:accountId/messages/:id?folder=inbox
apiRouter.get("/accounts/:accountId/messages/:id", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  try {
    const msg = await getProvider(acc.provider).getMessage(acc, req.params.id, req.query.folder);
    res.json({ message: msg });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// PATCH /api/accounts/:accountId/messages/:id  { read?, starred?, folder? }
apiRouter.patch("/accounts/:accountId/messages/:id", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  const provider = getProvider(acc.provider);
  const { read, starred, folder } = req.body;
  try {
    if (typeof read === "boolean") await provider.markRead(acc, req.params.id, read, folder);
    if (typeof starred === "boolean") await provider.toggleStar(acc, req.params.id, starred, folder);
    res.json({ ok: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// POST /api/accounts/:accountId/send  { to, subject, body }
apiRouter.post("/accounts/:accountId/send", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  const { to, subject, body } = req.body || {};
  if (!to) return res.status(400).json({ error: "to is required" });
  try {
    await getProvider(acc.provider).send(acc, { to, subject: subject || "", body: body || "" });
    res.json({ sent: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

/* ---- sync ---- */

// POST /api/accounts/:accountId/refresh → force an immediate poll, return counts
apiRouter.post("/accounts/:accountId/refresh", async (req, res) => {
  const acc = loadAccount(req, res);
  if (!acc) return;
  try {
    const folders = await syncEngine.refreshAccount(acc);
    res.json({ folders, inboxUnread: folders.inbox || 0 });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// GET /api/events → Server-Sent Events stream of unread-count changes for this user
apiRouter.get("/events", (req, res) => {
  const uid = userId(req);
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
  });
  res.write(`event: ready\ndata: {"ok":true}\n\n`);

  const onChange = (payload) => {
    if (payload.userId !== uid) return; // only this user's accounts
    res.write(`event: change\ndata: ${JSON.stringify(payload)}\n\n`);
  };
  syncEngine.on("change", onChange);

  // Heartbeat keeps proxies from closing the idle connection.
  const heartbeat = setInterval(() => res.write(`: ping\n\n`), 25_000);

  req.on("close", () => {
    clearInterval(heartbeat);
    syncEngine.off("change", onChange);
  });
});
