import express from "express";
import { getProvider } from "../providers/index.js";
import { store } from "../store.js";
import { syncEngine } from "../sync.js";

export const apiRouter = express.Router();

import crypto from "node:crypto";

function userId(req) {
  return req.session.userId || null;
}

// Mint a session identity on demand (used by profile creation, which is the
// entry point before any OAuth flow has run).
function ensureUserId(req) {
  if (!req.session.userId) req.session.userId = "u_" + crypto.randomBytes(6).toString("hex");
  return req.session.userId;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function validateProfile(body) {
  const firstName = (body?.firstName || "").trim();
  const lastName = (body?.lastName || "").trim();
  const email = (body?.email || "").trim().toLowerCase();
  if (!firstName || firstName.length > 50) return { error: "First name is required (max 50 chars)" };
  if (!lastName || lastName.length > 50) return { error: "Last name is required (max 50 chars)" };
  if (!EMAIL_RE.test(email) || email.length > 254) return { error: "A valid email is required" };
  return { profile: { firstName, lastName, email } };
}

/* ---- auth + profile (these run BEFORE the session guard) ----
   Passwords are hashed with Node scrypt (N=16384 default) + per-user salt and
   compared with timingSafeEqual. Production hardening notes: swap scryptSync
   for the async form under load, and add rate limiting on /auth/login. */

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return { salt, passwordHash: hash };
}

function verifyPassword(password, salt, passwordHash) {
  try {
    const candidate = crypto.scryptSync(password, salt, 64);
    const stored = Buffer.from(passwordHash, "hex");
    return candidate.length === stored.length && crypto.timingSafeEqual(candidate, stored);
  } catch {
    return false;
  }
}

function validPassword(pw) {
  return typeof pw === "string" && pw.length >= 8 && pw.length <= 128;
}

// POST /api/dev/reset — DEV ONLY: wipe all users, accounts, and tokens, and
// end the current session. Refuses to run in production.
apiRouter.post("/dev/reset", (req, res) => {
  if (process.env.NODE_ENV === "production") {
    return res.status(403).json({ error: "Disabled in production" });
  }
  store.reset();
  req.session.destroy(() => res.json({ ok: true, wiped: true }));
});

// GET /api/me → { user } or { user: null }
apiRouter.get("/me", (req, res) => {
  const uid = userId(req);
  res.json({ user: uid ? store.getUser(uid) : null });
});

// POST /api/auth/signup { firstName, lastName, email, password }
apiRouter.post("/auth/signup", (req, res) => {
  const { profile, error } = validateProfile(req.body);
  if (error) return res.status(400).json({ error });
  if (!validPassword(req.body?.password)) {
    return res.status(400).json({ error: "Password must be 8–128 characters" });
  }
  if (store.findUserByEmail(profile.email)) {
    return res.status(409).json({ error: "An account with that email already exists" });
  }
  // Adopt the existing session id if present, so anything connected
  // pre-signup (e.g. a Slack OAuth) stays attached to this user.
  const uid = ensureUserId(req);
  const user = store.createUser(uid, { ...profile, ...hashPassword(req.body.password) });
  res.json({ user });
});

// POST /api/auth/login { email, password } → binds this session to the user
apiRouter.post("/auth/login", (req, res) => {
  const found = store.findUserByEmail(req.body?.email);
  if (!found || !verifyPassword(req.body?.password || "", found.salt, found.passwordHash)) {
    return res.status(401).json({ error: "Invalid email or password" });
  }
  req.session.userId = found.id;
  res.json({ user: store.getUser(found.id) });
});

// POST /api/auth/logout
apiRouter.post("/auth/logout", (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// PATCH /api/me { firstName?, lastName?, email? } — profile updates
apiRouter.patch("/me", (req, res) => {
  const uid = userId(req);
  if (!uid || !store.getUser(uid)) return res.status(401).json({ error: "Not signed in" });
  const merged = { ...store.getUser(uid), ...req.body };
  const { profile, error } = validateProfile(merged);
  if (error) return res.status(400).json({ error });
  const existing = store.findUserByEmail(profile.email);
  if (existing && existing.id !== uid) {
    return res.status(409).json({ error: "That email is already in use" });
  }
  res.json({ user: store.updateUser(uid, profile) });
});

// PATCH /api/me/password { currentPassword, newPassword }
apiRouter.patch("/me/password", (req, res) => {
  const uid = userId(req);
  const auth = uid && store.getUserAuth(uid);
  if (!auth) return res.status(401).json({ error: "Not signed in" });
  if (!verifyPassword(req.body?.currentPassword || "", auth.salt, auth.passwordHash)) {
    return res.status(401).json({ error: "Current password is incorrect" });
  }
  if (!validPassword(req.body?.newPassword)) {
    return res.status(400).json({ error: "New password must be 8–128 characters" });
  }
  store.updateUser(uid, hashPassword(req.body.newPassword));
  res.json({ ok: true });
});

// Require a session for everything else under /api.
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
  const { read, starred, folder, move } = req.body;
  try {
    if (move) await provider.move(acc, req.params.id, move, folder);
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
  const { to, subject, body, html } = req.body || {};
  if (!to) return res.status(400).json({ error: "to is required" });
  try {
    await getProvider(acc.provider).send(acc, { to, subject: subject || "", body: body || "", html });
    res.json({ sent: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

/* ---- sync ---- */

/* ---- slack ---- */

function loadSlackAccount(req, res) {
  const acc = loadAccount(req, res);
  if (acc && acc.provider !== "slack") {
    res.status(400).json({ error: "Not a Slack account" });
    return null;
  }
  return acc;
}

// GET /api/slack/:accountId/conversations → channels + DMs (+ who you are)
apiRouter.get("/slack/:accountId/conversations", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  try {
    const slack = getProvider("slack");
    const conversations = await slack.listConversations(acc);
    const self = await slack.getSelf(acc).catch(() => null);
    res.json({ conversations, team: acc.secrets?.team_name, self });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// GET /api/slack/:accountId/members → workspace directory for @mentions
apiRouter.get("/slack/:accountId/members", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  try {
    const members = await getProvider("slack").listMembers(acc);
    res.json({ members });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// Slack rate limits (esp. conversations.history at ~1/min) surface as HTTP 429
// with a retryAfter the client uses to schedule an automatic retry.
function slackError(res, e) {
  if (e.rateLimited) {
    return res.status(429).json({ error: "Slack is rate-limiting this request", retryAfter: e.retryAfter || 60 });
  }
  res.status(502).json({ error: e.message });
}

// GET /api/slack/:accountId/conversations/:channelId/messages
apiRouter.get("/slack/:accountId/conversations/:channelId/messages", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  try {
    const { messages, stale } = await getProvider("slack").listMessages(acc, req.params.channelId, {
      limit: Math.min(Number(req.query.limit) || 30, 100),
    });
    res.json({ messages, stale });
  } catch (e) {
    slackError(res, e);
  }
});

// GET /api/slack/:accountId/conversations/:channelId/thread/:ts → parent + replies
apiRouter.get("/slack/:accountId/conversations/:channelId/thread/:ts", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  try {
    const messages = await getProvider("slack").listReplies(acc, req.params.channelId, req.params.ts);
    res.json({ messages });
  } catch (e) {
    slackError(res, e);
  }
});

// POST /api/slack/:accountId/conversations/:channelId/messages { text, thread_ts? }
apiRouter.post("/slack/:accountId/conversations/:channelId/messages", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  const { text, thread_ts } = req.body || {};
  if (!text) return res.status(400).json({ error: "text is required" });
  try {
    await getProvider("slack").postMessage(acc, req.params.channelId, text, thread_ts || null);
    res.json({ sent: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// POST /api/slack/:accountId/conversations/:channelId/reactions { timestamp, emoji, add }
apiRouter.post("/slack/:accountId/conversations/:channelId/reactions", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  const { timestamp, emoji, add } = req.body || {};
  if (!timestamp || !emoji) return res.status(400).json({ error: "timestamp and emoji are required" });
  try {
    await getProvider("slack").toggleReaction(acc, req.params.channelId, timestamp, emoji, add !== false);
    res.json({ ok: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// POST /api/slack/:accountId/conversations/:channelId/read → mark read up to now
apiRouter.post("/slack/:accountId/conversations/:channelId/read", async (req, res) => {
  const acc = loadSlackAccount(req, res);
  if (!acc) return;
  try {
    await getProvider("slack").markRead(acc, req.params.channelId);
    res.json({ ok: true });
  } catch (e) {
    // Surface this in the server log — a missing_scope here means the
    // workspace was connected before the *:write scopes were added and needs
    // to be reconnected, otherwise read state never syncs back to Slack.
    console.warn(`[slack] mark-read failed for ${req.params.channelId}: ${e.message}`);
    res.status(502).json({ error: e.message });
  }
});

/* ---- teams (chats) ----
   Teams is powered by the same Microsoft account as Outlook mail. Each connected
   Microsoft account surfaces as one Teams "org" whose chats come from Graph.
   Channels are omitted (they require admin-consented ChannelMessage.Read.All). */

function loadTeamsAccount(req, res) {
  const acc = loadAccount(req, res);
  if (acc && acc.provider !== "microsoft") {
    res.status(400).json({ error: "Not a Microsoft account" });
    return null;
  }
  return acc;
}

// GET /api/teams → one org per connected Microsoft account, with its chats
apiRouter.get("/teams", async (req, res) => {
  const accounts = store.listRawForUser(userId(req)).filter((a) => a.provider === "microsoft");
  const ms = getProvider("microsoft");
  const orgs = [];
  for (const acc of accounts) {
    try {
      const [identity, chats] = await Promise.all([ms.teamsIdentity(acc), ms.listChats(acc)]);
      orgs.push({
        id: acc.id,
        provider: "teams",
        org: identity.org,
        label: acc.label || identity.org,
        address: acc.address,
        you: identity.name,
        teams: [], // channels require admin consent — chats only for now
        chats,
      });
    } catch (e) {
      // A Microsoft account whose consent predates the Teams scopes will 403
      // here; surface it as an org with no chats rather than failing the list.
      orgs.push({
        id: acc.id, provider: "teams", org: acc.address, label: acc.label || acc.address,
        address: acc.address, you: "You", teams: [], chats: [], error: e.message,
      });
    }
  }
  res.json({ orgs });
});

// GET /api/teams/:accountId/chats/:chatId/messages
apiRouter.get("/teams/:accountId/chats/:chatId/messages", async (req, res) => {
  const acc = loadTeamsAccount(req, res);
  if (!acc) return;
  try {
    const messages = await getProvider("microsoft").listChatMessages(acc, req.params.chatId, {
      limit: Math.min(Number(req.query.limit) || 30, 50),
    });
    res.json({ messages });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// POST /api/teams/:accountId/chats/:chatId/messages { text }
apiRouter.post("/teams/:accountId/chats/:chatId/messages", async (req, res) => {
  const acc = loadTeamsAccount(req, res);
  if (!acc) return;
  const { text } = req.body || {};
  if (!text) return res.status(400).json({ error: "text is required" });
  try {
    await getProvider("microsoft").sendChatMessage(acc, req.params.chatId, text);
    res.json({ sent: true });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

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
