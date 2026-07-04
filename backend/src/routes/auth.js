import express from "express";
import crypto from "node:crypto";
import { config, providerConfigured } from "../config.js";
import { getProvider } from "../providers/index.js";
import { store } from "../store.js";

export const authRouter = express.Router();

/* A real app authenticates the human first (their own login). For this scaffold
   we derive a stable per-session userId so each browser gets its own mailboxes. */
function userId(req) {
  if (!req.session.userId) req.session.userId = "u_" + crypto.randomBytes(6).toString("hex");
  return req.session.userId;
}

/* ---- OAuth providers (Google, Microsoft) ---- */

// GET /auth/:provider  → redirect the browser to the provider's consent screen
authRouter.get("/:provider", (req, res) => {
  const { provider } = req.params;
  if (provider === "apple") {
    return res.status(400).json({ error: "Apple connects via POST /auth/apple/password" });
  }
  if (!providerConfigured(provider)) {
    return res.status(500).json({ error: `${provider} is not configured on the server` });
  }

  userId(req);
  // CSRF state, tied to the session.
  const state = crypto.randomBytes(16).toString("hex");
  req.session.oauthState = state;
  req.session.oauthProvider = provider;

  try {
    const url = getProvider(provider).authUrl(state);
    res.redirect(url);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// GET /auth/:provider/callback  → exchange code, store account, bounce to client
authRouter.get("/:provider/callback", async (req, res) => {
  const { provider } = req.params;
  const { code, state, error } = req.query;

  if (error) return res.redirect(`${config.clientOrigin}/?connect=error&reason=${error}`);
  if (!state || state !== req.session.oauthState) {
    return res.redirect(`${config.clientOrigin}/?connect=error&reason=bad_state`);
  }

  try {
    const { address, secrets } = await getProvider(provider).handleCallback(code);
    const acc = store.upsert({ userId: userId(req), provider, address, secrets });
    delete req.session.oauthState;
    res.redirect(`${config.clientOrigin}/?connect=ok&account=${acc.id}`);
  } catch (e) {
    console.error(`[auth/${provider}] callback failed:`, e.message);
    res.redirect(`${config.clientOrigin}/?connect=error&reason=exchange_failed`);
  }
});

/* ---- Apple / iCloud (app-specific password, no redirect) ---- */

// POST /auth/apple/password  { address, appPassword }
authRouter.post("/apple/password", async (req, res) => {
  const { address, appPassword } = req.body || {};
  if (!address || !appPassword) {
    return res.status(400).json({ error: "address and appPassword are required" });
  }
  try {
    const apple = getProvider("apple");
    const { secrets } = await apple.verifyCredentials({ address, appPassword });
    const acc = store.upsert({ userId: userId(req), provider: "apple", address, secrets });
    res.json({ account: acc });
  } catch (e) {
    console.error("[auth/apple] verify failed:", e.message);
    res.status(401).json({ error: "Could not sign in to iCloud. Check the address and app-specific password." });
  }
});
