import dotenv from "dotenv";
dotenv.config();

function required(name, fallback) {
  const v = process.env[name] ?? fallback;
  if (v === undefined) {
    console.warn(`[config] Missing env var ${name}`);
  }
  return v;
}

export const config = {
  port: Number(process.env.PORT || 4000),
  clientOrigin: required("CLIENT_ORIGIN", "http://localhost:5173"),
  serverOrigin: required("SERVER_ORIGIN", "http://localhost:4000"),
  sessionSecret: required("SESSION_SECRET", "dev-insecure-secret"),
  tokenEncKey: required("TOKEN_ENC_KEY", ""),

  google: {
    clientId: process.env.GOOGLE_CLIENT_ID || "",
    clientSecret: process.env.GOOGLE_CLIENT_SECRET || "",
    // Gmail read + modify (mark read/star) + send. Drop gmail.send if you only read.
    scopes: [
      "https://www.googleapis.com/auth/gmail.modify",
      "https://www.googleapis.com/auth/gmail.send",
      "https://www.googleapis.com/auth/userinfo.email",
      "openid",
    ],
  },

  microsoft: {
    clientId: process.env.MS_CLIENT_ID || "",
    clientSecret: process.env.MS_CLIENT_SECRET || "",
    tenant: process.env.MS_TENANT || "common",
    scopes: [
      "openid",
      "email",
      "offline_access",
      "User.Read",
      "Mail.ReadWrite",
      "Mail.Send",
      // NOTE: no Teams scopes here. Chat.ReadWrite is normally user-consentable,
      // but this tenant blocks all user consent ("needs admin approval"), and
      // including it would block the MAIL connect too. Re-add once a tenant
      // admin grants the Beacon app consent — the Teams chat backend
      // (providers/microsoft.js + /api/teams) is built and dormant.
    ],
  },

  slack: {
    clientId: process.env.SLACK_CLIENT_ID || "",
    clientSecret: process.env.SLACK_CLIENT_SECRET || "",
    // User-token scopes — Beacon acts as the signed-in person, not a bot.
    // The *:write scopes are required by conversations.mark (moving the read
    // cursor when a channel is opened) — without them, mark-as-read silently
    // fails with missing_scope and unread state never clears on Slack's side.
    // Workspaces connected before these were added must be RECONNECTED to
    // grant them (Connect account → Slack; tokens refresh in place).
    userScopes: [
      "channels:read", "channels:history", "channels:write",
      "groups:read", "groups:history", "groups:write",
      "im:read", "im:history", "im:write",
      "mpim:read", "mpim:history", "mpim:write",
      "chat:write", "reactions:read", "reactions:write",
      "users:read",
    ],
  },

  apple: {
    imapHost: process.env.IMAP_HOST || "imap.mail.me.com",
    imapPort: Number(process.env.IMAP_PORT || 993),
    smtpHost: process.env.SMTP_HOST || "smtp.mail.me.com",
    smtpPort: Number(process.env.SMTP_PORT || 587),
  },
};

export function providerConfigured(provider) {
  if (provider === "google") return !!(config.google.clientId && config.google.clientSecret);
  if (provider === "microsoft") return !!(config.microsoft.clientId && config.microsoft.clientSecret);
  if (provider === "apple") return true; // per-account credentials, always "available"
  if (provider === "slack") return !!(config.slack.clientId && config.slack.clientSecret);
  return false;
}
