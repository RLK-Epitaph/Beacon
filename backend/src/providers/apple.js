import { ImapFlow } from "imapflow";
import { simpleParser } from "mailparser";
import nodemailer from "nodemailer";
import { config } from "../config.js";

/**
 * Apple has NO mail API and NO OAuth path for mail. iCloud mail is reached over
 * IMAP (read) and SMTP (send), authenticated with an APP-SPECIFIC PASSWORD the
 * user creates at appleid.apple.com → Sign-In & Security → App-Specific Passwords.
 *
 * So "connect" here is a form post (address + app password), not a redirect.
 * Credentials are stored encrypted in the same account store as the OAuth tokens.
 */

const FOLDER_MAP = {
  inbox: "INBOX",
  sent: "Sent Messages",
  drafts: "Drafts",
  spam: "Junk",
  trash: "Deleted Messages",
  archive: "Archive",
};

async function withClient(account, fn) {
  const client = new ImapFlow({
    host: config.apple.imapHost,
    port: config.apple.imapPort,
    secure: true,
    auth: { user: account.secrets.address, pass: account.secrets.app_password },
    logger: false,
  });
  await client.connect();
  try {
    return await fn(client);
  } finally {
    await client.logout().catch(() => {});
  }
}

function mapEnvelope(msg, folder) {
  const from = msg.envelope?.from?.[0];
  return {
    id: String(msg.uid),
    from: from?.name || from?.address || "Unknown",
    fromAddr: from ? `${from.mailbox}@${from.host}` : "",
    subject: msg.envelope?.subject || "(no subject)",
    preview: "",
    date: msg.envelope?.date ? new Date(msg.envelope.date).toLocaleString() : "",
    fullDate: msg.envelope?.date ? new Date(msg.envelope.date).toLocaleString() : "",
    unread: !msg.flags?.has("\\Seen"),
    starred: msg.flags?.has("\\Flagged") || false,
    folder,
    attachments: [],
  };
}

export const appleProvider = {
  id: "apple",

  /**
   * No redirect-based auth. The route layer calls verifyCredentials() instead of
   * authUrl()/handleCallback(). We expose a marker so routes can branch.
   */
  usesPassword: true,

  async verifyCredentials({ address, appPassword }) {
    const client = new ImapFlow({
      host: config.apple.imapHost,
      port: config.apple.imapPort,
      secure: true,
      auth: { user: address, pass: appPassword },
      logger: false,
    });
    await client.connect(); // throws if credentials are wrong
    await client.logout().catch(() => {});
    return { address, secrets: { address, app_password: appPassword } };
  },

  async listFolders(account) {
    return withClient(account, async (client) => {
      const out = [];
      for (const [id, name] of Object.entries(FOLDER_MAP)) {
        try {
          const status = await client.status(name, { unseen: true, messages: true });
          out.push({ id, name: id[0].toUpperCase() + id.slice(1), unread: status.unseen || 0 });
        } catch {
          out.push({ id, name: id, unread: 0 });
        }
      }
      out.push({ id: "starred", name: "Starred", unread: 0 });
      return out;
    });
  },

  // Returns { messages, nextCursor }. Cursor is a numeric offset counting back
  // from the newest message in the mailbox.
  async listMessages(account, folder, { limit = 30, cursor } = {}) {
    const mailbox = folder === "starred" ? "INBOX" : FOLDER_MAP[folder] || "INBOX";
    const offset = Number(cursor) || 0;
    return withClient(account, async (client) => {
      const lock = await client.getMailboxLock(mailbox);
      try {
        const total = client.mailbox.exists;
        if (!total) return { messages: [], nextCursor: null };
        // Window of sequence numbers: newest-first, skipping `offset` already seen.
        const end = total - offset; // highest seq in this page
        if (end < 1) return { messages: [], nextCursor: null };
        const start = Math.max(1, end - limit + 1);
        const range = `${start}:${end}`;
        const results = [];
        for await (const msg of client.fetch(range, { envelope: true, flags: true })) {
          const mapped = mapEnvelope(msg, folder);
          if (folder === "starred" && !mapped.starred) continue;
          results.push(mapped);
        }
        results.reverse(); // newest first
        const consumed = end - start + 1;
        const nextCursor = start > 1 ? String(offset + consumed) : null;
        return { messages: results, nextCursor };
      } finally {
        lock.release();
      }
    });
  },

  async getMessage(account, id, folder = "inbox") {
    const mailbox = FOLDER_MAP[folder] || "INBOX";
    return withClient(account, async (client) => {
      const lock = await client.getMailboxLock(mailbox);
      try {
        const msg = await client.fetchOne(id, { source: true, flags: true, envelope: true }, { uid: true });
        const parsed = await simpleParser(msg.source);
        return {
          id: String(id),
          from: parsed.from?.value?.[0]?.name || "",
          fromAddr: parsed.from?.value?.[0]?.address || "",
          subject: parsed.subject || "(no subject)",
          fullDate: parsed.date ? parsed.date.toLocaleString() : "",
          body: parsed.text || stripHtml(parsed.html || ""),
          unread: !msg.flags?.has("\\Seen"),
          starred: msg.flags?.has("\\Flagged") || false,
          attachments: (parsed.attachments || []).map((a) => ({
            name: a.filename || "attachment",
            size: a.size ? `${Math.round(a.size / 1024)} KB` : "",
          })),
        };
      } finally {
        lock.release();
      }
    });
  },

  async markRead(account, id, read = true, folder = "inbox") {
    const mailbox = FOLDER_MAP[folder] || "INBOX";
    return withClient(account, async (client) => {
      const lock = await client.getMailboxLock(mailbox);
      try {
        if (read) await client.messageFlagsAdd(id, ["\\Seen"], { uid: true });
        else await client.messageFlagsRemove(id, ["\\Seen"], { uid: true });
      } finally {
        lock.release();
      }
    });
  },

  async toggleStar(account, id, starred, folder = "inbox") {
    const mailbox = FOLDER_MAP[folder] || "INBOX";
    return withClient(account, async (client) => {
      const lock = await client.getMailboxLock(mailbox);
      try {
        if (starred) await client.messageFlagsAdd(id, ["\\Flagged"], { uid: true });
        else await client.messageFlagsRemove(id, ["\\Flagged"], { uid: true });
      } finally {
        lock.release();
      }
    });
  },

  async send(account, { to, subject, body }) {
    const transport = nodemailer.createTransport({
      host: config.apple.smtpHost,
      port: config.apple.smtpPort,
      secure: false,
      requireTLS: true,
      auth: { user: account.secrets.address, pass: account.secrets.app_password },
    });
    await transport.sendMail({ from: account.secrets.address, to, subject, text: body });
  },
};

function stripHtml(html) {
  return (html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .trim();
}
