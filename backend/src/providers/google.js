import { google } from "googleapis";
import { config } from "../config.js";
import { store } from "../store.js";

/* Maps our generic folder ids to Gmail system labels. */
const FOLDER_TO_LABEL = {
  inbox: "INBOX",
  sent: "SENT",
  drafts: "DRAFT",
  spam: "SPAM",
  trash: "TRASH",
  archive: null, // Gmail "archive" = in All Mail but not Inbox; handled via query
  starred: "STARRED",
};

function oauthClient() {
  return new google.auth.OAuth2(
    config.google.clientId,
    config.google.clientSecret,
    `${config.serverOrigin}/auth/google/callback`
  );
}

function clientForAccount(account) {
  const o = oauthClient();
  o.setCredentials({
    access_token: account.secrets.access_token,
    refresh_token: account.secrets.refresh_token,
    expiry_date: account.secrets.expiry_date,
  });
  // Persist refreshed tokens automatically.
  o.on("tokens", (tokens) => {
    store.updateSecrets(account.id, {
      ...(tokens.access_token ? { access_token: tokens.access_token } : {}),
      ...(tokens.refresh_token ? { refresh_token: tokens.refresh_token } : {}),
      ...(tokens.expiry_date ? { expiry_date: tokens.expiry_date } : {}),
    });
  });
  return google.gmail({ version: "v1", auth: o });
}

function headerVal(headers, name) {
  return headers.find((h) => h.name.toLowerCase() === name.toLowerCase())?.value || "";
}

function parseAddrList(value) {
  // "A <a@x.com>, b@y.com" → ["a@x.com", "b@y.com"]
  return (value || "")
    .split(",")
    .map((part) => {
      const m = part.match(/<([^>]+)>/);
      return (m ? m[1] : part).trim();
    })
    .filter(Boolean);
}

function parseFrom(value) {
  // "Jane Doe <jane@x.com>" → { name, addr }
  const m = value.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/);
  if (m) return { name: m[1].trim() || m[2], addr: m[2] };
  return { name: value, addr: value };
}

export const googleProvider = {
  id: "google",

  authUrl(state) {
    return oauthClient().generateAuthUrl({
      access_type: "offline", // get a refresh_token
      prompt: "consent",
      scope: config.google.scopes,
      state,
    });
  },

  async handleCallback(code) {
    const o = oauthClient();
    const { tokens } = await o.getToken(code);
    o.setCredentials(tokens);
    const oauth2 = google.oauth2({ version: "v2", auth: o });
    const { data: me } = await oauth2.userinfo.get();
    return {
      address: me.email,
      secrets: {
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token,
        expiry_date: tokens.expiry_date,
      },
    };
  },

  async listFolders(account) {
    // Return our canonical folder set with unread counts from Gmail.
    const gmail = clientForAccount(account);
    const ids = ["inbox", "sent", "drafts", "spam", "trash", "starred"];
    const out = [];
    for (const id of ids) {
      const label = FOLDER_TO_LABEL[id];
      if (!label) continue;
      try {
        const { data } = await gmail.users.labels.get({ userId: "me", id: label });
        out.push({ id, name: id[0].toUpperCase() + id.slice(1), unread: data.messagesUnread || 0 });
      } catch {
        out.push({ id, name: id, unread: 0 });
      }
    }
    out.push({ id: "archive", name: "Archive", unread: 0 });
    return out;
  },

  // Returns { messages, nextCursor }. Pass the prior nextCursor as `cursor`
  // (Gmail's opaque pageToken) to fetch the next page.
  async listMessages(account, folder, { limit = 30, cursor } = {}) {
    const gmail = clientForAccount(account);
    let q, labelIds;
    if (folder === "archive") {
      q = "-in:inbox -in:sent -in:trash -in:spam -in:drafts in:all";
    } else if (folder === "starred") {
      labelIds = ["STARRED"];
    } else {
      labelIds = [FOLDER_TO_LABEL[folder] || "INBOX"];
    }

    const { data } = await gmail.users.messages.list({
      userId: "me",
      maxResults: limit,
      ...(cursor ? { pageToken: cursor } : {}),
      ...(labelIds ? { labelIds } : {}),
      ...(q ? { q } : {}),
    });

    const messages = data.messages || [];
    // Batch the metadata fetches.
    const detailed = await Promise.all(
      messages.map(async ({ id }) => {
        const { data: msg } = await gmail.users.messages.get({
          userId: "me",
          id,
          format: "metadata",
          metadataHeaders: ["From", "Subject", "Date"],
        });
        const headers = msg.payload?.headers || [];
        const from = parseFrom(headerVal(headers, "From"));
        return {
          id: msg.id,
          from: from.name,
          fromAddr: from.addr,
          subject: headerVal(headers, "Subject") || "(no subject)",
          preview: msg.snippet || "",
          date: new Date(Number(msg.internalDate)).toLocaleString(),
          unread: (msg.labelIds || []).includes("UNREAD"),
          starred: (msg.labelIds || []).includes("STARRED"),
          folder,
          attachments: [],
        };
      })
    );
    return { messages: detailed, nextCursor: data.nextPageToken || null };
  },

  async getMessage(account, id) {
    const gmail = clientForAccount(account);
    const { data: msg } = await gmail.users.messages.get({ userId: "me", id, format: "full" });
    const headers = msg.payload?.headers || [];
    const from = parseFrom(headerVal(headers, "From"));

    // Walk parts for text/plain AND text/html bodies + attachment names.
    let body = "";
    let bodyHtml = "";
    const attachments = [];
    const walk = (part) => {
      if (!part) return;
      if (part.filename && part.body?.attachmentId) {
        attachments.push({ name: part.filename, size: `${Math.round((part.body.size || 0) / 1024)} KB` });
      }
      if (part.mimeType === "text/plain" && part.body?.data) {
        body += Buffer.from(part.body.data, "base64").toString("utf8");
      }
      if (part.mimeType === "text/html" && part.body?.data) {
        bodyHtml += Buffer.from(part.body.data, "base64").toString("utf8");
      }
      (part.parts || []).forEach(walk);
    };
    walk(msg.payload);
    if (!body && !bodyHtml && msg.payload?.body?.data) {
      const raw = Buffer.from(msg.payload.body.data, "base64").toString("utf8");
      if (msg.payload.mimeType === "text/html") bodyHtml = raw;
      else body = raw;
    }

    return {
      id: msg.id,
      from: from.name,
      fromAddr: from.addr,
      to: parseAddrList(headerVal(headers, "To")),
      cc: parseAddrList(headerVal(headers, "Cc")),
      subject: headerVal(headers, "Subject") || "(no subject)",
      fullDate: headerVal(headers, "Date"),
      body: body || msg.snippet || "",
      bodyHtml: bodyHtml || null,
      unread: (msg.labelIds || []).includes("UNREAD"),
      starred: (msg.labelIds || []).includes("STARRED"),
      attachments,
    };
  },

  async markRead(account, id, read = true) {
    const gmail = clientForAccount(account);
    await gmail.users.messages.modify({
      userId: "me",
      id,
      requestBody: read ? { removeLabelIds: ["UNREAD"] } : { addLabelIds: ["UNREAD"] },
    });
  },

  // Move a message: Gmail models folders as labels, trash/spam as endpoints.
  async move(account, id, dest) {
    const gmail = clientForAccount(account);
    if (dest === "trash") {
      await gmail.users.messages.trash({ userId: "me", id });
      return;
    }
    if (dest === "inbox") {
      try { await gmail.users.messages.untrash({ userId: "me", id }); } catch { /* not in trash */ }
      await gmail.users.messages.modify({
        userId: "me", id,
        requestBody: { addLabelIds: ["INBOX"], removeLabelIds: ["SPAM"] },
      });
      return;
    }
    const addLabelIds = dest === "spam" ? ["SPAM"] : [];
    await gmail.users.messages.modify({
      userId: "me", id,
      requestBody: { addLabelIds, removeLabelIds: ["INBOX"] }, // archive = leave INBOX
    });
  },

  async toggleStar(account, id, starred) {
    const gmail = clientForAccount(account);
    await gmail.users.messages.modify({
      userId: "me",
      id,
      requestBody: starred ? { addLabelIds: ["STARRED"] } : { removeLabelIds: ["STARRED"] },
    });
  },

  async send(account, { to, subject, body, html }) {
    const gmail = clientForAccount(account);
    const raw = Buffer.from(
      [
        `To: ${to}`,
        `Subject: ${subject}`,
        "MIME-Version: 1.0",
        `Content-Type: ${html ? "text/html" : "text/plain"}; charset=utf-8`,
        "",
        html || body,
      ].join("\r\n")
    ).toString("base64url");
    await gmail.users.messages.send({ userId: "me", requestBody: { raw } });
  },
};
