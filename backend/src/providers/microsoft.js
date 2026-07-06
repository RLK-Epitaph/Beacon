import { request } from "undici";
import { config } from "../config.js";
import { store } from "../store.js";

const AUTHORITY = (tenant) => `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
const GRAPH = "https://graph.microsoft.com/v1.0";

/* Generic folder id → Graph well-known folder name. */
const FOLDER_MAP = {
  inbox: "inbox",
  sent: "sentitems",
  drafts: "drafts",
  spam: "junkemail",
  trash: "deleteditems",
  archive: "archive",
};

function redirectUri() {
  return `${config.serverOrigin}/auth/microsoft/callback`;
}

async function tokenRequest(params) {
  const res = await request(`${AUTHORITY(config.microsoft.tenant)}/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params).toString(),
  });
  const json = await res.body.json();
  if (res.statusCode >= 400) throw new Error(`MS token error: ${JSON.stringify(json)}`);
  return json;
}

/* Returns a valid access token, refreshing + persisting if expired. */
async function accessToken(account) {
  const now = Date.now();
  if (account.secrets.access_token && account.secrets.expiry_date > now + 60000) {
    return account.secrets.access_token;
  }
  const tok = await tokenRequest({
    client_id: config.microsoft.clientId,
    client_secret: config.microsoft.clientSecret,
    grant_type: "refresh_token",
    refresh_token: account.secrets.refresh_token,
    scope: config.microsoft.scopes.join(" "),
    redirect_uri: redirectUri(),
  });
  const secrets = {
    access_token: tok.access_token,
    refresh_token: tok.refresh_token || account.secrets.refresh_token,
    expiry_date: now + tok.expires_in * 1000,
  };
  store.updateSecrets(account.id, secrets);
  account.secrets = { ...account.secrets, ...secrets };
  return secrets.access_token;
}

async function graph(account, pathStr, { method = "GET", body } = {}) {
  const token = await accessToken(account);
  const res = await request(`${GRAPH}${pathStr}`, {
    method,
    headers: {
      authorization: `Bearer ${token}`,
      ...(body ? { "content-type": "application/json" } : {}),
    },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  if (res.statusCode === 204) return null;
  const json = await res.body.json();
  if (res.statusCode >= 400) throw new Error(`Graph error ${res.statusCode}: ${JSON.stringify(json)}`);
  return json;
}

function mapMessage(m, folder) {
  return {
    id: m.id,
    from: m.from?.emailAddress?.name || m.from?.emailAddress?.address || "Unknown",
    fromAddr: m.from?.emailAddress?.address || "",
    subject: m.subject || "(no subject)",
    preview: m.bodyPreview || "",
    date: new Date(m.receivedDateTime || m.sentDateTime).toLocaleString(),
    fullDate: new Date(m.receivedDateTime || m.sentDateTime).toLocaleString(),
    unread: m.isRead === false,
    starred: m.flag?.flagStatus === "flagged",
    folder,
    attachments: m.hasAttachments ? [{ name: "attachment", size: "" }] : [],
  };
}

export const microsoftProvider = {
  id: "microsoft",

  authUrl(state) {
    const params = new URLSearchParams({
      client_id: config.microsoft.clientId,
      response_type: "code",
      redirect_uri: redirectUri(),
      response_mode: "query",
      scope: config.microsoft.scopes.join(" "),
      state,
    });
    return `${AUTHORITY(config.microsoft.tenant)}/authorize?${params}`;
  },

  async handleCallback(code) {
    const tok = await tokenRequest({
      client_id: config.microsoft.clientId,
      client_secret: config.microsoft.clientSecret,
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri(),
      scope: config.microsoft.scopes.join(" "),
    });
    const secrets = {
      access_token: tok.access_token,
      refresh_token: tok.refresh_token,
      expiry_date: Date.now() + tok.expires_in * 1000,
    };
    // Identify the mailbox.
    const me = await graph({ secrets }, "/me");
    return { address: me.mail || me.userPrincipalName, secrets };
  },

  async listFolders(account) {
    const out = [];
    for (const [id, gname] of Object.entries(FOLDER_MAP)) {
      try {
        const f = await graph(account, `/me/mailFolders/${gname}`);
        out.push({ id, name: id[0].toUpperCase() + id.slice(1), unread: f.unreadItemCount || 0 });
      } catch {
        out.push({ id, name: id, unread: 0 });
      }
    }
    out.push({ id: "starred", name: "Starred", unread: 0 });
    return out;
  },

  // Returns { messages, nextCursor }. Cursor is a numeric $skip offset.
  async listMessages(account, folder, { limit = 30, cursor } = {}) {
    const skip = Number(cursor) || 0;
    const select = "id,subject,bodyPreview,from,receivedDateTime,sentDateTime,isRead,flag,hasAttachments";
    let path;
    if (folder === "starred") {
      path = `/me/messages?$filter=flag/flagStatus eq 'flagged'&$top=${limit}&$skip=${skip}&$select=${select}`;
    } else {
      const gname = FOLDER_MAP[folder] || "inbox";
      path = `/me/mailFolders/${gname}/messages?$top=${limit}&$skip=${skip}&$orderby=receivedDateTime desc&$select=${select}`;
    }
    const res = await graph(account, path);
    const messages = (res.value || []).map((m) => mapMessage(m, folder));
    // If Graph returned a full page (or signals more via nextLink), advance the offset.
    const hasMore = !!res["@odata.nextLink"] || messages.length === limit;
    const nextCursor = hasMore ? String(skip + limit) : null;
    return { messages, nextCursor };
  },

  async getMessage(account, id) {
    const m = await graph(account, `/me/messages/${id}`);
    let attachments = [];
    if (m.hasAttachments) {
      const att = await graph(account, `/me/messages/${id}/attachments?$select=name,size`);
      attachments = (att.value || []).map((a) => ({
        name: a.name,
        size: a.size ? `${Math.round(a.size / 1024)} KB` : "",
      }));
    }
    return {
      id: m.id,
      from: m.from?.emailAddress?.name || "",
      fromAddr: m.from?.emailAddress?.address || "",
      to: (m.toRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean),
      cc: (m.ccRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean),
      subject: m.subject || "(no subject)",
      fullDate: new Date(m.receivedDateTime).toLocaleString(),
      body: m.body?.contentType === "html" ? stripHtml(m.body.content) : m.body?.content || "",
      bodyHtml: m.body?.contentType === "html" ? m.body.content : null,
      unread: m.isRead === false,
      starred: m.flag?.flagStatus === "flagged",
      attachments,
    };
  },

  async markRead(account, id, read = true) {
    await graph(account, `/me/messages/${id}`, { method: "PATCH", body: { isRead: read } });
  },

  // Move via Graph; note Graph assigns the moved message a NEW id, so the
  // client should drop its local copy rather than track it.
  async move(account, id, dest) {
    await graph(account, `/me/messages/${id}/move`, {
      method: "POST",
      body: { destinationId: FOLDER_MAP[dest] || "inbox" },
    });
  },

  async toggleStar(account, id, starred) {
    await graph(account, `/me/messages/${id}`, {
      method: "PATCH",
      body: { flag: { flagStatus: starred ? "flagged" : "notFlagged" } },
    });
  },

  async send(account, { to, subject, body, html }) {
    await graph(account, "/me/sendMail", {
      method: "POST",
      body: {
        message: {
          subject,
          body: { contentType: html ? "HTML" : "Text", content: html || body },
          toRecipients: to.split(",").map((a) => ({ emailAddress: { address: a.trim() } })),
        },
      },
    });
  },

  /* ---- Microsoft Teams (chats) ----
     Teams runs on the same Microsoft identity as Outlook mail, so these reuse
     this account's Graph token. Scoped to CHATS (1:1 + group DMs) via
     Chat.ReadWrite — team CHANNELS need ChannelMessage.Read.All (admin consent)
     and are intentionally excluded here. */

  // Basic identity for the Teams "org" header + resolving 1:1 chat names.
  async teamsIdentity(account) {
    const me = await graph(account, "/me?$select=id,displayName,mail,userPrincipalName");
    const addr = me.mail || me.userPrincipalName || account.address;
    return { id: me.id, name: me.displayName || addr, org: orgFromAddress(addr) };
  },

  // List the signed-in user's chats. Names mirror the Teams client: the OTHER
  // person for 1:1s, the topic (or member first-names) for groups. `unread` is
  // derived from the last-read vs last-message timestamps in the chat viewpoint.
  async listChats(account) {
    const me = await graph(account, "/me?$select=id");
    const myId = me.id;
    const res = await graph(
      account,
      "/me/chats?$expand=members,lastMessagePreview&$top=50&$orderby=lastMessagePreview/createdDateTime desc"
    );
    const chats = [];
    for (const c of res.value || []) {
      const others = (c.members || []).filter((m) => m.userId && m.userId !== myId);
      let name;
      let kind;
      if (c.chatType === "oneOnOne") {
        kind = "chat";
        name = others[0]?.displayName || "Direct chat";
      } else {
        kind = "group";
        name =
          c.topic ||
          others.map((m) => (m.displayName || "").split(" ")[0]).filter(Boolean).join(", ") ||
          "Group chat";
      }
      const last = c.lastMessagePreview;
      const lastRead = c.viewpoint?.lastMessageReadDateTime;
      const unread =
        last?.createdDateTime &&
        lastRead &&
        new Date(last.createdDateTime) > new Date(lastRead) &&
        last.from?.user?.id !== myId
          ? 1
          : 0;
      chats.push({ id: c.id, name, kind, unread, messages: null });
    }
    return chats;
  },

  async listChatMessages(account, chatId, { limit = 30 } = {}) {
    const me = await graph(account, "/me?$select=id");
    const res = await graph(account, `/chats/${chatId}/messages?$top=${limit}`);
    const msgs = (res.value || [])
      .filter((m) => m.messageType === "message" && (m.body?.content || "").trim())
      .map((m) => ({
        id: m.id,
        author: m.from?.user?.displayName || "Unknown",
        text: (m.body?.contentType || "").toLowerCase() === "html" ? stripHtml(m.body.content) : m.body?.content || "",
        time: formatTeamsTime(m.createdDateTime),
        reactions: groupTeamsReactions(m.reactions, me.id),
      }));
    return msgs.reverse(); // Graph returns newest-first; the stream wants oldest-first
  },

  async sendChatMessage(account, chatId, text) {
    return graph(account, `/chats/${chatId}/messages`, {
      method: "POST",
      body: { body: { contentType: "text", content: text } },
    });
  },
};

function stripHtml(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function orgFromAddress(addr) {
  const domain = (addr.split("@")[1] || "").split(".")[0];
  return domain ? domain[0].toUpperCase() + domain.slice(1) : "Microsoft";
}

// Show a clock time for today's messages, otherwise a short date — matching how
// the Teams client labels messages.
function formatTeamsTime(iso) {
  const d = new Date(iso);
  const today = new Date();
  const sameDay = d.toDateString() === today.toDateString();
  return sameDay
    ? d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
    : d.toLocaleDateString([], { month: "short", day: "numeric" });
}

const TEAMS_REACTION_EMOJI = {
  like: "👍", heart: "❤️", laugh: "😆", surprised: "😮", sad: "😢", angry: "😠",
};
function groupTeamsReactions(list = [], myId) {
  const map = new Map();
  for (const r of list) {
    const emoji = TEAMS_REACTION_EMOJI[r.reactionType] || "👍";
    const e = map.get(emoji) || { emoji, count: 0, mine: false };
    e.count++;
    if (r.user?.user?.id === myId) e.mine = true;
    map.set(emoji, e);
  }
  return [...map.values()];
}
