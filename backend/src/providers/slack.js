import { request } from "undici";
import { config } from "../config.js";

const SLACK_API = "https://slack.com/api";

/**
 * Slack provider.
 *
 * Auth: OAuth v2 with USER token scopes (`user_scope=` in the authorize URL),
 * so Beacon reads and posts as the signed-in person, not as a bot. The token
 * exchange returns the user token under `authed_user.access_token` (xoxp-...).
 *
 * Frontend emoji are unicode; Slack's reactions API wants emoji NAMES
 * (e.g. "thumbsup"), so a small map covers the quick-picker set.
 */

const EMOJI_TO_NAME = {
  "👍": "thumbsup", "🎉": "tada", "❤️": "heart", "😄": "smile",
  "🙏": "pray", "🚀": "rocket", "👀": "eyes", "✅": "white_check_mark",
  "🔥": "fire", "💡": "bulb", "😅": "sweat_smile", "💯": "100",
};
const NAME_TO_EMOJI = Object.fromEntries(
  Object.entries(EMOJI_TO_NAME).map(([e, n]) => [n, e])
);

function redirectUri() {
  return `${config.serverOrigin}/auth/slack/callback`;
}

/* Call a Slack Web API method with the account's user token. */
async function slackApi(token, method, params = {}, { post = false } = {}) {
  const url = post
    ? `${SLACK_API}/${method}`
    : `${SLACK_API}/${method}?${new URLSearchParams(params)}`;
  const res = await request(url, {
    method: post ? "POST" : "GET",
    headers: {
      authorization: `Bearer ${token}`,
      ...(post ? { "content-type": "application/json; charset=utf-8" } : {}),
    },
    ...(post ? { body: JSON.stringify(params) } : {}),
  });
  const json = await res.body.json();
  if (!json.ok) throw new Error(`Slack ${method} error: ${json.error}`);
  return json;
}

/* Resolve user IDs → display names, cached per call chain. */
async function makeNameResolver(token) {
  const cache = new Map();
  return async (userId) => {
    if (!userId) return "Unknown";
    if (cache.has(userId)) return cache.get(userId);
    try {
      const { user } = await slackApi(token, "users.info", { user: userId });
      const name = user.profile?.display_name || user.real_name || user.name;
      cache.set(userId, name);
      return name;
    } catch {
      cache.set(userId, userId);
      return userId;
    }
  };
}

function mapReactions(reactions = [], selfId) {
  return reactions.map((r) => ({
    emoji: NAME_TO_EMOJI[r.name] || `:${r.name}:`,
    count: r.count,
    mine: (r.users || []).includes(selfId),
  }));
}

export const slackProvider = {
  id: "slack",

  authUrl(state) {
    const params = new URLSearchParams({
      client_id: config.slack.clientId,
      user_scope: config.slack.userScopes.join(","),
      redirect_uri: redirectUri(),
      state,
    });
    return `https://slack.com/oauth/v2/authorize?${params}`;
  },

  async handleCallback(code) {
    const res = await request(`${SLACK_API}/oauth.v2.access`, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: config.slack.clientId,
        client_secret: config.slack.clientSecret,
        code,
        redirect_uri: redirectUri(),
      }).toString(),
    });
    const json = await res.body.json();
    if (!json.ok) throw new Error(`Slack OAuth error: ${json.error}`);
    const token = json.authed_user?.access_token;
    if (!token) throw new Error("Slack OAuth: no user token returned (check user_scope)");
    return {
      // address doubles as the display identity in the client
      address: json.team?.name || json.team?.id || "Slack workspace",
      secrets: {
        user_token: token,
        user_id: json.authed_user.id,
        team_id: json.team?.id,
        team_name: json.team?.name,
      },
    };
  },

  /* Channels, DMs, and group DMs with unread counts where Slack exposes them. */
  async listConversations(account) {
    const token = account.secrets.user_token;
    const resolveName = await makeNameResolver(token);
    const { channels } = await slackApi(token, "users.conversations", {
      types: "public_channel,private_channel,im,mpim",
      exclude_archived: "true",
      limit: "200",
    });
    const out = [];
    for (const c of channels) {
      let name = c.name;
      let kind = "channel";
      if (c.is_im) {
        kind = "dm";
        name = await resolveName(c.user);
      } else if (c.is_mpim) {
        kind = "dm";
        name = (c.purpose?.value || c.name || "Group DM").replace(/^Group messaging with:\s*/, "");
      }
      // unread_count_display requires conversations.info per channel; fetch
      // lazily only for ims to keep this fast. Channels report 0 here —
      // upgrade path: track last_read via conversations.info per channel.
      out.push({ id: c.id, name, kind, unread: 0 });
    }
    return out;
  },

  // Per-conversation unread counts for the notification poller. Slack only
  // exposes unread_count_display via conversations.info (one call per channel),
  // so we cap the fan-out and only check channels the user is actually in.
  // Returns { total, byChannel: { channelId: count }, dmUnread }.
  async listUnreadCounts(account) {
    const token = account.secrets.user_token;
    const { channels } = await slackApi(token, "users.conversations", {
      types: "public_channel,private_channel,im,mpim",
      exclude_archived: "true",
      limit: "200",
    });
    let total = 0;
    let dmUnread = 0;
    const byChannel = {};
    // Cap at 50 conversations per tick to stay well under rate limits.
    for (const c of channels.slice(0, 50)) {
      try {
        const { channel } = await slackApi(token, "conversations.info", { channel: c.id });
        const n = channel?.unread_count_display || 0;
        if (n > 0) {
          byChannel[c.id] = n;
          total += n;
          if (c.is_im || c.is_mpim) dmUnread += n;
        }
      } catch {
        // Skip channels that error (e.g. not_in_channel) — they contribute 0.
      }
    }
    return { total, byChannel, dmUnread };
  },

  async listMessages(account, channelId, { limit = 30 } = {}) {
    const token = account.secrets.user_token;
    const selfId = account.secrets.user_id;
    const resolveName = await makeNameResolver(token);
    const { messages } = await slackApi(token, "conversations.history", {
      channel: channelId,
      limit: String(limit),
    });
    const mapped = [];
    for (const m of messages) {
      if (m.subtype && m.subtype !== "thread_broadcast") continue; // skip joins etc.
      mapped.push({
        id: m.ts,
        author: await resolveName(m.user),
        text: m.text || "",
        time: new Date(Number(m.ts) * 1000).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }),
        reactions: mapReactions(m.reactions, selfId),
        replyCount: m.reply_count || 0,
      });
    }
    return mapped.reverse(); // oldest first for the stream
  },

  async listReplies(account, channelId, threadTs) {
    const token = account.secrets.user_token;
    const selfId = account.secrets.user_id;
    const resolveName = await makeNameResolver(token);
    const { messages } = await slackApi(token, "conversations.replies", {
      channel: channelId,
      ts: threadTs,
    });
    const out = [];
    for (const m of messages) {
      out.push({
        id: m.ts,
        author: await resolveName(m.user),
        text: m.text || "",
        time: new Date(Number(m.ts) * 1000).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }),
        reactions: mapReactions(m.reactions, selfId),
      });
    }
    return out; // [0] is the parent
  },

  async postMessage(account, channelId, text, threadTs = null) {
    return slackApi(account.secrets.user_token, "chat.postMessage", {
      channel: channelId,
      text,
      ...(threadTs ? { thread_ts: threadTs } : {}),
    }, { post: true });
  },

  async toggleReaction(account, channelId, timestamp, emoji, add) {
    const name = EMOJI_TO_NAME[emoji] || emoji.replace(/:/g, "");
    return slackApi(
      account.secrets.user_token,
      add ? "reactions.add" : "reactions.remove",
      { channel: channelId, timestamp, name },
      { post: true }
    );
  },
};
