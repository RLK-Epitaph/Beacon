import { request } from "undici";
import { config } from "../config.js";
import { getActivity, setActivity } from "../slackActivity.js";

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

/* Call a Slack Web API method with the account's user token. HTTP 429s are
   thrown as errors carrying `rateLimited` + `retryAfter` (seconds) so routes
   can tell the client exactly when to try again instead of failing opaquely. */
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
  if (res.statusCode === 429) {
    await res.body.dump();
    const err = new Error(`Slack ${method} rate limited`);
    err.rateLimited = true;
    err.retryAfter = Number(res.headers["retry-after"] || 60);
    throw err;
  }
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

// Drop only conversations Slack itself never surfaces anywhere: IMs with a
// since-deactivated user. Everything else (including quiet channels and DMs
// with no recent activity) is kept and returned to the client — the "active
// in the last 30 days" default vs. "all" is a filter applied in the UI, not a
// hard exclusion, so nothing is permanently hidden.
function filterDeadConversations(channels) {
  return channels.filter((c) => !(c.is_im && c.is_user_deleted));
}

const THIRTY_DAYS_S = 30 * 24 * 60 * 60;

// In-memory unread badges, refreshed by the poller.
// accountId → { channelId: { n, at } } where n = unread count and at = when we
// last checked (ms) — `at` drives the poller's round-robin so every
// conversation gets re-checked eventually, not just the recently-active ones.
// (Unread must be near-real-time; last-activity, which changes slowly, lives in
// the persistent slackActivity cache instead.)
const unreadCache = new Map();
function getUnreadMap(accountId) {
  let m = unreadCache.get(accountId);
  if (!m) { m = {}; unreadCache.set(accountId, m); }
  return m;
}

// Workspace member directory (for @mention autocomplete) + the signed-in
// user's own identity. Membership changes rarely — cache for 15 minutes.
const memberCache = new Map(); // accountId → { at, list }
const selfCache = new Map();   // accountId → { id, name }
const MEMBER_TTL_MS = 15 * 60 * 1000;

// Last successful history fetch per conversation. conversations.history is
// rate-limited to ~1 call/minute for this app, so when a fetch gets 429'd we
// serve the last-known messages (marked stale) rather than erroring — a user
// flipping between channels should never see a false "no messages".
const messagesCache = new Map(); // `${accountId}:${channelId}` → messages[]

// One conversations.info call → unread state for a conversation.
// IMPORTANT (verified against the live API): Slack only returns
// unread_count_display for IMs/MPIMs. For CHANNELS the response carries no
// unread fields at all — only `last_read` — so channel unread must be derived
// by comparing last_read with the channel's real last-message timestamp (which
// we record in the activity cache whenever history is fetched). For DMs the
// response also includes `latest`, the newest message — its ts keeps the
// activity cache fresh for free.
// Returns { ok, unread, lastRead, latestTs } | { rateLimited } | { ok:false }.
async function convInfoUnread(token, channelId) {
  const res = await request(
    `${SLACK_API}/conversations.info?${new URLSearchParams({ channel: channelId })}`,
    { headers: { authorization: `Bearer ${token}` } }
  );
  if (res.statusCode === 429) {
    await res.body.dump();
    return { rateLimited: true };
  }
  const json = await res.body.json();
  if (!json.ok) return { ok: false };
  const ch = json.channel || {};
  return {
    ok: true,
    unread: ch.unread_count_display || 0,
    lastRead: ch.last_read ? Number(ch.last_read) : 0,
    latestTs: ch.latest?.ts ? Number(ch.latest.ts) : 0,
  };
}

// Best last-activity timestamp (unix seconds) we have for a conversation:
//   - a real last-message ts recorded when the user opened the channel (accurate), or
//   - users.conversations' `updated` field as a fallback (cheap but approximate —
//     it tracks metadata changes and tends to lag real message activity).
// Returns { ts, exact }. exact=false means it's the approximate `updated` value.
function activityTs(conv, cache) {
  const cached = cache[conv.id];
  if (cached) return { ts: cached.ts, exact: true };
  return { ts: conv.updated ? conv.updated / 1000 : 0, exact: false };
}

// Group DM names fall back to Slack's internal "mpdm-you--alice--bob-1"
// slug (which includes your own name) unless a custom purpose is set.
// Slack's real client instead labels these with just the OTHER members'
// display names — resolve the actual membership list to match that.
async function resolveMpimName(token, channel, selfId, resolveName) {
  try {
    const { members } = await slackApi(token, "conversations.members", { channel: channel.id });
    const others = members.filter((id) => id !== selfId);
    const names = await Promise.all(others.map(resolveName));
    return names.join(", ") || "Group DM";
  } catch {
    return (channel.purpose?.value || "Group DM").replace(/^Group messaging with:\s*/, "");
  }
}

/* Slack message text arrives in raw mrkdwn: user mentions as <@U123>, special
   mentions as <!here>, channel refs as <#C123|name>, links as <url|label>, and
   &-escaped entities. Convert those into lightweight tokens the client renders
   as styled chips — {{@id|Name}} / {{@id|Name|me}} / {{#|name}} / {{L|url|label}}
   — and report whether the message mentions the signed-in user. */
async function slackifyText(raw, selfId, resolveName) {
  if (!raw) return { text: "", mentionsMe: false };
  let mentionsMe = false;

  // Resolve every distinct mentioned user id up front (resolveName caches).
  const ids = [...raw.matchAll(/<@([A-Z0-9]+)(?:\|[^>]*)?>/g)].map((m) => m[1]);
  const names = {};
  for (const id of new Set(ids)) names[id] = await resolveName(id);

  let text = raw.replace(/<@([A-Z0-9]+)(?:\|[^>]*)?>/g, (_, id) => {
    if (id === selfId) mentionsMe = true;
    return `{{@${id}|${names[id] || id}${id === selfId ? "|me" : ""}}}`;
  });
  text = text
    .replace(/<!(channel|here|everyone)(?:\|[^>]*)?>/g, "{{@|$1}}")
    .replace(/<#[A-Z0-9]+\|([^>]+)>/g, "{{#|$1}}")
    .replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, "{{L|$1|$2}}")
    .replace(/<(https?:\/\/[^>]+)>/g, "{{L|$1|$1}}")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  return { text, mentionsMe };
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

  /* Channels, DMs, and group DMs. Each is annotated with an `active` flag —
     activity in the last 30 days — so the client defaults to Slack's "recently
     active" view while still offering an "all" filter. Activity comes from the
     real last-message time recorded when a channel was last opened, falling back
     to Slack's `updated` field (approximate) for channels not yet visited. A
     conversation with unread is always active. `lastTs` (unix seconds) drives
     ordering. */
  async listConversations(account) {
    const token = account.secrets.user_token;
    const selfId = account.secrets.user_id;
    const resolveName = await makeNameResolver(token);
    const { channels } = await slackApi(token, "users.conversations", {
      types: "public_channel,private_channel,im,mpim",
      exclude_archived: "true",
      limit: "200",
    });
    const live = filterDeadConversations(channels);
    const cache = getActivity(account.id);
    const unreadMap = getUnreadMap(account.id);
    const nowS = Date.now() / 1000;
    const out = [];
    for (const c of live) {
      let name = c.name;
      let kind = "channel";
      if (c.is_im) {
        kind = "dm";
        name = await resolveName(c.user);
      } else if (c.is_mpim) {
        kind = "dm";
        name = await resolveMpimName(token, c, selfId, resolveName);
      }
      const unread = unreadMap[c.id]?.n || 0;
      const { ts } = activityTs(c, cache);
      const active = unread > 0 || (ts > 0 && nowS - ts < THIRTY_DAYS_S);
      out.push({ id: c.id, name, kind, unread, lastTs: ts, active });
    }
    // Order by most recent activity so the sidebar (and the default channel
    // picked on open) lands on something with real recent history — matching
    // how Slack surfaces recently-active conversations first.
    out.sort((a, b) => (b.lastTs || 0) - (a.lastTs || 0));
    return out;
  },

  // Per-conversation unread counts for the notification poller. conversations.info
  // (one rate-limited call per channel) is the only unread source, so each tick
  // checks a capped batch: conversations already showing unread first (to notice
  // them being read elsewhere), then round-robin by least-recently-checked so a
  // new message in ANY conversation — even a long-quiet one — is noticed within a
  // few ticks. Previously only "recently active" channels were ever polled, which
  // meant a fresh message in a quiet channel produced no notification at all.
  // Totals are built from the whole cache (not just this tick's batch) so counts
  // seen in earlier ticks don't flicker to zero.
  // Returns { total, byChannel: { channelId: count }, dmUnread }.
  async listUnreadCounts(account) {
    const token = account.secrets.user_token;
    const { channels } = await slackApi(token, "users.conversations", {
      types: "public_channel,private_channel,im,mpim",
      exclude_archived: "true",
      limit: "200",
    });
    const live = filterDeadConversations(channels);
    const unreadMap = getUnreadMap(account.id);

    const batch = [...live]
      .sort((a, b) => {
        const ea = unreadMap[a.id];
        const eb = unreadMap[b.id];
        const ua = ea?.n > 0 ? 0 : 1;
        const ub = eb?.n > 0 ? 0 : 1;
        if (ua !== ub) return ua - ub;               // unread ones first
        return (ea?.at || 0) - (eb?.at || 0);        // then stalest check first
      })
      .slice(0, 10);
    const activity = getActivity(account.id);
    for (const c of batch) {
      const r = await convInfoUnread(token, c.id);
      if (r.rateLimited) break; // stop early; next tick resumes where we left off
      if (!r.ok) {
        // not_in_channel etc. → record a 0 so the rotation moves past it.
        unreadMap[c.id] = { n: 0, at: Date.now() };
        continue;
      }
      let n = r.unread;
      if (!(c.is_im || c.is_mpim)) {
        // Channels expose NO unread count via the API (verified live) — only
        // last_read. Derive "has new": a real last-message ts we've recorded
        // that's newer than the user's read cursor means unread activity.
        // (Understates when we haven't seen the channel's history yet; never
        // false-positives, since both timestamps are real.)
        const knownTs = activity[c.id]?.ts || 0;
        if (n === 0 && knownTs > r.lastRead) n = 1;
      } else if (r.latestTs > 0) {
        // DM info includes the latest message — keep the activity cache fresh
        // so newly-messaged DMs surface in the Active view and sort correctly.
        setActivity(account.id, c.id, r.latestTs);
      }
      unreadMap[c.id] = { n, at: Date.now() };
    }

    const byChannel = {};
    let total = 0;
    let dmUnread = 0;
    for (const c of live) {
      const n = unreadMap[c.id]?.n || 0;
      if (n > 0) {
        byChannel[c.id] = n;
        total += n;
        if (c.is_im || c.is_mpim) dmUnread += n;
      }
    }
    return { total, byChannel, dmUnread };
  },

  // Mark a channel read up to now, so Slack's own read cursor moves —
  // without this, opening a channel in Beacon only clears the badge locally
  // and the very next poll tick puts the "unread" count right back.
  async markRead(account, channelId) {
    return slackApi(
      account.secrets.user_token,
      "conversations.mark",
      { channel: channelId, ts: String(Date.now() / 1000) },
      { post: true }
    );
  },

  // Returns { messages, stale }. When history is rate-limited and we have a
  // previous fetch cached, that cache is served with stale=true; with nothing
  // cached the rate-limit error (with retryAfter) propagates to the route.
  async listMessages(account, channelId, { limit = 30 } = {}) {
    const token = account.secrets.user_token;
    const selfId = account.secrets.user_id;
    const cacheKey = `${account.id}:${channelId}`;
    let messages;
    try {
      ({ messages } = await slackApi(token, "conversations.history", {
        channel: channelId,
        limit: String(limit),
      }));
    } catch (e) {
      const cached = messagesCache.get(cacheKey);
      if (e.rateLimited && cached) return { messages: cached, stale: true };
      throw e;
    }
    // conversations.history returns newest-first, so messages[0] is the last
    // message. Record its timestamp as this conversation's real last-activity —
    // this is the accurate signal for the 30-day filter, captured for free from
    // a call the user's channel-open already made (no extra rate-limited calls).
    setActivity(account.id, channelId, messages?.[0]?.ts ? Number(messages[0].ts) : 0);
    const resolveName = await makeNameResolver(token);
    const mapped = [];
    for (const m of messages) {
      if (m.subtype && m.subtype !== "thread_broadcast") continue; // skip joins etc.
      const { text, mentionsMe } = await slackifyText(m.text || "", selfId, resolveName);
      mapped.push({
        id: m.ts,
        author: await resolveName(m.user),
        text,
        mentionsMe,
        time: new Date(Number(m.ts) * 1000).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }),
        reactions: mapReactions(m.reactions, selfId),
        replyCount: m.reply_count || 0,
      });
    }
    const out = mapped.reverse(); // oldest first for the stream
    messagesCache.set(cacheKey, out);
    return { messages: out, stale: false };
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
      const { text, mentionsMe } = await slackifyText(m.text || "", selfId, resolveName);
      out.push({
        id: m.ts,
        author: await resolveName(m.user),
        text,
        mentionsMe,
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

  // Workspace member directory for @mention autocomplete (humans + app bots,
  // minus deleted users and the Slackbot pseudo-user). Cached per account.
  async listMembers(account) {
    const cached = memberCache.get(account.id);
    if (cached && Date.now() - cached.at < MEMBER_TTL_MS) return cached.list;
    const token = account.secrets.user_token;
    const list = [];
    let cursor;
    do {
      const res = await slackApi(token, "users.list", {
        limit: "200",
        ...(cursor ? { cursor } : {}),
      });
      for (const u of res.members || []) {
        if (u.deleted || u.id === "USLACKBOT") continue;
        list.push({ id: u.id, name: u.profile?.display_name || u.real_name || u.name });
      }
      cursor = res.response_metadata?.next_cursor;
    } while (cursor && list.length < 500);
    memberCache.set(account.id, { at: Date.now(), list });
    return list;
  },

  // The signed-in user's identity (id + display name), cached indefinitely —
  // used by the client for "you" labels and mention highlighting.
  async getSelf(account) {
    const cached = selfCache.get(account.id);
    if (cached) return cached;
    const selfId = account.secrets.user_id;
    const resolveName = await makeNameResolver(account.secrets.user_token);
    const self = { id: selfId, name: await resolveName(selfId) };
    selfCache.set(account.id, self);
    return self;
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
