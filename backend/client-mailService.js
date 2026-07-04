/**
 * Drop-in replacement for the mock `mailService` in Beacon.jsx.
 *
 * Copy this object over the mock one (or import it) and the existing UI talks to
 * the real backend. Every request sends the session cookie (credentials:"include").
 *
 * The backend base URL — point at wherever the Express server runs.
 */
const BACKEND = "http://localhost:4000";

async function api(path, options = {}) {
  const res = await fetch(`${BACKEND}${path}`, {
    credentials: "include",
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Request failed: ${res.status}`);
  }
  return res.json();
}

export const mailService = {
  // List connected accounts. Then hydrate each with its inbox messages so the
  // sidebar unread counts work exactly like the mock did.
  async listAccounts() {
    const { accounts } = await api("/api/accounts");
    return Promise.all(
      accounts.map(async (a) => {
        let messages = [];
        try {
          const r = await api(`/api/accounts/${a.id}/folders/inbox/messages?limit=30`);
          messages = r.messages;
        } catch {
          /* token may need refresh; leave empty */
        }
        return { ...a, messages };
      })
    );
  },

  // Connect kicks off OAuth. For Google/Microsoft we redirect the whole window;
  // the backend bounces back to the client with ?connect=ok. For Apple we POST
  // the app-specific password (collect address + appPassword in your modal).
  connectRedirect(provider) {
    window.location.href = `${BACKEND}/auth/${provider}`;
  },

  async connectApple(address, appPassword) {
    const { account } = await api("/auth/apple/password", {
      method: "POST",
      body: JSON.stringify({ address, appPassword }),
    });
    const r = await api(`/api/accounts/${account.id}/folders/inbox/messages?limit=30`);
    return { ...account, messages: r.messages };
  },

  // Returns { messages, nextCursor }. Pass nextCursor back as `cursor` for the
  // next page; nextCursor === null means you've reached the end.
  async listMessages(accountId, folder, { cursor, limit = 30 } = {}) {
    const qs = new URLSearchParams({ limit: String(limit) });
    if (cursor) qs.set("cursor", cursor);
    const { messages, nextCursor } = await api(
      `/api/accounts/${accountId}/folders/${folder}/messages?${qs}`
    );
    return { messages, nextCursor };
  },

  async getMessage(accountId, id, folder) {
    const { message } = await api(`/api/accounts/${accountId}/messages/${id}?folder=${folder}`);
    return message;
  },

  async markRead(accountId, id, read = true, folder = "inbox") {
    await api(`/api/accounts/${accountId}/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ read, folder }),
    });
  },

  async toggleStar(accountId, id, starred, folder = "inbox") {
    await api(`/api/accounts/${accountId}/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ starred, folder }),
    });
  },

  async setLabel(accountId, label) {
    const { account } = await api(`/api/accounts/${accountId}`, {
      method: "PATCH",
      body: JSON.stringify({ label }),
    });
    return account;
  },

  async send(accountId, { to, subject, body }) {
    await api(`/api/accounts/${accountId}/send`, {
      method: "POST",
      body: JSON.stringify({ to, subject, body }),
    });
  },

  // Force the backend to re-poll one account now. Returns { folders, inboxUnread }.
  async refresh(accountId) {
    return api(`/api/accounts/${accountId}/refresh`, { method: "POST" });
  },

  async removeAccount(accountId) {
    await api(`/api/accounts/${accountId}`, { method: "DELETE" });
  },
};

/**
 * Subscribe to live unread-count updates over Server-Sent Events.
 * The callback fires with { accountId, folders, inboxUnread } whenever the
 * backend's poller detects a change. Returns an unsubscribe function.
 *
 *   useEffect(() => subscribeToUpdates(({ accountId, inboxUnread }) => {
 *     setAccounts(prev => prev.map(a =>
 *       a.id === accountId ? { ...a, _liveUnread: inboxUnread } : a));
 *   }), []);
 */
export function subscribeToUpdates(onChange) {
  const es = new EventSource(`${BACKEND}/api/events`, { withCredentials: true });
  es.addEventListener("change", (e) => {
    try {
      onChange(JSON.parse(e.data));
    } catch {
      /* ignore malformed frame */
    }
  });
  es.onerror = () => {
    /* EventSource auto-reconnects; nothing to do */
  };
  return () => es.close();
}

/**
 * Handle the OAuth return. Call this once on app mount:
 *   useEffect(() => { handleConnectReturn(refresh); }, []);
 */
export function handleConnectReturn(onConnected) {
  const params = new URLSearchParams(window.location.search);
  if (params.get("connect") === "ok") {
    window.history.replaceState({}, "", window.location.pathname);
    onConnected?.(params.get("account"));
  } else if (params.get("connect") === "error") {
    console.error("Connect failed:", params.get("reason"));
    window.history.replaceState({}, "", window.location.pathname);
  }
}
