import { EventEmitter } from "node:events";
import { getProvider } from "./providers/index.js";
import { store } from "./store.js";

/**
 * Sync engine.
 *
 * Polls each connected account's folder unread counts on an interval and emits
 * a "change" event when anything differs from the last snapshot. The API layer
 * relays those events to the browser over Server-Sent Events (SSE), so the
 * sidebar unread badges update without the client polling on its own.
 *
 * This is deliberately lightweight (folder unread counts, not full message
 * diffing). For true push you'd use Gmail watch + Pub/Sub and Graph webhook
 * subscriptions; IMAP has IDLE. Those need public callback URLs, so interval
 * polling is the portable default. Swap pollAccount() internals to upgrade.
 */

const POLL_INTERVAL_MS = Number(process.env.SYNC_INTERVAL_MS || 60_000);

class SyncEngine extends EventEmitter {
  constructor() {
    super();
    this.timer = null;
    this.snapshots = new Map(); // accountId → { folderId: unread }
    this.inFlight = new Set(); // accountIds currently being polled
  }

  start() {
    if (this.timer) return;
    this.timer = setInterval(() => this.tick().catch(() => {}), POLL_INTERVAL_MS);
    // Kick once shortly after boot so first data arrives fast.
    setTimeout(() => this.tick().catch(() => {}), 3_000);
    console.log(`[sync] polling every ${POLL_INTERVAL_MS / 1000}s`);
  }

  stop() {
    if (this.timer) clearInterval(this.timer);
    this.timer = null;
  }

  // Poll every account across all users.
  async tick() {
    const all = store.listAllRaw();
    await Promise.allSettled(all.map((acc) => this.pollAccount(acc)));
  }

  // Poll a single account; emit a change event if its unread map shifted.
  // Providers without listFolders (e.g. Slack) are skipped — their clients
  // fetch conversation state directly.
  async pollAccount(acc) {
    if (typeof getProvider(acc.provider).listFolders !== "function") return;
    if (this.inFlight.has(acc.id)) return;
    this.inFlight.add(acc.id);
    try {
      const folders = await getProvider(acc.provider).listFolders(acc);
      const next = Object.fromEntries(folders.map((f) => [f.id, f.unread]));
      const prev = this.snapshots.get(acc.id);
      this.snapshots.set(acc.id, next);

      if (!prev || changed(prev, next)) {
        this.emit("change", {
          userId: acc.userId,
          accountId: acc.id,
          folders: next,
          inboxUnread: next.inbox || 0,
          at: Date.now(),
        });
      }
    } catch (e) {
      this.emit("error", { accountId: acc.id, error: e.message });
    } finally {
      this.inFlight.delete(acc.id);
    }
  }

  // Force an immediate poll of one account (used by the manual refresh route).
  async refreshAccount(acc) {
    await this.pollAccount(acc);
    return this.snapshots.get(acc.id) || {};
  }
}

function changed(a, b) {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const k of keys) if ((a[k] || 0) !== (b[k] || 0)) return true;
  return false;
}

export const syncEngine = new SyncEngine();
