import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Persistent Slack per-conversation last-activity cache.
 *
 * Determining a conversation's last-message time requires conversations.history,
 * which Slack rate-limits to ~1 request/minute for this app tier AND is the same
 * endpoint used to actually open a channel — so we must not spend it scanning in
 * the background. Instead we record the real last-message timestamp for free
 * every time the user opens a channel (listMessages already fetches it), and fall
 * back to the cheap-but-approximate `updated` field for conversations not yet
 * visited. Persisting to disk means a conversation, once seen, stays accurately
 * classified across restarts.
 *
 * Shape: { [accountId]: { [channelId]: { ts, at } } }
 *   ts = unix seconds of the last message (0 = reachable but empty)
 *   at = when we recorded it (ms)
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FILE = path.join(__dirname, "..", ".data", "slack-activity.json");

let data = load();
function load() {
  try {
    return JSON.parse(fs.readFileSync(FILE, "utf8"));
  } catch {
    return {};
  }
}

let writeTimer = null;
function scheduleSave() {
  if (writeTimer) return;
  writeTimer = setTimeout(() => {
    writeTimer = null;
    try {
      fs.mkdirSync(path.dirname(FILE), { recursive: true });
      fs.writeFileSync(FILE, JSON.stringify(data));
    } catch {
      /* best-effort cache; a failed write just means we re-record it later */
    }
  }, 1000);
}

export function getActivity(accountId) {
  return data[accountId] || {};
}

export function setActivity(accountId, channelId, ts) {
  if (!data[accountId]) data[accountId] = {};
  const prev = data[accountId][channelId];
  // Never move a known timestamp backwards (an older page shouldn't overwrite a
  // newer last-message time we already recorded).
  if (prev && ts <= prev.ts) {
    prev.at = Date.now();
  } else {
    data[accountId][channelId] = { ts, at: Date.now() };
  }
  scheduleSave();
}
