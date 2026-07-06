import session from "express-session";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * File-backed session store.
 *
 * express-session defaults to MemoryStore, which is wiped whenever the
 * process restarts — including every save while `npm run dev` runs the
 * backend under `node --watch`. Losing sessions mid-development meant every
 * restart orphaned the current user's connected accounts (a fresh anonymous
 * session has no userId, so signing back in without realizing the old one
 * exists just mints a new user and starts empty).
 *
 * Persists alongside accounts.enc using the same reference-implementation
 * philosophy: a local JSON file, zero extra dependencies. Swap for
 * connect-redis/connect-pg-simple etc. in production.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SESS_PATH = path.join(__dirname, "..", ".data", "sessions.json");

function load() {
  try {
    return JSON.parse(fs.readFileSync(SESS_PATH, "utf8"));
  } catch {
    return {};
  }
}

function save(sessions) {
  fs.mkdirSync(path.dirname(SESS_PATH), { recursive: true });
  fs.writeFileSync(SESS_PATH, JSON.stringify(sessions));
}

export class FileSessionStore extends session.Store {
  constructor() {
    super();
    this.sessions = load();
  }

  get(sid, cb) {
    const rec = this.sessions[sid];
    if (!rec) return cb(null, null);
    if (rec.expires && rec.expires < Date.now()) {
      delete this.sessions[sid];
      save(this.sessions);
      return cb(null, null);
    }
    cb(null, rec.data);
  }

  set(sid, sessionData, cb) {
    const maxAge = sessionData.cookie?.maxAge;
    this.sessions[sid] = { data: sessionData, expires: maxAge ? Date.now() + maxAge : null };
    save(this.sessions);
    cb && cb();
  }

  destroy(sid, cb) {
    delete this.sessions[sid];
    save(this.sessions);
    cb && cb();
  }

  touch(sid, sessionData, cb) {
    const rec = this.sessions[sid];
    if (rec) {
      rec.data = sessionData;
      const maxAge = sessionData.cookie?.maxAge;
      rec.expires = maxAge ? Date.now() + maxAge : null;
      save(this.sessions);
    }
    cb && cb();
  }
}
