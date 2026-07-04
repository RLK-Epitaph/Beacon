import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { config } from "./config.js";

/**
 * Account/token store.
 *
 * This reference implementation persists to an encrypted JSON file so the
 * project runs with zero external dependencies. Tokens are encrypted at rest
 * with AES-256-GCM. For production, swap the load()/save() internals for a real
 * database (Postgres, Redis, etc.) — the public methods are the contract.
 *
 * Account shape:
 * {
 *   id, userId, provider, address, label,
 *   secrets: { ...provider-specific tokens... }  // never sent to the client
 * }
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DB_PATH = path.join(__dirname, "..", ".data", "accounts.enc");

function getKey() {
  const hex = config.tokenEncKey;
  if (!hex || hex.length !== 64) {
    // Dev fallback: derive a key so the app boots, but warn loudly.
    console.warn("[store] TOKEN_ENC_KEY missing/invalid — using insecure derived key. Set it for production.");
    return crypto.createHash("sha256").update("beacon-dev-key").digest();
  }
  return Buffer.from(hex, "hex");
}

function encrypt(plaintext) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", getKey(), iv);
  const enc = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, enc]).toString("base64");
}

function decrypt(b64) {
  const raw = Buffer.from(b64, "base64");
  const iv = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const data = raw.subarray(28);
  const decipher = crypto.createDecipheriv("aes-256-gcm", getKey(), iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(data), decipher.final()]).toString("utf8");
}

let accounts = [];

function load() {
  try {
    if (fs.existsSync(DB_PATH)) {
      accounts = JSON.parse(decrypt(fs.readFileSync(DB_PATH, "utf8")));
    }
  } catch (e) {
    console.error("[store] failed to load accounts:", e.message);
    accounts = [];
  }
}

function save() {
  fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
  fs.writeFileSync(DB_PATH, encrypt(JSON.stringify(accounts)));
}

load();

// Strip secrets before anything reaches the client.
function publicView(a) {
  return { id: a.id, provider: a.provider, address: a.address, label: a.label || "" };
}

export const store = {
  listForUser(userId) {
    return accounts.filter((a) => a.userId === userId).map(publicView);
  },

  listRawForUser(userId) {
    return accounts.filter((a) => a.userId === userId);
  },

  // Every account across all users — used by the sync engine's poll loop.
  listAllRaw() {
    return accounts.slice();
  },

  getRaw(userId, accountId) {
    return accounts.find((a) => a.userId === userId && a.id === accountId) || null;
  },

  upsert({ userId, provider, address, label, secrets }) {
    // De-dupe on (userId, provider, address) so re-connecting refreshes tokens.
    let acc = accounts.find(
      (a) => a.userId === userId && a.provider === provider && a.address === address
    );
    if (acc) {
      acc.secrets = { ...acc.secrets, ...secrets };
      if (label !== undefined) acc.label = label;
    } else {
      acc = {
        id: "acc_" + crypto.randomBytes(6).toString("hex"),
        userId,
        provider,
        address,
        label: label || "",
        secrets,
      };
      accounts.push(acc);
    }
    save();
    return publicView(acc);
  },

  updateSecrets(accountId, secrets) {
    const acc = accounts.find((a) => a.id === accountId);
    if (acc) {
      acc.secrets = { ...acc.secrets, ...secrets };
      save();
    }
  },

  setLabel(userId, accountId, label) {
    const acc = accounts.find((a) => a.userId === userId && a.id === accountId);
    if (!acc) return null;
    acc.label = label;
    save();
    return publicView(acc);
  },

  remove(userId, accountId) {
    const before = accounts.length;
    accounts = accounts.filter((a) => !(a.userId === userId && a.id === accountId));
    save();
    return accounts.length < before;
  },
};
