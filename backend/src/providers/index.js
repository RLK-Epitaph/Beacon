import { googleProvider } from "./google.js";
import { microsoftProvider } from "./microsoft.js";
import { appleProvider } from "./apple.js";

export const providers = {
  google: googleProvider,
  microsoft: microsoftProvider,
  apple: appleProvider,
};

export function getProvider(id) {
  const p = providers[id];
  if (!p) throw new Error(`Unknown provider: ${id}`);
  return p;
}
