import express from "express";
import session from "express-session";
import cookieParser from "cookie-parser";
import cors from "cors";
import { config, providerConfigured } from "./config.js";
import { authRouter } from "./routes/auth.js";
import { apiRouter } from "./routes/api.js";
import { syncEngine } from "./sync.js";
import { FileSessionStore } from "./sessionStore.js";

const app = express();
app.set("trust proxy", 1);

app.use(
  cors({
    origin: config.clientOrigin,
    credentials: true, // allow the session cookie cross-origin
  })
);
app.use(express.json());
app.use(cookieParser());
// Behind ngrok or any HTTPS origin, the frontend (localhost:5173) is a
// DIFFERENT site from the backend, so the session cookie must be
// SameSite=None + Secure or browsers will drop it on cross-site fetches —
// which would make OAuth-connected accounts invisible to the app.
const secureCookies = config.serverOrigin.startsWith("https");
app.use(
  session({
    store: new FileSessionStore(),
    secret: config.sessionSecret,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: secureCookies ? "none" : "lax",
      secure: secureCookies,
      maxAge: 1000 * 60 * 60 * 24 * 14, // 14 days
    },
  })
);

// Health + which providers are usable right now.
app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    providers: {
      google: providerConfigured("google"),
      microsoft: providerConfigured("microsoft"),
      apple: providerConfigured("apple"),
    },
  });
});

app.use("/auth", authRouter);
app.use("/api", apiRouter);

app.use((err, _req, res, _next) => {
  console.error("[unhandled]", err);
  res.status(500).json({ error: "Internal server error" });
});

app.listen(config.port, () => {
  console.log(`\n  Beacon backend → ${config.serverOrigin}`);
  console.log(`  Client origin    → ${config.clientOrigin}`);
  console.log("  Providers configured:");
  for (const p of ["google", "microsoft", "apple"]) {
    console.log(`    ${providerConfigured(p) ? "✓" : "✗"} ${p}`);
  }
  console.log("");
  syncEngine.start();
});
