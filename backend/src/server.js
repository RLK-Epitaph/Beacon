import express from "express";
import session from "express-session";
import cookieParser from "cookie-parser";
import cors from "cors";
import { config, providerConfigured } from "./config.js";
import { authRouter } from "./routes/auth.js";
import { apiRouter } from "./routes/api.js";
import { syncEngine } from "./sync.js";

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
app.use(
  session({
    secret: config.sessionSecret,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: config.serverOrigin.startsWith("https"),
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
