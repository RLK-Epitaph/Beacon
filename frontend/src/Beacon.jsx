import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Mail, Inbox, Send, FileText, Trash2, Star, Archive, AlertOctagonIcon,
  Plus, Search, Settings, ChevronDown, Check, Pencil, Reply, ReplyAll,
  Forward, MoreHorizontal, Paperclip, X, RefreshCw, CircleDot,
  ChevronsLeft, ChevronsRight, SmilePlus,
  Home, Calendar as CalendarIcon, Cloud, CloudSun, Sun, CloudRain,
  CheckSquare, Square, ListTodo, MapPin, Clock, Users, MessageSquare,
  List, ListOrdered, AlignLeft, AlignCenter, AlignRight,
  IndentIncrease, IndentDecrease, Quote, RemoveFormatting, MailOpen
} from "lucide-react";

/* =============================================================================
   Beacon — multi-account email client
   -----------------------------------------------------------------------------
   ARCHITECTURE / WHERE REAL PROVIDER APIS PLUG IN
   -----------------------------------------------------------------------------
   This is the FRONTEND. All data here comes from the mock store below. To make
   it real, you replace the functions in `mailService` with calls to a backend
   you control. A browser cannot talk to Gmail/Graph/iCloud IMAP directly
   (OAuth client secrets must stay server-side; CORS + IMAP block direct calls).

   OAuth, per provider (handled by YOUR backend, not here):
     • Google     → OAuth 2.0 + Gmail API     (scope: gmail.readonly / gmail.modify)
                    https://developers.google.com/gmail/api
     • Microsoft  → MSAL + Microsoft Graph     (scope: Mail.Read / Mail.ReadWrite)
                    https://learn.microsoft.com/graph/api/resources/mail-api-overview
     • Apple      → "Sign in with Apple" gets you identity only; mail itself is
                    iCloud IMAP/SMTP via app-specific passwords, proxied server-side.

   FLOW: button → window.open(`${BACKEND}/auth/${provider}`) → backend runs the
   OAuth dance, stores tokens, returns an account id → frontend calls
   mailService.listFolders/listMessages against your backend. Each function below
   is marked // REPLACE: ...
============================================================================= */

/* ----------------------------- mock data ---------------------------------- */

const PROVIDERS = {
  google: {
    name: "Google",
    badge: "G",
    color: "#EA4335",
    gradient: "linear-gradient(135deg,#4285F4,#EA4335)",
  },
  apple: {
    name: "Apple",
    badge: "",
    color: "#1d1d1f",
    gradient: "linear-gradient(135deg,#555,#1d1d1f)",
  },
  microsoft: {
    name: "Microsoft",
    badge: "M",
    color: "#0078D4",
    gradient: "linear-gradient(135deg,#0078D4,#00A4EF)",
  },
  slack: {
    name: "Slack",
    badge: "S",
    color: "#4A154B",
    gradient: "linear-gradient(135deg,#611f69,#4A154B)",
  },
  teams: {
    name: "Microsoft Teams",
    badge: "T",
    color: "#6264A7",
    gradient: "linear-gradient(135deg,#7B83EB,#4B53BC)",
  },
};

const FOLDERS = [
  { id: "inbox", name: "Inbox", icon: Inbox },
  { id: "starred", name: "Starred", icon: Star },
  { id: "sent", name: "Sent", icon: Send },
  { id: "drafts", name: "Drafts", icon: FileText },
  { id: "archive", name: "Archive", icon: Archive },
  { id: "spam", name: "Spam", icon: AlertOctagonIcon },
  { id: "trash", name: "Trash", icon: Trash2 },
];

let _id = 0;
const uid = () => `m${++_id}`;

function msg(from, fromAddr, subject, preview, body, opts = {}) {
  return {
    id: uid(),
    from,
    fromAddr,
    subject,
    preview,
    body,
    date: opts.date || "10:24 AM",
    fullDate: opts.fullDate || "Today, 10:24 AM",
    unread: opts.unread ?? false,
    starred: opts.starred ?? false,
    folder: opts.folder || "inbox",
    attachments: opts.attachments || [],
  };
}

// Seed accounts. In production this list comes from your backend after OAuth.
const SEED_ACCOUNTS = [
  {
    id: "acc_g1",
    provider: "google",
    address: "alex.rivera@gmail.com",
    label: "Personal",
    messages: [
      msg("Spotify", "no-reply@spotify.com", "Your 2025 Wrapped is here",
        "See the artists and songs that defined your year…",
        "Hi Alex,\n\nYour year in music is ready. You listened to 41,203 minutes across 1,847 different artists. Your top genre was indie folk.\n\nTap below to see the full breakdown and share your story.\n\n— The Spotify Team",
        { unread: true, date: "9:41 AM", fullDate: "Today, 9:41 AM" }),
      msg("Mom", "linda.rivera@gmail.com", "Sunday dinner?",
        "Are you coming over this weekend? I'm making the…",
        "Hi sweetie,\n\nAre you free Sunday? I'm making lasagna and your sister is bringing the kids. Let me know by Friday so I can plan.\n\nLove,\nMom",
        { unread: true, starred: true, date: "8:15 AM", fullDate: "Today, 8:15 AM" }),
      msg("REI Co-op", "members@rei.com", "Members get 20% off one item",
        "Your Co-op dividend plus an extra savings event…",
        "Hello Alex,\n\nAs a Co-op member you've earned a $48.20 dividend this year. Plus, members save 20% on one full-price item through Monday.\n\nHappy trails,\nREI",
        { date: "Yesterday", fullDate: "Yesterday, 4:02 PM" }),
      msg("GitHub", "noreply@github.com", "[octo/web] PR #482 was merged",
        "tessa-w merged 3 commits into main…",
        "Your pull request #482 'Fix race condition in sync queue' was merged into main by tessa-w.\n\nView the diff and deployment status on GitHub.",
        { folder: "archive", date: "Mon", fullDate: "Monday, 11:30 AM" }),
    ],
  },
  {
    id: "acc_m1",
    provider: "microsoft",
    address: "a.rivera@northwind.com",
    label: "Work",
    messages: [
      msg("Priya Shah", "priya.shah@northwind.com", "Q3 planning deck — review by EOD",
        "Hey Alex, dropped the latest version in the shared…",
        "Hey Alex,\n\nThe Q3 planning deck is ready for your review — it's in the shared drive under /Planning/Q3. Can you look at slides 8–14 (the roadmap section) and leave comments by end of day? Leadership review is tomorrow at 9.\n\nThanks,\nPriya",
        { unread: true, starred: true, date: "11:02 AM", fullDate: "Today, 11:02 AM",
          attachments: [{ name: "Q3-Planning-v4.pptx", size: "8.2 MB" }] }),
      msg("IT Service Desk", "itdesk@northwind.com", "Action required: password expires in 3 days",
        "Your network password will expire soon. Update it…",
        "Your Northwind network password will expire in 3 days. To avoid losing access, update it through the self-service portal.\n\nThis is an automated message.",
        { unread: true, date: "10:30 AM", fullDate: "Today, 10:30 AM" }),
      msg("Calendar", "calendar@northwind.com", "Accepted: Design sync (2:00–2:30 PM)",
        "Tessa Wong accepted your invitation…",
        "Tessa Wong accepted your meeting invitation for Design sync, today 2:00–2:30 PM, Room 4B / Teams.",
        { date: "9:50 AM", fullDate: "Today, 9:50 AM" }),
      msg("Alex Rivera", "a.rivera@northwind.com", "Re: Budget numbers",
        "Sending these over now — let me know if the…",
        "Hi Marcus,\n\nSending these over now — let me know if the Q2 figures line up with what finance has. I rounded the travel line to the nearest hundred.\n\nBest,\nAlex",
        { folder: "sent", date: "Tue", fullDate: "Tuesday, 3:15 PM" }),
    ],
  },
  {
    id: "acc_a1",
    provider: "apple",
    address: "alex@icloud.com",
    label: "Side projects",
    messages: [
      msg("Stripe", "support@stripe.com", "You received a payment of $129.00",
        "A payment was successfully processed for Lumen…",
        "Nice work!\n\nA payment of $129.00 USD was successfully processed for your account 'Lumen Studio'. Funds will be deposited to your bank account in 2 business days.\n\n— Stripe",
        { unread: true, date: "7:20 AM", fullDate: "Today, 7:20 AM" }),
      msg("Vercel", "notifications@vercel.com", "Deployment ready: lumen-app",
        "Your deployment to production succeeded…",
        "Your deployment of lumen-app to Production succeeded in 38s.\n\nCommit: 'Add dark mode toggle'\nURL: lumen-app.vercel.app",
        { date: "Yesterday", fullDate: "Yesterday, 10:11 PM" }),
      msg("Namecheap", "support@namecheap.com", "lumenstudio.io renews in 14 days",
        "Auto-renew is on. No action needed unless…",
        "Heads up — your domain lumenstudio.io will auto-renew in 14 days for $13.98/yr. No action needed unless you'd like to make changes.",
        { folder: "archive", date: "Sun", fullDate: "Sunday, 1:00 PM" }),
    ],
  },
];

// Tag every seeded email account so the rail/router can tell kinds apart.
for (const a of SEED_ACCOUNTS) a.kind = "email";

/* ------------------------------ slack data -------------------------------- */

// Slack message factory. Replies live inline on a message via `replies`.
function smsg(author, text, opts = {}) {
  return {
    id: uid(),
    author,
    text,
    time: opts.time || "10:24 AM",
    reactions: opts.reactions || [], // [{ emoji, count }]
    replies: opts.replies || [],     // nested smsg() objects (one level: a thread)
  };
}

const SLACK_WORKSPACES = [
  {
    id: "ws_acme",
    kind: "slack",
    provider: "slack",
    workspace: "Acme Corp",
    label: "Acme Corp",
    address: "alex@acme.com",
    accent: "#4A154B",
    you: "Alex Rivera",
    channels: [
      {
        id: "c_general", name: "general", kind: "channel", unread: 2,
        messages: [
          smsg("Dana Lee", "Morning everyone — reminder that the all-hands is at 11 today.", {
            time: "8:30 AM",
            reactions: [{ emoji: "👍", count: 4 }, { emoji: "🎉", count: 1 }],
          }),
          smsg("Marcus Chen", "Thanks Dana. Will the recording be shared after?", {
            time: "8:34 AM",
            replies: [
              smsg("Dana Lee", "Yes, I'll post it in here by EOD.", { time: "8:36 AM" }),
              smsg("Marcus Chen", "Perfect, thank you!", { time: "8:37 AM" }),
            ],
          }),
          smsg("Priya Shah", "Quick heads up: the office will be closed Friday for maintenance.", {
            time: "9:15 AM",
            reactions: [{ emoji: "🙏", count: 6 }],
          }),
        ],
      },
      {
        id: "c_eng", name: "engineering", kind: "channel", unread: 0,
        messages: [
          smsg("Tessa Wong", "Deployed v2.4.1 to staging. Please smoke-test the sync queue.", {
            time: "Yesterday",
            reactions: [{ emoji: "🚀", count: 3 }],
            replies: [
              smsg("Alex Rivera", "Looks good on my end — race condition is gone.", { time: "Yesterday" }),
              smsg("Tessa Wong", "🙌 shipping to prod tomorrow then.", { time: "Yesterday" }),
            ],
          }),
          smsg("Sam Okafor", "Anyone seeing flaky CI on the auth tests?", { time: "Yesterday" }),
        ],
      },
      {
        id: "c_design", name: "design", kind: "channel", unread: 5,
        messages: [
          smsg("Jordan Park", "New icon set is in Figma, link in the thread.", {
            time: "7:50 AM",
            replies: [smsg("Jordan Park", "figma.com/file/acme-icons", { time: "7:50 AM" })],
          }),
        ],
      },
      {
        id: "c_random", name: "random", kind: "channel", unread: 0,
        messages: [
          smsg("Marcus Chen", "Coffee run — anyone want anything? ☕", {
            time: "10:02 AM",
            reactions: [{ emoji: "☕", count: 2 }],
          }),
        ],
      },
      {
        id: "dm_dana", name: "Dana Lee", kind: "dm", unread: 1,
        messages: [
          smsg("Dana Lee", "Hey, can you review my PR when you get a sec?", { time: "9:40 AM" }),
        ],
      },
      {
        id: "dm_tessa", name: "Tessa Wong", kind: "dm", unread: 0,
        messages: [
          smsg("Tessa Wong", "ty for the help earlier 🙏", { time: "Yesterday" }),
        ],
      },
    ],
  },
  {
    id: "ws_lumen",
    kind: "slack",
    provider: "slack",
    workspace: "Lumen Studio",
    label: "Lumen Studio",
    address: "alex@lumenstudio.io",
    accent: "#2EB67D",
    you: "Alex Rivera",
    channels: [
      {
        id: "l_general", name: "general", kind: "channel", unread: 3,
        messages: [
          smsg("Robin Diaz", "Stripe payment came through for the new client 🎉", {
            time: "7:20 AM",
            reactions: [{ emoji: "💰", count: 2 }, { emoji: "🎉", count: 3 }],
          }),
          smsg("Robin Diaz", "Kicking off the project Monday — kickoff doc incoming.", { time: "7:22 AM" }),
        ],
      },
      {
        id: "l_launch", name: "launch-plan", kind: "channel", unread: 0,
        messages: [
          smsg("Alex Rivera", "Landing page copy is final. Dev handoff ready.", {
            time: "Yesterday",
            replies: [smsg("Robin Diaz", "Amazing, I'll start building tonight.", { time: "Yesterday" })],
          }),
        ],
      },
      {
        id: "l_dm_robin", name: "Robin Diaz", kind: "dm", unread: 0,
        messages: [smsg("Robin Diaz", "see you at standup", { time: "8:00 AM" })],
      },
    ],
  },
];

/* ------------------------------ teams data -------------------------------- */

// Teams post factory. In Teams every post IS a thread: a root message with
// inline replies shown beneath it in the channel.
function tpost(author, text, opts = {}) {
  return {
    id: uid(),
    author,
    text,
    time: opts.time || "10:24 AM",
    reactions: opts.reactions || [], // [{ emoji, count, mine? }]
    replies: opts.replies || [],     // tpost() objects (no nested threads)
  };
}

const TEAMS_ACCOUNTS = [
  {
    id: "tm_northwind",
    kind: "teams",
    provider: "teams",
    org: "Northwind",
    label: "Northwind",
    address: "a.rivera@northwind.com",
    you: "Alex Rivera",
    teams: [
      {
        id: "team_product", name: "Product", initials: "PR", color: "#4B53BC",
        channels: [
          {
            id: "tc_prod_general", name: "General", unread: 2,
            posts: [
              tpost("Priya Shah", "Q3 planning deck is finalized — thanks everyone for the feedback. Leadership review is tomorrow at 9.", {
                time: "11:05 AM",
                reactions: [{ emoji: "👍", count: 5 }, { emoji: "🎉", count: 2 }],
                replies: [
                  tpost("Marcus Chen", "Great work. Are we presenting the roadmap section too?", { time: "11:12 AM" }),
                  tpost("Priya Shah", "Yes — slides 8–14. Alex reviewed them yesterday.", { time: "11:15 AM" }),
                ],
              }),
              tpost("Dana Lee", "Reminder: submit your OKR drafts by Friday EOD.", {
                time: "9:30 AM",
                reactions: [{ emoji: "👀", count: 3 }],
                replies: [],
              }),
            ],
          },
          {
            id: "tc_prod_roadmap", name: "Roadmap", unread: 0,
            posts: [
              tpost("Alex Rivera", "Posted the updated H2 roadmap in Files. Major change: sync engine moves up to August.", {
                time: "Yesterday",
                reactions: [{ emoji: "🚀", count: 4 }],
                replies: [
                  tpost("Tessa Wong", "That works — v2.4.1 clears the way for it.", { time: "Yesterday" }),
                ],
              }),
            ],
          },
        ],
      },
      {
        id: "team_eng", name: "Engineering", initials: "EN", color: "#038387",
        channels: [
          {
            id: "tc_eng_general", name: "General", unread: 1,
            posts: [
              tpost("Sam Okafor", "CI is green again — the auth test flakiness was a clock-skew issue in the container.", {
                time: "10:40 AM",
                reactions: [{ emoji: "🙌", count: 6 }],
                replies: [
                  tpost("Tessa Wong", "Amazing. Merging my PR then.", { time: "10:44 AM" }),
                ],
              }),
            ],
          },
          {
            id: "tc_eng_deploys", name: "Deploys", unread: 0,
            posts: [
              tpost("Deploy Bot", "✅ v2.4.1 deployed to production. 0 errors in the first hour.", {
                time: "8:00 AM",
                reactions: [{ emoji: "🎉", count: 8 }],
                replies: [],
              }),
            ],
          },
        ],
      },
    ],
    // Teams "Chat" — 1:1 and group chats, separate from team channels.
    chats: [
      {
        id: "chat_priya", name: "Priya Shah", kind: "chat", unread: 1,
        messages: [
          tpost("Priya Shah", "Did you get a chance to look at slides 8–14?", { time: "10:50 AM" }),
          tpost("Alex Rivera", "Yes — left comments on 9 and 12, rest looks great.", { time: "10:58 AM" }),
          tpost("Priya Shah", "Perfect. One more thing — can you present slide 12 tomorrow?", { time: "11:20 AM" }),
        ],
      },
      {
        id: "chat_standup", name: "Platform standup", kind: "group", unread: 2,
        messages: [
          tpost("Tessa Wong", "Standup moved to 9:45 today, conflict with the all-hands.", { time: "8:12 AM" }),
          tpost("Sam Okafor", "👍 works for me", { time: "8:15 AM" }),
          tpost("Marcus Chen", "Same. I'll update the invite.", { time: "8:16 AM", reactions: [{ emoji: "🙏", count: 2 }] }),
        ],
      },
      {
        id: "chat_it", name: "IT Service Desk", kind: "chat", unread: 0,
        messages: [
          tpost("IT Service Desk", "Your password reset ticket #4821 has been resolved.", { time: "Yesterday" }),
          tpost("Alex Rivera", "Confirmed working, thanks!", { time: "Yesterday" }),
        ],
      },
    ],
  },
  {
    id: "tm_contoso",
    kind: "teams",
    provider: "teams",
    org: "Contoso Consulting",
    label: "Contoso (client)",
    address: "alex.r@contoso.com",
    you: "Alex Rivera",
    teams: [
      {
        id: "team_migration", name: "Cloud Migration", initials: "CM", color: "#C239B3",
        channels: [
          {
            id: "tc_mig_general", name: "General", unread: 3,
            posts: [
              tpost("Jordan Blake", "Phase 2 kickoff is confirmed for Monday. Agenda attached in Files tab.", {
                time: "9:00 AM",
                reactions: [{ emoji: "👍", count: 2 }],
                replies: [
                  tpost("Alex Rivera", "I'll have the data-mapping doc ready by then.", { time: "9:20 AM" }),
                  tpost("Jordan Blake", "Perfect, thank you Alex.", { time: "9:22 AM" }),
                ],
              }),
            ],
          },
        ],
      },
    ],
    chats: [
      {
        id: "chat_jordan", name: "Jordan Blake", kind: "chat", unread: 1,
        messages: [
          tpost("Jordan Blake", "Quick one — can we push our sync 30 min later on Monday?", { time: "9:40 AM" }),
        ],
      },
    ],
  },
];

/* ------------------------------ home data --------------------------------- */

// Live weather snapshot for the user's location (Locust, NC) captured at build.
// REPLACE: fetch from a weather API (Open-Meteo, OpenWeather) keyed on the
// user's geolocation; shape: { temp, condition, high, location, forecast[] }.
const WEATHER = {
  location: "Locust, NC",
  temp: 87,
  condition: "Partly sunny",
  icon: "cloud-sun",
  high: 89,
  low: 71,
  forecast: [
    { day: "Tue", high: 90, icon: "cloud-sun" },
    { day: "Wed", high: 93, icon: "cloud-rain" },
    { day: "Thu", high: 94, icon: "sun" },
    { day: "Fri", high: 95, icon: "sun" },
  ],
};

// Calendar events, each tied to the email account whose calendar it lives on.
// REPLACE: per account, pull from that provider's calendar API
//   google    → Google Calendar API (calendar.events.list)
//   microsoft → Microsoft Graph (/me/events or /me/calendarView)
//   apple     → CalDAV (iCloud calendars)
// Merge the per-account results client-side, as the widget does below.
const CALENDAR_EVENTS = [
  // Personal (Google)
  { id: "ev1", accountId: "acc_g1", title: "Dentist appointment", time: "9:00 AM", duration: "9:00 – 9:45 AM", today: true },
  { id: "ev2", accountId: "acc_g1", title: "Dinner with Mom", time: "6:30 PM", duration: "6:30 – 8:00 PM", today: true },
  { id: "ev3", accountId: "acc_g1", title: "Weekend hike", time: "Sat, 8:00 AM", duration: "Saturday, 8:00 AM", today: false },
  // Work (Microsoft)
  { id: "ev4", accountId: "acc_m1", title: "All-hands meeting", time: "11:00 AM", duration: "11:00 – 11:45 AM", today: true },
  { id: "ev5", accountId: "acc_m1", title: "Design sync", time: "2:00 PM", duration: "2:00 – 2:30 PM", today: true },
  { id: "ev6", accountId: "acc_m1", title: "1:1 with Priya", time: "4:00 PM", duration: "4:00 – 4:30 PM", today: true },
  { id: "ev7", accountId: "acc_m1", title: "Sprint planning", time: "Mon, 1:00 PM", duration: "Monday, 1:00 PM", today: false },
  // Side projects (Apple)
  { id: "ev8", accountId: "acc_a1", title: "Client kickoff (Lumen)", time: "Mon, 9:00 AM", duration: "Monday, 9:00 AM", today: false },
  { id: "ev9", accountId: "acc_a1", title: "Stripe payout review", time: "Wed, 10:00 AM", duration: "Wednesday, 10:00 AM", today: false },
];

// Slack Lists — Slack's lightweight project/task lists, aggregated across
// workspaces for the home view.
// REPLACE: Slack API (lists are accessed via the Slack Lists API / slackLists.*).
const SLACK_LISTS = [
  {
    id: "list_sprint",
    name: "Q3 Sprint board",
    workspace: "Acme Corp",
    accent: "#4A154B",
    items: [
      { id: "t1", title: "Fix sync queue race condition", status: "done", assignee: "Tessa Wong", due: "Today" },
      { id: "t2", title: "Ship v2.4.1 to production", status: "in_progress", assignee: "Alex Rivera", due: "Tomorrow" },
      { id: "t3", title: "Review auth test flakiness", status: "todo", assignee: "Sam Okafor", due: "Wed" },
      { id: "t4", title: "Update API docs", status: "todo", assignee: "Alex Rivera", due: "Fri" },
    ],
  },
  {
    id: "list_launch",
    name: "Launch checklist",
    workspace: "Lumen Studio",
    accent: "#2EB67D",
    items: [
      { id: "t5", title: "Finalize landing page copy", status: "done", assignee: "Alex Rivera", due: "Yesterday" },
      { id: "t6", title: "Build landing page", status: "in_progress", assignee: "Robin Diaz", due: "Today" },
      { id: "t7", title: "Set up Stripe webhooks", status: "todo", assignee: "Robin Diaz", due: "Thu" },
    ],
  },
];

/* --------------------- service layer (swap for real APIs) ------------------ */

const mailService = {
  // REPLACE: GET ${BACKEND}/accounts  → returns connected accounts after OAuth
  async listAccounts() {
    return structuredClone(SEED_ACCOUNTS);
  },
  // REPLACE: window.open(`${BACKEND}/auth/${provider}`) then poll/await the result.
  // Returns a newly connected account object.
  async connect(provider) {
    await new Promise((r) => setTimeout(r, 900)); // simulate the OAuth round-trip
    const sample = {
      google: "new.user@gmail.com",
      microsoft: "new.user@outlook.com",
      apple: "new.user@icloud.com",
    }[provider];
    return {
      id: `acc_${Math.random().toString(36).slice(2, 8)}`,
      provider,
      address: sample,
      label: "",
      messages: [
        msg("Welcome", "team@beacon.app", "Your inbox is connected",
          "Messages will appear here once syncing finishes…",
          "This account is connected. In a production build, your real messages would sync in from the provider here.",
          { unread: true }),
      ],
    };
  },
  // REPLACE: PATCH ${BACKEND}/messages/:id  { read: true }  (gmail.modify / Mail.ReadWrite)
  async markRead() { return true; },
  // REPLACE: PATCH ${BACKEND}/messages/:id  { starred }
  async toggleStar() { return true; },
};

/* ----- slack service (swap for Slack Web API + Socket Mode later) --------- */

const slackService = {
  // REPLACE: GET ${BACKEND}/slack/workspaces  → workspaces after Slack OAuth
  async listWorkspaces() {
    return structuredClone(SLACK_WORKSPACES);
  },
  // REPLACE: kick off Slack OAuth (https://slack.com/oauth/v2/authorize),
  // exchange code server-side, return the workspace + channels.
  async connect() {
    await new Promise((r) => setTimeout(r, 900));
    const n = Math.random().toString(36).slice(2, 6);
    return {
      id: `ws_${n}`,
      kind: "slack",
      provider: "slack",
      workspace: "New Workspace",
      label: "New Workspace",
      address: `you@workspace-${n}.com`,
      accent: "#4A154B",
      you: "You",
      channels: [
        {
          id: `c_${n}`, name: "general", kind: "channel", unread: 1,
          messages: [
            smsg("Slackbot", "Welcome to your new workspace! Channels will sync here.", { time: "now" }),
          ],
        },
      ],
    };
  },
  // REPLACE: POST chat.postMessage { channel, text }
  async postMessage() { return true; },
  // REPLACE: POST chat.postMessage { channel, thread_ts, text }
  async postReply() { return true; },
  // REPLACE: POST reactions.add / reactions.remove { channel, timestamp, name }
  async toggleReaction() { return true; },
};

/* ----- live slack service (talks to the Beacon backend) ------------------- */
// Enabled when the frontend runs with VITE_BACKEND_URL set (frontend/.env.local).
// Without it, Beacon runs entirely on the mock data above.
const BACKEND = import.meta.env?.VITE_BACKEND_URL || null;

async function backendApi(path, options = {}) {
  const res = await fetch(`${BACKEND}${path}`, {
    credentials: "include",
    headers: {
      "Content-Type": "application/json",
      // ngrok's free tier intercepts requests with an HTML warning page unless
      // this header is present — without it, API calls randomly "fail" after
      // cold starts (looking like logouts or stale/empty data).
      "ngrok-skip-browser-warning": "true",
      ...(options.headers || {}),
    },
    ...options,
  });
  const isJson = (res.headers.get("content-type") || "").includes("application/json");
  if (!res.ok || !isJson) {
    const body = isJson ? await res.json().catch(() => ({})) : {};
    throw new Error(
      body.error ||
        (!isJson
          ? "Backend returned a non-JSON response (tunnel interstitial or server offline)"
          : `Request failed: ${res.status}`)
    );
  }
  return res.json();
}

const liveSlackService = {
  // Slack accounts come from the shared /api/accounts store; each is hydrated
  // with its conversation list. Messages load lazily per channel (null = not
  // loaded yet), unlike the mock which ships everything upfront.
  async listWorkspaces() {
    const { accounts } = await backendApi("/api/accounts");
    const slackAccounts = accounts.filter((a) => a.provider === "slack");
    return Promise.all(
      slackAccounts.map(async (a) => {
        let channels = [];
        try {
          const r = await backendApi(`/api/slack/${a.id}/conversations`);
          channels = r.conversations.map((c) => ({ ...c, messages: null }));
        } catch (e) {
          console.warn("Slack conversations failed:", e.message);
        }
        return {
          id: a.id, kind: "slack", provider: "slack",
          workspace: a.address, label: a.label || a.address, address: a.address,
          accent: "#4A154B", you: "You", channels,
        };
      })
    );
  },
  // Full-page OAuth redirect; the backend bounces back with ?connect=ok.
  connect() {
    window.location.href = `${BACKEND}/auth/slack`;
    return new Promise(() => {});
  },
  async listMessages(accountId, channelId) {
    const { messages } = await backendApi(
      `/api/slack/${accountId}/conversations/${channelId}/messages`
    );
    return messages.map((m) => ({ ...m, reactions: m.reactions || [], replies: [] }));
  },
  async listReplies(accountId, channelId, ts) {
    const { messages } = await backendApi(
      `/api/slack/${accountId}/conversations/${channelId}/thread/${ts}`
    );
    return messages.map((m) => ({ ...m, reactions: m.reactions || [], replies: [] }));
  },
  async postMessage(accountId, channelId, text, threadTs = null) {
    await backendApi(`/api/slack/${accountId}/conversations/${channelId}/messages`, {
      method: "POST",
      body: JSON.stringify({ text, ...(threadTs ? { thread_ts: threadTs } : {}) }),
    });
  },
  async toggleReaction(accountId, channelId, timestamp, emoji, add) {
    await backendApi(`/api/slack/${accountId}/conversations/${channelId}/reactions`, {
      method: "POST",
      body: JSON.stringify({ timestamp, emoji, add }),
    });
  },
};

const profileService = {
  // Live mode: session-backed auth on the backend. Mock mode: a built-in
  // demo identity so the app works without a backend.
  async get() {
    if (!BACKEND) return { firstName: "Alex", lastName: "Rivera", email: "alex.rivera@gmail.com" };
    const { user } = await backendApi("/api/me");
    return user;
  },
  async signup(fields) {
    const { user } = await backendApi("/api/auth/signup", { method: "POST", body: JSON.stringify(fields) });
    return user;
  },
  async login(email, password) {
    const { user } = await backendApi("/api/auth/login", { method: "POST", body: JSON.stringify({ email, password }) });
    return user;
  },
  async logout() {
    await backendApi("/api/auth/logout", { method: "POST" });
  },
  async update(fields) {
    if (!BACKEND) return fields;
    const { user } = await backendApi("/api/me", { method: "PATCH", body: JSON.stringify(fields) });
    return user;
  },
  async changePassword(currentPassword, newPassword) {
    await backendApi("/api/me/password", { method: "PATCH", body: JSON.stringify({ currentPassword, newPassword }) });
  },
  async disconnectAccount(accountId) {
    if (!BACKEND) return;
    await backendApi(`/api/accounts/${accountId}`, { method: "DELETE" }).catch(() => {});
  },
};

const liveMailService = {
  // Mail accounts from the shared backend store (Google/Microsoft OAuth,
  // Apple app-password). Messages load lazily per folder.
  async listAccounts() {
    const { accounts } = await backendApi("/api/accounts");
    return accounts
      .filter((a) => ["google", "microsoft", "apple"].includes(a.provider))
      .map((a) => ({ ...a, kind: "email", messages: [], _loadedFolders: {} }));
  },
  async listFolderMessages(accountId, folder) {
    const { messages } = await backendApi(`/api/accounts/${accountId}/folders/${folder}/messages?limit=30`);
    return messages;
  },
  async getMessage(accountId, id, folder) {
    const { message } = await backendApi(`/api/accounts/${accountId}/messages/${id}?folder=${folder}`);
    return message;
  },
  connectRedirect(provider) {
    window.location.href = `${BACKEND}/auth/${provider}`;
    return new Promise(() => {});
  },
  async connectApple(address, appPassword) {
    const { account } = await backendApi("/auth/apple/password", {
      method: "POST",
      body: JSON.stringify({ address, appPassword }),
    });
    return { ...account, kind: "email", messages: [], _loadedFolders: {} };
  },
  async markRead(accountId, id, read, folder) {
    await backendApi(`/api/accounts/${accountId}/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ read, folder }),
    });
  },
  async toggleStar(accountId, id, starred, folder) {
    await backendApi(`/api/accounts/${accountId}/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ starred, folder }),
    });
  },
  async send(accountId, { to, subject, body, html }) {
    await backendApi(`/api/accounts/${accountId}/send`, {
      method: "POST",
      body: JSON.stringify({ to, subject, body, html }),
    });
  },
  async refresh(accountId) {
    return backendApi(`/api/accounts/${accountId}/refresh`, { method: "POST" });
  },
  async move(accountId, id, dest, folder) {
    await backendApi(`/api/accounts/${accountId}/messages/${id}`, {
      method: "PATCH",
      body: JSON.stringify({ move: dest, folder }),
    });
  },
};

const slackAPI = BACKEND ? liveSlackService : slackService;

/* ----- teams service (swap for Microsoft Graph later) --------------------- */

const teamsService = {
  // REPLACE: GET ${BACKEND}/teams/accounts → orgs after Microsoft OAuth.
  // Graph: /me/joinedTeams, then /teams/{id}/channels per team.
  async listAccounts() {
    return structuredClone(TEAMS_ACCOUNTS);
  },
  // REPLACE: Microsoft identity OAuth (same flow as Outlook mail, with
  // Team.ReadBasic.All + ChannelMessage.Read.All/Send scopes).
  async connect() {
    await new Promise((r) => setTimeout(r, 900));
    const n = Math.random().toString(36).slice(2, 6);
    return {
      id: `tm_${n}`,
      kind: "teams",
      provider: "teams",
      org: "New Organization",
      label: "New Organization",
      address: `you@org-${n}.com`,
      you: "You",
      teams: [
        {
          id: `team_${n}`, name: "General", initials: "GE", color: "#4B53BC",
          channels: [
            {
              id: `tc_${n}`, name: "General", unread: 0,
              posts: [tpost("Teams", "Welcome! Your teams and channels will sync here.", { time: "now" })],
            },
          ],
        },
      ],
      // REPLACE: Graph /me/chats (+ /chats/{id}/messages) for 1:1 and group chats.
      chats: [],
    };
  },
  // REPLACE: POST /teams/{id}/channels/{id}/messages
  async postMessage() { return true; },
  // REPLACE: POST /teams/{id}/channels/{id}/messages/{id}/replies
  async postReply() { return true; },
  // REPLACE: POST .../messages/{id}/setReaction | unsetReaction
  async toggleReaction() { return true; },
};

/* ---- calendar service (aggregates events across email accounts) ---------- */

const calendarService = {
  // REPLACE: for each connected email account, call that provider's calendar
  // API (Google Calendar / Microsoft Graph / iCloud CalDAV) and merge results.
  // Returns only events whose accountId matches a currently-configured account.
  async listEvents(accountIds) {
    const allowed = new Set(accountIds);
    return CALENDAR_EVENTS.filter((e) => allowed.has(e.accountId));
  },
};

/* -------------------------------- helpers --------------------------------- */

// Wrap an email's HTML for a sandboxed iframe: scripts are blocked by the
// sandbox, links open in a new tab, and images are constrained to the pane.
function emailHtmlDoc(html) {
  return `<!doctype html><html><head><meta charset="utf-8"><base target="_blank">
<style>
  body{margin:0;padding:4px 2px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
       font-size:14px;line-height:1.6;color:#2c3138;word-break:break-word}
  img{max-width:100%;height:auto}
  table{max-width:100%}
  a{color:#3b5bdb}
  blockquote{border-left:3px solid #e7eaee;margin:8px 0;padding:2px 0 2px 12px;color:#5b6470}
</style></head><body>${html}</body></html>`;
}

function initials(addr) {
  const name = addr.split("@")[0].replace(/[._]/g, " ");
  return name.split(" ").filter(Boolean).slice(0, 2).map((s) => s[0].toUpperCase()).join("");
}

function avatarColor(seed) {
  const palette = ["#6366F1", "#0EA5E9", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#EC4899", "#14B8A6"];
  let h = 0;
  for (const c of seed) h = (h * 31 + c.charCodeAt(0)) % palette.length;
  return palette[h];
}

/* ================================ app ===================================== */

export default function Beacon() {
  const [accounts, setAccounts] = useState([]);
  const [workspaces, setWorkspaces] = useState([]);
  const [events, setEvents] = useState([]);
  const [loading, setLoading] = useState(true);
  const [activeAccountId, setActiveAccountId] = useState(null);
  const [activeFolder, setActiveFolder] = useState("inbox");
  const [selectedId, setSelectedId] = useState(null);
  const [query, setQuery] = useState("");
  const [connecting, setConnecting] = useState(null);
  const [showConnect, setShowConnect] = useState(false);
  const [appleForm, setAppleForm] = useState(null);
  // Live unread counts pushed by the backend poller over SSE: { accId: { folderId: n } }
  const [liveUnread, setLiveUnread] = useState({});
  // Slack workspace unread from the backend poller: { accountId: { total, byChannel } }
  const [liveSlackUnread, setLiveSlackUnread] = useState({});
  const [editingLabel, setEditingLabel] = useState(null);
  const [composing, setComposing] = useState(false);
  const [railExpanded, setRailExpanded] = useState(false);
  // Slack-specific UI state
  const [activeChannelId, setActiveChannelId] = useState(null);
  const [openThreadId, setOpenThreadId] = useState(null);
  // User profile: null until loaded; live mode may require onboarding.
  const [profile, setProfile] = useState(null);
  const [profileLoaded, setProfileLoaded] = useState(false);
  const [showProfile, setShowProfile] = useState(false);
  // Teams-specific UI state: either a (team, channel) pair or a chat is active.
  const [teamsOrgs, setTeamsOrgs] = useState([]);
  const [teamsSel, setTeamsSel] = useState({ teamId: null, channelId: null, chatId: null });

  useEffect(() => {
    Promise.all([
      BACKEND ? liveMailService.listAccounts().catch((e) => { console.warn("Mail unavailable:", e.message); return []; }) : mailService.listAccounts(),
      slackAPI.listWorkspaces().catch((e) => { console.warn("Slack unavailable:", e.message); return []; }),
      BACKEND ? Promise.resolve([]) : teamsService.listAccounts(),
    ]).then(async ([accs, wss, tms]) => {
      setAccounts(accs);
      setWorkspaces(wss);
      setTeamsOrgs(tms);
      // Aggregate calendar events across the configured email accounts.
      const evs = await calendarService.listEvents(accs.map((a) => a.id));
      setEvents(evs);
      // Start on the home screen (activeAccountId stays null).
      setLoading(false);
    });
    profileService
      .get()
      .then(setProfile)
      .catch((e) => console.warn("Profile load failed:", e.message))
      .finally(() => setProfileLoaded(true));
  }, []);

  // Track the active view in a ref so the long-lived SSE handler always sees
  // the current selection without re-subscribing.
  const activeViewRef = useRef({});
  activeViewRef.current = { activeAccountId, activeFolder };

  // Real-time: subscribe to the backend's unread-change stream. When the
  // account being viewed changes, refetch its open folder so new mail appears
  // without any clicking around.
  useEffect(() => {
    if (!BACKEND) return;
    const es = new EventSource(`${BACKEND}/api/events`, { withCredentials: true });
    es.addEventListener("change", (e) => {
      try {
        const payload = JSON.parse(e.data);
        if (payload.kind === "slack") {
          // Notification: update the workspace badge + per-channel dots, and
          // reflect counts onto the loaded channel list so unread shows inline.
          setLiveSlackUnread((prev) => ({
            ...prev,
            [payload.accountId]: { total: payload.slackUnread || 0, byChannel: payload.byChannel || {} },
          }));
          setWorkspaces((prev) =>
            prev.map((w) =>
              w.id === payload.accountId
                ? {
                    ...w,
                    channels: w.channels.map((c) => ({
                      ...c,
                      unread: payload.byChannel?.[c.id] ?? 0,
                    })),
                  }
                : w
            )
          );
          return;
        }
        // Mail change event
        setLiveUnread((prev) => ({ ...prev, [payload.accountId]: payload.folders || {} }));
        const { activeAccountId: aid, activeFolder: af } = activeViewRef.current;
        if (payload.accountId === aid) refetchFolder(aid, af);
      } catch { /* malformed frame */ }
    });
    return () => es.close();
  }, []);

  // Handle the OAuth bounce-back (?connect=ok&account=...) from the backend.
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (!params.get("connect")) return;
    const status = params.get("connect");
    window.history.replaceState({}, "", window.location.pathname);
    if (status === "ok" && BACKEND) {
      liveSlackService.listWorkspaces().then(setWorkspaces).catch(() => {});
      liveMailService.listAccounts().then(setAccounts).catch(() => {});
    } else if (status === "error") {
      console.warn("Account connect failed:", params.get("reason"));
    }
  }, []);

  // No active account/workspace → show the home dashboard.
  const isHome = activeAccountId === null;

  // The active item can be an email account, a Slack workspace, or a Teams org.
  const activeWorkspace = useMemo(
    () => workspaces.find((w) => w.id === activeAccountId) || null,
    [workspaces, activeAccountId]
  );
  const isSlack = !!activeWorkspace;

  const activeTeamsOrg = useMemo(
    () => teamsOrgs.find((t) => t.id === activeAccountId) || null,
    [teamsOrgs, activeAccountId]
  );
  const isTeams = !!activeTeamsOrg;

  // Active Teams selection: a chat, or a team + channel (with fallbacks).
  const activeTeamsChannel = useMemo(() => {
    if (!activeTeamsOrg) return null;
    if (teamsSel.chatId) {
      const chat = (activeTeamsOrg.chats || []).find((c) => c.id === teamsSel.chatId);
      if (chat) return { chat };
    }
    const team =
      activeTeamsOrg.teams.find((t) => t.id === teamsSel.teamId) || activeTeamsOrg.teams[0];
    if (!team) return null;
    const channel =
      team.channels.find((c) => c.id === teamsSel.channelId) || team.channels[0];
    return channel ? { team, channel } : null;
  }, [activeTeamsOrg, teamsSel]);

  const unreadByTeamsOrg = useMemo(() => {
    const map = {};
    for (const org of teamsOrgs) {
      const channelUnread = org.teams.reduce(
        (sum, t) => sum + t.channels.reduce((s, c) => s + (c.unread || 0), 0),
        0
      );
      const chatUnread = (org.chats || []).reduce((s, c) => s + (c.unread || 0), 0);
      map[org.id] = channelUnread + chatUnread;
    }
    return map;
  }, [teamsOrgs]);

  const activeAccount = useMemo(
    () => accounts.find((a) => a.id === activeAccountId) || null,
    [accounts, activeAccountId]
  );

  const unreadByAccount = useMemo(() => {
    const map = {};
    for (const a of accounts) {
      map[a.id] =
        BACKEND && liveUnread[a.id]
          ? liveUnread[a.id].inbox || 0
          : (a.messages || []).filter((m) => m.unread && m.folder === "inbox").length;
    }
    return map;
  }, [accounts, liveUnread]);

  const folderMessages = useMemo(() => {
    if (!activeAccount) return [];
    let list = (activeAccount.messages || []).filter((m) =>
      activeFolder === "starred" ? m.starred : m.folder === activeFolder
    );
    if (query.trim()) {
      const q = query.toLowerCase();
      list = list.filter(
        (m) =>
          m.subject.toLowerCase().includes(q) ||
          m.from.toLowerCase().includes(q) ||
          m.preview.toLowerCase().includes(q)
      );
    }
    return list;
  }, [activeAccount, activeFolder, query]);

  const selected = folderMessages.find((m) => m.id === selectedId) || null;

  const folderCounts = useMemo(() => {
    const c = {};
    if (!activeAccount) return c;
    for (const f of FOLDERS) {
      c[f.id] =
        f.id === "starred"
          ? (activeAccount.messages || []).filter((m) => m.starred).length
          : BACKEND && liveUnread[activeAccount.id]
          ? liveUnread[activeAccount.id][f.id] || 0
          : (activeAccount.messages || []).filter((m) => m.unread && m.folder === f.id).length;
    }
    return c;
  }, [activeAccount, liveUnread]);

  // Total unread per Slack workspace (sum across its channels).
  // Email accounts grouped by their user-set label; labeled groups sorted
  // alphabetically, unlabeled accounts together as one group at the end.
  const emailGroups = useMemo(() => {
    const byLabel = new Map();
    for (const a of accounts) {
      const key = (a.label || "").trim();
      if (!byLabel.has(key)) byLabel.set(key, []);
      byLabel.get(key).push(a);
    }
    const labeled = [...byLabel.entries()]
      .filter(([k]) => k)
      .sort((x, y) => x[0].localeCompare(y[0]))
      .map(([k, accs]) => ({ key: k, title: k, accounts: accs }));
    const unlabeled = byLabel.get("") || [];
    if (unlabeled.length) {
      labeled.push({ key: "__none", title: labeled.length ? "Other" : "Email", accounts: unlabeled });
    }
    return labeled;
  }, [accounts]);

  const unreadByWorkspace = useMemo(() => {
    const map = {};
    for (const w of workspaces) {
      // Prefer the backend poller's total when we have it; fall back to summing
      // whatever channel data is currently loaded.
      map[w.id] =
        BACKEND && liveSlackUnread[w.id]
          ? liveSlackUnread[w.id].total
          : w.channels.reduce((sum, c) => sum + (c.unread || 0), 0);
    }
    return map;
  }, [workspaces, liveSlackUnread]);

  // Active Slack channel + its thread, when a workspace is open.
  const activeChannel = useMemo(() => {
    if (!activeWorkspace) return null;
    return (
      activeWorkspace.channels.find((c) => c.id === activeChannelId) ||
      activeWorkspace.channels[0] ||
      null
    );
  }, [activeWorkspace, activeChannelId]);

  const openThread = useMemo(() => {
    if (!activeChannel || !openThreadId) return null;
    return (activeChannel.messages || []).find((m) => m.id === openThreadId) || null;
  }, [activeChannel, openThreadId]);

  function goHome() {
    setShowProfile(false);
    setActiveAccountId(null);
    setSelectedId(null);
    setQuery("");
    setOpenThreadId(null);
    setComposing(false);
  }

  async function loadMailFolder(accountId, folder) {
    if (!BACKEND) return;
    const acc = accounts.find((a) => a.id === accountId);
    if (!acc || acc._loadedFolders?.[folder]) return;
    try {
      const fetched = await liveMailService.listFolderMessages(accountId, folder);
      setAccounts((prev) =>
        prev.map((a) => {
          if (a.id !== accountId) return a;
          const known = new Set((a.messages || []).map((m) => m.id));
          return {
            ...a,
            messages: [...(a.messages || []), ...fetched.filter((m) => !known.has(m.id))],
            _loadedFolders: { ...(a._loadedFolders || {}), [folder]: true },
          };
        })
      );
    } catch (e) {
      console.warn("Failed to load folder:", e.message);
    }
  }

  // Force-fetch a folder and REPLACE its messages (used by SSE + refresh).
  async function refetchFolder(accountId, folder) {
    if (!BACKEND) return;
    try {
      const fetched = await liveMailService.listFolderMessages(accountId, folder);
      setAccounts((prev) =>
        prev.map((a) =>
          a.id === accountId
            ? {
                ...a,
                messages: [
                  ...(a.messages || []).filter((m) => (m.folder || "inbox") !== folder),
                  ...fetched,
                ],
                _loadedFolders: { ...(a._loadedFolders || {}), [folder]: true },
              }
            : a
        )
      );
    } catch (e) {
      console.warn("Refresh failed:", e.message);
    }
  }

  function refreshCurrentFolder() {
    if (!BACKEND || !activeAccount) return;
    refetchFolder(activeAccountId, activeFolder);
    liveMailService
      .refresh(activeAccountId)
      .then((r) => setLiveUnread((prev) => ({ ...prev, [activeAccountId]: r.folders || {} })))
      .catch(() => {});
  }

  function selectItem(id) {
    setShowProfile(false); // leaving profile/settings for an account view
    setActiveAccountId(id);
    setSelectedId(null);
    setQuery("");
    const ws = workspaces.find((w) => w.id === id);
    const org = teamsOrgs.find((t) => t.id === id);
    if (!ws && !org) loadMailFolder(id, "inbox");
    if (ws) {
      setActiveChannelId(ws.channels[0]?.id ?? null);
      setOpenThreadId(null);
    } else if (org) {
      const firstTeam = org.teams[0];
      setTeamsSel({
        teamId: firstTeam?.id ?? null,
        channelId: firstTeam?.channels[0]?.id ?? null,
        chatId: null,
      });
    } else {
      setActiveFolder("inbox");
    }
  }

  async function openChannel(channelId) {
    setActiveChannelId(channelId);
    setOpenThreadId(null);
    // Live mode: fetch this channel's history the first time it's opened.
    if (BACKEND) {
      const ws = workspaces.find((w) => w.id === activeAccountId);
      const ch = ws?.channels.find((c) => c.id === channelId);
      if (ch && ch.messages == null) {
        try {
          const messages = await liveSlackService.listMessages(activeAccountId, channelId);
          setWorkspaces((prev) =>
            prev.map((w) =>
              w.id === activeAccountId
                ? { ...w, channels: w.channels.map((c) => (c.id === channelId ? { ...c, messages } : c)) }
                : w
            )
          );
        } catch (e) {
          console.warn("Failed to load channel:", e.message);
        }
      }
    }
    // Mark channel read locally.
    setWorkspaces((prev) =>
      prev.map((w) =>
        w.id === activeAccountId
          ? { ...w, channels: w.channels.map((c) => (c.id === channelId ? { ...c, unread: 0 } : c)) }
          : w
      )
    );
  }

  async function openSlackThread(messageId) {
    // Live mode: pull the thread's replies before showing the pane.
    if (BACKEND && activeChannel) {
      try {
        const thread = await liveSlackService.listReplies(activeAccountId, activeChannel.id, messageId);
        const replies = thread.slice(1); // [0] is the parent
        setWorkspaces((prev) =>
          prev.map((w) =>
            w.id === activeAccountId
              ? {
                  ...w,
                  channels: w.channels.map((c) =>
                    c.id === activeChannel.id
                      ? { ...c, messages: (c.messages || []).map((m) => (m.id === messageId ? { ...m, replies } : m)) }
                      : c
                  ),
                }
              : w
          )
        );
      } catch (e) {
        console.warn("Failed to load thread:", e.message);
      }
    }
    setOpenThreadId(messageId);
  }

  function sendSlackMessage(text) {
    if (!text.trim() || !activeChannel) return;
    if (BACKEND) liveSlackService.postMessage(activeAccountId, activeChannel.id, text).catch((e) => console.warn(e.message));
    else slackService.postMessage(activeChannel.id, text);
    const newMsg = smsg(activeWorkspace.you, text, { time: "now" });
    setWorkspaces((prev) =>
      prev.map((w) =>
        w.id === activeAccountId
          ? {
              ...w,
              channels: w.channels.map((c) =>
                c.id === activeChannel.id ? { ...c, messages: [...c.messages, newMsg] } : c
              ),
            }
          : w
      )
    );
  }

  function sendSlackReply(parentId, text) {
    if (!text.trim() || !activeChannel) return;
    if (BACKEND) liveSlackService.postMessage(activeAccountId, activeChannel.id, text, parentId).catch((e) => console.warn(e.message));
    else slackService.postReply(activeChannel.id, parentId, text);
    const reply = smsg(activeWorkspace.you, text, { time: "now" });
    setWorkspaces((prev) =>
      prev.map((w) =>
        w.id === activeAccountId
          ? {
              ...w,
              channels: w.channels.map((c) =>
                c.id === activeChannel.id
                  ? {
                      ...c,
                      messages: c.messages.map((m) =>
                        m.id === parentId ? { ...m, replies: [...m.replies, reply] } : m
                      ),
                    }
                  : c
              ),
            }
          : w
      )
    );
  }

  function toggleReaction(messageId, emoji) {
    if (!activeChannel) return;
    if (BACKEND) {
      // Determine whether this toggle adds or removes before mutating state.
      const msgs = activeChannel.messages || [];
      const target =
        msgs.find((m) => m.id === messageId) ||
        msgs.flatMap((m) => m.replies || []).find((r) => r.id === messageId);
      const willAdd = !(target?.reactions || []).some((r) => r.emoji === emoji && r.mine);
      liveSlackService
        .toggleReaction(activeAccountId, activeChannel.id, messageId, emoji, willAdd)
        .catch((e) => console.warn(e.message));
    } else {
      slackService.toggleReaction(activeChannel.id, messageId, emoji);
    }

    // Apply the toggle to a single message's reactions array.
    const applyToReactions = (reactions = []) => {
      const existing = reactions.find((r) => r.emoji === emoji);
      if (existing) {
        if (existing.mine) {
          // Remove my reaction; drop the chip entirely if count hits zero.
          const count = existing.count - 1;
          if (count <= 0) return reactions.filter((r) => r.emoji !== emoji);
          return reactions.map((r) => (r.emoji === emoji ? { ...r, count, mine: false } : r));
        }
        // Add my reaction to an existing chip.
        return reactions.map((r) =>
          r.emoji === emoji ? { ...r, count: r.count + 1, mine: true } : r
        );
      }
      // Brand-new emoji on this message.
      return [...reactions, { emoji, count: 1, mine: true }];
    };

    // Walk a message and its (one level of) replies looking for the target id.
    const applyToMessage = (m) => {
      if (m.id === messageId) return { ...m, reactions: applyToReactions(m.reactions) };
      if (m.replies?.length) {
        const replies = m.replies.map((r) =>
          r.id === messageId ? { ...r, reactions: applyToReactions(r.reactions) } : r
        );
        return { ...m, replies };
      }
      return m;
    };

    setWorkspaces((prev) =>
      prev.map((w) =>
        w.id === activeAccountId
          ? {
              ...w,
              channels: w.channels.map((c) =>
                c.id === activeChannel.id
                  ? { ...c, messages: c.messages.map(applyToMessage) }
                  : c
              ),
            }
          : w
      )
    );
  }

  /* ---- Teams handlers ---- */

  // Immutably update one channel inside the active Teams org.
  function updateTeamsChannel(channelId, updater) {
    setTeamsOrgs((prev) =>
      prev.map((org) =>
        org.id === activeAccountId
          ? {
              ...org,
              teams: org.teams.map((t) => ({
                ...t,
                channels: t.channels.map((c) => (c.id === channelId ? updater(c) : c)),
              })),
            }
          : org
      )
    );
  }

  function openTeamsChannel(teamId, channelId) {
    setTeamsSel({ teamId, channelId, chatId: null });
    updateTeamsChannel(channelId, (c) => ({ ...c, unread: 0 }));
  }

  // Immutably update one chat inside the active Teams org.
  function updateTeamsChat(chatId, updater) {
    setTeamsOrgs((prev) =>
      prev.map((org) =>
        org.id === activeAccountId
          ? { ...org, chats: (org.chats || []).map((c) => (c.id === chatId ? updater(c) : c)) }
          : org
      )
    );
  }

  function openTeamsChat(chatId) {
    setTeamsSel({ teamId: null, channelId: null, chatId });
    updateTeamsChat(chatId, (c) => ({ ...c, unread: 0 }));
  }

  function sendTeamsChat(text) {
    if (!text.trim() || !activeTeamsChannel?.chat) return;
    teamsService.postMessage(activeTeamsChannel.chat.id, text);
    const msg = tpost(activeTeamsOrg.you, text, { time: "now" });
    updateTeamsChat(activeTeamsChannel.chat.id, (c) => ({
      ...c,
      messages: [...c.messages, msg],
    }));
  }

  function sendTeamsPost(text) {
    if (!text.trim() || !activeTeamsChannel?.channel) return;
    teamsService.postMessage(activeTeamsChannel.channel.id, text);
    const post = tpost(activeTeamsOrg.you, text, { time: "now" });
    updateTeamsChannel(activeTeamsChannel.channel.id, (c) => ({
      ...c,
      posts: [...c.posts, post],
    }));
  }

  function sendTeamsReply(postId, text) {
    if (!text.trim() || !activeTeamsChannel?.channel) return;
    teamsService.postReply(activeTeamsChannel.channel.id, postId, text);
    const reply = tpost(activeTeamsOrg.you, text, { time: "now" });
    updateTeamsChannel(activeTeamsChannel.channel.id, (c) => ({
      ...c,
      posts: c.posts.map((p) =>
        p.id === postId ? { ...p, replies: [...p.replies, reply] } : p
      ),
    }));
  }

  function toggleTeamsReaction(messageId, emoji) {
    if (!activeTeamsChannel) return;
    const target = activeTeamsChannel.chat || activeTeamsChannel.channel;
    teamsService.toggleReaction(target.id, messageId, emoji);

    const applyToReactions = (reactions = []) => {
      const existing = reactions.find((r) => r.emoji === emoji);
      if (existing) {
        if (existing.mine) {
          const count = existing.count - 1;
          if (count <= 0) return reactions.filter((r) => r.emoji !== emoji);
          return reactions.map((r) => (r.emoji === emoji ? { ...r, count, mine: false } : r));
        }
        return reactions.map((r) =>
          r.emoji === emoji ? { ...r, count: r.count + 1, mine: true } : r
        );
      }
      return [...reactions, { emoji, count: 1, mine: true }];
    };

    // Chat: flat message list.
    if (activeTeamsChannel.chat) {
      updateTeamsChat(activeTeamsChannel.chat.id, (c) => ({
        ...c,
        messages: c.messages.map((m) =>
          m.id === messageId ? { ...m, reactions: applyToReactions(m.reactions) } : m
        ),
      }));
      return;
    }

    // Channel: posts with one level of replies.
    const applyToPost = (p) => {
      if (p.id === messageId) return { ...p, reactions: applyToReactions(p.reactions) };
      if (p.replies?.length) {
        const replies = p.replies.map((r) =>
          r.id === messageId ? { ...r, reactions: applyToReactions(r.reactions) } : r
        );
        return { ...p, replies };
      }
      return p;
    };

    updateTeamsChannel(activeTeamsChannel.channel.id, (c) => ({
      ...c,
      posts: c.posts.map(applyToPost),
    }));
  }

  function openMessage(m) {
    setSelectedId(m.id);
    // Live mode: pull the full body if the list row only has a preview.
    if (BACKEND && m.body === undefined) {
      liveMailService
        .getMessage(activeAccountId, m.id, m.folder || activeFolder)
        .then((full) =>
          setAccounts((prev) =>
            prev.map((a) =>
              a.id === activeAccountId
                ? { ...a, messages: (a.messages || []).map((x) => (x.id === m.id ? { ...x, ...full, folder: x.folder } : x)) }
                : a
            )
          )
        )
        .catch((e) => console.warn("Failed to load message:", e.message));
    }
    if (m.unread) {
      if (BACKEND) liveMailService.markRead(activeAccountId, m.id, true, m.folder || activeFolder).catch(() => {});
      else mailService.markRead(m.id);
      setAccounts((prev) =>
        prev.map((a) =>
          a.id === activeAccountId
            ? { ...a, messages: a.messages.map((x) => (x.id === m.id ? { ...x, unread: false } : x)) }
            : a
        )
      );
    }
  }

  function toggleStar(e, m) {
    e.stopPropagation();
    if (BACKEND) liveMailService.toggleStar(activeAccountId, m.id, !m.starred, m.folder || activeFolder).catch(() => {});
    else mailService.toggleStar(m.id, !m.starred);
    setAccounts((prev) =>
      prev.map((a) =>
        a.id === activeAccountId
          ? { ...a, messages: a.messages.map((x) => (x.id === m.id ? { ...x, starred: !x.starred } : x)) }
          : a
      )
    );
  }

  async function connect(provider) {
    if (BACKEND) {
      // Live routing: OAuth redirects for Google/Microsoft/Slack; Apple uses
      // the in-modal app-password form; Teams isn't wired to a backend yet.
      if (provider === "google" || provider === "microsoft") {
        setConnecting(provider);
        liveMailService.connectRedirect(provider);
        return;
      }
      if (provider === "apple") {
        setAppleForm({ address: "", appPassword: "", error: null });
        return;
      }
      if (provider === "teams") return; // button disabled in live mode
    }
    setConnecting(provider);
    if (provider === "slack") {
      if (BACKEND) { liveSlackService.connect(); return; } // full-page OAuth redirect
      const ws = await slackService.connect();
      setWorkspaces((prev) => [...prev, ws]);
      setActiveAccountId(ws.id);
      setActiveChannelId(ws.channels[0]?.id ?? null);
      setOpenThreadId(null);
    } else if (provider === "teams") {
      const org = await teamsService.connect();
      setTeamsOrgs((prev) => [...prev, org]);
      setActiveAccountId(org.id);
      const firstTeam = org.teams[0];
      setTeamsSel({
        teamId: firstTeam?.id ?? null,
        channelId: firstTeam?.channels[0]?.id ?? null,
        chatId: null,
      });
    } else {
      const acc = await mailService.connect(provider);
      setAccounts((prev) => [...prev, acc]);
      setActiveAccountId(acc.id);
      setActiveFolder("inbox");
      // Fold the new account's calendar into the aggregated set.
      const evs = await calendarService.listEvents([...accounts.map((a) => a.id), acc.id]);
      setEvents(evs);
    }
    setSelectedId(null);
    setConnecting(null);
    setShowConnect(false);
  }

  function disconnectAccount(id) {
    // Backend delete only matters for live-stored accounts (Slack OAuth for
    // now); mock accounts just vanish locally. Errors are swallowed upstream.
    profileService.disconnectAccount(id);
    setAccounts((prev) => prev.filter((a) => a.id !== id));
    setWorkspaces((prev) => prev.filter((w) => w.id !== id));
    setTeamsOrgs((prev) => prev.filter((t) => t.id !== id));
    if (activeAccountId === id) goHome();
  }

  async function logout() {
    try { await profileService.logout(); } catch { /* session may be gone */ }
    window.location.reload();
  }

  async function submitAppleForm() {
    if (!appleForm?.address || !appleForm?.appPassword) return;
    setConnecting("apple");
    try {
      const acc = await liveMailService.connectApple(appleForm.address, appleForm.appPassword);
      setAccounts((prev) => [...prev.filter((a) => a.id !== acc.id), acc]);
      setActiveAccountId(acc.id);
      setActiveFolder("inbox");
      setAppleForm(null);
      setShowConnect(false);
      loadMailFolder(acc.id, "inbox");
    } catch (e) {
      setAppleForm((f) => ({ ...f, error: e.message }));
    } finally {
      setConnecting(null);
    }
  }

  // Move a message to archive/spam/trash/inbox: optimistic local re-folder,
  // then the provider call. Closes the reader if it was showing this message.
  function moveMessage(m, dest) {
    if (BACKEND) {
      liveMailService.move(activeAccountId, m.id, dest, m.folder || activeFolder).catch((e) => console.warn(e.message));
    }
    setAccounts((prev) =>
      prev.map((a) =>
        a.id === activeAccountId
          ? { ...a, messages: (a.messages || []).map((x) => (x.id === m.id ? { ...x, folder: dest } : x)) }
          : a
      )
    );
    if (selectedId === m.id) setSelectedId(null);
  }

  function markMessageRead(m, read) {
    if (BACKEND) {
      liveMailService.markRead(activeAccountId, m.id, read, m.folder || activeFolder).catch(() => {});
    } else {
      mailService.markRead(m.id);
    }
    setAccounts((prev) =>
      prev.map((a) =>
        a.id === activeAccountId
          ? { ...a, messages: (a.messages || []).map((x) => (x.id === m.id ? { ...x, unread: !read } : x)) }
          : a
      )
    );
    if (!read && selectedId === m.id) setSelectedId(null);
  }

  async function sendEmailReply(orig, { to, subject, html, text }) {
    if (BACKEND) {
      await liveMailService.send(activeAccountId, { to, subject, body: text, html });
    } else {
      // Mock mode: drop a copy in the Sent folder so the flow feels real.
      const sentMsg = msg("Me", activeAccount.address, subject, text.slice(0, 80), text, {
        folder: "sent", date: "now", fullDate: "Just now",
      });
      setAccounts((prev) =>
        prev.map((a) => (a.id === activeAccountId ? { ...a, messages: [...a.messages, sentMsg] } : a))
      );
    }
  }

  function saveLabel(id, value) {
    setAccounts((prev) => prev.map((a) => (a.id === id ? { ...a, label: value } : a)));
    setWorkspaces((prev) => prev.map((w) => (w.id === id ? { ...w, label: value } : w)));
    setTeamsOrgs((prev) => prev.map((t) => (t.id === id ? { ...t, label: value } : t)));
    setEditingLabel(null);
    // Persist for backend-stored accounts (mail + Slack). Teams is mock-only,
    // and unknown ids 404 harmlessly.
    if (BACKEND) {
      backendApi(`/api/accounts/${id}`, { method: "PATCH", body: JSON.stringify({ label: value }) })
        .catch(() => {});
    }
  }

  return (
    <div className={`md-root ${railExpanded ? "rail-open" : ""}`}>
      <style>{CSS}</style>
      {BACKEND && !profileLoaded && (
        <div className="auth-overlay"><div className="auth-splash">Loading Beacon…</div></div>
      )}
      {BACKEND && profileLoaded && !profile && (
        <AuthScreen onDone={setProfile} />
      )}

      {/* ---------- account rail ---------- */}
      <aside className={`rail ${railExpanded ? "is-expanded" : ""}`}>
        <div className="rail-top">
          <div className="rail-brand">
            <Mail size={20} strokeWidth={2.4} />
          </div>
          {railExpanded && <span className="rail-wordmark">Beacon</span>}
        </div>

        <button
          className="rail-toggle"
          onClick={() => setRailExpanded((v) => !v)}
          title={railExpanded ? "Collapse" : "Expand"}
          aria-label={railExpanded ? "Collapse sidebar" : "Expand sidebar"}
        >
          {railExpanded ? <ChevronsLeft size={16} /> : <ChevronsRight size={16} />}
          {railExpanded && <span>Collapse</span>}
        </button>

        <div className="rail-accounts">
          {emailGroups.map((g, gi) => (
            <React.Fragment key={g.key}>
              {!railExpanded && gi > 0 && <div className="rail-group-gap" />}
              {railExpanded && <span className="rail-section-label">{g.title}</span>}
              {g.accounts.map((a) => {
            const p = PROVIDERS[a.provider];
            const unread = unreadByAccount[a.id] || 0;
            const active = a.id === activeAccountId;
            return (
              <button
                key={a.id}
                className={`rail-acc ${active ? "is-active" : ""}`}
                onClick={() => selectItem(a.id)}
                title={railExpanded ? undefined : `${a.label || p.name} — ${a.address}`}
              >
                <span className="rail-avatar" style={{ background: avatarColor(a.address) }}>
                  {initials(a.address)}
                  <span className="rail-provider" style={{ background: p.gradient }}>
                    {p.badge || <AppleGlyph />}
                  </span>
                </span>

                {railExpanded && (
                  <span className="rail-acc-text">
                    <span className="rail-acc-label">{a.label || p.name}</span>
                    <span className="rail-acc-addr">{a.address}</span>
                  </span>
                )}

                {unread > 0 && (
                  <span className={`rail-dot ${railExpanded ? "inline" : ""}`}>
                    {unread > 9 ? "9+" : unread}
                  </span>
                )}
              </button>
            );
              })}
            </React.Fragment>
          ))}

          {/* separator between email and slack */}
          <div className="rail-sep" />
          {railExpanded && <span className="rail-section-label">Slack</span>}

          {workspaces.map((w) => {
            const unread = unreadByWorkspace[w.id] || 0;
            const active = w.id === activeAccountId;
            return (
              <button
                key={w.id}
                className={`rail-acc rail-ws ${active ? "is-active" : ""}`}
                onClick={() => selectItem(w.id)}
                title={railExpanded ? undefined : `${w.label} (Slack)`}
              >
                <span className="rail-ws-logo" style={{ background: w.accent }}>
                  {w.workspace.slice(0, 1)}
                  <span className="rail-provider" style={{ background: PROVIDERS.slack.gradient }}>
                    S
                  </span>
                </span>

                {railExpanded && (
                  <span className="rail-acc-text">
                    <span className="rail-acc-label">{w.label}</span>
                    <span className="rail-acc-addr">{w.channels.length} channels</span>
                  </span>
                )}

                {unread > 0 && (
                  <span className={`rail-dot ${railExpanded ? "inline" : ""}`}>
                    {unread > 9 ? "9+" : unread}
                  </span>
                )}
              </button>
            );
          })}

          {/* separator between slack and teams */}
          <div className="rail-sep" />
          {railExpanded && <span className="rail-section-label">Teams</span>}

          {teamsOrgs.map((org) => {
            const unread = unreadByTeamsOrg[org.id] || 0;
            const active = org.id === activeAccountId;
            const teamCount = org.teams.length;
            return (
              <button
                key={org.id}
                className={`rail-acc rail-ws ${active ? "is-active" : ""}`}
                onClick={() => selectItem(org.id)}
                title={railExpanded ? undefined : `${org.label} (Teams)`}
              >
                <span className="rail-ws-logo rail-tm-logo" style={{ background: PROVIDERS.teams.gradient }}>
                  {org.org.slice(0, 1)}
                  <span className="rail-provider" style={{ background: PROVIDERS.teams.gradient }}>
                    T
                  </span>
                </span>

                {railExpanded && (
                  <span className="rail-acc-text">
                    <span className="rail-acc-label">{org.label}</span>
                    <span className="rail-acc-addr">
                      {teamCount} {teamCount === 1 ? "team" : "teams"}
                      {(org.chats?.length || 0) > 0 && ` · ${org.chats.length} ${org.chats.length === 1 ? "chat" : "chats"}`}
                    </span>
                  </span>
                )}

                {unread > 0 && (
                  <span className={`rail-dot ${railExpanded ? "inline" : ""}`}>
                    {unread > 9 ? "9+" : unread}
                  </span>
                )}
              </button>
            );
          })}

          <button
            className={`rail-add ${railExpanded ? "wide" : ""}`}
            onClick={() => setShowConnect(true)}
            title="Connect account"
          >
            <Plus size={20} />
            {railExpanded && <span>Connect account</span>}
          </button>
        </div>

        <button
          className={`rail-settings ${railExpanded ? "wide" : ""}`}
          title="Profile & settings"
          onClick={() => {
            setActiveAccountId(null);
            setSelectedId(null);
            setShowProfile(true);
          }}
        >
          <Settings size={18} />
          {railExpanded && <span>Profile</span>}
        </button>
      </aside>

      {/* ---------- folder / channel column ---------- */}
      <nav className="sidebar">
        {isHome ? (
          <div className="home-side">
            <div className="home-side-head">
              <span className="home-side-logo"><Home size={18} /></span>
              <div>
                <strong>Home</strong>
                <span className="home-side-sub">Your dashboard</span>
              </div>
            </div>
            <p className="home-side-hint">
              Pick an account or workspace from the left to dive in, or stay here for
              your daily overview.
            </p>
            <div className="sidebar-foot">
              <span className="prov-tag" style={{ background: "linear-gradient(135deg,#4f6ef7,#3b5bdb)" }}>
                <Home size={12} />
              </span>
              Dashboard
            </div>
          </div>
        ) : isSlack ? (
          <SlackChannelList
            workspace={activeWorkspace}
            activeChannel={activeChannel}
            onOpenChannel={openChannel}
            editingLabel={editingLabel}
            onEditLabel={setEditingLabel}
            onSaveLabel={saveLabel}
            onGoHome={goHome}
          />
        ) : isTeams ? (
          <TeamsSidebar
            org={activeTeamsOrg}
            selection={activeTeamsChannel}
            onOpenChannel={openTeamsChannel}
            onOpenChat={openTeamsChat}
            editingLabel={editingLabel}
            onEditLabel={setEditingLabel}
            onSaveLabel={saveLabel}
            onGoHome={goHome}
          />
        ) : (
          activeAccount && (
          <>
            <div className="acc-head">
              <span className="acc-avatar" style={{ background: avatarColor(activeAccount.address) }}>
                {initials(activeAccount.address)}
              </span>
              <div className="acc-meta">
                {editingLabel === activeAccount.id ? (
                  <LabelEditor
                    initial={activeAccount.label}
                    onSave={(v) => saveLabel(activeAccount.id, v)}
                    onCancel={() => setEditingLabel(null)}
                  />
                ) : (
                  <button className="acc-label" onClick={() => setEditingLabel(activeAccount.id)}>
                    <span>{activeAccount.label || "Add a label"}</span>
                    <Pencil size={12} />
                  </button>
                )}
                <span className="acc-addr">{activeAccount.address}</span>
              </div>
            </div>

            <div className="compose-row">
              <button className="compose" onClick={() => setComposing(true)}>
                <Pencil size={15} /> Compose
              </button>
              <button className="home-btn" onClick={goHome} title="Home" aria-label="Home">
                <Home size={17} />
              </button>
            </div>

            <ul className="folders">
              {FOLDERS.map((f) => {
                const Icon = f.icon;
                const count = folderCounts[f.id] || 0;
                return (
                  <li key={f.id}>
                    <button
                      className={`folder ${activeFolder === f.id ? "is-active" : ""}`}
                      onClick={() => {
                        setActiveFolder(f.id);
                        setSelectedId(null);
                        loadMailFolder(activeAccount.id, f.id);
                      }}
                    >
                      <Icon size={17} strokeWidth={2} />
                      <span className="folder-name">{f.name}</span>
                      {count > 0 && (
                        <span className={`folder-count ${f.id === "starred" ? "muted" : ""}`}>
                          {count}
                        </span>
                      )}
                    </button>
                  </li>
                );
              })}
            </ul>

            <div className="sidebar-foot">
              <span className="prov-tag" style={{ background: PROVIDERS[activeAccount.provider].gradient }}>
                {PROVIDERS[activeAccount.provider].badge || <AppleGlyph />}
              </span>
              {PROVIDERS[activeAccount.provider].name}
            </div>
          </>
          )
        )}
      </nav>

      {/* ---------- main panes ---------- */}
      {showProfile ? (
        <ProfileScreen
          profile={profile}
          onSaveProfile={async (fields) => setProfile(await profileService.update(fields))}
          onChangePassword={(cur, next) => profileService.changePassword(cur, next)}
          accounts={accounts}
          workspaces={workspaces}
          teamsOrgs={teamsOrgs}
          onDisconnect={disconnectAccount}
          onSetLabel={saveLabel}
          onLogout={logout}
          onBack={goHome}
          live={!!BACKEND}
        />
      ) : isHome ? (
        <HomeScreen
          accounts={accounts}
          workspaces={workspaces}
          teamsOrgs={teamsOrgs}
          events={events}
          unreadByAccount={unreadByAccount}
          unreadByWorkspace={unreadByWorkspace}
          unreadByTeamsOrg={unreadByTeamsOrg}
          onSelect={selectItem}
          onConnect={() => setShowConnect(true)}
          profile={profile}
        />
      ) : isSlack ? (
        <SlackView
          workspace={activeWorkspace}
          channel={activeChannel}
          openThread={openThread}
          onOpenThread={openSlackThread}
          onCloseThread={() => setOpenThreadId(null)}
          onSend={sendSlackMessage}
          onReply={sendSlackReply}
          onReact={toggleReaction}
          loading={loading}
        />
      ) : isTeams ? (
        <TeamsView
          org={activeTeamsOrg}
          selection={activeTeamsChannel}
          onSend={sendTeamsPost}
          onSendChat={sendTeamsChat}
          onReply={sendTeamsReply}
          onReact={toggleTeamsReaction}
          loading={loading}
        />
      ) : (
        <>
      {/* ---------- message list ---------- */}
      <section className="list">
        <header className="list-head">
          <div className="list-title">
            <h1>{FOLDERS.find((f) => f.id === activeFolder)?.name}</h1>
            <button className="refresh" title="Refresh" onClick={refreshCurrentFolder}>
              <RefreshCw size={15} />
            </button>
          </div>
          <div className="search">
            <Search size={15} />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search mail"
            />
            {query && (
              <button onClick={() => setQuery("")} className="search-clear">
                <X size={14} />
              </button>
            )}
          </div>
        </header>

        <div className="list-scroll">
          {loading ? (
            <div className="empty">Loading accounts…</div>
          ) : folderMessages.length === 0 ? (
            <div className="empty">
              <Inbox size={40} strokeWidth={1.3} />
              <p>{query ? "No messages match your search." : "Nothing here yet."}</p>
            </div>
          ) : (
            folderMessages.map((m) => (
              <button
                key={m.id}
                className={`row ${m.unread ? "is-unread" : ""} ${selectedId === m.id ? "is-selected" : ""}`}
                onClick={() => openMessage(m)}
              >
                <span className="row-avatar" style={{ background: avatarColor(m.fromAddr) }}>
                  {initials(m.fromAddr)}
                </span>
                <div className="row-body">
                  <div className="row-top">
                    <span className="row-from">{m.from}</span>
                    <span className="row-date">{m.date}</span>
                  </div>
                  <div className="row-subject">{m.subject}</div>
                  <div className="row-preview">{m.preview}</div>
                </div>
                <div className="row-marks">
                  <span
                    className={`row-star ${m.starred ? "on" : ""}`}
                    onClick={(e) => toggleStar(e, m)}
                    role="button"
                    aria-label="Star"
                  >
                    <Star size={15} fill={m.starred ? "currentColor" : "none"} />
                  </span>
                  {m.attachments.length > 0 && <Paperclip size={13} className="row-clip" />}
                  {m.unread && <CircleDot size={9} className="row-unread-dot" />}
                </div>

                {/* hover quick actions — spans with button role (the row itself is a button) */}
                <div className="row-actions" onClick={(e) => e.stopPropagation()}>
                  <span
                    role="button" tabIndex={0} className="row-act"
                    title={m.unread ? "Mark as read" : "Mark as unread"}
                    aria-label={m.unread ? "Mark as read" : "Mark as unread"}
                    onClick={() => markMessageRead(m, m.unread)}
                  >
                    {m.unread ? <MailOpen size={15} /> : <Mail size={15} />}
                  </span>
                  <span
                    role="button" tabIndex={0} className="row-act" title="Archive" aria-label="Archive"
                    onClick={() => moveMessage(m, "archive")}
                  >
                    <Archive size={15} />
                  </span>
                  <span
                    role="button" tabIndex={0} className="row-act" title="Mark as spam" aria-label="Mark as spam"
                    onClick={() => moveMessage(m, "spam")}
                  >
                    <AlertOctagonIcon size={15} />
                  </span>
                  <span
                    role="button" tabIndex={0} className="row-act danger" title="Delete" aria-label="Delete"
                    onClick={() => moveMessage(m, "trash")}
                  >
                    <Trash2 size={15} />
                  </span>
                </div>
              </button>
            ))
          )}
        </div>
      </section>

      {/* ---------- reading pane ---------- */}
      <section className="reader">
        {selected ? (
          <Reader
            m={selected}
            account={activeAccount}
            onStar={(e) => toggleStar(e, selected)}
            onSendReply={sendEmailReply}
            onMove={(dest) => moveMessage(selected, dest)}
            onMarkUnread={() => markMessageRead(selected, false)}
          />
        ) : (
          <div className="reader-empty">
            <Mail size={48} strokeWidth={1.1} />
            <p>Select a message to read</p>
          </div>
        )}
      </section>
        </>
      )}

      {/* ---------- connect modal ---------- */}
      {showConnect && (
        <Modal onClose={() => { if (!connecting) { setShowConnect(false); setAppleForm(null); } }}>
          <h2 className="modal-title">Connect an account</h2>
          <p className="modal-sub">Sign in with your provider to add a mailbox or Slack workspace.</p>
          {appleForm ? (
            <div className="apple-form">
              <p className="apple-form-hint">
                iCloud uses an app-specific password. Create one at account.apple.com →
                Sign-In and Security, then enter it below.
              </p>
              <input
                type="email"
                placeholder="iCloud email"
                value={appleForm.address}
                onChange={(e) => setAppleForm((f) => ({ ...f, address: e.target.value }))}
                autoFocus
              />
              <input
                type="password"
                placeholder="App-specific password"
                value={appleForm.appPassword}
                onChange={(e) => setAppleForm((f) => ({ ...f, appPassword: e.target.value }))}
                onKeyDown={(e) => e.key === "Enter" && submitAppleForm()}
              />
              {appleForm.error && <div className="auth-error">{appleForm.error}</div>}
              <div className="apple-form-actions">
                <button className="auth-submit sm" onClick={submitAppleForm} disabled={connecting === "apple"}>
                  {connecting === "apple" ? "Connecting…" : "Connect iCloud"}
                </button>
                <button className="auth-switch" onClick={() => setAppleForm(null)}>Back</button>
              </div>
            </div>
          ) : (
          <div className="provider-grid">
            {Object.entries(PROVIDERS).map(([key, p]) => {
              const teamsDisabled = !!BACKEND && key === "teams";
              return (
              <button
                key={key}
                className="provider-btn"
                disabled={!!connecting || teamsDisabled}
                onClick={() => connect(key)}
              >
                <span className="provider-mark" style={{ background: p.gradient }}>
                  {p.badge || <AppleGlyph size={18} />}
                </span>
                <span className="provider-name">
                  {connecting === key
                    ? "Connecting…"
                    : teamsDisabled
                    ? `${p.name} — coming soon`
                    : `Continue with ${p.name}`}
                </span>
                {connecting === key && <RefreshCw size={16} className="spin" />}
              </button>
              );
            })}
          </div>
          )}
          <p className="modal-note">
            Real sign-in runs an OAuth flow on your backend (Gmail API, Microsoft Graph,
            iCloud IMAP, Slack Web API, Teams via Graph). This demo simulates the result.
          </p>
        </Modal>
      )}

      {/* ---------- compose ---------- */}
      {composing && activeAccount && (
        <Compose account={activeAccount} onClose={() => setComposing(false)} />
      )}
    </div>
  );
}

/* ------------------------------ subcomponents ------------------------------ */

/* ---- Home: weather icon mapper ---- */
function WeatherGlyph({ icon, size = 22 }) {
  const map = {
    sun: Sun,
    "cloud-sun": CloudSun,
    cloud: Cloud,
    "cloud-rain": CloudRain,
  };
  const Ico = map[icon] || CloudSun;
  return <Ico size={size} />;
}

/* ---- Home: live date + time widget ---- */
function DateTimeWidget() {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const time = now.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  const seconds = now.toLocaleTimeString(undefined, { second: "2-digit" }).padStart(2, "0");
  const weekday = now.toLocaleDateString(undefined, { weekday: "long" });
  const date = now.toLocaleDateString(undefined, { month: "long", day: "numeric" });
  const year = now.getFullYear();

  return (
    <div className="widget widget-clock">
      <div className="clock-time">
        {time.replace(/\s?[AP]M/, "")}
        <span className="clock-sec">:{seconds}</span>
        <span className="clock-ampm">{/[AP]M/.exec(time)?.[0]}</span>
      </div>
      <div className="clock-date">
        <CalendarIcon size={15} />
        <span><strong>{weekday}</strong>, {date}, {year}</span>
      </div>
    </div>
  );
}

/* ---- Home: weather widget ---- */
function WeatherWidget() {
  return (
    <div className="widget widget-weather">
      <div className="weather-main">
        <div className="weather-now">
          <WeatherGlyph icon={WEATHER.icon} size={40} />
          <div>
            <div className="weather-temp">{WEATHER.temp}°</div>
            <div className="weather-cond">{WEATHER.condition}</div>
          </div>
        </div>
        <div className="weather-loc">
          <MapPin size={12} /> {WEATHER.location}
          <span className="weather-hilo">H:{WEATHER.high}° L:{WEATHER.low}°</span>
        </div>
      </div>
      <div className="weather-forecast">
        {WEATHER.forecast.map((d) => (
          <div className="weather-day" key={d.day}>
            <span className="weather-day-name">{d.day}</span>
            <WeatherGlyph icon={d.icon} size={18} />
            <span className="weather-day-high">{d.high}°</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---- Home: calendar widget (mini month + aggregated agenda) ---- */
// Each provider's calendar gets a distinct color so merged events stay readable.
const PROVIDER_CAL_COLOR = {
  google: "#EA4335",
  microsoft: "#0078D4",
  apple: "#1d1d1f",
};

function CalendarWidget({ events = [], accounts = [] }) {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const today = now.getDate();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthName = now.toLocaleDateString(undefined, { month: "long" });

  const cells = [];
  for (let i = 0; i < firstDay; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  // Resolve each event's source account → color + label.
  const acctById = Object.fromEntries(accounts.map((a) => [a.id, a]));
  const decorate = (e) => {
    const acct = acctById[e.accountId];
    const provider = acct?.provider;
    return {
      ...e,
      color: PROVIDER_CAL_COLOR[provider] || "#6b7688",
      source: acct ? acct.label || PROVIDERS[provider]?.name || acct.address : "Unknown",
    };
  };

  const todayEvents = events.filter((e) => e.today).map(decorate);
  const upcoming = events.filter((e) => !e.today).map(decorate);

  // Only show calendars (legend) for accounts that contributed events.
  const contributing = accounts.filter((a) => events.some((e) => e.accountId === a.id));

  const renderEvent = (e) => (
    <div className="cal-event" key={e.id}>
      <span className="cal-event-dot" style={{ background: e.color }} />
      <span className="cal-event-title">{e.title}</span>
      <span className="cal-event-source" style={{ color: e.color }}>{e.source}</span>
      <span className="cal-event-time">{e.time}</span>
    </div>
  );

  return (
    <div className="widget widget-calendar">
      <div className="widget-head">
        <h3><CalendarIcon size={15} /> Calendar</h3>
        <span className="widget-sub">{monthName} {year}</span>
      </div>

      {accounts.length === 0 && (
        <span className="cal-empty">Connect an email account to see its calendar here.</span>
      )}
      {contributing.length > 0 && (
        <div className="cal-legend">
          {contributing.map((a) => (
            <span className="cal-legend-item" key={a.id}>
              <span className="cal-legend-dot" style={{ background: PROVIDER_CAL_COLOR[a.provider] }} />
              {a.label || PROVIDERS[a.provider]?.name}
            </span>
          ))}
        </div>
      )}

      <div className="cal-grid">
        {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => (
          <span className="cal-dow" key={i}>{d}</span>
        ))}
        {cells.map((d, i) => (
          <span
            key={i}
            className={`cal-day ${d === today ? "is-today" : ""} ${d ? "" : "empty"}`}
          >
            {d || ""}
          </span>
        ))}
      </div>

      <div className="cal-agenda">
        <span className="cal-agenda-label">Today</span>
        {todayEvents.length > 0 ? todayEvents.map(renderEvent) : (
          <span className="cal-empty">No events today</span>
        )}
        <span className="cal-agenda-label">Upcoming</span>
        {upcoming.length > 0 ? upcoming.map(renderEvent) : (
          <span className="cal-empty">Nothing upcoming</span>
        )}
      </div>
    </div>
  );
}

/* ---- Home: Slack lists widget ---- */
const LIST_STATUS = {
  done: { label: "Done", icon: CheckSquare, cls: "done" },
  in_progress: { label: "In progress", icon: Square, cls: "progress" },
  todo: { label: "To do", icon: Square, cls: "todo" },
};

function SlackListsWidget({ lists = [] }) {
  return (
    <div className="widget widget-lists">
      <div className="widget-head">
        <h3><ListTodo size={15} /> Slack Lists</h3>
        <span className="widget-sub">{lists.length === 1 ? "1 list" : `${lists.length} lists`}</span>
      </div>
      {lists.length === 0 && (
        <span className="cal-empty">Connect a Slack workspace to see your lists here.</span>
      )}
      <div className="lists-scroll">
        {lists.map((list) => {
          const done = list.items.filter((i) => i.status === "done").length;
          return (
            <div className="list-card" key={list.id}>
              <div className="list-card-head">
                <span className="list-accent" style={{ background: list.accent }} />
                <span className="list-name">{list.name}</span>
                <span className="list-ws">{list.workspace}</span>
                <span className="list-progress">{done}/{list.items.length}</span>
              </div>
              <div className="list-items">
                {list.items.map((item) => {
                  const st = LIST_STATUS[item.status];
                  const Ico = st.icon;
                  return (
                    <div className={`list-item ${st.cls}`} key={item.id}>
                      <Ico size={15} className="list-check" />
                      <span className="list-item-title">{item.title}</span>
                      <span className="list-item-meta">
                        <span className="list-assignee">{item.assignee.split(" ")[0]}</span>
                        <span className="list-due">{item.due}</span>
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---- Home screen ---- */
function HomeScreen({ accounts, workspaces, teamsOrgs = [], events, unreadByAccount, unreadByWorkspace, unreadByTeamsOrg = {}, onSelect, onConnect, profile }) {
  const now = new Date();
  const hour = now.getHours();
  const greeting = hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening";

  const totalEmailUnread = accounts.reduce((s, a) => s + (unreadByAccount[a.id] || 0), 0);
  const totalSlackUnread = workspaces.reduce((s, w) => s + (unreadByWorkspace[w.id] || 0), 0);
  const totalTeamsUnread = teamsOrgs.reduce((s, t) => s + (unreadByTeamsOrg[t.id] || 0), 0);

  return (
    <section className="home">
      <div className="home-scroll">
        <header className="home-header">
          <h1>{greeting}{profile?.firstName ? `, ${profile.firstName}` : ""}</h1>
          <p className="home-summary">
            You have <strong>{totalEmailUnread}</strong> unread {totalEmailUnread === 1 ? "email" : "emails"},{" "}
            <strong>{totalSlackUnread}</strong> unread Slack {totalSlackUnread === 1 ? "message" : "messages"}, and{" "}
            <strong>{totalTeamsUnread}</strong> unread Teams {totalTeamsUnread === 1 ? "message" : "messages"}.
          </p>
        </header>

        {/* top row: clock + weather */}
        <div className="home-row home-row-top">
          <DateTimeWidget />
          <WeatherWidget />
        </div>

        {/* quick access to accounts */}
        <div className="widget widget-accounts">
          <div className="widget-head">
            <h3><Mail size={15} /> Quick access</h3>
          </div>
          {accounts.length + workspaces.length + teamsOrgs.length === 0 ? (
            <div className="qa-cta">
              <span className="qa-cta-title">Connect your first account</span>
              <span className="qa-cta-sub">Email, Slack, and Teams accounts appear here once connected.</span>
              <button className="qa-cta-btn" onClick={onConnect}>
                <Plus size={15} /> Connect account
              </button>
            </div>
          ) : (
          <div className="qa-grid">
            {accounts.map((a) => {
              const p = PROVIDERS[a.provider];
              const unread = unreadByAccount[a.id] || 0;
              return (
                <button className="qa-item" key={a.id} onClick={() => onSelect(a.id)}>
                  <span className="qa-avatar" style={{ background: avatarColor(a.address) }}>
                    {initials(a.address)}
                    <span className="qa-prov" style={{ background: p.gradient }}>
                      {p.badge || <AppleGlyph size={9} />}
                    </span>
                  </span>
                  <span className="qa-text">
                    <span className="qa-label">{a.label || p.name}</span>
                    <span className="qa-sub">{a.address}</span>
                  </span>
                  {unread > 0 && <span className="qa-badge">{unread}</span>}
                </button>
              );
            })}
            {workspaces.map((w) => {
              const unread = unreadByWorkspace[w.id] || 0;
              return (
                <button className="qa-item" key={w.id} onClick={() => onSelect(w.id)}>
                  <span className="qa-avatar qa-ws" style={{ background: w.accent }}>
                    {w.workspace.slice(0, 1)}
                    <span className="qa-prov" style={{ background: PROVIDERS.slack.gradient }}>S</span>
                  </span>
                  <span className="qa-text">
                    <span className="qa-label">{w.label}</span>
                    <span className="qa-sub">{w.channels.length} channels · Slack</span>
                  </span>
                  {unread > 0 && <span className="qa-badge">{unread}</span>}
                </button>
              );
            })}
            {teamsOrgs.map((org) => {
              const unread = unreadByTeamsOrg[org.id] || 0;
              const teamCount = org.teams.length;
              return (
                <button className="qa-item" key={org.id} onClick={() => onSelect(org.id)}>
                  <span className="qa-avatar qa-ws" style={{ background: PROVIDERS.teams.gradient }}>
                    {org.org.slice(0, 1)}
                    <span className="qa-prov" style={{ background: PROVIDERS.teams.gradient }}>T</span>
                  </span>
                  <span className="qa-text">
                    <span className="qa-label">{org.label}</span>
                    <span className="qa-sub">{teamCount} {teamCount === 1 ? "team" : "teams"} · Teams</span>
                  </span>
                  {unread > 0 && <span className="qa-badge">{unread}</span>}
                </button>
              );
            })}
            <button className="qa-add" onClick={onConnect} title="Connect another account">
              <Plus size={16} /> <span>Connect account</span>
            </button>
          </div>
          )}
        </div>

        {/* main row: calendar + slack lists */}
        <div className="home-row home-row-main">
          <CalendarWidget events={events} accounts={accounts} />
          <SlackListsWidget lists={BACKEND ? [] : SLACK_LISTS} />
        </div>
      </div>
    </section>
  );
}

/* ---- Auth: signup / login (live mode only) ---- */
function AuthScreen({ onDone }) {
  const [mode, setMode] = useState("signup");
  const [form, setForm] = useState({ firstName: "", lastName: "", email: "", password: "" });
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async () => {
    setError(null);
    setBusy(true);
    try {
      const user =
        mode === "signup"
          ? await profileService.signup(form)
          : await profileService.login(form.email, form.password);
      onDone(user);
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="auth-overlay">
      <div className="auth-card">
        <div className="auth-brand"><Mail size={22} strokeWidth={2.4} /></div>
        <h1>{mode === "signup" ? "Create your Beacon account" : "Welcome back"}</h1>
        <p className="auth-sub">
          {mode === "signup"
            ? "One profile for all your email, Slack, and Teams accounts."
            : "Log in to get back to your connected accounts."}
        </p>

        {mode === "signup" && (
          <div className="auth-row">
            <input placeholder="First name" value={form.firstName} onChange={set("firstName")} autoFocus />
            <input placeholder="Last name" value={form.lastName} onChange={set("lastName")} />
          </div>
        )}
        <input type="email" placeholder="Email" value={form.email} onChange={set("email")} autoFocus={mode === "login"} />
        <input
          type="password"
          placeholder={mode === "signup" ? "Password (8+ characters)" : "Password"}
          value={form.password}
          onChange={set("password")}
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />

        {error && <div className="auth-error">{error}</div>}

        <button className="auth-submit" onClick={submit} disabled={busy}>
          {busy ? "One moment…" : mode === "signup" ? "Create account" : "Log in"}
        </button>

        <button className="auth-switch" onClick={() => { setMode(mode === "signup" ? "login" : "signup"); setError(null); }}>
          {mode === "signup" ? "Already have an account? Log in" : "New here? Create an account"}
        </button>
      </div>
    </div>
  );
}

/* ---- Profile: personal info + connected account management ---- */
function ProfileScreen({ profile, onSaveProfile, onChangePassword, accounts, workspaces, teamsOrgs, onDisconnect, onSetLabel, onLogout, onBack, live }) {
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({
    firstName: profile?.firstName || "",
    lastName: profile?.lastName || "",
    email: profile?.email || "",
  });
  const [saved, setSaved] = useState(false);
  const [err, setErr] = useState(null);
  const [pw, setPw] = useState({ current: "", next: "" });
  const [pwMsg, setPwMsg] = useState(null);

  const set = (k) => (e) => { setForm((f) => ({ ...f, [k]: e.target.value })); setSaved(false); };

  const save = async () => {
    setErr(null);
    try {
      await onSaveProfile(form);
      setSaved(true);
    } catch (e) {
      setErr(e.message);
    }
  };

  const changePw = async () => {
    setPwMsg(null);
    try {
      await onChangePassword(pw.current, pw.next);
      setPw({ current: "", next: "" });
      setPwMsg({ ok: true, text: "Password updated." });
    } catch (e) {
      setPwMsg({ ok: false, text: e.message });
    }
  };

  const connected = [
    ...accounts.map((a) => ({ ...a, group: "Email", prov: PROVIDERS[a.provider] })),
    ...workspaces.map((w) => ({ ...w, group: "Slack", prov: PROVIDERS.slack })),
    ...teamsOrgs.map((t) => ({ ...t, group: "Teams", prov: PROVIDERS.teams })),
  ];

  return (
    <section className="home prof">
      <div className="home-scroll">
        <header className="home-header prof-header">
          <h1>Profile & settings</h1>
          <button className="prof-back" onClick={onBack}><Home size={15} /> Back to home</button>
        </header>

        <div className="widget prof-card">
          <div className="widget-head"><h3><Settings size={15} /> Personal information</h3></div>
          <div className="prof-grid">
            <label>First name<input value={form.firstName} onChange={set("firstName")} /></label>
            <label>Last name<input value={form.lastName} onChange={set("lastName")} /></label>
            <label className="prof-wide">Email<input type="email" value={form.email} onChange={set("email")} /></label>
          </div>
          {err && <div className="auth-error">{err}</div>}
          <div className="prof-actions">
            <button className="auth-submit sm" onClick={save}>Save changes</button>
            {saved && <span className="prof-saved">Saved ✓</span>}
          </div>
        </div>

        {live && (
          <div className="widget prof-card">
            <div className="widget-head"><h3><Settings size={15} /> Change password</h3></div>
            <div className="prof-grid">
              <label>Current password<input type="password" value={pw.current} onChange={(e) => setPw((v) => ({ ...v, current: e.target.value }))} /></label>
              <label>New password (8+ chars)<input type="password" value={pw.next} onChange={(e) => setPw((v) => ({ ...v, next: e.target.value }))} /></label>
            </div>
            {pwMsg && <div className={pwMsg.ok ? "prof-saved" : "auth-error"}>{pwMsg.text}</div>}
            <div className="prof-actions">
              <button className="auth-submit sm" onClick={changePw} disabled={!pw.current || !pw.next}>Update password</button>
            </div>
          </div>
        )}

        <div className="widget prof-card">
          <div className="widget-head">
            <h3><Mail size={15} /> Connected accounts</h3>
            <span className="widget-sub">{connected.length} connected</span>
          </div>
          <div className="prof-accounts">
            {connected.map((c) => (
              <div className="prof-acc" key={c.id}>
                <span className="qa-avatar" style={{ background: c.group === "Email" ? avatarColor(c.address) : c.accent || c.prov.gradient }}>
                  {c.group === "Email" ? initials(c.address) : (c.workspace || c.org || "?").slice(0, 1)}
                  <span className="qa-prov" style={{ background: c.prov.gradient }}>
                    {c.prov.badge || <AppleGlyph size={9} />}
                  </span>
                </span>
                <span className="prof-acc-text">
                  {editingId === c.id ? (
                    <LabelEditor
                      initial={c.label}
                      onSave={(v) => { onSetLabel(c.id, v); setEditingId(null); }}
                      onCancel={() => setEditingId(null)}
                    />
                  ) : (
                    <span className="qa-label prof-label-row">
                      {c.label || c.prov.name}
                      <button className="prof-label-edit" onClick={() => setEditingId(c.id)} title="Edit label" aria-label="Edit label">
                        <Pencil size={12} />
                      </button>
                    </span>
                  )}
                  <span className="qa-sub">{c.address} · {c.group}</span>
                </span>
                <button className="prof-disconnect" onClick={() => onDisconnect(c.id)}>Disconnect</button>
              </div>
            ))}
            {connected.length === 0 && <span className="cal-empty">No accounts connected yet.</span>}
          </div>
        </div>

        {live && (
          <button className="prof-logout" onClick={onLogout}>Log out</button>
        )}
      </div>
    </section>
  );
}

/* ---- Teams: sidebar (teams → channels hierarchy) ---- */
function TeamsSidebar({ org, selection, onOpenChannel, onOpenChat, editingLabel, onEditLabel, onSaveLabel, onGoHome }) {
  const chats = org.chats || [];
  return (
    <div className="tm-side">
      <div className="tm-head">
        <div className="tm-head-row">
          {editingLabel === org.id ? (
            <LabelEditor
              initial={org.label}
              onSave={(v) => onSaveLabel(org.id, v)}
              onCancel={() => onEditLabel(null)}
            />
          ) : (
            <button className="tm-org-name" onClick={() => onEditLabel(org.id)}>
              <span>{org.label}</span>
              <Pencil size={12} />
            </button>
          )}
          <button className="tm-home-btn" onClick={onGoHome} title="Home" aria-label="Home">
            <Home size={16} />
          </button>
        </div>
        <span className="tm-you">{org.address}</span>
      </div>

      <div className="tm-teams">
        {chats.length > 0 && (
          <>
            <span className="tm-section">Chat</span>
            <div className="tm-chats">
              {chats.map((c) => {
                const active = selection?.chat?.id === c.id;
                const last = c.messages[c.messages.length - 1];
                return (
                  <button
                    key={c.id}
                    className={`tm-chat ${active ? "is-active" : ""} ${c.unread ? "has-unread" : ""}`}
                    onClick={() => onOpenChat(c.id)}
                  >
                    <span className={`tm-chat-avatar ${c.kind === "group" ? "group" : ""}`}
                      style={c.kind === "group" ? {} : { background: avatarColor(c.name) }}>
                      {c.kind === "group"
                        ? <Users size={15} />
                        : c.name.split(" ").map((s) => s[0]).slice(0, 2).join("")}
                    </span>
                    <span className="tm-chat-text">
                      <span className="tm-chat-name">{c.name}</span>
                      {last && <span className="tm-chat-preview">{last.author.split(" ")[0]}: {last.text}</span>}
                    </span>
                    {c.unread > 0 && <span className="tm-badge">{c.unread}</span>}
                  </button>
                );
              })}
            </div>
          </>
        )}

        <span className="tm-section">Your teams</span>
        {org.teams.map((team) => (
          <div className="tm-team" key={team.id}>
            <div className="tm-team-head">
              <span className="tm-team-avatar" style={{ background: team.color }}>
                {team.initials}
              </span>
              <span className="tm-team-name">{team.name}</span>
            </div>
            <div className="tm-channels">
              {team.channels.map((c) => {
                const active =
                  selection?.team?.id === team.id && selection?.channel?.id === c.id;
                return (
                  <button
                    key={c.id}
                    className={`tm-chan ${active ? "is-active" : ""} ${c.unread ? "has-unread" : ""}`}
                    onClick={() => onOpenChannel(team.id, c.id)}
                  >
                    <span className="tm-chan-name">{c.name}</span>
                    {c.unread > 0 && <span className="tm-badge">{c.unread}</span>}
                  </button>
                );
              })}
            </div>
          </div>
        ))}
      </div>

      <div className="sidebar-foot">
        <span className="prov-tag" style={{ background: PROVIDERS.teams.gradient }}>T</span>
        Microsoft Teams
      </div>
    </div>
  );
}

/* ---- Teams: reactions row (chips + picker), Teams-styled ---- */
function TeamsReactions({ m, onReact }) {
  const [picking, setPicking] = useState(false);
  return (
    <div className="sl-reactions tm-reactions">
      {(m.reactions || []).map((r, i) => (
        <button
          key={i}
          className={`sl-reaction ${r.mine ? "mine" : ""}`}
          onClick={() => onReact?.(m.id, r.emoji)}
          title={r.mine ? "Click to remove your reaction" : "Click to add your reaction"}
        >
          <span>{r.emoji}</span>
          <span className="sl-reaction-count">{r.count}</span>
        </button>
      ))}
      <div className="sl-react-add-wrap">
        <button
          className="sl-react-add"
          onClick={() => setPicking((v) => !v)}
          aria-label="Add reaction"
          title="Add reaction"
        >
          <SmilePlus size={15} />
        </button>
        {picking && (
          <>
            <div className="sl-emoji-backdrop" onClick={() => setPicking(false)} />
            <div className="sl-emoji-pop" role="menu">
              {QUICK_EMOJI.map((e) => (
                <button
                  key={e}
                  className="sl-emoji-opt"
                  onClick={() => {
                    onReact?.(m.id, e);
                    setPicking(false);
                  }}
                >
                  {e}
                </button>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ---- Teams: a single post card with inline replies + reply box ---- */
function TeamsPost({ post, onReply, onReact }) {
  const [replyOpen, setReplyOpen] = useState(false);
  const [text, setText] = useState("");

  const submit = () => {
    if (!text.trim()) return;
    onReply(post.id, text);
    setText("");
    setReplyOpen(false);
  };

  return (
    <div className="tm-post">
      <div className="tm-post-main">
        <span className="tm-avatar" style={{ background: avatarColor(post.author) }}>
          {post.author.split(" ").map((s) => s[0]).slice(0, 2).join("")}
        </span>
        <div className="tm-post-body">
          <div className="tm-post-head">
            <span className="tm-author">{post.author}</span>
            <span className="tm-time">{post.time}</span>
          </div>
          <div className="tm-text">{post.text}</div>
          <TeamsReactions m={post} onReact={onReact} />
        </div>
      </div>

      {post.replies.length > 0 && (
        <div className="tm-replies">
          {post.replies.map((r) => (
            <div className="tm-reply" key={r.id}>
              <span className="tm-avatar sm" style={{ background: avatarColor(r.author) }}>
                {r.author.split(" ").map((s) => s[0]).slice(0, 2).join("")}
              </span>
              <div className="tm-post-body">
                <div className="tm-post-head">
                  <span className="tm-author">{r.author}</span>
                  <span className="tm-time">{r.time}</span>
                </div>
                <div className="tm-text">{r.text}</div>
                <TeamsReactions m={r} onReact={onReact} />
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="tm-reply-box">
        {replyOpen ? (
          <div className="tm-reply-input">
            <textarea
              autoFocus
              value={text}
              onChange={(e) => setText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  submit();
                }
                if (e.key === "Escape") setReplyOpen(false);
              }}
              placeholder="Type a reply…"
              rows={1}
            />
            <button className="sl-send" onClick={submit} disabled={!text.trim()} aria-label="Send reply">
              <Send size={15} />
            </button>
          </div>
        ) : (
          <button className="tm-reply-toggle" onClick={() => setReplyOpen(true)}>
            <Reply size={14} /> Reply
          </button>
        )}
      </div>
    </div>
  );
}

/* ---- Teams: main posts pane (spans both right columns, like real Teams) ---- */
function TeamsView({ org, selection, onSend, onSendChat, onReply, onReact, loading }) {
  const [postText, setPostText] = useState("");

  if (loading) {
    return (
      <section className="tm-main">
        <div className="empty">Loading Teams…</div>
      </section>
    );
  }
  if (!selection) {
    return (
      <section className="tm-main">
        <div className="reader-empty">
          <Mail size={48} strokeWidth={1.1} />
          <p>No channel selected</p>
        </div>
      </section>
    );
  }

  /* ---- chat mode: flat bubble stream, like Teams Chat ---- */
  if (selection.chat) {
    const chat = selection.chat;
    const submitChat = () => {
      if (!postText.trim()) return;
      onSendChat(postText);
      setPostText("");
    };
    return (
      <section className="tm-main">
        <header className="tm-chan-head">
          <div className="tm-chan-title">
            <span className={`tm-chat-avatar hd ${chat.kind === "group" ? "group" : ""}`}
              style={chat.kind === "group" ? {} : { background: avatarColor(chat.name) }}>
              {chat.kind === "group"
                ? <Users size={16} />
                : chat.name.split(" ").map((s) => s[0]).slice(0, 2).join("")}
            </span>
            <div>
              <h1>{chat.name}</h1>
              <span className="tm-crumb">
                {chat.kind === "group" ? "Group chat" : "Chat"} · {org.org}
              </span>
            </div>
          </div>
          <div className="tm-tabs">
            <span className="tm-tab is-active">Chat</span>
            <span className="tm-tab">Files</span>
          </div>
        </header>

        <div className="tm-stream tm-chat-stream">
          {chat.messages.map((m) => {
            const mine = m.author === org.you;
            return (
              <div className={`tm-bubble-row ${mine ? "mine" : ""}`} key={m.id}>
                {!mine && (
                  <span className="tm-avatar sm" style={{ background: avatarColor(m.author) }}>
                    {m.author.split(" ").map((s) => s[0]).slice(0, 2).join("")}
                  </span>
                )}
                <div className="tm-bubble-col">
                  <div className="tm-bubble-meta">
                    {!mine && <span className="tm-author">{m.author}</span>}
                    <span className="tm-time">{m.time}</span>
                  </div>
                  <div className={`tm-bubble ${mine ? "mine" : ""}`}>{m.text}</div>
                  <TeamsReactions m={m} onReact={onReact} />
                </div>
              </div>
            );
          })}
        </div>

        <div className="tm-composer">
          <div className="tm-reply-input new-post">
            <textarea
              value={postText}
              onChange={(e) => setPostText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  submitChat();
                }
              }}
              placeholder="Type a message"
              rows={1}
            />
            <button className="sl-send" onClick={submitChat} disabled={!postText.trim()} aria-label="Send">
              <Send size={16} />
            </button>
          </div>
        </div>
      </section>
    );
  }

  /* ---- channel mode: post cards with inline replies ---- */
  const { team, channel } = selection;
  const submitPost = () => {
    if (!postText.trim()) return;
    onSend(postText);
    setPostText("");
  };

  return (
    <section className="tm-main">
      <header className="tm-chan-head">
        <div className="tm-chan-title">
          <span className="tm-team-avatar sm" style={{ background: team.color }}>
            {team.initials}
          </span>
          <div>
            <h1>{channel.name}</h1>
            <span className="tm-crumb">{team.name} · {org.org}</span>
          </div>
        </div>
        <div className="tm-tabs">
          <span className="tm-tab is-active">Posts</span>
          <span className="tm-tab">Files</span>
          <span className="tm-tab">Wiki</span>
        </div>
      </header>

      <div className="tm-stream">
        {channel.posts.map((p) => (
          <TeamsPost key={p.id} post={p} onReply={onReply} onReact={onReact} />
        ))}
      </div>

      <div className="tm-composer">
        <div className="tm-reply-input new-post">
          <textarea
            value={postText}
            onChange={(e) => setPostText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                submitPost();
              }
            }}
            placeholder={`Start a post in ${channel.name}`}
            rows={1}
          />
          <button className="sl-send" onClick={submitPost} disabled={!postText.trim()} aria-label="Post">
            <Send size={16} />
          </button>
        </div>
      </div>
    </section>
  );
}

/* ---- Slack: channel sidebar ---- */
function SlackChannelList({ workspace, activeChannel, onOpenChannel, editingLabel, onEditLabel, onSaveLabel, onGoHome }) {
  const channels = workspace.channels.filter((c) => c.kind === "channel");
  const dms = workspace.channels.filter((c) => c.kind === "dm");
  return (
    <div className="sl-side">
      <div className="sl-ws-head" style={{ background: workspace.accent }}>
        <div className="sl-ws-head-row">
          {editingLabel === workspace.id ? (
            <LabelEditor
              initial={workspace.label}
              onSave={(v) => onSaveLabel(workspace.id, v)}
              onCancel={() => onEditLabel(null)}
            />
          ) : (
            <button className="sl-ws-name" onClick={() => onEditLabel(workspace.id)}>
              <span>{workspace.label}</span>
              <Pencil size={12} />
            </button>
          )}
          <button className="sl-home-btn" onClick={onGoHome} title="Home" aria-label="Home">
            <Home size={16} />
          </button>
        </div>
        <span className="sl-ws-you">{workspace.you}</span>
      </div>

      <div className="sl-groups">
        <div className="sl-group">
          <span className="sl-group-title">Channels</span>
          {channels.map((c) => (
            <button
              key={c.id}
              className={`sl-chan ${activeChannel?.id === c.id ? "is-active" : ""} ${c.unread ? "has-unread" : ""}`}
              onClick={() => onOpenChannel(c.id)}
            >
              <span className="sl-hash">#</span>
              <span className="sl-chan-name">{c.name}</span>
              {c.unread > 0 && <span className="sl-badge">{c.unread}</span>}
            </button>
          ))}
        </div>

        <div className="sl-group">
          <span className="sl-group-title">Direct messages</span>
          {dms.map((c) => (
            <button
              key={c.id}
              className={`sl-chan ${activeChannel?.id === c.id ? "is-active" : ""} ${c.unread ? "has-unread" : ""}`}
              onClick={() => onOpenChannel(c.id)}
            >
              <span className="sl-dm-dot" style={{ background: avatarColor(c.name) }} />
              <span className="sl-chan-name">{c.name}</span>
              {c.unread > 0 && <span className="sl-badge">{c.unread}</span>}
            </button>
          ))}
        </div>
      </div>

      <div className="sidebar-foot">
        <span className="prov-tag" style={{ background: PROVIDERS.slack.gradient }}>S</span>
        Slack workspace
      </div>
    </div>
  );
}

/* ---- Slack: message list + optional thread pane ---- */
function SlackView({ workspace, channel, openThread, onOpenThread, onCloseThread, onSend, onReply, onReact, loading }) {
  if (loading) {
    return (
      <section className="list">
        <div className="empty">Loading workspace…</div>
      </section>
    );
  }
  if (!channel) {
    return (
      <section className="reader">
        <div className="reader-empty">
          <Mail size={48} strokeWidth={1.1} />
          <p>No channel selected</p>
        </div>
      </section>
    );
  }

  return (
    <>
      <section className="list sl-main">
        <header className="sl-chan-head">
          <div className="sl-chan-head-title">
            {channel.kind === "channel" ? <span className="sl-hash big">#</span> : null}
            <h1>{channel.name}</h1>
          </div>
          <span className="sl-chan-meta">
            {channel.kind === "channel" ? `${workspace.workspace}` : "Direct message"}
          </span>
        </header>

        <div className="sl-stream">
          {channel.messages == null ? (
            <div className="empty">Loading messages…</div>
          ) : (
          channel.messages.map((m) => (
            <SlackMessage
              key={m.id}
              m={m}
              you={workspace.you}
              accent={workspace.accent}
              onOpenThread={() => onOpenThread(m.id)}
              onReact={onReact}
            />
          ))
          )}
        </div>

        <SlackComposer placeholder={`Message ${channel.kind === "channel" ? "#" + channel.name : channel.name}`}
          onSend={onSend} />
      </section>

      <section className="reader sl-thread-pane">
        {openThread ? (
          <SlackThread
            parent={openThread}
            you={workspace.you}
            accent={workspace.accent}
            channelName={channel.name}
            channelKind={channel.kind}
            onClose={onCloseThread}
            onReply={(text) => onReply(openThread.id, text)}
            onReact={onReact}
          />
        ) : (
          <div className="reader-empty">
            <Reply size={44} strokeWidth={1.1} />
            <p>Open a thread to see replies</p>
          </div>
        )}
      </section>
    </>
  );
}

// Common reactions offered in the quick picker.
const QUICK_EMOJI = ["👍", "🎉", "❤️", "😄", "🙏", "🚀", "👀", "✅", "🔥", "💡", "😅", "💯"];

function SlackMessage({ m, accent, onOpenThread, onReact, compact }) {
  const [picking, setPicking] = useState(false);

  const addReaction = (emoji) => {
    onReact?.(m.id, emoji);
    setPicking(false);
  };

  return (
    <div className={`sl-msg ${compact ? "compact" : ""}`}>
      <span className="sl-avatar" style={{ background: avatarColor(m.author) }}>
        {m.author.split(" ").map((s) => s[0]).slice(0, 2).join("")}
      </span>
      <div className="sl-msg-body">
        <div className="sl-msg-head">
          <span className="sl-author">{m.author}</span>
          <span className="sl-time">{m.time}</span>
        </div>
        <div className="sl-text">{m.text}</div>

        {/* reactions row: existing chips (clickable to toggle) + add button */}
        {(m.reactions?.length > 0 || true) && (
          <div className="sl-reactions">
            {(m.reactions || []).map((r, i) => (
              <button
                key={i}
                className={`sl-reaction ${r.mine ? "mine" : ""}`}
                onClick={() => onReact?.(m.id, r.emoji)}
                title={r.mine ? "Click to remove your reaction" : "Click to add your reaction"}
              >
                <span>{r.emoji}</span>
                <span className="sl-reaction-count">{r.count}</span>
              </button>
            ))}

            <div className="sl-react-add-wrap">
              <button
                className="sl-react-add"
                onClick={() => setPicking((v) => !v)}
                aria-label="Add reaction"
                title="Add reaction"
              >
                <SmilePlus size={15} />
              </button>
              {picking && (
                <>
                  <div className="sl-emoji-backdrop" onClick={() => setPicking(false)} />
                  <div className="sl-emoji-pop" role="menu">
                    {QUICK_EMOJI.map((e) => (
                      <button
                        key={e}
                        className="sl-emoji-opt"
                        onClick={() => addReaction(e)}
                      >
                        {e}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {!compact && ((m.replies?.length || 0) > 0 || (m.replyCount || 0) > 0) && (
          <button className="sl-thread-link" onClick={onOpenThread} style={{ color: accent }}>
            <span className="sl-thread-avatars">
              {m.replies.slice(0, 3).map((r, i) => (
                <span key={i} className="sl-thread-av" style={{ background: avatarColor(r.author) }}>
                  {r.author[0]}
                </span>
              ))}
            </span>
            {m.replies?.length || m.replyCount} {(m.replies?.length || m.replyCount) === 1 ? "reply" : "replies"}
            <span className="sl-thread-last">View thread</span>
          </button>
        )}
      </div>
    </div>
  );
}

function SlackThread({ parent, you, accent, channelName, channelKind, onClose, onReply, onReact }) {
  return (
    <>
      <div className="sl-thread-head">
        <div>
          <strong>Thread</strong>
          <span className="sl-thread-sub">
            {channelKind === "channel" ? "#" + channelName : channelName}
          </span>
        </div>
        <button className="sl-thread-close" onClick={onClose} aria-label="Close thread">
          <X size={18} />
        </button>
      </div>

      <div className="sl-thread-scroll">
        <SlackMessage m={parent} accent={accent} onReact={onReact} compact />
        <div className="sl-thread-divider">
          {parent.replies.length} {parent.replies.length === 1 ? "reply" : "replies"}
        </div>
        {parent.replies.map((r) => (
          <SlackMessage key={r.id} m={r} accent={accent} onReact={onReact} compact />
        ))}
      </div>

      <SlackComposer placeholder="Reply…" onSend={onReply} />
    </>
  );
}

function SlackComposer({ placeholder, onSend }) {
  const [text, setText] = useState("");
  const submit = () => {
    if (!text.trim()) return;
    onSend(text);
    setText("");
  };
  return (
    <div className="sl-composer">
      <div className="sl-composer-box">
        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              submit();
            }
          }}
          placeholder={placeholder}
          rows={1}
        />
        <button className="sl-send" onClick={submit} disabled={!text.trim()} aria-label="Send">
          <Send size={16} />
        </button>
      </div>
    </div>
  );
}

/* ---- Email reply composer: WYSIWYG via contentEditable + execCommand ---- */
function ReplyComposer({ mode, m, account, onSend, onClose }) {
  const initialTo = useMemo(() => {
    if (mode === "forward") return "";
    const set = new Set([m.fromAddr]);
    if (mode === "replyAll") {
      for (const addr of [...(m.to || []), ...(m.cc || [])]) set.add(addr);
    }
    set.delete(account?.address); // don't reply to yourself
    return [...set].filter(Boolean).join(", ");
  }, [mode, m, account]);

  const [to, setTo] = useState(initialTo);
  const [subject, setSubject] = useState(
    (mode === "forward" ? "Fwd: " : "Re: ") + (m.subject || "").replace(/^(Re|Fwd):\s*/i, "")
  );
  const editorRef = useRef(null);
  const [sending, setSending] = useState(false);
  const [sent, setSent] = useState(false);
  const [error, setError] = useState(null);

  // Dropdowns and popovers steal focus, which collapses the text selection —
  // so we snapshot the range before they open and restore it before applying.
  const savedRange = useRef(null);
  const [colorOpen, setColorOpen] = useState(false);
  const saveSel = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) savedRange.current = sel.getRangeAt(0).cloneRange();
  };
  const restoreSel = () => {
    const r = savedRange.current;
    if (!r) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  };
  const exec = (cmd, val = null) => {
    editorRef.current?.focus();
    document.execCommand(cmd, false, val);
  };
  const execPreserving = (cmd, val = null) => {
    restoreSel();
    exec(cmd, val);
  };

  const send = async () => {
    const html = editorRef.current?.innerHTML || "";
    const text = editorRef.current?.innerText || "";
    if (!to.trim() || !text.trim()) return;
    setSending(true);
    setError(null);
    try {
      await onSend({ to: to.trim(), subject, html, text });
      setSent(true);
      setTimeout(onClose, 900);
    } catch (e) {
      setError(e.message);
    } finally {
      setSending(false);
    }
  };

  const FONTS = [
    ["Sans Serif", "Arial, sans-serif"],
    ["Serif", "Georgia, serif"],
    ["Fixed width", "monospace"],
  ];
  const SIZES = [["Small", "1"], ["Normal", "3"], ["Large", "5"], ["Huge", "7"]];
  const COLORS = [
    "#1a1d21", "#5b6470", "#B91C1C", "#C2410C", "#B45309",
    "#15803D", "#0E7490", "#1D4ED8", "#7C3AED", "#DB2777",
  ];
  const btn = (cmd, node, title, val = null) => (
    <button
      key={title}
      className="rc-tool"
      title={title}
      aria-label={title}
      onMouseDown={(e) => { e.preventDefault(); exec(cmd, val); }}
    >
      {node}
    </button>
  );

  return (
    <div className="rc">
      <div className="rc-head">
        <strong>{mode === "forward" ? "Forward" : mode === "replyAll" ? "Reply all" : "Reply"}</strong>
        <button className="rc-close" onClick={onClose} aria-label="Discard"><X size={15} /></button>
      </div>
      <div className="rc-field">
        <label>To</label>
        <input value={to} onChange={(e) => setTo(e.target.value)} placeholder="Recipients (comma separated)" />
      </div>
      <div className="rc-field">
        <label>Subject</label>
        <input value={subject} onChange={(e) => setSubject(e.target.value)} />
      </div>
      <div className="rc-toolbar">
        <select
          className="rc-select"
          title="Font"
          aria-label="Font"
          onMouseDown={saveSel}
          onChange={(e) => execPreserving("fontName", e.target.value)}
        >
          {FONTS.map(([name, val]) => <option key={val} value={val}>{name}</option>)}
        </select>
        <select
          className="rc-select rc-select-size"
          title="Text size"
          aria-label="Text size"
          defaultValue="3"
          onMouseDown={saveSel}
          onChange={(e) => execPreserving("fontSize", e.target.value)}
        >
          {SIZES.map(([name, val]) => <option key={val} value={val}>{name}</option>)}
        </select>
        <span className="rc-sep" />
        {btn("bold", <strong>B</strong>, "Bold")}
        {btn("italic", <em>I</em>, "Italic")}
        {btn("underline", <u>U</u>, "Underline")}
        {btn("strikeThrough", <s>S</s>, "Strikethrough")}
        <span className="rc-color-wrap">
          <button
            className="rc-tool rc-color-btn"
            title="Text color"
            aria-label="Text color"
            onMouseDown={(e) => { e.preventDefault(); saveSel(); setColorOpen((v) => !v); }}
          >
            A<span className="rc-color-bar" />
          </button>
          {colorOpen && (
            <>
              <div className="sl-emoji-backdrop" onMouseDown={() => setColorOpen(false)} />
              <div className="rc-color-pop" role="menu">
                {COLORS.map((c) => (
                  <button
                    key={c}
                    className="rc-swatch"
                    style={{ background: c }}
                    aria-label={`Color ${c}`}
                    onMouseDown={(e) => { e.preventDefault(); execPreserving("foreColor", c); setColorOpen(false); }}
                  />
                ))}
              </div>
            </>
          )}
        </span>
        <span className="rc-sep" />
        {btn("justifyLeft", <AlignLeft size={15} />, "Align left")}
        {btn("justifyCenter", <AlignCenter size={15} />, "Align center")}
        {btn("justifyRight", <AlignRight size={15} />, "Align right")}
        <span className="rc-sep" />
        {btn("insertOrderedList", <ListOrdered size={15} />, "Numbered list")}
        {btn("insertUnorderedList", <List size={15} />, "Bullet list")}
        {btn("outdent", <IndentDecrease size={15} />, "Decrease indent")}
        {btn("indent", <IndentIncrease size={15} />, "Increase indent")}
        {btn("formatBlock", <Quote size={15} />, "Quote", "blockquote")}
        <span className="rc-sep" />
        {btn("removeFormat", <RemoveFormatting size={15} />, "Clear formatting")}
      </div>
      <div
        ref={editorRef}
        className="rc-editor"
        contentEditable
        suppressContentEditableWarning
        role="textbox"
        aria-multiline="true"
        aria-label="Message body"
      />
      {error && <div className="auth-error rc-error">{error}</div>}
      <div className="rc-foot">
        <button className="reply-btn" onClick={send} disabled={sending || sent}>
          <Send size={14} /> {sent ? "Sent ✓" : sending ? "Sending…" : "Send"}
        </button>
        <button className="reply-btn ghost" onClick={onClose}>Discard</button>
      </div>
    </div>
  );
}

function Reader({ m, account, onStar, onSendReply, onMove, onMarkUnread }) {
  const [replyMode, setReplyMode] = useState(null);
  const [moreOpen, setMoreOpen] = useState(false);
  return (
    <>
      <div className="reader-bar">
        <div className="reader-actions">
          <button title="Archive" aria-label="Archive" onClick={() => onMove("archive")}><Archive size={17} /></button>
          <button title="Delete" aria-label="Delete" onClick={() => onMove("trash")}><Trash2 size={17} /></button>
          <button title="Mark as spam" aria-label="Mark as spam" onClick={() => onMove("spam")}><AlertOctagonIcon size={17} /></button>
          <span className="bar-sep" />
          <button className={m.starred ? "on" : ""} onClick={onStar} title="Star" aria-label="Star">
            <Star size={17} fill={m.starred ? "currentColor" : "none"} />
          </button>
          <span className="reader-more-wrap">
            <button title="More" aria-label="More actions" onClick={() => setMoreOpen((v) => !v)}>
              <MoreHorizontal size={17} />
            </button>
            {moreOpen && (
              <>
                <div className="sl-emoji-backdrop" onClick={() => setMoreOpen(false)} />
                <div className="reader-more-menu" role="menu">
                  <button onClick={() => { setMoreOpen(false); onMarkUnread(); }}>
                    <MailOpen size={14} /> Mark as unread
                  </button>
                  <button onClick={() => { setMoreOpen(false); onMove("inbox"); }}>
                    <Inbox size={14} /> Move to inbox
                  </button>
                </div>
              </>
            )}
          </span>
        </div>
        <span className="reader-acct">{account?.label || account?.address}</span>
      </div>

      <div className="reader-scroll">
        <h2 className="reader-subject">{m.subject}</h2>
        <div className="reader-from">
          <span className="reader-avatar" style={{ background: avatarColor(m.fromAddr) }}>
            {initials(m.fromAddr)}
          </span>
          <div className="reader-from-meta">
            <div className="reader-from-line">
              <strong>{m.from}</strong>
              <span className="reader-fulldate">{m.fullDate}</span>
            </div>
            <span className="reader-fromaddr">
              {m.fromAddr} &nbsp;to&nbsp; me
            </span>
          </div>
        </div>

        {m.bodyHtml ? (
          <iframe
            className="reader-html"
            title="Email content"
            sandbox="allow-popups allow-popups-to-escape-sandbox"
            srcDoc={emailHtmlDoc(m.bodyHtml)}
          />
        ) : (
          <div className="reader-body">
            {(m.body || m.preview || "").split("\n").map((line, i) =>
              line.trim() === "" ? <br key={i} /> : <p key={i}>{line}</p>
            )}
          </div>
        )}

        {m.attachments.length > 0 && (
          <div className="attachments">
            {m.attachments.map((a, i) => (
              <div className="attachment" key={i}>
                <Paperclip size={15} />
                <span className="att-name">{a.name}</span>
                <span className="att-size">{a.size}</span>
              </div>
            ))}
          </div>
        )}

        <div className="reader-reply">
          <button className="reply-btn" onClick={() => setReplyMode("reply")}><Reply size={15} /> Reply</button>
          <button className="reply-btn ghost" onClick={() => setReplyMode("replyAll")}><ReplyAll size={15} /> Reply all</button>
          <button className="reply-btn ghost" onClick={() => setReplyMode("forward")}><Forward size={15} /> Forward</button>
        </div>

        {replyMode && (
          <ReplyComposer
            key={replyMode + m.id}
            mode={replyMode}
            m={m}
            account={account}
            onSend={(payload) => onSendReply(m, payload)}
            onClose={() => setReplyMode(null)}
          />
        )}
      </div>
    </>
  );
}

function LabelEditor({ initial, onSave, onCancel }) {
  const [val, setVal] = useState(initial || "");
  const ref = useRef(null);
  useEffect(() => ref.current?.focus(), []);
  return (
    <div className="label-edit">
      <input
        ref={ref}
        value={val}
        maxLength={24}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") onSave(val.trim());
          if (e.key === "Escape") onCancel();
        }}
        placeholder="e.g. Personal"
      />
      <button onClick={() => onSave(val.trim())} title="Save"><Check size={14} /></button>
    </div>
  );
}

function Compose({ account, onClose }) {
  return (
    <div className="compose-window">
      <div className="compose-head">
        <span>New message</span>
        <button onClick={onClose}><X size={16} /></button>
      </div>
      <div className="compose-field">
        <label>From</label>
        <span className="compose-from">{account.address}</span>
      </div>
      <div className="compose-field"><label>To</label><input placeholder="Recipients" /></div>
      <div className="compose-field"><label>Subject</label><input placeholder="Subject" /></div>
      <textarea className="compose-body" placeholder="Write your message…" />
      <div className="compose-foot">
        <button className="compose-send"><Send size={14} /> Send</button>
        <button className="compose-attach"><Paperclip size={15} /></button>
      </div>
    </div>
  );
}

function Modal({ children, onClose }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <button className="modal-close" onClick={onClose}><X size={18} /></button>
        {children}
      </div>
    </div>
  );
}

function AppleGlyph({ size = 13 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="#fff" aria-hidden>
      <path d="M17.05 12.04c-.03-2.6 2.12-3.85 2.22-3.91-1.21-1.77-3.1-2.01-3.77-2.04-1.6-.16-3.13.94-3.94.94-.81 0-2.07-.92-3.4-.89-1.75.03-3.36 1.02-4.26 2.58-1.82 3.15-.46 7.8 1.3 10.35.86 1.25 1.88 2.65 3.22 2.6 1.29-.05 1.78-.83 3.34-.83 1.56 0 2 .83 3.37.81 1.39-.03 2.27-1.27 3.12-2.53.98-1.45 1.39-2.86 1.41-2.93-.03-.01-2.71-1.04-2.74-4.13l-.01.01zM14.53 4.4c.71-.86 1.19-2.06 1.06-3.25-1.02.04-2.26.68-3 1.54-.66.76-1.24 1.98-1.08 3.15 1.14.09 2.3-.58 3.02-1.44z" />
    </svg>
  );
}

/* --------------------------------- styles --------------------------------- */

const CSS = `
:root{
  --bg:#f6f7f9; --panel:#ffffff; --ink:#1a1d21; --ink-2:#5b6470; --ink-3:#8a929e;
  --line:#e7eaee; --line-2:#eef1f4; --accent:#3b5bdb; --accent-soft:#eef2ff;
  --unread:#1a1d21; --star:#f59e0b; --shadow:0 1px 3px rgba(16,24,40,.06),0 1px 2px rgba(16,24,40,.04);
  --rail:#1b2230;
}
*{box-sizing:border-box}
.md-root{
  position:fixed; inset:0; display:grid;
  grid-template-columns:64px 232px 360px 1fr;
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,Roboto,sans-serif;
  background:var(--bg); color:var(--ink); font-size:14px; overflow:hidden;
  transition:grid-template-columns .22s cubic-bezier(.4,0,.2,1);
}
.md-root.rail-open{grid-template-columns:248px 232px 360px 1fr}
button{font-family:inherit; cursor:pointer; border:none; background:none; color:inherit}

/* ---- rail ---- */
.rail{
  background:var(--rail); display:flex; flex-direction:column; align-items:center;
  padding:14px 10px 12px; gap:6px; overflow:hidden;
}
.rail.is-expanded{align-items:stretch}
.rail-top{
  display:flex; align-items:center; gap:10px; width:100%;
  justify-content:center; min-height:38px;
}
.rail.is-expanded .rail-top{justify-content:flex-start; padding:0 2px}
.rail-brand{
  width:38px; height:38px; border-radius:11px; display:grid; place-items:center;
  background:linear-gradient(135deg,#4f6ef7,#3b5bdb); color:#fff; flex-shrink:0;
}
.rail-wordmark{
  color:#fff; font-weight:700; font-size:15px; letter-spacing:-.01em; flex:1;
  white-space:nowrap;
}
.rail-toggle{
  color:#8893a7; width:38px; height:30px; border-radius:8px; display:grid;
  place-items:center; flex-shrink:0; transition:.15s; align-self:center;
  margin:8px 0; border:1px solid rgba(255,255,255,.08);
}
.rail-toggle:hover{color:#fff; background:rgba(255,255,255,.08)}
.rail.is-expanded .rail-toggle{
  width:100%; display:flex; gap:9px; padding:0 12px; justify-content:flex-start;
  align-self:stretch; font-size:12.5px; font-weight:600;
}
.rail.is-expanded .rail-toggle span{white-space:nowrap}

.rail-accounts{
  display:flex; flex-direction:column; align-items:center; gap:10px; flex:1; width:100%;
  border-top:1px solid rgba(255,255,255,.07); padding-top:12px;
}
.rail.is-expanded .rail-accounts{align-items:stretch; gap:4px}

.rail-acc{position:relative; width:44px; height:44px; border-radius:50%; transition:transform .12s; align-self:center}
.rail-acc:hover{transform:translateY(-1px)}
.rail-acc.is-active::before{
  content:""; position:absolute; left:-10px; top:50%; transform:translateY(-50%);
  width:4px; height:26px; border-radius:0 4px 4px 0; background:#fff;
}
/* expanded: account becomes a full-width row */
.rail.is-expanded .rail-acc{
  width:100%; height:auto; border-radius:11px; display:flex; align-items:center;
  gap:11px; padding:8px 10px; align-self:stretch; text-align:left;
}
.rail.is-expanded .rail-acc:hover{transform:none; background:rgba(255,255,255,.05)}
.rail.is-expanded .rail-acc.is-active{background:rgba(79,110,247,.18)}
.rail.is-expanded .rail-acc.is-active::before{left:-10px; height:60%}

.rail-avatar{
  width:44px; height:44px; border-radius:50%; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:14px; position:relative; flex-shrink:0;
  box-shadow:0 0 0 2px var(--rail), 0 0 0 2px transparent;
}
.rail.is-expanded .rail-avatar{width:40px; height:40px; font-size:13px}
.rail-acc.is-active .rail-avatar{box-shadow:0 0 0 2px var(--rail),0 0 0 4px #4f6ef7}
.rail.is-expanded .rail-acc.is-active .rail-avatar{box-shadow:none}
.rail-provider{
  position:absolute; bottom:-2px; right:-3px; width:18px; height:18px; border-radius:50%;
  display:grid; place-items:center; color:#fff; font-size:10px; font-weight:700;
  border:2px solid var(--rail);
}
.rail-acc-text{display:flex; flex-direction:column; min-width:0; flex:1; gap:1px}
.rail-acc-label{
  color:#fff; font-weight:600; font-size:13.5px; white-space:nowrap;
  overflow:hidden; text-overflow:ellipsis;
}
.rail-acc-addr{
  color:#9aa4b6; font-size:11.5px; white-space:nowrap; overflow:hidden;
  text-overflow:ellipsis;
}
.rail-dot{
  position:absolute; top:-3px; right:-3px; min-width:18px; height:18px; padding:0 4px;
  border-radius:9px; background:#ef4444; color:#fff; font-size:10px; font-weight:700;
  display:grid; place-items:center; border:2px solid var(--rail);
}
.rail-dot.inline{
  position:static; border:none; flex-shrink:0; height:19px; min-width:19px;
}

.rail-add{
  width:44px; height:44px; border-radius:50%; border:1.5px dashed #44506a; color:#aab4c6;
  display:grid; place-items:center; transition:.15s; align-self:center;
}
.rail-add:hover{border-color:#6b7896; color:#fff; background:rgba(255,255,255,.05)}
.rail-add.wide{
  width:100%; height:auto; border-radius:11px; display:flex; gap:10px; padding:9px 12px;
  justify-content:flex-start; align-self:stretch; font-size:13px; font-weight:600; margin-top:4px;
}
.rail-add.wide span{white-space:nowrap}

.rail-settings{color:#8893a7; padding:8px; border-radius:8px; align-self:center; display:grid; place-items:center}
.rail-settings:hover{color:#fff; background:rgba(255,255,255,.06)}
.rail-settings.wide{
  width:100%; display:flex; gap:11px; padding:9px 12px; border-radius:11px;
  justify-content:flex-start; align-self:stretch; font-size:13px; font-weight:500;
}
.rail-settings.wide span{white-space:nowrap}

/* ---- sidebar ---- */
.sidebar{
  background:var(--panel); border-right:1px solid var(--line);
  display:flex; flex-direction:column; padding:16px 12px;
}
.acc-head{display:flex; gap:10px; align-items:center; padding:4px 6px 14px}
.acc-avatar{
  width:38px; height:38px; border-radius:11px; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:14px; flex-shrink:0;
}
.acc-meta{min-width:0; display:flex; flex-direction:column; gap:2px}
.acc-label{
  display:inline-flex; align-items:center; gap:5px; font-weight:650; font-size:14px;
  color:var(--ink); padding:0; border-radius:6px;
}
.acc-label span{max-width:120px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.acc-label svg{opacity:.5; transition:.15s; flex-shrink:0; color:var(--ink-3)}
.acc-label:hover svg{opacity:1}
.acc-addr{font-size:11.5px; color:var(--ink-3); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:170px}
.label-edit{display:flex; gap:4px; align-items:center}
.label-edit input{
  width:118px; border:1px solid var(--accent); border-radius:6px; padding:3px 7px;
  font-size:13px; font-weight:600; outline:none; font-family:inherit;
}
.label-edit button{background:var(--accent); color:#fff; border-radius:6px; padding:4px; display:grid; place-items:center}

.compose{
  display:flex; align-items:center; justify-content:center; gap:8px; margin:2px 0 16px;
  background:var(--accent); color:#fff; font-weight:600; font-size:13.5px;
  padding:10px; border-radius:10px; box-shadow:var(--shadow); transition:.15s;
}
.compose:hover{background:#324bc0}

.folders{list-style:none; margin:0; padding:0; display:flex; flex-direction:column; gap:1px}
.folder{
  width:100%; display:flex; align-items:center; gap:11px; padding:8px 10px;
  border-radius:8px; color:var(--ink-2); font-size:13.5px; font-weight:500; transition:.12s;
}
.folder:hover{background:var(--line-2)}
.folder.is-active{background:var(--accent-soft); color:var(--accent); font-weight:600}
.folder.is-active svg{color:var(--accent)}
.folder-name{flex:1; text-align:left}
.folder-count{
  font-size:11.5px; font-weight:700; background:var(--accent); color:#fff;
  min-width:19px; height:19px; padding:0 6px; border-radius:10px; display:grid; place-items:center;
}
.folder-count.muted{background:var(--line); color:var(--ink-2)}
.folder.is-active .folder-count{background:var(--accent); color:#fff}

.sidebar-foot{
  margin-top:auto; padding:12px 8px 2px; display:flex; align-items:center; gap:8px;
  font-size:12px; color:var(--ink-3); border-top:1px solid var(--line-2);
}
.prov-tag{
  width:20px; height:20px; border-radius:6px; display:grid; place-items:center;
  color:#fff; font-size:11px; font-weight:700;
}

/* ---- list ---- */
.list{background:var(--panel); border-right:1px solid var(--line); display:flex; flex-direction:column; min-width:0}
.list-head{padding:16px 16px 12px; border-bottom:1px solid var(--line-2)}
.list-title{display:flex; align-items:center; justify-content:space-between; margin-bottom:12px}
.list-title h1{margin:0; font-size:18px; font-weight:700; letter-spacing:-.01em}
.refresh{color:var(--ink-3); padding:6px; border-radius:7px}
.refresh:hover{background:var(--line-2); color:var(--ink)}
.search{
  display:flex; align-items:center; gap:8px; background:var(--bg); border:1px solid var(--line);
  border-radius:9px; padding:8px 11px; color:var(--ink-3);
}
.search input{flex:1; border:none; background:none; outline:none; font-size:13.5px; color:var(--ink); font-family:inherit}
.search-clear{color:var(--ink-3); display:grid; place-items:center; padding:2px; border-radius:5px}
.search-clear:hover{background:var(--line); color:var(--ink)}

.list-scroll{overflow-y:auto; flex:1}
.row{
  width:100%; display:flex; gap:11px; padding:13px 16px; text-align:left;
  border-bottom:1px solid var(--line-2); transition:background .1s; position:relative;
}
.row:hover{background:#fafbfc}
.row.is-selected{background:var(--accent-soft)}
.row.is-selected:hover{background:var(--accent-soft)}
.row-avatar{
  width:36px; height:36px; border-radius:50%; flex-shrink:0; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:13px;
}
.row-body{min-width:0; flex:1}
.row-top{display:flex; align-items:baseline; justify-content:space-between; gap:8px}
.row-from{font-size:13.5px; color:var(--ink); overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.row.is-unread .row-from{font-weight:700}
.row-date{font-size:11.5px; color:var(--ink-3); flex-shrink:0}
.row.is-unread .row-date{color:var(--accent); font-weight:600}
.row-subject{
  font-size:13px; color:var(--ink); margin:2px 0 1px; overflow:hidden;
  text-overflow:ellipsis; white-space:nowrap;
}
.row.is-unread .row-subject{font-weight:650}
.row-preview{font-size:12.5px; color:var(--ink-3); overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.row-marks{display:flex; flex-direction:column; align-items:center; gap:5px; flex-shrink:0}
.row-star{color:var(--ink-3); display:grid; place-items:center}
.row-star:hover{color:var(--star)}
.row-star.on{color:var(--star)}
.row-clip{color:var(--ink-3)}
.row-unread-dot{color:var(--accent)}
.row{position:relative}
.row-actions{
  position:absolute; right:10px; top:50%; transform:translateY(-50%);
  display:none; gap:2px; background:var(--panel); border:1px solid var(--line);
  border-radius:9px; padding:3px; box-shadow:var(--shadow);
}
.row:hover .row-actions{display:flex}
.row-act{
  width:27px; height:27px; border-radius:7px; display:grid; place-items:center;
  color:var(--ink-2); cursor:pointer; transition:.1s;
}
.row-act:hover{background:var(--line-2); color:var(--ink)}
.row-act.danger:hover{background:#FEF2F2; color:#B91C1C}

.reader-more-wrap{position:relative; display:inline-flex}
.reader-more-menu{
  position:absolute; top:calc(100% + 6px); right:0; z-index:31; background:var(--panel);
  border:1px solid var(--line); border-radius:11px; box-shadow:0 10px 30px rgba(16,24,40,.16);
  padding:5px; display:flex; flex-direction:column; min-width:180px; animation:pop .14s;
}
.reader-more-menu button{
  display:flex; align-items:center; gap:9px; padding:8px 11px; border-radius:8px;
  font-size:13px; color:var(--ink); text-align:left; width:100%;
}
.reader-more-menu button:hover{background:var(--line-2)}

.reader-html{
  width:100%; min-height:420px; border:none; background:var(--panel);
  border-radius:10px; margin:20px 0; display:block;
}

.empty,.reader-empty{
  display:flex; flex-direction:column; align-items:center; justify-content:center;
  gap:12px; height:100%; color:var(--ink-3); padding:40px; text-align:center;
}
.empty p,.reader-empty p{margin:0; font-size:13.5px}

/* ---- reader ---- */
.reader{background:var(--bg); display:flex; flex-direction:column; min-width:0}
.reader-bar{
  display:flex; align-items:center; justify-content:space-between; padding:12px 22px;
  border-bottom:1px solid var(--line); background:var(--panel);
}
.reader-actions{display:flex; align-items:center; gap:2px}
.reader-actions button{color:var(--ink-2); padding:8px; border-radius:8px; display:grid; place-items:center}
.reader-actions button:hover{background:var(--line-2); color:var(--ink)}
.reader-actions button.on{color:var(--star)}
.bar-sep{width:1px; height:20px; background:var(--line); margin:0 6px}
.reader-acct{font-size:12px; color:var(--ink-3); font-weight:500}

.reader-scroll{overflow-y:auto; flex:1; padding:28px 40px 60px; max-width:860px; width:100%}
.reader-subject{margin:0 0 22px; font-size:22px; font-weight:700; letter-spacing:-.015em; line-height:1.25}
.reader-from{display:flex; gap:13px; align-items:center; padding-bottom:22px; border-bottom:1px solid var(--line)}
.reader-avatar{
  width:44px; height:44px; border-radius:50%; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:16px; flex-shrink:0;
}
.reader-from-meta{min-width:0}
.reader-from-line{display:flex; align-items:baseline; gap:12px}
.reader-from-line strong{font-size:14.5px}
.reader-fulldate{font-size:12px; color:var(--ink-3)}
.reader-fromaddr{font-size:12.5px; color:var(--ink-3)}
.reader-body{padding:24px 0; font-size:14.5px; line-height:1.7; color:#2c3138}
.reader-body p{margin:0 0 4px}
.attachments{display:flex; flex-wrap:wrap; gap:10px; padding-bottom:24px}
.attachment{
  display:flex; align-items:center; gap:8px; background:var(--panel); border:1px solid var(--line);
  border-radius:9px; padding:9px 13px; font-size:12.5px;
}
.att-name{font-weight:600; color:var(--ink)}
.att-size{color:var(--ink-3)}
.reader-reply{display:flex; gap:10px; padding-top:20px; border-top:1px solid var(--line)}
.reply-btn{
  display:flex; align-items:center; gap:7px; background:var(--accent); color:#fff;
  font-weight:600; font-size:13px; padding:9px 16px; border-radius:9px; transition:.15s;
}
.reply-btn:hover{background:#324bc0}
.reply-btn.ghost{background:var(--panel); color:var(--ink-2); border:1px solid var(--line)}
.reply-btn.ghost:hover{background:var(--line-2); color:var(--ink)}

/* ---- modal ---- */
.modal-overlay{position:fixed; inset:0; background:rgba(16,22,33,.45); backdrop-filter:blur(2px); display:grid; place-items:center; z-index:50; animation:fade .15s}
.modal{background:var(--panel); border-radius:18px; padding:30px; width:400px; max-width:92vw; position:relative; box-shadow:0 20px 60px rgba(0,0,0,.25); animation:pop .18s}
.modal-close{position:absolute; top:18px; right:18px; color:var(--ink-3); padding:5px; border-radius:7px}
.modal-close:hover{background:var(--line-2); color:var(--ink)}
.modal-title{margin:0 0 6px; font-size:19px; font-weight:700}
.modal-sub{margin:0 0 22px; font-size:13.5px; color:var(--ink-2)}
.provider-grid{display:flex; flex-direction:column; gap:10px}
.provider-btn{
  display:flex; align-items:center; gap:13px; padding:13px 15px; border:1px solid var(--line);
  border-radius:12px; transition:.13s; font-weight:600; font-size:14px;
}
.provider-btn:hover:not(:disabled){border-color:var(--accent); background:var(--accent-soft)}
.provider-btn:disabled{opacity:.6; cursor:default}
.provider-mark{width:32px; height:32px; border-radius:9px; display:grid; place-items:center; color:#fff; font-weight:700; font-size:15px; flex-shrink:0}
.provider-name{flex:1; text-align:left}
.modal-note{margin:20px 0 0; font-size:11.5px; color:var(--ink-3); line-height:1.55}
.spin{animation:spin 1s linear infinite}

/* ---- compose ---- */
.compose-window{
  position:fixed; right:24px; bottom:0; width:440px; max-width:92vw; background:var(--panel);
  border-radius:14px 14px 0 0; box-shadow:0 -4px 30px rgba(0,0,0,.18); z-index:40;
  display:flex; flex-direction:column; overflow:hidden; animation:slideup .2s;
}
.compose-head{display:flex; align-items:center; justify-content:space-between; padding:12px 16px; background:#222b3a; color:#fff; font-size:13.5px; font-weight:600}
.compose-head button{color:#fff; padding:3px; border-radius:6px}
.compose-head button:hover{background:rgba(255,255,255,.12)}
.compose-field{display:flex; align-items:center; gap:10px; padding:10px 16px; border-bottom:1px solid var(--line-2)}
.compose-field label{font-size:12px; color:var(--ink-3); width:54px; flex-shrink:0}
.compose-field input{flex:1; border:none; outline:none; font-size:13.5px; font-family:inherit}
.compose-from{font-size:13px; color:var(--ink-2); font-weight:500}
.compose-body{border:none; outline:none; resize:none; padding:16px; min-height:200px; font-size:14px; font-family:inherit; line-height:1.6}
.compose-foot{display:flex; align-items:center; gap:12px; padding:12px 16px; border-top:1px solid var(--line-2)}
.compose-send{display:flex; align-items:center; gap:7px; background:var(--accent); color:#fff; font-weight:600; font-size:13px; padding:9px 18px; border-radius:9px}
.compose-send:hover{background:#324bc0}
.compose-attach{color:var(--ink-3); padding:8px; border-radius:7px}
.compose-attach:hover{background:var(--line-2); color:var(--ink)}

@keyframes fade{from{opacity:0}to{opacity:1}}
@keyframes pop{from{opacity:0;transform:scale(.96)}to{opacity:1;transform:scale(1)}}
@keyframes slideup{from{transform:translateY(40px);opacity:0}to{transform:translateY(0);opacity:1}}
@keyframes spin{to{transform:rotate(360deg)}}

@media (max-width:1080px){
  .md-root{grid-template-columns:64px 200px 320px 1fr}
  .md-root.rail-open{grid-template-columns:248px 200px 320px 1fr}
}
@media (prefers-reduced-motion:reduce){*{animation:none!important; transition:none!important}}

/* ---- rail: separator + slack workspaces ---- */
.rail-sep{
  height:1px; background:rgba(255,255,255,.1); margin:10px 4px; align-self:stretch;
  flex-shrink:0;
}
.rail.is-expanded .rail-sep{margin:10px 0}
.rail-section-label{
  display:none; color:#6b7688; font-size:10px; font-weight:700; letter-spacing:.08em;
  text-transform:uppercase; padding:2px 6px 4px;
}
.rail.is-expanded .rail-section-label{display:block}
.rail-ws-logo{
  width:44px; height:44px; border-radius:13px; display:grid; place-items:center;
  color:#fff; font-weight:700; font-size:18px; position:relative; flex-shrink:0;
}
.rail.is-expanded .rail-ws-logo{width:40px; height:40px; border-radius:11px; font-size:16px}
.rail-ws.is-active .rail-ws-logo{box-shadow:0 0 0 2px var(--rail),0 0 0 4px #611f69}
.rail.is-expanded .rail-ws.is-active .rail-ws-logo{box-shadow:none}

/* ===== Slack: channel sidebar ===== */
.sl-side{display:flex; flex-direction:column; height:100%; margin:-16px -12px; }
.sl-ws-head{padding:16px 16px 14px; color:#fff}
.sl-ws-head-row{display:flex; align-items:center}
.sl-ws-name{
  display:inline-flex; align-items:center; gap:7px; color:#fff; font-weight:700;
  font-size:15.5px; letter-spacing:-.01em;
}
.sl-ws-name svg{opacity:0; transition:.15s}
.sl-ws-name:hover svg{opacity:.85}
.sl-ws-you{font-size:12px; color:rgba(255,255,255,.7); display:flex; align-items:center; gap:6px; margin-top:3px}
.sl-ws-you::before{content:""; width:8px; height:8px; border-radius:50%; background:#2EB67D; box-shadow:0 0 0 2px rgba(255,255,255,.25)}

.sl-groups{flex:1; overflow-y:auto; padding:12px 8px}
.sl-group{margin-bottom:18px}
.sl-group-title{
  display:block; font-size:11px; font-weight:700; color:var(--ink-3); letter-spacing:.04em;
  text-transform:uppercase; padding:0 8px 6px;
}
.sl-chan{
  width:100%; display:flex; align-items:center; gap:9px; padding:6px 8px; border-radius:7px;
  color:var(--ink-2); font-size:14px; transition:.1s; text-align:left;
}
.sl-chan:hover{background:var(--line-2)}
.sl-chan.is-active{background:var(--accent); color:#fff}
.sl-chan.has-unread{color:var(--ink); font-weight:700}
.sl-chan.is-active.has-unread{color:#fff}
.sl-hash{color:var(--ink-3); font-weight:600; flex-shrink:0}
.sl-chan.is-active .sl-hash{color:rgba(255,255,255,.85)}
.sl-hash.big{font-size:20px; color:var(--ink-3)}
.sl-dm-dot{width:9px; height:9px; border-radius:50%; flex-shrink:0; margin:0 1px}
.sl-chan-name{flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.sl-badge{
  background:#ef4444; color:#fff; font-size:11px; font-weight:700; min-width:18px; height:18px;
  border-radius:9px; padding:0 5px; display:grid; place-items:center; flex-shrink:0;
}
.sl-chan.is-active .sl-badge{background:rgba(255,255,255,.25)}

/* ===== Slack: main message pane ===== */
.sl-main{display:flex; flex-direction:column}
.sl-chan-head{
  padding:14px 22px; border-bottom:1px solid var(--line); display:flex;
  flex-direction:column; gap:2px;
}
.sl-chan-head-title{display:flex; align-items:center; gap:8px}
.sl-chan-head-title h1{margin:0; font-size:18px; font-weight:700; letter-spacing:-.01em}
.sl-chan-meta{font-size:12px; color:var(--ink-3)}

.sl-stream{flex:1; overflow-y:auto; padding:16px 0}
.sl-msg{display:flex; gap:11px; padding:7px 22px; transition:background .1s}
.sl-msg:hover{background:#fafbfc}
.sl-msg.compact{padding:7px 20px}
.sl-avatar{
  width:38px; height:38px; border-radius:9px; flex-shrink:0; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:14px;
}
.sl-msg.compact .sl-avatar{width:32px; height:32px; font-size:12px; border-radius:7px}
.sl-msg-body{min-width:0; flex:1}
.sl-msg-head{display:flex; align-items:baseline; gap:9px}
.sl-author{font-weight:700; font-size:14px; color:var(--ink)}
.sl-time{font-size:11px; color:var(--ink-3)}
.sl-text{font-size:14px; line-height:1.5; color:#2c3138; margin-top:1px; white-space:pre-wrap; word-break:break-word}

.sl-reactions{display:flex; gap:6px; margin-top:6px; flex-wrap:wrap; align-items:center}
.sl-reaction{
  display:inline-flex; align-items:center; gap:5px; background:var(--bg);
  border:1px solid var(--line); border-radius:11px; padding:2px 8px; font-size:12px;
  cursor:pointer; transition:.12s; font-family:inherit;
}
.sl-reaction:hover{border-color:var(--accent); background:var(--accent-soft)}
.sl-reaction.mine{background:var(--accent-soft); border-color:var(--accent)}
.sl-reaction.mine .sl-reaction-count{color:var(--accent)}
.sl-reaction-count{font-weight:600; color:var(--ink-2)}

.sl-react-add-wrap{position:relative; display:inline-flex}
.sl-react-add{
  display:inline-flex; align-items:center; justify-content:center; width:28px; height:24px;
  border:1px solid var(--line); border-radius:11px; background:var(--bg); color:var(--ink-3);
  cursor:pointer; transition:.12s;
}
.sl-react-add:hover{border-color:var(--accent); color:var(--accent); background:var(--accent-soft)}
.sl-emoji-backdrop{position:fixed; inset:0; z-index:30}
.sl-emoji-pop{
  position:absolute; bottom:calc(100% + 6px); left:0; z-index:31; background:var(--panel);
  border:1px solid var(--line); border-radius:12px; box-shadow:0 10px 30px rgba(16,24,40,.16);
  padding:8px; display:grid; grid-template-columns:repeat(6,1fr); gap:2px; width:236px;
  animation:pop .14s;
}
.sl-emoji-opt{
  width:34px; height:34px; border-radius:8px; font-size:18px; display:grid; place-items:center;
  cursor:pointer; transition:.1s; background:none; border:none; line-height:1;
}
.sl-emoji-opt:hover{background:var(--line-2)}

.sl-thread-link{
  display:inline-flex; align-items:center; gap:8px; margin-top:7px; font-size:13px;
  font-weight:600; padding:3px 8px 3px 4px; border-radius:14px; transition:.12s;
}
.sl-thread-link:hover{background:var(--accent-soft)}
.sl-thread-avatars{display:flex}
.sl-thread-av{
  width:20px; height:20px; border-radius:5px; display:grid; place-items:center; color:#fff;
  font-size:10px; font-weight:700; margin-right:-4px; border:2px solid var(--panel);
}
.sl-thread-last{color:var(--ink-3); font-weight:500; margin-left:2px}

/* composer (shared by channel + thread) */
.sl-composer{padding:10px 18px 16px}
.sl-composer-box{
  display:flex; align-items:flex-end; gap:8px; border:1px solid var(--line); border-radius:11px;
  padding:8px 8px 8px 14px; background:var(--panel); box-shadow:var(--shadow);
}
.sl-composer-box:focus-within{border-color:var(--accent)}
.sl-composer textarea{
  flex:1; border:none; outline:none; resize:none; font-family:inherit; font-size:14px;
  line-height:1.5; max-height:140px; padding:4px 0; background:none;
}
.sl-send{
  width:34px; height:34px; border-radius:8px; background:var(--accent); color:#fff;
  display:grid; place-items:center; flex-shrink:0; transition:.15s;
}
.sl-send:hover:not(:disabled){background:#324bc0}
.sl-send:disabled{background:var(--line); color:var(--ink-3); cursor:default}

/* ===== Slack: thread pane ===== */
.sl-thread-pane{background:var(--panel)}
.sl-thread-head{
  display:flex; align-items:center; justify-content:space-between; padding:15px 20px;
  border-bottom:1px solid var(--line);
}
.sl-thread-head strong{font-size:16px}
.sl-thread-sub{font-size:12.5px; color:var(--ink-3); margin-left:9px}
.sl-thread-close{color:var(--ink-3); padding:6px; border-radius:7px; display:grid; place-items:center}
.sl-thread-close:hover{background:var(--line-2); color:var(--ink)}
.sl-thread-scroll{flex:1; overflow-y:auto; padding:14px 0}
.sl-thread-divider{
  display:flex; align-items:center; gap:12px; font-size:12px; color:var(--ink-3); font-weight:600;
  padding:8px 20px; white-space:nowrap;
}
.sl-thread-divider::after{content:""; flex:1; height:1px; background:var(--line)}

/* ===== Home screen ===== */
.home{grid-column:3 / span 2; background:var(--bg); overflow:hidden; display:flex}
.home-scroll{flex:1; overflow-y:auto; padding:28px 32px 48px; max-width:1100px; margin:0 auto; width:100%}
.home-header{margin-bottom:22px}
.home-header h1{margin:0 0 6px; font-size:26px; font-weight:750; letter-spacing:-.02em}
.home-summary{margin:0; font-size:14px; color:var(--ink-2)}
.home-summary strong{color:var(--ink); font-weight:700}

.home-row{display:grid; gap:18px; margin-bottom:18px}
.home-row-top{grid-template-columns:1fr 1.3fr}
.home-row-main{grid-template-columns:1fr 1.25fr; align-items:start}

.widget{
  background:var(--panel); border:1px solid var(--line); border-radius:16px;
  padding:18px 20px; box-shadow:var(--shadow);
}
.widget-head{display:flex; align-items:center; justify-content:space-between; margin-bottom:14px}
.widget-head h3{margin:0; font-size:14px; font-weight:700; display:flex; align-items:center; gap:8px}
.widget-head h3 svg{color:var(--accent)}
.widget-sub{font-size:12px; color:var(--ink-3); font-weight:500}

/* clock */
.widget-clock{display:flex; flex-direction:column; justify-content:center; background:linear-gradient(135deg,#2b3556,#1b2138); border:none; color:#fff}
.clock-time{font-size:48px; font-weight:750; letter-spacing:-.03em; line-height:1; display:flex; align-items:baseline; gap:2px}
.clock-sec{font-size:24px; font-weight:600; color:rgba(255,255,255,.55)}
.clock-ampm{font-size:18px; font-weight:600; color:rgba(255,255,255,.7); margin-left:6px}
.clock-date{display:flex; align-items:center; gap:8px; margin-top:12px; font-size:13.5px; color:rgba(255,255,255,.8)}
.clock-date strong{color:#fff; font-weight:700}

/* weather */
.widget-weather{display:flex; flex-direction:column; gap:14px}
.weather-main{display:flex; flex-direction:column; gap:8px}
.weather-now{display:flex; align-items:center; gap:14px}
.weather-now svg{color:#E8912D}
.weather-temp{font-size:38px; font-weight:750; letter-spacing:-.02em; line-height:1}
.weather-cond{font-size:13px; color:var(--ink-2); margin-top:2px}
.weather-loc{display:flex; align-items:center; gap:5px; font-size:12.5px; color:var(--ink-3)}
.weather-hilo{margin-left:auto; font-weight:600; color:var(--ink-2)}
.weather-forecast{display:flex; gap:8px; border-top:1px solid var(--line-2); padding-top:12px}
.weather-day{flex:1; display:flex; flex-direction:column; align-items:center; gap:5px; font-size:12px}
.weather-day-name{color:var(--ink-3); font-weight:600}
.weather-day svg{color:#E8912D}
.weather-day-high{font-weight:700}

/* quick access */
.qa-grid{display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:8px}
.qa-item{
  display:flex; align-items:center; gap:11px; padding:10px 12px; border-radius:11px;
  border:1px solid var(--line); background:var(--panel); transition:.13s; text-align:left;
}
.qa-item:hover{border-color:var(--accent); background:var(--accent-soft)}
.qa-avatar{
  width:38px; height:38px; border-radius:10px; display:grid; place-items:center; color:#fff;
  font-weight:600; font-size:14px; position:relative; flex-shrink:0;
}
.qa-ws{border-radius:11px; font-weight:700; font-size:17px}
.qa-prov{
  position:absolute; bottom:-3px; right:-3px; width:17px; height:17px; border-radius:50%;
  display:grid; place-items:center; color:#fff; font-size:9px; font-weight:700; border:2px solid var(--panel);
}
.qa-text{display:flex; flex-direction:column; min-width:0; flex:1; gap:1px}
.qa-label{font-weight:650; font-size:13.5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
.qa-sub{font-size:11.5px; color:var(--ink-3); white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
.qa-badge{
  background:#ef4444; color:#fff; font-size:11px; font-weight:700; min-width:19px; height:19px;
  border-radius:10px; padding:0 5px; display:grid; place-items:center; flex-shrink:0;
}

/* calendar */
.cal-grid{display:grid; grid-template-columns:repeat(7,1fr); gap:2px; margin-bottom:16px}
.cal-dow{text-align:center; font-size:10.5px; font-weight:700; color:var(--ink-3); padding:2px 0 6px}
.cal-day{
  aspect-ratio:1; display:grid; place-items:center; font-size:12.5px; border-radius:8px;
  color:var(--ink-2); font-weight:500;
}
.cal-day.is-today{background:var(--accent); color:#fff; font-weight:700}
.cal-day.empty{color:transparent}
.cal-agenda{display:flex; flex-direction:column; gap:4px; border-top:1px solid var(--line-2); padding-top:12px}
.cal-agenda-label{font-size:10.5px; font-weight:700; color:var(--ink-3); text-transform:uppercase; letter-spacing:.05em; margin-top:6px}
.cal-agenda-label:first-child{margin-top:0}
.cal-event{display:flex; align-items:center; gap:9px; padding:5px 4px}
.cal-event-dot{width:8px; height:8px; border-radius:50%; flex-shrink:0}
.cal-event-title{font-size:13px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; min-width:0}
.cal-event-source{
  font-size:10.5px; font-weight:700; flex-shrink:0; padding:1px 6px; border-radius:6px;
  background:var(--bg); border:1px solid var(--line); white-space:nowrap;
}
.cal-event-time{font-size:11.5px; color:var(--ink-3); font-weight:500; flex-shrink:0; margin-left:auto}
.cal-legend{display:flex; flex-wrap:wrap; gap:12px; margin-bottom:14px; padding-bottom:12px; border-bottom:1px solid var(--line-2)}
.cal-legend-item{display:flex; align-items:center; gap:6px; font-size:11.5px; color:var(--ink-2); font-weight:600}
.cal-legend-dot{width:9px; height:9px; border-radius:3px; flex-shrink:0}
.cal-empty{font-size:12.5px; color:var(--ink-3); padding:4px 4px 2px; font-style:italic}

/* slack lists */
.widget-lists{display:flex; flex-direction:column; max-height:520px}
.lists-scroll{overflow-y:auto; display:flex; flex-direction:column; gap:14px; margin:0 -4px; padding:0 4px}
.list-card{border:1px solid var(--line); border-radius:12px; overflow:hidden}
.list-card-head{display:flex; align-items:center; gap:9px; padding:10px 12px; background:var(--bg); border-bottom:1px solid var(--line-2)}
.list-accent{width:4px; height:18px; border-radius:3px; flex-shrink:0}
.list-name{font-weight:700; font-size:13.5px}
.list-ws{font-size:11px; color:var(--ink-3)}
.list-progress{margin-left:auto; font-size:11.5px; font-weight:700; color:var(--ink-2); background:var(--line-2); padding:2px 8px; border-radius:10px}
.list-items{display:flex; flex-direction:column}
.list-item{display:flex; align-items:center; gap:10px; padding:9px 12px; border-bottom:1px solid var(--line-2)}
.list-item:last-child{border-bottom:none}
.list-check{flex-shrink:0; color:var(--ink-3)}
.list-item.done .list-check{color:#2EB67D}
.list-item.progress .list-check{color:#E8912D}
.list-item-title{flex:1; font-size:13px; color:var(--ink)}
.list-item.done .list-item-title{color:var(--ink-3); text-decoration:line-through}
.list-item-meta{display:flex; align-items:center; gap:10px; flex-shrink:0}
.list-assignee{font-size:11.5px; color:var(--ink-2); font-weight:600}
.list-due{font-size:11px; color:var(--ink-3); background:var(--bg); padding:2px 7px; border-radius:8px; border:1px solid var(--line)}

/* home button next to compose */
.compose-row{display:flex; gap:8px; margin:2px 0 16px}
.compose-row .compose{flex:1; margin:0}
.home-btn{
  width:42px; flex-shrink:0; display:grid; place-items:center; border-radius:10px;
  border:1px solid var(--line); color:var(--ink-2); background:var(--panel); transition:.13s;
}
.home-btn:hover{border-color:var(--accent); color:var(--accent); background:var(--accent-soft)}

/* slack sidebar home button */
.sl-ws-head-row{display:flex; align-items:center; gap:8px}
.sl-home-btn{
  margin-left:auto; width:30px; height:30px; border-radius:8px; display:grid; place-items:center;
  color:rgba(255,255,255,.8); transition:.13s; flex-shrink:0;
}
.sl-home-btn:hover{background:rgba(255,255,255,.15); color:#fff}

/* home sidebar */
.home-side{display:flex; flex-direction:column; height:100%}
.home-side-head{display:flex; align-items:center; gap:11px; padding:4px 4px 16px}
.home-side-logo{
  width:40px; height:40px; border-radius:11px; display:grid; place-items:center; color:#fff;
  background:linear-gradient(135deg,#4f6ef7,#3b5bdb); flex-shrink:0;
}
.home-side-head strong{display:block; font-size:15px; font-weight:700}
.home-side-sub{font-size:12px; color:var(--ink-3)}
.home-side-hint{font-size:13px; color:var(--ink-2); line-height:1.6; padding:0 4px}

@media (max-width:1080px){
  .home-row-top,.home-row-main{grid-template-columns:1fr}
}

/* ===== Microsoft Teams ===== */
.rail-tm-logo{border-radius:9px}

/* sidebar */
.tm-side{display:flex; flex-direction:column; height:100%; margin:-16px -12px}
.tm-head{padding:16px 16px 14px; background:linear-gradient(135deg,#7B83EB,#4B53BC); color:#fff}
.tm-head-row{display:flex; align-items:center; gap:8px}
.tm-org-name{
  display:inline-flex; align-items:center; gap:7px; color:#fff; font-weight:700;
  font-size:15.5px; letter-spacing:-.01em;
}
.tm-org-name svg{opacity:0; transition:.15s}
.tm-org-name:hover svg{opacity:.85}
.tm-home-btn{
  margin-left:auto; width:30px; height:30px; border-radius:8px; display:grid; place-items:center;
  color:rgba(255,255,255,.85); transition:.13s; flex-shrink:0;
}
.tm-home-btn:hover{background:rgba(255,255,255,.15); color:#fff}
.tm-you{font-size:11.5px; color:rgba(255,255,255,.75); display:block; margin-top:3px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}

.tm-teams{flex:1; overflow-y:auto; padding:14px 10px}
.tm-section{
  display:block; font-size:11px; font-weight:700; color:var(--ink-3); letter-spacing:.04em;
  text-transform:uppercase; padding:0 6px 10px;
}
.tm-team{margin-bottom:14px}
.tm-team-head{display:flex; align-items:center; gap:9px; padding:4px 6px}
.tm-team-avatar{
  width:28px; height:28px; border-radius:7px; display:grid; place-items:center; color:#fff;
  font-size:11px; font-weight:700; flex-shrink:0;
}
.tm-team-avatar.sm{width:32px; height:32px; font-size:12px}
.tm-team-name{font-weight:700; font-size:13.5px}
.tm-channels{display:flex; flex-direction:column; margin-top:2px}
.tm-chan{
  display:flex; align-items:center; gap:8px; padding:6px 8px 6px 43px; border-radius:7px;
  color:var(--ink-2); font-size:13.5px; transition:.1s; text-align:left; position:relative;
}
.tm-chan:hover{background:var(--line-2)}
.tm-chan.is-active{background:#ECEBFA; color:#4B53BC; font-weight:700}
.tm-chan.is-active::before{
  content:""; position:absolute; left:0; top:20%; bottom:20%; width:3px; border-radius:0 3px 3px 0;
  background:#6264A7;
}
.tm-chan.has-unread{color:var(--ink); font-weight:700}
.tm-chan.is-active.has-unread{color:#4B53BC}
.tm-chan-name{flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.tm-badge{
  background:#C4314B; color:#fff; font-size:11px; font-weight:700; min-width:18px; height:18px;
  border-radius:9px; padding:0 5px; display:grid; place-items:center; flex-shrink:0;
}

/* main posts pane — spans both right columns like the real Teams client */
.tm-main{grid-column:3 / span 2; background:#F5F5F5; display:flex; flex-direction:column; overflow:hidden}
.tm-chan-head{
  background:var(--panel); border-bottom:1px solid var(--line); padding:14px 26px 0;
  display:flex; flex-direction:column; gap:10px;
}
.tm-chan-title{display:flex; align-items:center; gap:12px}
.tm-chan-title h1{margin:0; font-size:18px; font-weight:750; letter-spacing:-.01em}
.tm-crumb{font-size:12px; color:var(--ink-3)}
.tm-tabs{display:flex; gap:22px}
.tm-tab{
  font-size:13px; font-weight:600; color:var(--ink-3); padding:6px 2px 10px; cursor:pointer;
  border-bottom:2.5px solid transparent;
}
.tm-tab:hover{color:var(--ink)}
.tm-tab.is-active{color:#4B53BC; border-bottom-color:#6264A7}

.tm-stream{flex:1; overflow-y:auto; padding:20px 26px; display:flex; flex-direction:column; gap:14px}
.tm-post{
  background:var(--panel); border:1px solid var(--line); border-radius:10px;
  box-shadow:0 1px 2px rgba(16,24,40,.05); overflow:hidden;
}
.tm-post-main{display:flex; gap:11px; padding:14px 16px 10px}
.tm-avatar{
  width:36px; height:36px; border-radius:50%; flex-shrink:0; display:grid; place-items:center;
  color:#fff; font-weight:600; font-size:13px;
}
.tm-avatar.sm{width:28px; height:28px; font-size:11px}
.tm-post-body{min-width:0; flex:1}
.tm-post-head{display:flex; align-items:baseline; gap:9px}
.tm-author{font-weight:700; font-size:13.5px}
.tm-time{font-size:11px; color:var(--ink-3)}
.tm-text{font-size:14px; line-height:1.55; color:#242424; margin-top:2px; white-space:pre-wrap; word-break:break-word}
.tm-reactions{margin-top:7px}

.tm-replies{
  border-top:1px solid var(--line-2); padding:10px 16px 6px 30px;
  display:flex; flex-direction:column; gap:10px; background:#FAFAFC;
}
.tm-reply{display:flex; gap:10px}

.tm-reply-box{border-top:1px solid var(--line-2); padding:8px 16px 10px 30px; background:#FAFAFC}
.tm-reply-toggle{
  display:inline-flex; align-items:center; gap:7px; font-size:13px; font-weight:600;
  color:var(--ink-2); padding:5px 10px; border-radius:7px; transition:.12s;
}
.tm-reply-toggle:hover{background:var(--line-2); color:#4B53BC}
.tm-reply-input{
  display:flex; align-items:flex-end; gap:8px; border:1px solid var(--line); border-radius:9px;
  padding:7px 7px 7px 12px; background:var(--panel);
}
.tm-reply-input:focus-within{border-color:#6264A7}
.tm-reply-input textarea{
  flex:1; border:none; outline:none; resize:none; font-family:inherit; font-size:13.5px;
  line-height:1.5; max-height:120px; padding:3px 0; background:none;
}
.tm-reply-input .sl-send{background:#6264A7}
.tm-reply-input .sl-send:hover:not(:disabled){background:#4B53BC}

.tm-composer{padding:12px 26px 18px; background:#F5F5F5; border-top:1px solid var(--line)}
.tm-composer .tm-reply-input.new-post{box-shadow:var(--shadow)}

/* chat list in sidebar */
.tm-chats{display:flex; flex-direction:column; gap:1px; margin-bottom:16px}
.tm-chat{
  display:flex; align-items:center; gap:10px; padding:7px 8px; border-radius:8px;
  transition:.1s; text-align:left; position:relative;
}
.tm-chat:hover{background:var(--line-2)}
.tm-chat.is-active{background:#ECEBFA}
.tm-chat.is-active::before{
  content:""; position:absolute; left:0; top:20%; bottom:20%; width:3px; border-radius:0 3px 3px 0;
  background:#6264A7;
}
.tm-chat-avatar{
  width:32px; height:32px; border-radius:50%; display:grid; place-items:center; color:#fff;
  font-size:12px; font-weight:600; flex-shrink:0;
}
.tm-chat-avatar.group{background:#6264A7}
.tm-chat-avatar.hd{width:36px; height:36px; font-size:13px}
.tm-chat-text{display:flex; flex-direction:column; min-width:0; flex:1; gap:1px}
.tm-chat-name{font-size:13.5px; font-weight:600; color:var(--ink); white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
.tm-chat.has-unread .tm-chat-name{font-weight:750}
.tm-chat.is-active .tm-chat-name{color:#4B53BC}
.tm-chat-preview{font-size:11.5px; color:var(--ink-3); white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
.tm-chat.has-unread .tm-chat-preview{color:var(--ink-2); font-weight:600}

/* chat bubble stream */
.tm-chat-stream{gap:16px; padding:22px 30px}
.tm-bubble-row{display:flex; gap:10px; max-width:70%}
.tm-bubble-row.mine{margin-left:auto; flex-direction:row-reverse}
.tm-bubble-col{display:flex; flex-direction:column; min-width:0}
.tm-bubble-row.mine .tm-bubble-col{align-items:flex-end}
.tm-bubble-meta{display:flex; align-items:baseline; gap:8px; margin-bottom:3px; padding:0 2px}
.tm-bubble{
  background:var(--panel); border:1px solid var(--line); border-radius:10px;
  padding:9px 13px; font-size:14px; line-height:1.5; color:#242424;
  white-space:pre-wrap; word-break:break-word; box-shadow:0 1px 2px rgba(16,24,40,.04);
}
.tm-bubble.mine{background:#E8EBFA; border-color:#D6DAF5}
.tm-bubble-row.mine .tm-reactions{justify-content:flex-end}
.tm-bubble-row.mine .sl-emoji-pop{left:auto; right:0}

.rail-group-gap{height:8px; flex-shrink:0}

/* email reply composer */
.rc{
  margin-top:20px; background:var(--panel); border:1px solid var(--line); border-radius:14px;
  box-shadow:var(--shadow); overflow:hidden;
}
.rc-head{
  display:flex; align-items:center; justify-content:space-between; padding:12px 16px;
  border-bottom:1px solid var(--line-2); font-size:14px;
}
.rc-close{color:var(--ink-3); padding:4px; border-radius:6px; display:grid; place-items:center}
.rc-close:hover{background:var(--line-2); color:var(--ink)}
.rc-field{display:flex; align-items:center; gap:10px; padding:9px 16px; border-bottom:1px solid var(--line-2)}
.rc-field label{font-size:12px; color:var(--ink-3); width:52px; flex-shrink:0}
.rc-field input{flex:1; border:none; outline:none; font-size:13.5px; font-family:inherit; background:none}
.rc-toolbar{display:flex; align-items:center; gap:2px; padding:7px 12px; border-bottom:1px solid var(--line-2); flex-wrap:wrap}
.rc-select{
  border:1px solid var(--line); border-radius:7px; padding:4px 6px; font-size:12px;
  font-family:inherit; color:var(--ink-2); background:var(--panel); outline:none; cursor:pointer;
  max-width:104px;
}
.rc-select:hover{border-color:var(--accent)}
.rc-select-size{max-width:84px}
.rc-sep{width:1px; height:18px; background:var(--line); margin:0 5px; flex-shrink:0}
.rc-color-wrap{position:relative; display:inline-flex}
.rc-color-btn{flex-direction:column; gap:0; line-height:1; font-weight:700; font-size:13px}
.rc-color-btn{display:grid}
.rc-color-bar{width:14px; height:3px; background:#B91C1C; border-radius:2px; margin-top:1px}
.rc-color-pop{
  position:absolute; top:calc(100% + 6px); left:0; z-index:31; background:var(--panel);
  border:1px solid var(--line); border-radius:10px; box-shadow:0 10px 30px rgba(16,24,40,.16);
  padding:8px; display:grid; grid-template-columns:repeat(5,1fr); gap:6px; animation:pop .14s;
}
.rc-swatch{width:22px; height:22px; border-radius:6px; border:1px solid rgba(0,0,0,.08); cursor:pointer}
.rc-swatch:hover{transform:scale(1.12)}
.rc-tool{
  min-width:30px; height:28px; border-radius:7px; display:grid; place-items:center;
  color:var(--ink-2); font-size:13px; transition:.1s; padding:0 7px;
}
.rc-tool:hover{background:var(--line-2); color:var(--ink)}
.rc-editor{
  min-height:140px; max-height:320px; overflow-y:auto; padding:14px 16px; font-size:14px;
  line-height:1.6; outline:none;
}
.rc-editor:focus{background:#fdfdfe}
.rc-editor ul,.rc-editor ol{padding-left:22px; margin:6px 0}
.rc-editor blockquote{
  border-left:3px solid var(--line); margin:8px 0; padding:2px 0 2px 12px; color:var(--ink-2);
}
.rc-error{margin:0 16px 10px}
.rc-foot{display:flex; gap:10px; padding:12px 16px; border-top:1px solid var(--line-2)}

/* apple connect form */
.apple-form{display:flex; flex-direction:column; gap:10px}
.apple-form-hint{margin:0; font-size:12.5px; color:var(--ink-2); line-height:1.55}
.apple-form input{
  border:1px solid var(--line); border-radius:9px; padding:10px 12px; font-size:13.5px;
  font-family:inherit; outline:none;
}
.apple-form input:focus{border-color:var(--accent)}
.apple-form-actions{display:flex; align-items:center; gap:12px; margin-top:2px}

/* ===== home empty states ===== */
.qa-cta{
  display:flex; flex-direction:column; align-items:center; text-align:center;
  padding:22px 16px; gap:5px;
}
.qa-cta-title{font-size:14.5px; font-weight:700; color:var(--ink)}
.qa-cta-sub{font-size:12.5px; color:var(--ink-3); margin-bottom:10px}
.qa-cta-btn{
  display:inline-flex; align-items:center; gap:7px; background:var(--accent); color:#fff;
  font-weight:650; font-size:13px; padding:10px 20px; border-radius:9px; transition:.15s;
}
.qa-cta-btn:hover{background:#324bc0}
.qa-add{
  display:flex; align-items:center; justify-content:center; gap:8px; padding:10px 12px;
  border-radius:11px; border:1.5px dashed var(--line); color:var(--ink-3); font-size:13px;
  font-weight:600; transition:.13s; background:none;
}
.qa-add:hover{border-color:var(--accent); color:var(--accent); background:var(--accent-soft)}

/* ===== auth + profile ===== */
.auth-overlay{
  position:fixed; inset:0; z-index:100; background:var(--bg);
  display:grid; place-items:center;
}
.auth-splash{font-size:15px; color:var(--ink-2); font-weight:600}
.auth-card{
  width:400px; max-width:92vw; background:var(--panel); border:1px solid var(--line);
  border-radius:18px; padding:34px 32px; box-shadow:0 12px 40px rgba(16,24,40,.1);
  display:flex; flex-direction:column; gap:11px;
}
.auth-brand{
  width:44px; height:44px; border-radius:13px; display:grid; place-items:center;
  background:linear-gradient(135deg,#4f6ef7,#3b5bdb); color:#fff; margin-bottom:4px;
}
.auth-card h1{margin:0; font-size:21px; font-weight:750; letter-spacing:-.015em}
.auth-sub{margin:0 0 8px; font-size:13.5px; color:var(--ink-2); line-height:1.5}
.auth-card input{
  border:1px solid var(--line); border-radius:10px; padding:11px 13px; font-size:14px;
  font-family:inherit; outline:none; width:100%;
}
.auth-card input:focus{border-color:var(--accent)}
.auth-row{display:flex; gap:10px}
.auth-error{
  background:#FEF2F2; border:1px solid #FECACA; color:#B91C1C; border-radius:9px;
  padding:9px 12px; font-size:13px;
}
.auth-submit{
  background:var(--accent); color:#fff; font-weight:650; font-size:14px; padding:12px;
  border-radius:10px; transition:.15s; margin-top:4px;
}
.auth-submit:hover:not(:disabled){background:#324bc0}
.auth-submit:disabled{opacity:.6; cursor:default}
.auth-submit.sm{padding:9px 18px; font-size:13px; margin-top:0; width:auto}
.auth-switch{color:var(--accent); font-size:13px; font-weight:600; padding:4px}
.auth-switch:hover{text-decoration:underline}

.prof-header{display:flex; align-items:center; justify-content:space-between}
.prof-back{
  display:inline-flex; align-items:center; gap:7px; font-size:13px; font-weight:600;
  color:var(--ink-2); border:1px solid var(--line); border-radius:9px; padding:8px 14px;
  background:var(--panel); transition:.13s;
}
.prof-back:hover{border-color:var(--accent); color:var(--accent)}
.prof-card{margin-bottom:18px}
.prof-grid{display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:14px}
.prof-grid label{
  display:flex; flex-direction:column; gap:6px; font-size:12px; font-weight:650; color:var(--ink-2);
}
.prof-grid input{
  border:1px solid var(--line); border-radius:9px; padding:10px 12px; font-size:14px;
  font-family:inherit; outline:none;
}
.prof-grid input:focus{border-color:var(--accent)}
.prof-wide{grid-column:1 / -1}
.prof-actions{display:flex; align-items:center; gap:12px}
.prof-saved{color:#15803D; font-size:13px; font-weight:650}
.prof-accounts{display:flex; flex-direction:column; gap:10px}
.prof-acc{display:flex; align-items:center; gap:12px; padding:10px 12px; border:1px solid var(--line); border-radius:11px}
.prof-acc-text{display:flex; flex-direction:column; min-width:0; flex:1; gap:1px}
.prof-label-row{display:inline-flex; align-items:center; gap:6px}
.prof-label-edit{
  color:var(--ink-3); display:grid; place-items:center; padding:3px; border-radius:5px; opacity:.6;
}
.prof-label-edit:hover{background:var(--line-2); color:var(--ink); opacity:1}
.prof-disconnect{
  font-size:12.5px; font-weight:650; color:#B91C1C; border:1px solid #FECACA; border-radius:8px;
  padding:7px 13px; background:#FEF2F2; transition:.13s; flex-shrink:0;
}
.prof-disconnect:hover{background:#FEE2E2}
.prof-logout{
  font-size:13.5px; font-weight:650; color:var(--ink-2); border:1px solid var(--line);
  border-radius:10px; padding:10px 20px; background:var(--panel); transition:.13s;
}
.prof-logout:hover{border-color:#B91C1C; color:#B91C1C}

/* ===== responsive: narrow viewports =====
   Lives at the END of the stylesheet so it wins the cascade over pane rules
   (.sl-main, .tm-main, .home) declared above. At ≤900px the grid drops to
   three columns: rail + folder/channel sidebar + one content pane. */
@media (max-width:900px){
  .md-root{grid-template-columns:64px 200px 1fr}
  .md-root.rail-open{grid-template-columns:248px 200px 1fr}
  /* email: keep the message list, hide the reading pane */
  .reader{display:none !important}
  .list{grid-column:3}
  /* slack: stream fills the content column (thread pane is a .reader, hidden above) */
  .sl-main{grid-column:3}
  /* teams + home collapse their two-column span to the single content column */
  .tm-main,.home{grid-column:3 / span 1}
}
`;
