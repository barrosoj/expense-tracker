# Quadro de Tarefas / Expense Tracker

Offline-first web app (tasks, meetings, projects, expenses). Data lives in the browser (localStorage + IndexedDB) — nothing is sent to a server.

## Structure

```
index.html            markup
styles.css            styles (light/dark)
app.js                app logic
sw.js                 service worker (offline cache)
manifest.webmanifest  install-to-home-screen metadata
icon.png              app icon
```

## Login + database (Supabase)

Without configuration the app runs in local-only mode (no login). To enable login (email code + Google) and the database:

```sh
./scripts/setup-supabase.sh
```

The wizard walks through: create project (EU) → keys into `config.js` → run `supabase/schema.sql` → redirect URLs → Google OAuth (the 2 users as Google test users) → disable the Email provider → first login of both users (creates their accounts) → disable sign-ups.

Login is **Google only** (`emailLogin: false` in `config.js`), so no custom SMTP is needed. Google's Testing mode acts as the allow-list: only the listed test users can sign in. To re-enable the email-code form, set `emailLogin: true`, enable the Email provider, configure custom SMTP, and paste `supabase/email-template-otp.html` into the "Magic link or OTP" template.

Each user has separate data, both in the database (Row Level Security) and on the device (local data is stored per user ID). Data saved before login existed is offered once, to the first user who signs in on that device. Answers are remembered in `.env.supabase` (gitignored), so it's safe to re-run.

| File | Purpose |
|---|---|
| `config.js` | Supabase URL + publishable key (public by design; safe to commit) |
| `auth.js` | Login screen: email code + OAuth providers |
| `supabase/schema.sql` | `records` table, Row Level Security, `expenses_v` view, private `documents` bucket |
| `supabase/email-template-otp.html` | Login email with the one-time code |
| `vendor/supabase.js` | supabase-js 2.119.0 UMD build (vendored for offline use) |

Never put the `secret` / `service_role` key in this project.

## Run locally

```sh
python3 -m http.server 8000   # then open http://localhost:8000
```

## Deploy (GitHub Pages)

Push to GitHub → Settings → Pages → Source: `main` branch, `/ (root)`.

After changing `index.html`, `styles.css` or `app.js`, bump `CACHE_VERSION` in `sw.js` so installed copies pick up the update.

## Moving data from the old single-file version

Data is stored per address, so the hosted app starts empty:

1. Old file → **Exportar cópia** (saves a `.json`).
2. Hosted app → **Importar cópia** → pick that `.json`.

Project documents are not included in the backup and must be re-attached.
