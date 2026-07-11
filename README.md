# ShipX Billing — hosted team app

A web app where your team logs in and works from **shared rate cards and billing history that update live** for everyone. Built as static files (no build step) on top of **Supabase** (database + login + realtime) and deployed to **Vercel**.

Files:
- `index.html` — login screen + app shell
- `app.js` — the full billing engine (all 9 services) + Supabase data layer
- `styles.css` — styling
- `config.js` — **you paste your Supabase keys here**
- `supabase-schema.sql` — database tables + security, paste into Supabase once

---

## Setup (about 15 minutes, one time)

### 1. Create the Supabase project
1. Go to **supabase.com** → sign up → **New project**. Pick a name, a strong database password, a region near Singapore. Wait ~2 min for it to provision.
2. Left sidebar → **SQL Editor** → **New query**. Open `supabase-schema.sql`, copy everything, paste, click **Run**. You should see "Success".
3. Left sidebar → **Authentication → Providers → Email**: make sure it's **Enabled**, and turn **OFF** "Allow new users to sign up" (so only people you invite can log in).
4. **Authentication → Users → Add user** (or **Invite**): add yourself and each teammate (email + a password, or send an invite). These are the only people who can sign in.

### 2. Put your keys in `config.js`
1. Supabase → **Project Settings → API**.
2. Copy **Project URL** and the **anon public** key.
3. Open `config.js` and paste them in:
   ```js
   window.SHIPX_CONFIG = {
     SUPABASE_URL:      "https://xxxxxxxx.supabase.co",
     SUPABASE_ANON_KEY: "eyJhbGciOi....(long key)...."
   };
   ```
   (The anon key is safe to ship in the browser — the database rules block anyone who isn't logged in.)

### 3. Deploy to Vercel
This machine has no Node/CLI, so use the web flow:
1. Create a **GitHub** account (if needed) → **New repository** → name it e.g. `shipx-billing`.
2. On the repo page → **Add file → Upload files** → drag in `index.html`, `app.js`, `styles.css`, `config.js` (the whole `shipx-web` folder contents) → **Commit**.
3. Go to **vercel.com** → sign in with GitHub → **Add New → Project** → **Import** your `shipx-billing` repo → **Deploy**. No settings needed (it's static).
4. Vercel gives you a URL like `https://shipx-billing.vercel.app`. Share it with the team.

**Simpler alternative (no GitHub):** go to **app.netlify.com/drop** and drag the `shipx-web` folder onto the page — it deploys instantly and gives you a URL. (Netlify instead of Vercel, but identical result.)

> To change rates/logic later: edit the files and re-upload to GitHub (Vercel redeploys automatically), or re-drag the folder onto Netlify.

---

## How the team uses it
1. Open the URL → **sign in** with the email/password you set in Supabase.
2. Pick a service in the sidebar, check/adjust the **rate card**, **upload** the input file(s), **review**, **download** the billing (and reconciliation), and **Save to records**.
3. Everything is **shared and live**: if one person edits a rate or saves a run, everyone else sees it within about a second — no refresh. The **Dashboard** rolls up all saved runs by month/quarter/year/customer.

## Notes
- Your uploaded input files are read in the browser; only the **computed results** (billing lines, totals) are stored in Supabase.
- To remove someone's access, delete their user in Supabase → Authentication → Users.
- The offline single-file version (`../ShipX Billing Tool.html`) still works for solo use with no internet.
