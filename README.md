# CITE Internship Records — Streamlit edition

A rebuild of `CITE-Internship-Portal.html` as a Streamlit web app. The core
logic (duration/status calculation, batch/year derivation, per-student week
totals, merge-on-import, manual week entry, refined dedup) is ported from
the newer HTML design — see the top-of-file docstring in `app.py` for the
section map.

## Two storage backends, same app

The app runs in one of two modes, decided automatically by whether Supabase
credentials are configured:

- **Local mode (default)** — one Excel workbook *is* the database, read and
  written directly, in a folder you choose. This is what the desktop `.exe`
  build (see `PACKAGING.md`) uses. Good for one person, or a folder on a
  shared network drive.
- **Cloud mode** — when `SUPABASE_URL` / `SUPABASE_ANON_KEY` are set (see
  below), every session reads and writes the same shared [Supabase](https://supabase.com)
  database instead. This is what you want when a few people need to see and
  edit the same live data from a shared URL — e.g. the Streamlit Community
  Cloud deployment. The app re-checks the cloud on every interaction, so
  changes made by a teammate show up as soon as you click anything (or hit
  the sidebar's **Refresh now**).

## What's different from the original HTML file

| | Original (.html) | This (Streamlit) |
|---|---|---|
| Data storage | Browser localStorage, per device | Excel workbook (local mode) **or** a shared Supabase database (cloud mode) |
| Who can view it | Whoever has the file | Whoever can reach the running app (server/network access) |
| Read vs. edit | License key + admin password | Admin password only (see below) |
| Runs via | Double-click the .html file | `streamlit run app.py` (a server process) |

The license-key gate from the original was dropped — it existed to control
who could activate copies of a *distributed file*; it doesn't map onto a
centrally hosted app (anyone who can reach the URL already sees it). Control
who can reach this app at the network/deployment level instead. The
**admin password** stays, and still gates add/edit/delete/import/settings —
read-only browsing needs no login, matching the original. In cloud mode the
admin password is shared (stored in Supabase), so any admin's password
change applies to the whole team.

## Local mode: the workbook is the database

Every change (add, edit, delete, import) is written straight to one
`.xlsx` file with two sheets:

- **Internships** — one row per record, same columns as the original
  export/template, plus an internal `ID` column.
- **Settings** — owner/contact info, the admin password hash, and the
  "show full CNIC" flag.

You can open this file directly in Excel to inspect or bulk-edit it — just
**close it in Excel before editing in the app again** (Windows locks the
file while Excel has it open, and the app's write will fail with a clear
error until you close it).

## Cloud mode: shared Supabase database (2-3 people syncing)

1. Create a free project at [supabase.com](https://supabase.com).
2. In the project's **SQL Editor**, run:

   ```sql
   create table internships (
     id text primary key,
     name text, roll text, batch text, cnic text,
     semester text, org text,
     start_date text, end_date text,
     manual_weeks numeric,
     cert_on_file boolean default false,
     cert_link text
   );
   alter table internships disable row level security;

   create table app_settings (
     id text primary key default 'default',
     owner_name text, owner_title text, owner_inst text,
     owner_email text, owner_phone text,
     admin_hash text, reveal_cnic boolean default false
   );
   alter table app_settings disable row level security;
   ```

   (RLS is disabled because the anon key lives only in Streamlit's
   server-side secrets — it's never sent to anyone's browser, since Streamlit
   itself makes the Supabase calls. That's actually more private than the
   original HTML file, which embeds the key in client-side JS.)

3. In **Project Settings → API**, copy the **Project URL** and the **anon
   public** key.
4. Locally: copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`
   and paste them in. On Streamlit Community Cloud: paste them into the app's
   **Secrets** box (see Deployment below) instead — never commit the real
   `secrets.toml`.

With those two values set, the app switches to cloud mode automatically —
no code changes needed.

## Deploying (GitHub + Streamlit Community Cloud)

1. Push this `streamlit_app` folder to a GitHub repo (see the repo root for
   the exact commands used).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, **New app**, pick the repo/branch, and set the main file path to
   `app.py` (or `streamlit_app/app.py` if this folder isn't the repo root).
3. Before (or after) the first deploy, open the app's **Settings → Secrets**
   in the Streamlit Cloud dashboard and paste in `SUPABASE_URL` and
   `SUPABASE_ANON_KEY` (see above) so the deployed app runs in cloud mode —
   otherwise it'll try to use a local workbook path that doesn't persist
   between container restarts on Cloud, and everyone's changes would be lost
   whenever the app sleeps/redeploys.
4. Share the resulting `*.streamlit.app` URL with your 2-3 teammates. First
   admin to log in should set a real admin password in Settings (shared —
   applies for everyone) and update the owner/contact info.

## Running it

```bash
cd streamlit_app
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt      # Windows
# source .venv/bin/activate && pip install -r requirements.txt   # macOS/Linux
.venv\Scripts\streamlit run app.py
```

The first time it opens, the sidebar's **Data folder** defaults to
`streamlit_app/data`. Type a different folder path and click **Use this
folder** to point it anywhere else (a shared network drive, a Dropbox/OneDrive
folder, etc.) — it remembers your last choice for next time. You can also
set it via an environment variable before launch:

```bash
set CITE_DATA_DIR=D:\CITE\data   # Windows
streamlit run app.py
```

**Default admin password:** `cite-admin` — change it immediately from
**Settings** (sidebar, admin only) once you've logged in.

## Notes / honest limitations

- **No conflict resolution.** If two people save changes to the *same*
  record at nearly the same instant, the second write wins (local mode:
  whole-workbook rewrite; cloud mode: last upsert wins). Fine for a small
  team editing different records; not built for high-concurrency use.
- **Printing** relies on your browser's own Print (Ctrl/Cmd+P) rather than a
  custom print view.
- The first-run "how do you want to view records" chooser from the original
  was simplified to just the always-visible scope bar (All / By batch / By
  internship year) at the top of the page.
- The original HTML's multi-dataset **Download** dialog (records / students /
  by-batch / by-year as separate files) wasn't ported — Export CSV/Excel here
  always exports the full records table; use the **By student** view plus
  Export for a student-level look, or filter/scope first.

## Packaging as a Windows .exe

See `PACKAGING.md` in this same folder — same `app.py`, bundled with a real
Python + Streamlit runtime so it still gets full filesystem access to
whatever data folder you point it at. (The desktop build always uses local
mode, even if you've configured cloud secrets elsewhere — there's no secrets
store to read from on someone else's machine unless you deliberately embed
one.)
