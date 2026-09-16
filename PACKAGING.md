# Packaging as a Windows .exe

`desktop_launcher.py` runs the real Streamlit server in-process (via
`streamlit.web.bootstrap`) and opens your default browser to it — no
`streamlit run` command, no separate Python install needed by whoever runs
the .exe. PyInstaller bundles a full CPython + all dependencies, so the app
keeps **full, normal filesystem access** — pointing it at any folder for the
Excel database works exactly like running `app.py` directly.

## Build it yourself

```bash
cd streamlit_app
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\pip install pyinstaller pyinstaller-hooks-contrib

.venv\Scripts\python -m PyInstaller desktop_launcher.py ^
  --name CITE-Internship-Portal ^
  --onedir --noconfirm --console ^
  --add-data "app.py;." ^
  --collect-all streamlit ^
  --collect-all altair ^
  --collect-all pandas ^
  --collect-all openpyxl
```

This produces `dist/CITE-Internship-Portal/` — a folder containing
`CITE-Internship-Portal.exe` plus an `_internal` folder with everything it
needs. **Distribute the whole folder** (zip it up), not just the .exe —
`_internal` has to travel with it.

On a rebuild after changing `app.py` only (no new imports), it's faster to
reuse the generated `.spec` file instead of retyping the flags:

```bash
.venv\Scripts\python -m PyInstaller CITE-Internship-Portal.spec --noconfirm
```

First build takes 10-20 minutes (it's collecting Streamlit's frontend
assets, Altair's Vega schemas, pandas, etc.) and produces a ~230-250MB
folder — that's normal for a bundled Python + web-framework app. Rebuilds
from the `.spec` are quicker.

## Running the built app

Double-click `CITE-Internship-Portal.exe`. A console window opens (shows
server logs — closing it stops the app) and your browser opens to the app
automatically. The sidebar's **Data folder** defaults to a `data` folder
created next to the .exe, and remembers whatever folder you point it at
next (written to `.cite_last_datadir.txt`, also next to the .exe) — so
non-technical users can just point it at a shared/synced folder once and
it'll stick.

To pre-set the data folder without touching the UI (e.g. for a scripted
install), set `CITE_DATA_DIR` before launching, for example in a `.bat`
wrapper:

```bat
@echo off
set CITE_DATA_DIR=%~dp0data
start "" "%~dp0CITE-Internship-Portal.exe"
```

## Known trade-offs

- **Size.** ~230-250MB unzipped. This is the cost of bundling a full Python
  + Streamlit + pandas + altair runtime so nobody needs to install anything.
  Using `--onefile` instead of `--onedir` makes a single .exe but it
  re-extracts everything to a temp folder on *every* launch, which is
  noticeably slower to start — `--onedir` (a folder, double-click the exe
  inside it) is the better trade for an app people open often.
- **Antivirus / SmartScreen.** Unsigned PyInstaller executables commonly
  trigger a Windows SmartScreen warning ("Windows protected your PC") on
  first run for people you send it to, and can occasionally be flagged by
  antivirus software — this is a known PyInstaller false-positive pattern,
  not a sign anything is wrong. Code-signing the exe (a paid certificate)
  removes the SmartScreen prompt if this becomes a real distribution
  concern.
- **First launch is slower** than later ones (a few seconds) — Streamlit is
  spinning up a real local web server before your browser can load it.
