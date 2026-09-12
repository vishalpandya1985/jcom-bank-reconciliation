# RTGS/NEFT Reconciliation — Web App

A self-hosted web app for reconciling your HDFC nodal account statement
against the three internal GL suspense heads (3493 outward RTGS/NEFT, 3496
inward RTGS/NEFT, 345051 IMPS/UPI/POS/ACH/NACH/Returns).

- Login, so every run is tied to the person who ran it
- Upload the 4 CSVs in the browser, get results instantly
- Full history of every past run, downloadable as Excel at any time
- Admin can add teammates from a Users page (no public signup after the first account)

This runs on one computer (or a small server) on your office network. Everyone
on the team opens it in their browser using that computer's address — there's
no external hosting or cloud account involved, and your bank data never
leaves your network.

## 1. Requirements

- Python 3.9 or newer, installed on the computer that will run the server
- The 4 CSV files, same format as always (HDFC statement + GL 3493/3496/345051)

## 2. First-time setup

Open a terminal in this folder and run:

```bash
pip install -r requirements.txt
python3 app.py
```

You should see:

```
Uvicorn running on http://0.0.0.0:8000
```

Leave this window open — it's the server. On the SAME computer, open a
browser to:

```
http://localhost:8000
```

The very first visit asks you to create the admin account (your name,
a username, a password). That account can later add the rest of the team
from the "Team & Users" page — nobody else can self-register.

## 3. Letting your team access it from their own computers

1. Find this computer's local network IP address:
   - Windows: open Command Prompt, run `ipconfig`, look for "IPv4 Address" (e.g. `192.168.1.42`)
   - Mac/Linux: run `ifconfig` or `ip addr`, look for something like `192.168.1.42`
2. Make sure this computer stays on and connected to the office network/Wi-Fi
   while the server is running.
3. Teammates open, in their own browser, on the same network:
   ```
   http://192.168.1.42:8000
   ```
   (using the actual IP you found in step 1)
4. They log in with the account you create for them under "Team & Users".

If Windows Firewall blocks the connection the first time, allow Python
through the firewall when prompted (or allow inbound connections on port
8000).

## 4. Keeping it running

- The server needs to stay running for people to use it. For daily use,
  just leave the terminal window open on one office computer.
- To run it in the background so it survives a closed terminal (Linux/Mac):
  ```bash
  nohup python3 app.py > server.log 2>&1 &
  ```
- To stop it, close the terminal window (or `Ctrl+C`), or find and stop the
  process if running in the background.
- To run it automatically every time that computer starts, ask whoever
  manages that machine to set it up as a scheduled task (Windows Task
  Scheduler) or a systemd service (Linux) — happy to write that config too
  if useful.

## 5. Where your data lives

Everything is stored locally on this computer, in the `data/` folder:

- `data/recon.db` — SQLite database: user accounts (hashed passwords only)
  and the history/metadata of every run
- `reports/` — the generated Excel report for every run (so History downloads
  keep working)
- Uploaded CSVs are deleted immediately after each run finishes — only the
  generated Excel report and a summary are kept

**Back up the `data/` and `reports/` folders periodically** if you want to
keep your history — there's no cloud backup unless you add one yourself.

## 6. Changing the matching rules

The actual reconciliation logic lives in `reconcile_engine.py` (matching
rules) and `report_builder.py` (how the Excel report is laid out). These are
the same rules already validated against your real August data — reference
number matching, embedded transaction codes, batch-settlement labels, and
beneficiary-name matching, with duplicate-amount groups resolved in
chronological order. Anything genuinely ambiguous is left "Pending" rather
than guessed.

If your file formats ever change (new bank, new GL export layout), this file
is where to adjust the column parsing.

## 7. Security notes

- Passwords are stored as bcrypt hashes, never in plain text.
- Sessions are signed cookies; the signing key is auto-generated into
  `data/secret.key` on first run — don't share or commit that file.
- This app has no built-in HTTPS. That's fine for a private office LAN;
  if you ever expose it beyond your own network, put it behind a reverse
  proxy (e.g. Caddy or nginx) with HTTPS first.
