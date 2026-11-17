# PREREQUIREMENTS — what must be true before `xreg.py` can run

Verified against the working install in this repo. Anything marked **hard** stops
the run; anything marked **soft** degrades it.

---

## 1. Host

| Requirement | Status | Notes |
|---|---|---|
| Linux x86_64 | **hard** | tested on `Linux 7.0.0 x86_64`, glibc 2.39 |
| Python ≥ 3.12 | **hard** | repo venv is 3.12.14 |
| `uv` on PATH | **soft** | only needed to build the venv; plain `pip` works too |
| A display server (`DISPLAY`) | **hard** for default runs | headless gets 403 from X; a captcha may need a human |
| ~2.4 GB free disk | **hard** | Camoufox browser payload |
| Network egress to `x.com`, `login.microsoftonline.com`, `outlook.office365.com:993` | **hard** | |

---

## 2. Python packages

Install into the repo venv (`.venv/`):

```bash
uv venv .venv
uv pip install --python .venv/bin/python camoufox playwright
```

| Package | Version installed | Why |
|---|---|---|
| `camoufox` | v0.5.9 | patched Firefox driver — primary engine |
| `fpgen` | v1.3.0 | camoufox dependency (fingerprint generation) |
| `playwright` | v1.62.0 | fallback engine + camoufox's transport |
| stdlib only | — | IMAP, OAuth2, MIME parsing, argparse |

`camoufox` pins `playwright<1.63`; do not force a newer Playwright or camoufox
breaks.

### Browser binary

The Camoufox **browser** is a separate download from the Python package:

```bash
.venv/bin/python -m camoufox fetch     # required once
.venv/bin/python -m camoufox version   # verify: "Installed  Yes"
```

| Item | Value |
|---|---|
| Browser | `v158.0.1-beta.41` (channel `official/stable`) |
| Install path | `~/.cache/camoufox` (~2.4 GB) |
| GeoIP extra | **not installed** — the script passes `geoip=False`, so this is fine |

The Playwright fallback (`--engine chromium`) needs its own Chromium:
`.venv/bin/python -m playwright install chromium`. Not required when using
Camoufox, which is the default.

---

## 3. Credentials — `account.txt`

One line, pipe-separated, `chmod 600`:

```
email|password|msa_artifact|client_id
```

| Field | Requirement |
|---|---|
| `email` | a real Outlook.com mailbox |
| `password` | reused as the X account password |
| `msa_artifact` | MSA artifact (`M.C…`), which **is** an OAuth refresh token |
| `client_id` | 36-char Azure app id; `9e5f94bc-e8a4-4e73-b8be-63364c29d753` (Thunderbird) is pre-approved |

Overridable per field via `OUTLOOK_EMAIL`, `X_PASSWORD`, `OUTLOOK_MSA`,
`OUTLOOK_CLIENT_ID`.

### Mailbox must permit IMAP

Outlook.com → Settings → Mail → Forwarding and IMAP → enable **“Let devices and
apps use IMAP.”** Without it, auth fails with `AUTHENTICATE failed`.

### Token cache

`~/.cache/outlook_newest/<email>.json`, mode 600. Holds a live rotating refresh
token. **Never delete it** unless you intend to redo the device-code login
(`--mail-login`).

---

## 4. Proxies — `PROXIES.txt` (optional)

Only needed when direct egress is blocked or rate-limited.

- One URL per line: `http://user:pass@host:port`
- File is **CRLF**; `load_proxies()` strips `\r`. A stray CR makes both Playwright
  and curl reject the URL as malformed
- `--proxy ''` forces direct, bypassing the file entirely

---

## 5. What cannot be provisioned

| Gate | Automatable? |
|---|---|
| X face-liveness check (`#/s/face_liveness`) | **no** — physical scan; exit `6` |
| Arkose/FunCaptcha | **no** — run headed and solve by hand |
| X's per-IP signup reputation | **partly** — proxy rotation helps, not guarantees |

---

## 6. Smoke test

Proves every prerequisite without side effects:

```bash
.venv/bin/python xreg.py --self-check    # offline helpers        → "self-check OK"
.venv/bin/python xreg.py --mail-check    # credentials + IMAP     → "[mail-check] ok"
.venv/bin/python xreg.py --dry-run       # display + browser      → "email field: True"
```

All three exit `0` means the host is ready for a real run. `--mail-check` prints a
real verification code, so treat its output as sensitive.

This smoke test says nothing about which flow to run: it never contacts X with
the account. **Unregistered mailbox → `xreg.py` (signup). Already registered →
`xreg.py login`.** See README §1.
---

## 7. Sessions (login only)

`xreg.py login --save-session` writes `.session/<localpart>.json`, a live cookie
jar (`chmod 600`, never committed). `xreg.py check` reads it and exits `0`
(still signed in) or `9` (expired). No session file → `3`.

One mailbox in `account.txt` means one flow per run: signup creates, login signs
in. `check` is the cheap follow-up to login, not a substitute for either.
