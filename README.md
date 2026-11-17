# xreg — X account registration

One script, end to end: read the mailbox in `account.txt`, drive X's wizard in a
real browser, pull the verification code out of Outlook over IMAP, and walk the
flow to `x.com/home` — either registering a new account (**signup**) or signing
in to one that already exists (**login**).

Merged from a Playwright browser stub and `GRABBER_MAIL/outlook_newest.py`
(Outlook IMAP + OAuth2), then hardened against X's React onboarding wizard.

```bash
python3 xreg.py                 # signup: register a NEW account (default flow)
python3 xreg.py login           # login: sign in to an account that already exists
python3 xreg.py check           # verify a saved session is still valid
python3 xreg.py --dry-run       # just prove the signup page loads
python3 xreg.py --self-check    # offline helper checks, no network
python3 xreg.py --mail-check    # just prove the mailbox is readable
```

### Signup or login? — pick by what state the mailbox is in

The tool has two flows and they are **not interchangeable**. The mailbox in
`account.txt` decides which one applies:

| State of the account | Flow | Command |
|---|---|---|
| Never registered on X | **signup** | `xreg.py` (default) |
| Already registered | **login** | `xreg.py login` |
| Registered and you saved a session | **check** | `xreg.py check` |

Running signup on an already-registered mailbox fails: X refuses the address at
`email_entry` and the run exits `2` (`stage=blocked`). Running login on a mailbox
that was never registered fails the other way — the address is unknown to X and
no code is ever mailed. Neither flow can detect the other's
state for you, so **you** must know which account you are holding.

Rule of thumb: after a successful signup, that mailbox is a *login* mailbox from
then on. Re-registering it is never the right move; use `login` (and
`login --save-session` once) instead.

---

## 1. What it actually does

```
account.txt ──► x.com/i/flow/signup ──► code mailed ──► typed back ──► x.com/home
     │                  │                    ▲
     │                  └── Camoufox ────────┘
     └──> Outlook IMAP + XOAUTH2 (reads the code)
```

The wizard stages it walks, in order:

| Stage | What X asks | What the script does |
|---|---|---|
| `email_entry` | email or username | types the mailbox, presses Enter |
| `code` | 6-digit verification code | polls the inbox, types the code back |
| `dob` | birthday | random date in 1990–2005 |
| `setup` | name, username, password | name = mailbox name without digits, username = 5 chars + 4 digits |
| `preferences` | email discoverability toggle | turns it **off**, then clicks Continue |
| `optional` | optional prompts: follow suggestions, notifications | clicks **Skip**, **Skip for now**, or **Not now** — whichever the pane offers; repeats |
| `username` | handle (sometimes alone) | reuses the pinned handle |
| `done` | — | verifies a signed-in session at `x.com/home` |
| `liveness` | face scan | **stops** — no browser can pass this |
| `captcha` | Arkose challenge | stops headless, waits for a human otherwise |

Each stage is detected from what is on screen (`classify()`), so the flow adapts
instead of following a fixed script. An unrecognised pane is dumped to
`x_unknown.html` and named in the log rather than guessed at.

### The login flow

`xreg.py login` is **not** the mirror of signup. X serves
email → (username prompt) → password → `/home`, and mails a code only when it
decides it wants one. Password, code, and face-liveness are all conditional
branches handled if they appear and skipped if they do not.

| Stage | What X asks | What the script does |
|---|---|---|
| `knowledge_check` | username or phone for the account | clicks **Use password** (the route that always exists) |
| `password_create` | password | types the password from `account.txt` |
| `code` | 6-digit code | polls the inbox, types it back |
| `done` | — | verifies a signed-in session at `x.com/home` |
| `liveness` | face scan | **stops**, exit `6` |
| `captcha` | Arkose challenge | waits for a human (headed only) |
| `blocked` | X refused the attempt | exit `2` |

Useful flags, login only:

| Flag | Effect |
|---|---|
| `--save-session` | write the cookie jar to `.session/<localpart>.json` so `check` can reuse it |
| `--retries N` | restart the login on `2`, `6`, `8` (default 3) |
| `--code XXXXXX` | supply the code yourself instead of polling the mailbox |

The session file is a **live credential** (`chmod 600`); it is never written
unless you pass `--save-session`. `xreg.py check` reads it, opens `x.com/home`,
and reports `0` (still signed in) or `9` (expired).

---

## 2. Setup

Full environment requirements are in **[PREREQUIREMENTS.md](PREREQUIREMENTS.md)** —
host, package versions, browser binary, credentials, IMAP, and what cannot be
automated. The short version:

```bash
uv venv .venv
uv pip install --python .venv/bin/python camoufox playwright
.venv/bin/python -m camoufox fetch
```

`camoufox fetch` is required — the patched Firefox binary (~2.4 GB, into
`~/.cache/camoufox`) is separate from the Python package. Verify with
`.venv/bin/python -m camoufox version`.

### Credentials

`account.txt`, one line, pipe-separated:

```
email|password|msa_artifact|client_id
```

| Field | Meaning |
|---|---|
| `email` | the mailbox being registered |
| `password` | reused as the X account password |
| `msa_artifact` | Outlook MSA artifact; **is** an OAuth refresh token |
| `client_id` | Azure application id. `9e5f94bc-e8a4-4e73-b8be-63364c29d753` is Thunderbird's public client, pre-approved for `IMAP.AccessAsUser.All` |

```bash
chmod 600 account.txt
```

Environment variables override the file **per field** — setting only
`OUTLOOK_EMAIL` does not discard the file's other values:

| Variable | Overrides |
|---|---|
| `OUTLOOK_EMAIL` | field 1 |
| `X_PASSWORD` | field 2 |
| `OUTLOOK_MSA` | field 3 |
| `OUTLOOK_CLIENT_ID` | field 4 |

Never commit `account.txt` or `.env`. The refresh token cache at
`~/.cache/outlook_newest/<email>.json` is a live credential (`chmod 600`).

---

## 3. Why these choices

**Camoufox, not plain Playwright.** Headless Chromium gets HTTP 403 from `x.com`
behind Cloudflare; the full Chromium channel and Camoufox both get through. So
Camoufox is primary (`--engine camoufox`) with Chromium as fallback
(`--engine chromium`). This is empirical, not a guarantee — a 403 is also
possible from a bad IP, which is what proxy rotation handles.

**The DOB pane has no `<select>`.** Month/day/year are custom pills
(`role=button` / `aria-haspopup`). Their popover has no `role=listbox`, so option
rows are matched by **text**, and the trigger is re-read afterwards to confirm the
value took.

**Continue is `<div role="button" aria-disabled>`.** It has no `disabled`
attribute. Clicking while gated is a silent no-op, which is what makes the pane
look hung. `submit()` polls until the gate clears, then clicks.

**`classify()` order matters.** The setup pane contains a password field, so a
password probe would shadow it. Stage detection is ordered most-specific first.

**Display names cannot contain digits.** `OgdenIntrieri341@outlook.com` becomes
`Ogden Intrieri`. The mailbox local part is stripped of digits, split on camel
case, and capped at 50 chars.

**codes are matched to X only.** Only mail from `x.com`/`twitter.com`, addressed
to this mailbox, with "code"/"confirm" in the subject, can supply a code. A
baseline UID set is captured before the run, so an older code in the inbox is
never reused.

---

## 4. Proxies

`PROXIES.txt` is one proxy URL per line (`http://user:pass@host:port`). The file
is CRLF — the loader strips `\r`, because both Playwright and curl reject a URL
with a trailing carriage return.

- No `--proxy` → picks a live proxy, remembers the last good one in `.proxy_used`
- `--proxy ''` → direct, no proxy at all
- `--proxy 'http://…'` → that proxy
- `--rotate N` → on a block, try N more proxies (default 3)
- `--no-rotate` → never switch mid-run

Datacenter proxies vary a lot: some never hydrate the page, some sit behind X's
cookie notice. A proxy that cannot render the signup pane is dropped from
`.proxy_used` so the next run picks a different one.

---

## 5. Exit codes

Success is `0`. Anything else is a named failure:

| Code | Meaning | Retried? |
|---|---|---|
| `0` | registered and signed in at `x.com/home` | — |
| `2` | edge block / proxy never hydrated | yes, next proxy
| `3` | mailbox unreadable, or no code arrived | no |
| `4` | a wizard pane did not accept input | no |
| `5` | loop exhausted its steps | no |
| `6` | stopped at `face_liveness` — **not registered** | yes, full restart |
| `7` | captcha, headless | no |
| `8` | browser died mid-run | yes, full restart |
| `9` | reached `/home` but no signed-in identity | no |

Login and `check` share this contract, with two differences: `3` also means
*no saved session* for `check`, and login retries only on `2`, `6`, `8` — a
stage it cannot classify ends at `5` (ran out of steps) and is **not** retried.

`--retries N` (default 3) restarts the **whole flow** on `2`, `6`, and `8`. X
decides per session whether it wants a face scan, so a fresh run is a legitimate
retry and often lands on the code-only path.

Exit code `6` is a real outcome, not a bug: the account record exists but is
unusable until a human passes the face check. The printed resume URL is where to
resume in a normal browser.

---

## 6. Diagnostics

```bash
python3 xreg.py --self-check         # offline: code parsing, DOB, name derivation, submit gating
python3 xreg.py --mail-check         # is the mailbox readable?
python3 xreg.py --mail --json -n 5   # dump recent messages as JSON
python3 xreg.py --dry-run            # does the signup page load?
python3 xreg.py login --save-session # sign in to an existing account, keep the cookies
python3 xreg.py check                # is that saved session still valid?
python3 xreg.py --mail-login         # force OAuth device-code, rewrite token cache
```

`--mail-check` prints real verification codes — treat its output as sensitive
and use `--json` when piping into anything.

Artifacts on failure: `x_blocked.png`, `x_stuck.png`, `x_liveness.png`,
`x_captcha.png`, `x_unknown.html`. The HTML dump is the fastest way to see what a
new pane wants.

| Symptom | Cause |
|---|---|
| `AUTHENTICATE failed` | IMAP not enabled for that Outlook mailbox |
| `invalid_grant` | refresh token rotated or revoked → `--mail-login` |
| proxy URL "malformed" | a `\r` survived — check the file is being read via `load_proxies()` |
| `unknown pane` | a stage `classify()` does not know yet; read `x_unknown.html` |
| stuck on `setup` | name contains digits, or handle already taken |

---

## 7. Known limits

- **Face liveness cannot be automated.** That is a physical check.
- **Arkose captcha** needs a human; run without `--headless`.
- One account per run. `account.txt` holds a single line.
- **Signup and login are separate flows.** `login_attempt()` handles a smaller
  stage set than `attempt()`; a pane it does not know (e.g. `dob`) waits for a
  human instead of advancing, and burns the step budget to exit `5`.
- `check` needs a session file, which only `login --save-session` writes.
- X changes its wizard. When it does, `classify()` needs a new branch — the
  `x_unknown.html` dump is designed for exactly that.
