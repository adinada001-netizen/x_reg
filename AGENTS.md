# AGENTS.md — operating this repo

Instructions for AI agents working in `TOOLING_X_REG`. Read this before touching
anything.

---

## What this repo is

One tool: `xreg.py`. It takes an X account end to end — reads the mailbox in
`account.txt`, drives X's wizard with Camoufox, pulls the verification code out
of Outlook over IMAP, and walks the flow to `x.com/home`.

Two flows, chosen by the state of the account, not by a flag:

| Account state | Flow | Command |
|---|---|---|
| never registered on X | **signup** | `xreg.py` (default) |
| already registered | **login** | `xreg.py login` |
| logged in once with `--save-session` | **check** | `xreg.py check` |

Running signup on a registered mailbox fails at `email_entry` (X refuses the
address); running login on an unregistered one never receives a code. Neither
flow probes the other's state — the operator must know which account it is.

| Path | Role |
|---|---|
| `xreg.py` | the whole tool — single entry point, single file |
| `README.md` | operator documentation: setup, flags, exit codes, limits |
| `PREREQUIREMENTS.md` | host, package versions, browser binary, credentials, IMAP, hard limits |
| `account.txt` | credentials, one pipe-separated line. **Secret. Never print it.** |
| `.session/` | cookie jar per account, written only by `login --save-session`. **Secret.** |
| `PROXIES.txt` | proxy list, CRLF line endings (the loader strips `\r`) |
| `.proxy_used` | marker for the last working proxy. Regenerable; safe to delete |
| `x_*.png`, `x_unknown.html` | failure artifacts written by the tool |

Nothing else in this directory is load-bearing.

---

## Rules

1. **Never print, log, or copy the contents of `account.txt`.** It holds a live
   mailbox password and an OAuth refresh token. Read it only via `load_account()`.
   To confirm a field exists, check length or presence, not value.
2. **Never commit `account.txt`, `.session/`, `.env`, or `~/.cache/outlook_newest/`.**
3. **Do not edit `xreg.py` to make a run pass.** If a stage fails, the fix belongs
   in a new `classify()` branch plus a fill function, and it must be covered by
   `--self-check`.
4. **`--self-check` must stay green.** It is the only check that runs without
   network or a browser. Extend it when you add a helper.
5. **Never claim a run succeeded without `EXIT=0` and `x.com/home` in the log.**
   Exit `6` means *not registered* — do not report it as success.

---

## How to operate it

### Prove it is alive (no side effects, safe anytime)

If any of these fail, check `PREREQUIREMENTS.md` before debugging the script.

```bash
.venv/bin/python xreg.py --self-check     # offline helpers
.venv/bin/python xreg.py --help           # CLI surface: signup | login | check
.venv/bin/python xreg.py --mail-check     # mailbox readable (prints a real code — sensitive)
.venv/bin/python xreg.py --dry-run        # signup page loads
```

### Choose the flow first

`account.txt` holds one mailbox. Is it registered on X yet?

- **No → signup.** `.venv/bin/python xreg.py ...` (no subcommand).
- **Yes → login.** `.venv/bin/python xreg.py login ...`. Never re-run signup on a
  registered address; X refuses it and you learn nothing.

If in doubt, `--mail-check` only proves the mailbox — it does **not** tell you
whether X knows the account. There is no offline probe for that; a `login` run
that reaches `password_create` is the cheapest proof the account exists.

### Run the real registration (account NOT yet registered)

```bash
.venv/bin/python xreg.py --proxy '' --code-timeout 180 --retries 3; echo "EXIT=$?"
```

- **No `--headless`.** The visible window is the point; a human may need to solve
  a captcha.
- **No `--code`, no `--handle`.** The tool derives them; passing them hides bugs.
- **Do not add retries around it.** `--retries` owns restart-on-failure.

### Run the login (account ALREADY registered)

```bash
.venv/bin/python xreg.py login --proxy '' --code-timeout 180 --retries 3 --save-session; echo "EXIT=$?"
```

- Same rules: no `--headless`, no `--code`, no `--handle`.
- `--save-session` writes `.session/<localpart>.json` (live credential, 600) so
  `xreg.py check` can verify the session later without a code.
- Login is not the mirror of signup: email → username prompt → password →
  `/home`, with `code` and `liveness` appearing only if X decides to ask. A
  pane `login_attempt()` does not know waits for a human, then exits `5`.
- Retries cover `2`, `6`, `8` only. `5` is not retried — read the log.

### Later: is the saved session still good?

```bash
.venv/bin/python xreg.py check --proxy ''; echo "EXIT=$?"
```

`0` still signed in, `9` expired (rerun `login`), `3` no session file, `2` launch
failed.

### Reading the result

| Exit | Meaning | Your action |
|---|---|---|
| `0` | registered and signed in at `x.com/home` (or session still valid, for `check`) | done |
| `6` | stopped at face liveness — **not usable** | report as "registered up to the face check", quote the resume URL |
| `2` | edge block / launch failed | let `--rotate` try another proxy |
| `3` | mailbox unreadable, no code, or no saved session (`check`) | fix credentials, or run `login --save-session` |
| `5` | ran out of steps on a pane the flow does not know | read `x_unknown.html`; not retried |
| `8` | browser died | rerun; usually the window was closed |
| `9` | `/home` reached without a signed-in identity | session expired; rerun `login` |
| others | named failure, see README §5 | do not improvise a fix |

Success is proven only by `EXIT=0` **and** `[x] account created and signed in`
(signup) or `[login] signed in: https://x.com/home` (login).

---

## When a run stops on an unknown pane

The tool prints:

```
[x] unknown pane dumped to x_unknown.html — url=...#/s/<route>/...
[x] unexpected pane — finish it in the visible browser, then press Enter
```

1. Read `x_unknown.html` — do not guess from the URL alone.
2. Find the control: `role=switch` / `aria-checked`, `role=button` with the
   action label, or the form field.
3. Add a `classify()` branch **before** the generic checks (the setup pane
   contains a password field, so order matters), and a fill function.
4. Route actions through `submit()`. X gates buttons with `aria-disabled` and
   clicking while gated is a silent no-op — never click blind.
5. Add assertions to `--self-check`.
6. Re-run `--dry-run`, then the full flow.

Known stages, for reference: `email_entry`, `code`, `dob`, `setup`,
`preferences`, `optional` (xN: Skip / Skip for now / Not now), `username`, `done`, `liveness`, `captcha`, `blocked`,
`cookie_notice`, `knowledge_check` (login only).

---

## Things that look like bugs but are not

- **`EXIT=6`** — X requires a face scan. Physical, not automatable.
- **A pane that "won't click Continue"** — check `aria-disabled` before blaming
  the selector; `submit()` exists for exactly this.
- **Digits in a display name** — X rejects them. `derive_name()` strips them.
- **Proxies "malformed"** — the file is CRLF. Only `load_proxies()` reads it correctly.
- **`window.*` reads returning `undefined`** — Camoufox hides page globals from
  `evaluate()`. Assert on rendered text and attributes instead. The tool never
  reads page globals.
- **`EXIT=5` on login** — `login_attempt()` knows fewer stages than `attempt()`;
  a `dob` pane waits for a human rather than being filled. Finish it in the
  window, or extend the login branch list.
- **Signup on an already-registered mailbox** — not a bug: X refuses the address.
  Use `xreg.py login`.

---

## Don't

- Don't add a second entry point, a package, or a config file. One file, one `main()`.
  (`signup` / `login` / `check` are subcommands of that one `main()`.)
- Don't add a dependency without updating the README setup section.
- Don't leave `x_*.png` / `x_unknown.html` behind after a successful investigation.
- Don't rewrite the IMAP/OAuth code. It is ported from a working
  implementation; rotating refresh tokens are cached deliberately.


NEXT-STEP
adalah menambahkan automasi connection berdasarkan dari login x yang telah dilakukan 
jadi tambahkan satu command --login-codebuddy yang mana sama mengikuti alur login -> hingga selesai
kemudian membuka tab baru dengan tetap mempertahankan sesilogin X 
kemudian buat script lanjutan untuk  buka http://localhost:20128/dashboard/ di tab baru
lalau login dengan menggunakan password|123456 aman karena di lokal
selanjutnya masuk http://localhost:20128/dashboard/providers/codebuddy-intl
cari tombol oAuth -> setelah sudah membuka tab https://www.codebuddy.ai/login pilih menggunakan sign-up with X
click confirm ataupun yang lainya.  ini gambaran awal tapi coba analisa directory next-step dan gabungkan flow yang layak diadopsi
