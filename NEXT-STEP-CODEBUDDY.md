# NEXT STEP — `--login-codebuddy`

Connect an X account to 9Router's `codebuddy-intl` provider, reusing the X session
`xreg.py login` already produces.

Status: **design only, nothing implemented yet.**

---

## 1. Facts probed live (not assumed)

9Router is running on `0.0.0.0:20128` (`~/.9router/launcher.pid`, next-server pid 3778).

| Probe | Result |
|---|---|
| `GET /api/auth/status` | `{"requireLogin":true,"authMode":"password","hasPassword":true}` |
| `POST /api/auth/login {"password":…}` | `{"success":true,"mustChangePassword":false}` + `auth_token` cookie (HS256, 24 h) |
| `GET /api/oauth/codebuddy-intl/device-code` | `{"device_code":"<uuid>","verification_uri":"https://www.codebuddy.ai/login?platform=ide&state=<uuid>","interval":5,"codeVerifier":"…"}` |
| `POST /api/oauth/codebuddy-intl/poll {"deviceCode":…}` | `{"success":false,"error":"authorization_pending","pending":true}` → on authorize `{"success":true,"connection":{…}}` |
| same call without cookie | `401` |
| `providerConnections` schema | `id, provider, authType NOT NULL, name, email, priority, isActive, data, createdAt, updatedAt` |

The whole dashboard detour (open `/dashboard/` → type password → `/dashboard/providers/codebuddy-intl`
→ click OAuth) is **two HTTP calls**. The browser is only needed for the part that
requires a human and an X session: `codebuddy.ai/login` → *Sign up with X* → confirm.

---

## 2. What to adopt from `next-step/`

| Piece | Verdict | Why |
|---|---|---|
| `request_codebuddy_state` + `poll_codebuddy_token` | **adopt the state machine, not the URLs** | identical `state` → `authUrl` → poll loop, but 9Router already exposes it; let 9Router own the connection record |
| iframe + policy checkbox + social button | **adopt the shape, not the ids** | `#social-google` / `#agree-policy-account` no longer exist in today's bundle (`social-providers`, `agreement`) — match by role + accessible name |
| `insert_codebuddy_to_9router` (SQLite write) | **drop** | its `INSERT` omits `authType` (NOT NULL) → would fail today; also races the running server and skips `createProviderConnection`'s bookkeeping |
| `get_user_resource` (email lookup) | **drop** | 9Router fills `email`/`displayName` from the token itself |
| `get_existing_connected_emails` | **adopt the idea** | read `/api/providers`, skip an email that is already connected (poll would create a duplicate) |
| `fetch_webshare_proxies`, `config.json`, `Google.txt`, `logs/` | **drop** | `PROXIES.txt` + `--proxy`/`--rotate` already exist; no second config file (AGENTS.md) |
| `register_single_google`, `extract_google_2fa_code` | **drop** | Google path, out of scope for this tool |
| `webbrowser.open(auth_url)` | **drop** | must stay inside the same Camoufox context or the X session is not there |

---

## 3. Flow

```
xreg.py login --save-session      →  .session/<localpart>.json   (X cookies, 600)
xreg.py --login-codebuddy         →  same file, restored into a fresh context

  1. restore X cookies            page.context.add_cookies(state["cookies"])
  2. POST /api/auth/login         → auth_token cookie
  3. GET  …/device-code           → verification_uri + device_code
  4. ctx.new_page() → verification_uri        (X session already present)
  5. iframe: tick agreement → click "Sign up with X"
  6. X consent pane → Continue / Authorize    (submit(), aria-disabled aware)
  7. POST …/poll every 5 s        → {"success":true,"connection":…}
  8. verify GET /api/providers    → codebuddy-intl connection with this X email
```

No `.session` file → exit `3` and print the exact `login --save-session` command.
Already connected → exit `0`, no duplicate, unless `--force`.

---

## 4. Code plan (one file, no new dependency)

`urllib` is already imported; `page.context.add_cookies()` is stock Playwright.

| Add to `xreg.py` | Kind | Notes |
|---|---|---|
| `NINEROUTER`, `NINEROUTER_PW`, `CB_PROVIDER`, `CB_X_LABELS` | consts | `NINEROUTER_URL` / `NINEROUTER_PASSWORD` env override; default password is 9Router's own local default |
| `nr_login()` | http | returns cookie header string; raises on non-200 |
| `cb_device_code(cookie)` / `cb_poll(cookie, code, timeout)` | http | pure JSON in/out, easy to self-check |
| `cb_pick_label(page)` | pure | `CB_X_LABELS` order: `"Sign up with X"` → `"Sign in with X"` → `"X"` → `"Twitter"`; mirrors `skip_label()` |
| `restore_session(page, email)` | glue | 6 lines around `add_cookies` |
| `connect_codebuddy(args, page)` | browser | steps 5–7 |
| `run_codebuddy(args)` | entry | launch → restore → connect → close |
| `--login-codebuddy` flag in `main()` | CLI | a flag, not a subcommand: one `main()`, per AGENTS.md |

Exit codes — reuse the existing contract, add two:

| Code | Meaning |
|---|---|
| `0` | connected (or already connected) |
| `3` | no `.session` file → run `login --save-session` first |
| `4` | codebuddy pane did not accept input (X button not found / consent not confirmed) |
| `10` | 9Router unreachable or rejected the password |
| `11` | codebuddy authorization never completed (poll timeout) |

`--self-check` gains three offline asserts: label pick order (specific beats
substring), `cb_poll` parsing of pending vs done, `cb_device_code` rejecting a
payload without `device_code`/`verification_uri`.

---

## 5. Open questions to settle against the live page

1. **X button selector.** The bundle is i18n'd; confirm the control by dumping the
   pane (`x_unknown.html` method) rather than copying an id. Keep the label list.
2. **Does X re-prompt for the password during OAuth?** If yes, reuse
   `login_password()`; if it demands a face scan, that is exit `6` again.
3. **Duplicate policy.** Poll always calls `createProviderConnection`, so a second
   run adds a second row. Default: skip when the email already exists; `--force`
   to add anyway.
4. **Consent wording.** X may show "Authorize app" or "Continue as @handle" —
   route it through `submit()` with both labels, same as `optional`.

---

## 6. Definition of done

```bash
.venv/bin/python xreg.py --self-check                     # green, 3 new asserts
.venv/bin/python xreg.py login --save-session             # EXIT=0
.venv/bin/python xreg.py --login-codebuddy; echo $?       # 0
curl -s -b jar localhost:20128/api/providers | jq '.connections[]|select(.provider=="codebuddy-intl")|.email'
```

Success is `EXIT=0` **and** the connected email equals the mailbox in `account.txt`.
