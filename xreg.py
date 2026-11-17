#!/usr/bin/env python3
"""xreg.py — register an X account from the email in account.txt, verify it with
the code X mails to that mailbox, and stop once the account is created.

Merged from: open_x.py (Playwright browser control) and
GRABBER_MAIL/outlook_newest.py (Outlook IMAP + OAuth2 XOAUTH2 mail reader).

Flow:
    account.txt -> x.com/i/flow/signup -> code mailed to mailbox -> code typed back -> handle

Usage:
    python3 xreg.py                 # full flow (visible browser by default)
    python3 xreg.py --headless      # hide it (X's edge 403s headless more often)
    python3 xreg.py --dry-run       # load the signup page only
    python3 xreg.py --mail-check    # only prove the mailbox is readable
    python3 xreg.py --self-check    # offline helper checks
    python3 xreg.py --code 123456   # supply the code yourself

Account file (account.txt, one line):
    email|password|msa_artifact|client_id
    - password      reused as the X account password
    - msa_artifact  Outlook MSA artifact; IS an OAuth refresh token, reads the code mail
    - client_id     Thunderbird public client id, pre-approved for IMAP.AccessAsUser.All
    OUTLOOK_* / X_PASSWORD env vars win over the file.

Token cache: ~/.cache/outlook_newest/<email>.json  (chmod 600)
"""
import argparse
import email
import imaplib
import json
import os
import re
import stat
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from email.header import decode_header, make_header
from html.parser import HTMLParser
from pathlib import Path

# ── config ────────────────────────────────────────────────────────────────────
IMAP_HOST, IMAP_PORT = "outlook.office365.com", 993
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
CACHE = Path.home() / ".cache" / "outlook_newest"
SIGNUP_URL = "https://x.com/i/flow/signup"
LOGIN_URL = "https://x.com/i/flow/login"
TIMEZONE = "America/New_York"  # matches the proxy pool's Ashburn egress
UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/138.0.0.0 Safari/537.36")
PROXY_FILE = Path(__file__).with_name("PROXIES.txt")
MARKER = Path(__file__).with_name(".proxy_used")  # last good proxy, prevents run-to-run drift
SESSION_DIR = Path(__file__).with_name(".session")  # storage_state per account  # last good proxy, prevents run-to-run drift

# X's edge rejects the headless shell; the full chromium build passes.
# channel="chromium" selects it.
CHROME_ARGS = ["--disable-blink-features=AutomationControlled"]
STEALTH = "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"

CODE_RE = re.compile(r"(?<!\d)(\d{6,8})(?!\d)")
# X mail bodies carry boilerplate numbers; the code sits next to "code"/"confirm".
CODE_HINT_RE = re.compile(
    r"(?:code|confirm\w*)[^\n\d]{0,80}?(?<!\d)(\d{6,8})(?!\d)"
    r"|(?<!\d)(\d{6,8})(?!\d)[^\n]{0,80}?(?:code|confirm)", re.I)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


# ── account file ──────────────────────────────────────────────────────────────
def load_account(path="account.txt"):
    """Resolve credentials: OUTLOOK_* env vars win, account file is the fallback.

    Returns (email, password, msa_artifact_or_None, client_id).
    """
    env = {k: os.environ.get(k, "").strip() for k in
           ("OUTLOOK_EMAIL", "OUTLOOK_CLIENT_ID", "OUTLOOK_MSA", "X_PASSWORD")}

    # Per-field precedence: each env var overrides only its own field, so setting
    # OUTLOOK_EMAIL alone does not discard the file's client_id or MSA.
    fields = []
    p = Path(path)
    if p.exists():
        fields = [f.strip() for f in p.read_text(encoding="utf-8").strip().split("|")]
    elif not env["OUTLOOK_EMAIL"]:
        raise SystemExit(f"missing account file {path} (see GRABBER_MAIL/accont.txt.example)")
    elif len(fields) < 2 and fields:
        raise SystemExit(f"bad account file {path}: expected email|password|msa|client_id")

    email_addr = env["OUTLOOK_EMAIL"] or (fields[0] if fields else "")
    password = env["X_PASSWORD"] or (fields[1] if len(fields) > 1 else "")
    msa = env["OUTLOOK_MSA"] or (fields[2] if len(fields) >= 3 and fields[2] else None)
    client_id = env["OUTLOOK_CLIENT_ID"] or (fields[-1] if fields else "")

    if "@" not in email_addr:
        sys.exit("missing email (OUTLOOK_EMAIL or first field in account.txt)")
    if len(client_id) != 36:
        sys.exit("missing/invalid OUTLOOK_CLIENT_ID (must be an Azure application id)")
    if not password:
        sys.exit("missing password (second field in account.txt, or $X_PASSWORD)")
    return email_addr, password, msa, client_id


# ── Outlook IMAP over OAuth2 (from outlook_newest.py) ─────────────────────────
def post(url, data, timeout=30, tries=3):
    body = urllib.parse.urlencode(data).encode()
    hdrs = {"Content-Type": "application/x-www-form-urlencoded"}
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, body, hdrs), timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            # 400 carries the device-code poll status as error JSON, not a failure.
            if e.code == 400:
                return json.load(e)
            if e.code != 429 and e.code < 500:
                raise
            if i == tries - 1:
                raise
            time.sleep(1 + i)
        except urllib.error.URLError:
            if i == tries - 1:
                raise
            time.sleep(1 + i)


def access_token(email_addr, client_id, force_login=False, msa=None):
    """Return an IMAP access token, refreshing the cached refresh token as needed."""
    CACHE.mkdir(parents=True, exist_ok=True)
    cache = CACHE / f"{email_addr}.json"
    tokurl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"

    # The MSA artifact IS a refresh token: seed the cache, no browser needed.
    if not cache.exists() and msa:
        cache.write_text(json.dumps({"refresh_token": msa}))
        os.chmod(cache, stat.S_IRUSR | stat.S_IWUSR)

    if not force_login and cache.exists():
        rt = json.loads(cache.read_text())["refresh_token"]
        r = post(tokurl, {"client_id": client_id, "grant_type": "refresh_token",
                          "refresh_token": rt, "scope": SCOPE})
        if "access_token" in r:
            if r.get("refresh_token"):
                cache.write_text(json.dumps({"refresh_token": r["refresh_token"]}))
                os.chmod(cache, stat.S_IRUSR | stat.S_IWUSR)
            return r["access_token"]
        log(f"[auth] refresh failed ({r.get('error')}), falling back to device code")

    dc = post("https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode",
              {"client_id": client_id, "scope": SCOPE})
    log(f"\n1. Open {dc['verification_uri']}\n2. Enter code: {dc['user_code']}\n"
        f"3. Sign in as {email_addr}\nwaiting…")
    deadline = time.time() + int(dc.get("expires_in", 900))
    interval = int(dc.get("interval", 5))
    while time.time() < deadline:
        time.sleep(interval)
        r = post(tokurl, {"client_id": client_id,
                          "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                          "device_code": dc["device_code"]})
        if "access_token" in r:
            cache.write_text(json.dumps({"refresh_token": r["refresh_token"]}))
            os.chmod(cache, stat.S_IRUSR | stat.S_IWUSR)
            return r["access_token"]
        err = r.get("error", "")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
        else:
            sys.exit(f"[auth] {err}: {r.get('error_description', '')[:300]}")
    sys.exit("[auth] device code expired")


def xoauth2(email_addr, token):
    return f"user={email_addr}\x01auth=Bearer {token}\x01\x01"


class _Text(HTMLParser):
    """Collect visible text; drops script/style/head, breaks on block tags."""
    SKIP = {"script", "style", "head", "title"}
    BREAK = {"br", "p", "div", "tr", "li", "h1", "h2", "h3", "h4", "table"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in self.SKIP:
            self._skip += 1
        elif tag in self.BREAK:
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in self.SKIP and self._skip:
            self._skip -= 1
        elif tag in self.BREAK:
            self.out.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.out.append(data)


def html_to_text(html):
    t = _Text()
    t.feed(html)
    return "\n".join(ln for ln in (x.strip() for x in "".join(t.out).splitlines()) if ln)


def body_text(msg):
    """Text body of a message. Prefers text/plain, falls back to stripped HTML."""
    # msg.walk() is a GENERATOR: materialize it, or the text/html pass runs over
    # an already-exhausted iterator and silently returns "".
    parts = list(msg.walk()) if msg.is_multipart() else [msg]
    for ctype in ("text/plain", "text/html"):
        for part in parts:
            if part.get_content_type() != ctype or "attachment" in str(
                    part.get("Content-Disposition", "")):
                continue
            raw = part.get_payload(decode=True)
            if raw is None:
                continue
            txt = raw.decode(part.get_content_charset() or "utf-8", "replace")
            txt = html_to_text(txt) if ctype == "text/html" else txt
            if txt.strip():
                return txt
    return ""


def _hdr(msg, name):
    return str(make_header(decode_header(msg.get(name, ""))))


def fetch_messages(email_addr, client_id, msa, count=5, token=None):
    """Newest `count` messages, oldest-first."""
    token = token or access_token(email_addr, client_id, msa=msa)
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=30)
    try:
        # imaplib base64-encodes the SASL response itself -> pass raw bytes
        imap.authenticate("XOAUTH2", lambda _=None: xoauth2(email_addr, token).encode())
        imap.select("INBOX", readonly=True)
        typ, data = imap.search(None, "ALL")
        if typ != "OK" or not data[0].split():
            return []
        out = []
        for uid in list(reversed(data[0].split()))[:count]:
            typ, data = imap.fetch(uid, "(RFC822)")
            if typ != "OK":
                break
            msg = email.message_from_bytes(data[0][1])
            out.append({"uid": uid.decode(), "from": _hdr(msg, "From"), "to": _hdr(msg, "To"),
                        "subject": _hdr(msg, "Subject"), "date": _hdr(msg, "Date"),
                        "message_id": msg.get("Message-ID", ""), "body": body_text(msg)})
        return out
    except imaplib.IMAP4.error as e:
        sys.exit(f"[imap] {e}\nhint: Outlook.com → Settings → Mail → Forwarding and IMAP → "
                 "enable 'Let devices and apps use IMAP'.")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


def extract_code(text):
    """Pull the verification code out of a mail body."""
    m = CODE_HINT_RE.search(text)
    if m:
        return m.group(1) or m.group(2)
    m = CODE_RE.search(text)
    return m.group(1) if m else None


def uid_baseline(email_addr, client_id, msa, count=5):
    """UIDs already in the mailbox, so a stale mail is never read as the fresh code.

    Raises on failure: a failed baseline must not look like an empty mailbox, or
    any old code in the inbox becomes a candidate.
    """
    return {m["uid"] for m in fetch_messages(email_addr, client_id, msa, count=count)}


def is_signup_mail(m, email_addr):
    """Only X signup/confirmation mail addressed to this mailbox can supply a code."""
    sender = (m.get("from") or "").lower()
    rcpt = (m.get("to") or "").lower()
    subject = (m.get("subject") or "").lower()
    return ("x.com" in sender or "twitter.com" in sender) and \
        (email_addr.lower() in rcpt or not rcpt) and \
        ("code" in subject or "confirm" in subject)


def wait_for_code(email_addr, client_id, msa, baseline=(), timeout=180, poll=8):
    """Poll until a mail outside `baseline`, from X, for this mailbox, yields a code."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for m in fetch_messages(email_addr, client_id, msa, count=8):
            if m["uid"] in baseline or not is_signup_mail(m, email_addr):
                continue
            code = extract_code(f"{m['subject']}\n{m['body']}")
            if code:
                log(f"[mail] code {code} from <{m['from']}> «{m['subject']}»")
                return code
        log(f"[mail] no code yet, polling again in {poll}s")
        time.sleep(poll)
    return None


# ── X signup (from open_x.py) ─────────────────────────────────────────────────
# ── proxies ───────────────────────────────────────────────────────────────────
def load_proxies(path=PROXY_FILE):
    """Proxy URLs from PROXIES.txt. The file is CRLF — strip the CR or both
    Playwright and curl reject the URL as malformed."""
    if not path.exists():
        return []
    out = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip().strip('"').strip("'")
        if line and "://" in line:
            out.append(line)
    return out


def parse_proxy(url):
    """http://user:pass@host:port -> Playwright proxy dict."""
    u = urllib.parse.urlsplit(url)
    if not u.hostname or not u.port:
        sys.exit(f"bad proxy url {url!r} (want scheme://[user:pass@]host:port)")
    p = {"server": f"{u.scheme}://{u.hostname}:{u.port}"}
    if u.username:
        p["username"] = urllib.parse.unquote(u.username)
        p["password"] = urllib.parse.unquote(u.password or "")
    return p


def proxy_reachable(url, timeout=12):
    """Cheap liveness check: ask the proxy for an IP echo."""
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": url, "https": url}))
    try:
        with opener.open("https://api.ipify.org", timeout=timeout) as r:
            return bool(r.read(64).strip())
    except Exception:
        return False


def pick_proxy(explicit=None, timeout=10, tries=8):
    """explicit > last good (.proxy_used) > first in PROXIES.txt that answers.

    `--proxy ''` (empty string) means "go direct": no proxy anywhere.
    """
    if explicit is not None:
        return explicit.strip() or None
    candidates = load_proxies()
    if not candidates:
        return None
    if MARKER.exists():
        last = MARKER.read_text(encoding="utf-8").strip()
        if last in candidates and proxy_reachable(last, timeout=timeout):
            return last
    for url in candidates[:tries]:
        if proxy_reachable(url, timeout=timeout):
            MARKER.write_text(url, encoding="utf-8")
            return url
        log(f"[proxy] dead {urllib.parse.urlsplit(url).hostname}")
    sys.exit("no live proxy in PROXIES.txt")


def _visible(page, selector, timeout=20000):
    try:
        page.wait_for_selector(selector, state="visible", timeout=timeout)
        return page.query_selector(selector)
    except Exception:
        return None


def _click(page, text, timeout=15000):
    loc = page.get_by_role("button", name=text, exact=True).first
    try:
        loc.wait_for(state="visible", timeout=timeout)
        loc.click()
        return True
    except Exception:
        return False


def _body(page):
    try:
        return page.inner_text("body")
    except Exception:
        return ""


# X's markup differs per engine and A/B bucket: Chromium name=challenge_response,
# Firefox/Camoufox an unnamed input with autocomplete=one-time-code and maxlength 6.
CODE_INPUT_SELECTORS = (
    'input[name="challenge_response"]',
    'input[autocomplete="one-time-code"]',
    'input[name="verification_code"]',
    'input[inputmode="numeric"][maxlength="6"]',
    'input[aria-label="Verification code"]',
)


def NAME_INPUT(page):
    """Name field of the setup/username pane."""
    for sel in ('input[name="name"]', 'input[autocomplete="name"]',
                'input[aria-label="Name"]', 'input[data-testid="ocfEnterTextTextInput"]'):
        el = page.query_selector(sel)
        if el and el.is_visible():
            return el
    return None


def PASS_INPUT(page):
    """Password field."""
    for sel in ('input[autocomplete="new-password"]', 'input[type="password"]',
                'input[name="password"]'):
        el = page.query_selector(sel)
        if el and el.is_visible():
            return el
    return None


def CODE_INPUT(page, timeout=0):
    """First visible verification-code input, or None."""
    for sel in CODE_INPUT_SELECTORS:
        el = _visible(page, sel, timeout=timeout) if timeout else page.query_selector(sel)
        if el and el.is_visible():
            return el
    return None


def logged_in(page):
    """True only when X shows a signed-in session we can attribute."""
    body = _body(page)
    low = body.lower()
    if "sign in" in low and "create account" in low:
        return False  # still the anonymous shell
    for sel in ('[data-testid="SideNav_AccountSwitcher_Button"]',
                '[data-testid="AppTabBar_Profile_Link"]',
                'a[href="/home"][aria-label], [aria-label="Profile"]'):
        if page.query_selector(sel):
            return True
    return "/home" in page.url and "sign in" not in low


def classify(page):
    """Derive the wizard stage from whatever is on screen."""
    body = _body(page)
    low = body.lower()
    if "something is not right" in low or "try a different method" in low:
        return {"stage": "blocked", "why": "X refused this signup attempt"}
    if "did someone say" in low or ("cookies" in low and "accept all" in low):
        return {"stage": "cookie_notice"}
    if "something went wrong" in low:
        return {"stage": "blocked", "why": "X served an error pane (proxy or edge block)"}
    # X's login reassurance pane ("Confirm your account", #/s/knowledge_check)
    # reuses the SAME challenge_response input as the code pane but asks for the
    # account's username/email. Detect it before CODE_INPUT, or it is misread as a
    # verification-code prompt and the run stalls waiting for mail that never comes.
    if "knowledge_check" in page.url or "confirm your account" in low:
        return {"stage": "knowledge_check"}
    if CODE_INPUT(page, 800):
        return {"stage": "code"}
    if "face_liveness" in page.url or "confirm you're a human" in low or \
            "confirm with face" in low:
        return {"stage": "liveness",
                "why": "X requires a face-scan liveness check; no browser can pass it"}
    # "Set up your account" carries all three fields *including* a password, so it
    # has to be recognised before the generic password pane or it is misread.
    if "set up your account" in low or "choose a name, username" in low:
        return {"stage": "setup"}
    if NAME_INPUT(page) and PASS_INPUT(page):
        return {"stage": "setup"}
    if NAME_INPUT(page) or _visible(page, 'input[data-testid="ocfEnterTextTextInput"]', 500):
        return {"stage": "username"}
    if PASS_INPUT(page):
        return {"stage": "password_create"}
    if "arkose" in low or "captcha" in low or "funCaptcha" in page.content()[:4000]:
        return {"stage": "captcha", "why": "Arkose/FunCaptcha challenge"}
    if "verify_code" in page.url or "we sent you a security code" in low:
        return {"stage": "code"}
    if "privacy preferences" in low or "signup_preferences" in page.url:
        return {"stage": "preferences"}
    # X chains a few optional setup questions after preferences and each offers a
    # different dismissal label. They carry no stable heading, so detect the
    # control itself. Order matters: "Skip for now" contains "Skip".
    if any(_control_exists(page, label) for label in SKIP_LABELS):
        return {"stage": "optional"}
    if "birthday" in low or "signup_dob" in page.url:
        return {"stage": "dob"}
    if "confirm your account" in low or "verification code" in low:
        return {"stage": "code"}
    if "password" in low:
        return {"stage": "password_create"}
    if "create your account" in low or "username" in low:
        return {"stage": "username"}
    if page.url.rstrip("/").endswith("/home") or "what's happening" not in low and "/home" in page.url:
        return {"stage": "done"}
    return {"stage": "unknown", "why": page.url}


def settle(page, timeout=45):
    """Classify, waiting out X's pane animations before calling it unknown."""
    deadline = time.time() + timeout
    state = classify(page)
    while state["stage"] == "unknown" and time.time() < deadline:
        page.wait_for_timeout(2000)
        state = classify(page)
    return state


def ready(page, timeout=40):
    """Wait for the signup pane to render on this proxy.

    Datacenter proxies vary: some serve the login pane, some sit behind X's
    cookie notice, some never hydrate. The notice is dismissed opportunistically.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        body = _body(page)
        if "cookies" in body.lower() and _click(page, "Accept all cookies", timeout=3000):
            page.wait_for_timeout(1500)
            continue
        if _visible(page, 'input[name="username_or_email"]', timeout=1000):
            return True
        page.wait_for_timeout(2000)
    return False


# Plausible adult birth years. Age checks bite below 18 and X's face-liveness
# loop bites on brand-new same-day birthdates, so keep a wide real spread.
YEAR_RANGE = (1990, 2005)
MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December")


def parse_dob_arg(value):
    """argparse type for --dob: a real calendar date in YEAR_RANGE."""
    import datetime
    try:
        y, m, d = (int(x) for x in value.split("-"))
        datetime.date(y, m, d)
    except (ValueError, TypeError):
        raise argparse.ArgumentTypeError("expected YYYY-MM-DD, got %r" % (value,))
    if not YEAR_RANGE[0] <= y <= YEAR_RANGE[1]:
        raise argparse.ArgumentTypeError("year must be %d-%d" % YEAR_RANGE)
    return (y, m, d)


def random_dob(year_range=YEAR_RANGE):
    """(year, month, day) inside year_range, day valid for that month."""
    import calendar, random
    year = random.randint(*year_range)
    month = random.randint(1, 12)
    return year, month, random.randint(1, calendar.monthrange(year, month)[1])


def _dob_controls(page):
    """The three birthday controls in page order: (month, day, year) elements.

    X renders these as custom pills - a clickable trigger that opens a listbox of
    plain divs. Native <select> only appears in some buckets, so both are handled.
    """
    selects = [s for s in page.query_selector_all("select") if s.is_visible()]
    if len(selects) >= 3:
        return selects[:3]
    for sel in ('[role="combobox"]', '[aria-haspopup="listbox"]',
                '[aria-haspopup="true"][role="button"]'):
        combos = [c for c in page.query_selector_all(sel) if c.is_visible()]
        if len(combos) >= 3:
            return combos[:3]
    labelled = [page.query_selector('[aria-label="%s"]' % l) for l in ("Month", "Day", "Year")]
    if all(labelled):
        return labelled
    return []


def _same_value(actual, wanted):
    """Exact compare after normalizing: wanted "1" must not match trigger "21"."""
    a = " ".join(str(actual or "").split())
    w = " ".join(str(wanted or "").split())
    return a == w or (a.split(" ")[0] if a else "") == w


def _pick_option(page, combo, wanted):
    """Open a custom combobox, choose `wanted`, confirm the trigger text changed.

    Option rows are bare divs in a popover with no role=listbox, and Continue only
    re-enables when X's own change handler fires - so the click must land on the
    row and the trigger's text has to be verified afterwards.
    """
    before = (combo.inner_text() or "").strip()
    combo.click()
    page.wait_for_timeout(600)

    rows = [r for r in page.query_selector_all(
        '[role="option"], [data-testid*="option"], ul li, '
        '[class*="popover"] div, [class*="menu"] div, div') if r.is_visible()]
    wanted_n = " ".join(str(wanted).split())

    for row in rows:
        try:
            txt = " ".join((row.inner_text() or "").split())
        except Exception:
            continue
        if txt != wanted_n:
            continue
        try:
            row.click()
            page.wait_for_timeout(500)
            if _same_value(combo.inner_text(), wanted_n):
                return True
        except Exception:
            continue

    try:
        page.keyboard.type(wanted_n)
        page.wait_for_timeout(300)
        page.keyboard.press("Enter")
        page.wait_for_timeout(400)
        if wanted_n in " ".join((combo.inner_text() or "").split()):
            return True
    except Exception:
        pass

    log("[dob] option %r not applied (trigger still %r)" % (wanted_n, before))
    return False



def fill_dob(page, year=None, month=None, day=None, dob=None):
    """Fill "When's your birthday?" with a random 1990-2005 date unless given one.

    Handles both renderings: native <select> (light pane, Chromium) and X's
    custom combobox (dark pane, Firefox/Camoufox).
    """
    year, month, day = dob or (year, month, day)
    if year is None or month is None or day is None:
        year, month, day = random_dob()
    log(f"[dob] {year}-{month:02d}-{day:02d}")

    ctrls = _dob_controls(page)
    if len(ctrls) < 3:
        return False
    wants = (MONTHS[month - 1], str(day), str(year))
    applied = []
    for ctrl, wanted in zip(ctrls, wants):
        if ctrl.evaluate("e => e.tagName") == "SELECT":
            try:
                ctrl.select_option(label=wanted)
            except Exception:
                ctrl.select_option(value=str(int(wanted)) if wanted.isdigit()
                                   else "%02d" % month)
            ok = _same_value(ctrl.evaluate("e => e.options[e.selectedIndex].text"), wanted)
        else:
            ok = _pick_option(page, ctrl, wanted)
        applied.append(ok)
        page.wait_for_timeout(400)

    if not all(applied):
        log("[dob] controls not all set: %r - not submitting" % (applied,))
        return False

    # Continue is <div role=button aria-disabled>: clicking while disabled is a
    # silent no-op, so wait for the flag to clear before firing.
    btn = None
    for cand in page.query_selector_all('[role="button"], button'):
        try:
            label = (cand.inner_text() or "").strip().lower()
            if cand.is_visible() and "cont" "inue" in label:
                btn = cand
                break
        except Exception:
            pass
    if not btn:
        log("[dob] Continue button not found")
        return False

    deadline = time.time() + 10
    while time.time() < deadline and btn.get_attribute("aria-disabled") == "true":
        page.wait_for_timeout(500)
    if btn.get_attribute("aria-disabled") == "true":
        log("[dob] Continue stayed disabled after all three fields were set")
        return False

    btn.click()
    page.wait_for_timeout(4000)
    return True


def submit_email(page, email_addr):
    if not ready(page):
        return False
    page.fill('input[name="username_or_email"]', email_addr)
    # Enter is what triggers the flow; the Continue button is inert until X has
    # validated the field client-side.
    page.keyboard.press("Enter")
    page.wait_for_timeout(3000)
    if not _click(page, "Continue", timeout=5000):
        page.keyboard.press("Enter")
    return True


def fill_code(page, code):
    inp = CODE_INPUT(page)
    if not inp:
        return False
    inp.fill(code)
    page.wait_for_timeout(300)
    if not _click(page, "Continue", timeout=6000):
        page.keyboard.press("Enter")
    page.wait_for_timeout(4000)
    return True


def fill_password(page, password):
    for sel in ('input[autocomplete="new-password"]', 'input[data-testid="password"]',
                'input[name="password"]'):
        inp = page.query_selector(sel)
        if inp and inp.is_visible():
            inp.fill(password)
            page.wait_for_timeout(300)
            if not _click(page, "Continue", timeout=6000):
                page.keyboard.press("Enter")
            page.wait_for_timeout(4000)
            return True
    return False


def fill_username(page, handle):
    """Username-only pane (X asks for it alone when it edits the suggestion)."""
    inp = page.query_selector('input[name="username"]') or \
        page.query_selector('input[autocomplete="username"]') or \
        page.query_selector('input[data-testid="ocfEnterTextTextInput"]')
    if not inp or not inp.is_visible():
        return False
    inp.fill(handle)
    page.wait_for_timeout(500)
    if not _click(page, "Continue", timeout=6000):
        page.keyboard.press("Enter")
    page.wait_for_timeout(4000)
    return True


def derive_handle(email_addr):
    """Handle seed from the mailbox local part; X lets the user edit it."""
    return re.sub(r"[^A-Za-z0-9_]", "", email_addr.split("@")[0])[:15] or "user"


# Dismissal labels on X's optional onboarding panes, longest first so the
# specific match wins over its own substring.
SKIP_LABELS = ("Skip for now", "Not now", "Skip")

NAME_FALLBACK = "User"      # X rejects an empty display name
NAME_MAX = 50              # X's own cap on the name field


def derive_name(email_addr):
    """"Set up your account" display name.

    X's name field rejects digits outright, so strip the numeric tail the mailbox
    carries (OgdenIntrieri341 -> OgdenIntrieri) and split camel case into words.
    Falls back to a fixed token when nothing usable survives.
    """
    local = email_addr.split("@")[0]

    # 1. drop the trailing digit run most mailboxes end with
    stem = re.sub(r"\d+$", "", local)
    # 2. any remaining digits are illegal in the field
    stem = re.sub(r"\d+", "", stem)
    # 3. separators and punctuation the field will not take
    stem = re.sub(r"[^A-Za-z]+", " ", stem).strip()
    if not stem:
        return NAME_FALLBACK

    # 4. "ogdenintrieri" -> "Ogdenintrieri"; "OgdenIntrieri" stays split
    words = re.findall(r"[A-Z]?[a-z]+|[A-Z]+(?![a-z])", stem) or [stem]
    name = " ".join(w.capitalize() for w in words[:2]) if len(words) > 1 else words[0].capitalize()

    name = name[:NAME_MAX].strip()
    # 5. a one-letter name is not accepted either
    return name if len(name) >= 2 else NAME_FALLBACK


def derive_handle_from_name(email_addr, digits=4, salt=None):
    """First 5 chars of the mailbox local part + random digits — 'Ogden' + '4731'."""
    import random
    stem = re.sub(r"[^A-Za-z0-9_]", "", email_addr.split("@")[0]) or "user"
    rnd = random.Random(salt) if salt is not None else random
    return f"{stem[:5]}{''.join(str(rnd.randint(0, 9)) for _ in range(digits))}"


def skip_label(page):
    """Which dismissal label this pane offers, or None.

    Longest first: "Skip for now" must win over "Skip", or the bare-Skip branch
    would fire on a substring and click the wrong control.
    """
    for label in SKIP_LABELS:
        if _enabled_control(page, label) is not None:
            return label
    for label in SKIP_LABELS:
        if _control_exists(page, label):
            return label
    return None


def skip_optional(page):
    """Dismiss an optional onboarding pane via whatever dismissal it offers.

    The panes differ only in their label — "Skip", then "Skip for now", then
    "Not now" — so the action is the same and only the wording changes.
    """
    label = skip_label(page)
    if label is None:
        log("[optional] pane offers no Skip/Not now control")
        return False
    if not submit(page, label):
        log(f"[optional] '{label}' never became actionable")
        return False
    log(f"[optional] clicked '{label}'")
    return True


def fill_preferences(page):
    """Privacy preferences pane: opt out of email discoverability, then submit.

    X ships "Let people find your account by your email" ON by default (a
    role=switch with aria-checked="true"). Turning it off is the safer default,
    so flip every switch that is still on before continuing.
    """
    switches = page.query_selector_all('[role="switch"], input[type="checkbox"]')
    flipped = 0
    for sw in switches:
        try:
            if not sw.is_visible():
                continue  # hidden switch, nothing to flip
            state = sw.get_attribute("aria-checked")
            if state is None:
                state = "true" if sw.is_checked() else "false"
            if str(state).lower() == "true":
                sw.click()
                page.wait_for_timeout(400)
                flipped += 1
        except Exception:
            pass  # node went stale between scan and click
    log(f"[prefs] turned off {flipped} discoverability toggle(s)")
    return submit(page, "Continue")


def submit(page, label="Continue", timeout=15):
    """Click a wizard affordance once it is actually actionable.

    X gates these controls with aria-disabled / disabled / a CSS class rather than
    one reliable attribute, so clicking a not-yet-enabled control is a silent no-op
    and the pane looks stuck. Poll until the gate clears, then click; fall back to
    Enter for panes that advance on a focused text field.
    """
    deadline = time.time() + timeout
    saw_disabled = False
    while time.time() < deadline:
        target = _enabled_control(page, label)
        if target is None:
            saw_disabled = saw_disabled or _control_exists(page, label)
        else:
            try:
                target.click()
            except Exception:
                pass  # node went stale between lookup and click; retry next sweep
            else:
                page.wait_for_timeout(2500)
                return True
        page.wait_for_timeout(500)

    if saw_disabled:
        log(f"[x] '{label}' never became enabled")
        return False
    page.keyboard.press("Enter")
    page.wait_for_timeout(2500)
    return True


def _controls(page, label):
    """Visible buttons whose text contains `label`."""
    out = []
    for cand in page.query_selector_all('[role="button"], button'):
        try:
            if cand.is_visible() and label.lower() in (cand.inner_text() or "").strip().lower():
                out.append(cand)
        except Exception:
            pass  # node went stale mid-scan
    return out


def _control_exists(page, label):
    return bool(_controls(page, label))


def _enabled_control(page, label):
    """First matching control that is actually actionable, else None."""
    for cand in _controls(page, label):
        try:
            gate = (cand.get_attribute("aria-disabled") or "").lower()
            disabled = cand.get_attribute("disabled") is not None or gate == "true"
            if not disabled:
                return cand
        except Exception:
            pass  # node went stale mid-check
    return None


def fill_setup(page, email_addr, password, handle=None):
    """"Set up your account": name = mailbox local part, username = 5 chars + digits,
    password = the mailbox password. Returns the username that was used."""
    name = derive_name(email_addr)
    user = handle or derive_handle_from_name(email_addr)

    def user_input():
        for sel in ('input[name="username"]', 'input[autocomplete="username"]',
                    'input[aria-label="Username"]'):
            el = page.query_selector(sel)
            if el and el.is_visible():
                return el
        return None

    def settle_for(getter, timeout=6000):
        """X reveals the next field a beat after the previous one is typed."""
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            el = getter()
            if el:
                return el
            page.wait_for_timeout(300)
        return None

    name_inp = settle_for(lambda: NAME_INPUT(page), 5000)
    if not name_inp:
        return None
    name_inp.fill(name)
    page.wait_for_timeout(500)

    # X sanitizes this field on change: whatever it kept is the truth, and it is
    # what the caller has to reuse if the pane is retried.
    landed = (name_inp.input_value() or "").strip()
    if landed != name:
        log(f"[setup] name {name!r} -> {landed!r} (X trimmed it)")
    if not landed:
        log("[setup] name field empty after fill; X rejected the value")
        return None
    name = landed

    user_inp = settle_for(user_input, 8000)
    if not user_inp:
        log("[setup] username field never appeared")
        return None
    user_inp.fill(user)
    page.wait_for_timeout(800)  # X validates availability as you type

    pass_inp = settle_for(lambda: PASS_INPUT(page), 8000)
    if not pass_inp:
        log("[setup] password field never appeared")
        return None
    pass_inp.fill(password)
    page.wait_for_timeout(300)

    log(f"[setup] name={name!r} username={user!r}")
    return user


def run_signup(args):
    email_addr, password, msa, client_id = load_account(args.account)
    log(f"[acct] {email_addr}")

    if args.mail_login:
        # --mail-login, carried over from outlook_newest.py: force the OAuth
        # device-code flow, ignoring (and replacing) the cached refresh token.
        access_token(email_addr, client_id, force_login=True, msa=msa)
        log("[mail] new refresh token cached")
        return 0

    if args.mail_check or args.mail:
        msgs = fetch_messages(email_addr, client_id, msa, count=args.count)
        if args.json:
            print(json.dumps(msgs, indent=2, ensure_ascii=False))
            return 0
        for m in reversed(msgs):
            log(f"[mail] «{m['subject']}» from <{m['from']}> {m['date']}")
        log(f"[mail-check] ok, {len(msgs)} message(s)")
        return 0

    # X weights signup by IP reputation. Failover across a dead/blocked proxy is
    # cheaper than rerunning by hand, so rotate on a 403/blocked stage.
    proxy = pick_proxy(args.proxy)
    rotation = [] if args.no_rotate else [p for p in load_proxies() if p != proxy]
    attempts = [(proxy, None)] + [(p, "rotate") for p in rotation[:args.rotate]]
    # Exit codes that mean "start over from the beginning":
    #   2 = proxy/edge blocked, 6 = face-liveness wall, 8 = browser died.
    # X decides per-session whether it demands a face scan, so a fresh run is a
    # legitimate retry and often lands on the code-only path.
    RETRY_FROM_START = {2, 6, 8}
    rc = 2
    runs = 0
    while runs < max(1, args.retries):
        runs += 1
        if runs > 1:
            log(f"[retry] restarting the flow from the beginning "
                f"(attempt {runs}/{args.retries}, last exit {rc})")
            time.sleep(args.retry_delay)
        for i, (proxy, _) in enumerate(attempts):
            rc = attempt(args, email_addr, password, msa, client_id, proxy)
            if rc != 2 or i + 1 == len(attempts):
                break
            log(f"[proxy] rotating off {urllib.parse.urlsplit(proxy).hostname} after block")
        if rc not in RETRY_FROM_START:
            break
        if runs >= max(1, args.retries):
            log(f"[retry] giving up after {runs} attempt(s); last exit {rc}")
    return rc


def launch(args, proxy):
    """Open a browser session. Returns (ctx, close_fn).

    Camoufox (patched Firefox) is the primary engine: it is not fingerprinted as
    automation at all, so X's Cloudflare edge serves the real signup pane where
    headless Chromium gets 403. Chromium stays as the fallback engine because
    Camoufox cannot exist on hosts that have no Firefox build.
    """
    px = parse_proxy(proxy) if proxy else None
    if args.engine == "camoufox":
        try:
            from camoufox.sync_api import Camoufox
            cm = Camoufox(headless=args.headless, humanize=not args.headless,
                          canvas_noise=False,  # stock canvas; noise reads as a hardened browser
                          proxy=px, geoip=False,
                          locale=["en-US", "en"], os="linux",
                          window=(1280, 900))
            b = cm.__enter__()
            # Camoufox hands back a Browser. Its implicit context (the one
            # browser.new_page() creates) refuses a second tab with "Please use
            # browser.new_context()", so make the context explicit up front:
            # Camoufox attaches its own defaults to new_context() too.
            ctx = b if hasattr(b, "pages") else b.new_context()
            return ctx, lambda: cm.__exit__(None, None, None)
        except ImportError:
            log("[engine] camoufox not installed -> falling back to chromium")
        except Exception as e:
            # missing binary, bad proxy, unlaunchable window: all mean "use the fallback"
            log(f"[engine] camoufox failed to launch ({type(e).__name__}) -> chromium")
        args.engine = "chromium"

    from playwright.sync_api import sync_playwright
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=args.headless, channel="chromium",
                                 args=CHROME_ARGS, proxy=px)
    ctx = browser.new_context(user_agent=UA, locale="en-US", timezone_id=TIMEZONE,
                              viewport={"width": 1280, "height": 900})
    ctx.add_init_script(STEALTH)

    def close():
        try:
            ctx.close()
            browser.close()
        finally:
            pw.stop()

    return ctx, close


def run_check(args):
    """Report whether a saved session for this account is still authenticated."""
    email_addr, password, msa, client_id = load_account(args.account)
    path = SESSION_DIR / f"{email_addr.split('@')[0]}.json"
    if not path.exists():
        log(f"[check] no saved session at {path}")
        return 3

    proxy = pick_proxy(args.proxy)
    from playwright.sync_api import sync_playwright
    try:
        ctx, close = launch(args, proxy)
    except Exception as e:
        log(f"[check] launch failed: {type(e).__name__}")
        return 2
    try:
        page = ctx.pages[0] if getattr(ctx, "pages", None) else None
        page = page or ctx.new_page()
        page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(4000)
        ok = logged_in(page)
        log(f"[check] {'still signed in' if ok else 'session expired'}")
        return 0 if ok else 9
    finally:
        try:
            close()
        except Exception:
            pass


def run_login(args):
    """Sign in to an existing account, then prove the session.

    X's login is NOT the mirror of signup: it serves email -> code -> /home.
    Password and face-liveness are conditional branches X decides per session,
    so they are handled if they appear and skipped if they do not.
    """
    email_addr, password, msa, client_id = load_account(args.account)
    log(f"[acct] {email_addr}")

    proxy = pick_proxy(args.proxy)
    rc = 2
    for attempt_no in range(1, max(1, args.retries) + 1):
        if attempt_no > 1:
            log(f"[retry] restarting login (attempt {attempt_no}/{args.retries})")
            time.sleep(args.retry_delay)
        rc = login_attempt(args, email_addr, password, msa, client_id, proxy)
        if rc not in (2, 6, 8):
            break
    return rc


def login_attempt(args, email_addr, password, msa, client_id, proxy):
    """One login attempt. Exit codes match run_signup's contract."""
    from playwright.sync_api import Error as PWError

    log(f"[proxy] {urllib.parse.urlsplit(proxy).hostname if proxy else 'direct'}")
    log(f"[engine] {args.engine}")

    try:
        ctx, close = launch(args, proxy)
    except Exception as e:
        log(f"[engine] launch failed: {type(e).__name__}: {e}")
        return 2

    try:
        page = ctx.pages[0] if getattr(ctx, "pages", None) else None
        page = page or ctx.new_page()
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90000)

        # The login SPA hydrates a beat after domcontentloaded.
        if not _visible(page, 'input[name="username_or_email"]', timeout=30000):
            log("[login] email field never appeared")
            page.screenshot(path="x_blocked.png")
            return 2

        # Baseline BEFORE submitting the address: X can mail the code within a
        # second, and anything that arrives after the UID set is captured must
        # still count as fresh.
        try:
            baseline = uid_baseline(email_addr, client_id, msa)
        except SystemExit:
            raise
        except Exception as e:
            log(f"[mail] cannot read the mailbox baseline ({type(e).__name__})")
            return 3

        page.fill('input[name="username_or_email"]', email_addr)
        page.keyboard.press("Enter")
        page.wait_for_timeout(3000)

        tried_password = False
        tried_knowledge = False
        for step in range(16):
            state = settle(page)
            stage = state["stage"]
            log(f"[login] stage={stage}" +
                (f" ({state['why']})" if state.get("why") else ""))

            if stage == "done":
                page.screenshot(path="x_login.png")
                if not logged_in(page):
                    log("[login] /home reached but no signed-in identity found")
                    return 9
                log(f"[login] signed in: {page.url}")
                if args.save_session:
                    save_session(page, email_addr)
                return 0
            if stage == "liveness":
                page.screenshot(path="x_liveness.png")
                log("[login] stopped: X requires a face-scan liveness check")
                log(f"[login] resume in a normal browser: {page.url}")
                return 6
            if stage == "blocked":
                page.screenshot(path="x_blocked.png")
                log(f"[login] blocked: {state.get('why')}")
                return 2
            if stage == "cookie_notice":
                _click(page, "Accept all cookies", timeout=6000)
                page.wait_for_timeout(2000)
            elif stage == "code":
                code = args.code or wait_for_code(email_addr, client_id, msa,
                                                  baseline=baseline,
                                                  timeout=args.code_timeout)
                if not code:
                    log("[login] no code arrived; pass --code XXXXXX once you have it")
                    page.screenshot(path="x_stuck.png")
                    return 3
                if not fill_code(page, code):
                    log("[login] verification-code field not found")
                    return 4
            elif stage == "knowledge_check":
                if tried_knowledge:
                    log("[login] username prompt did not advance")
                    page.screenshot(path="x_stuck.png")
                    return 4
                tried_knowledge = True
                if not fill_knowledge_check(page, email_addr, password,
                                            handle=args.handle):
                    log("[login] no way past the username prompt")
                    return 4
            elif stage == "password_create":
                if tried_password:
                    log("[login] password rejected")
                    page.screenshot(path="x_stuck.png")
                    return 4
                tried_password = True
                if not login_password(page, password):
                    log("[login] password field not found")
                    return 4
            elif stage == "captcha":
                page.screenshot(path="x_captcha.png")
                if args.headless:
                    log("[login] captcha, headless — rerun without --headless")
                    return 7
                log("[login] captcha on screen — solve it, then press Enter")
                try:
                    input()
                except EOFError:
                    pass
            else:
                page.screenshot(path="x_stuck.png")
                try:
                    Path("x_unknown.html").write_text(page.content(), encoding="utf-8")
                except Exception:
                    pass
                if not args.headless:
                    log("[login] unexpected pane — finish it in the window, then press Enter")
                    try:
                        input()
                    except EOFError:
                        pass
                else:
                    log(f"[login] unexpected pane at {page.url}")
                    return 4
        log("[login] ran out of steps")
        return 5
    except PWError as e:
        log(f"[login] browser died: {str(e).splitlines()[0]}")
        return 8
    finally:
        try:
            close()
        except Exception:
            pass


def report_result(email_addr, handle, dob, password):
    """Print the created account's identity. Never prints the password."""
    log("")
    log("[x] ================ REGISTERED ================")
    log(f"[x] email    : {email_addr}")
    log(f"[x] username : @{handle}")
    log(f"[x] profile  : https://x.com/{handle}")
    log(f"[x] birthday : {dob[0]}-{dob[1]:02d}-{dob[2]:02d}")
    log("[x] password : %d chars (from account.txt, value never printed)" % len(password))
    log("[x] ==========================================")


def save_session(page, email_addr):
    """Persist the authenticated cookie jar so later runs skip the code step.

    storage_state lives on the BrowserContext, which launch() does not hand back
    directly on the Camoufox path - reach it through the page.
    """
    SESSION_DIR.mkdir(exist_ok=True)
    path = SESSION_DIR / f"{email_addr.split('@')[0]}.json"
    try:
        page.context.storage_state(path=str(path))
        os.chmod(path, 0o600)
        log(f"[login] session saved to {path} (chmod 600)")
    except Exception as e:
        log(f"[login] could not save session: {type(e).__name__}")


def fill_knowledge_check(page, email_addr, password, handle=None):
    """Answer X's "Confirm your account" reassurance prompt.

    The field is labelled **Username** and X rejects the email address in it, so a
    candidate handle is used only when the caller knows one. Without it, take the
    "Use password" route, which is the path that always exists.

    Returns the stage X moved to, or None.
    """
    if handle:
        inp = _visible(page, 'input[name="challenge_response"]', 2000) or \
            _visible(page, 'input[autocomplete="username"]', 1000)
        if inp:
            inp.fill(handle)
            page.wait_for_timeout(400)
            if submit(page, "Continue"):
                log(f"[login] answered the username prompt with {handle!r}")
                nxt = settle(page)["stage"]
                if nxt != "knowledge_check":
                    return nxt
                log("[login] username rejected; falling back to the password route")

    if _click(page, "Use password", timeout=6000):
        page.wait_for_timeout(2500)
        log("[login] switched to the password route")
        return settle(page)["stage"]
    return None


def login_password(page, password):
    """Password pane. X uses new-password on signup, plain password on login."""
    inp = PASS_INPUT(page)
    if not inp:
        return False
    inp.fill(password)
    page.wait_for_timeout(300)
    if not submit(page, "Log in") and not submit(page, "Continue"):
        page.keyboard.press("Enter")
    page.wait_for_timeout(4000)
    return True


def attempt(args, email_addr, password, msa, client_id, proxy):
    """One signup attempt on one proxy. Returns an exit code; 2 means 'try another proxy'."""
    log(f"[proxy] {urllib.parse.urlsplit(proxy).hostname if proxy else 'direct'}")
    log(f"[engine] {args.engine}")

    # Baseline before any page interaction or browser start: X can mail the code
    # within a second of the address being typed, and this run only earned exit 3
    # once because the fresh code landed inside the baseline.
    try:
        baseline = uid_baseline(email_addr, client_id, msa)
    except SystemExit:
        raise
    except Exception as e:
        log(f"[mail] cannot read the mailbox baseline ({type(e).__name__}); "
            "refusing to guess freshness")
        return 3

    from playwright.sync_api import Error as PWError
    try:
        ctx, close = launch(args, proxy)
    except PWError as e:
        log(f"[engine] launch failed: {str(e).splitlines()[0]}")
        return 2  # let the caller rotate proxy/engine
    except Exception as e:
        log(f"[engine] launch failed: {type(e).__name__}: {e}")
        return 2
    try:
        page = ctx.pages[0] if getattr(ctx, "pages", None) else None
        page = page or ctx.new_page()
        page.goto(SIGNUP_URL, wait_until="domcontentloaded", timeout=90000)
        page.wait_for_timeout(3500)

        if args.dry_run:
            log(f"[dry-run] title={page.title()!r} url={page.url}")
            log(f"[dry-run] email field: {bool(page.query_selector('input[name=username_or_email]'))}")
            page.screenshot(path="x_signup.png")
            return 0

        if not submit_email(page, email_addr):
            log("[x] could not reach the email field (proxy never hydrated)")
            page.screenshot(path="x_blocked.png")
            if proxy and MARKER.exists() and MARKER.read_text(encoding="utf-8").strip() == proxy:
                MARKER.unlink()  # stop re-picking a proxy that cannot render X
                log("[proxy] dropped from marker; next run picks another")
            return 2

        # Pin once per attempt: a regenerated handle changes mid-flow and X then
        # rejects the "already taken" name on the retry.
        handle = args.handle or derive_handle_from_name(email_addr)
        dob = args.dob or random_dob()
        log(f"[plan] handle={handle!r} dob={dob[0]}-{dob[1]:02d}-{dob[2]:02d}")
        # Each pane gets one action, then the loop re-classifies. Progress is
        # tracked by pane signature so a stalled pane is not re-filled forever.
        seen_progress = {}  # (stage, url) -> attempts
        for _ in range(20):  # extra room: preferences + a chain of Skip panes
            state = settle(page)
            stage = state["stage"]
            log(f"[x] stage={stage}" + (f" ({state['why']})" if state.get("why") else ""))

            # A pane that repeats without advancing means our action did not take.
            # Allow one retry per pane, then stop instead of burning 14 iterations.
            sig = (stage, page.url)
            if stage not in ("done", "blocked", "liveness", "captcha"):
                seen_progress[sig] = seen_progress.get(sig, 0) + 1
                # Skip panes are a chain of near-identical screens on the same
                # route, so give them more room than a real form.
                budget = 4 if stage == "optional" else 2
                if seen_progress[sig] > budget:
                    log(f"[x] stuck on {stage}: no progress after {budget} attempts")
                    page.screenshot(path="x_stuck.png")
                    return 4

            if stage == "done":
                page.screenshot(path="x_done.png")
                if not logged_in(page):
                    log("[x] /home reached but no account identity found; "
                        "treating as unverified, not created")
                    return 9
                log(f"[x] account created and signed in: {page.url}")
                report_result(email_addr, handle, dob, password)
                return 0
            elif stage == "blocked":
                page.screenshot(path="x_blocked.png")
                log(f"[x] blocked: {state.get('why')}")
                return 2
            elif stage == "liveness":
                page.screenshot(path="x_liveness.png")
                log(f"[x] stopped: {state['why']}")
                log("[x] registration is NOT complete: X requires the face check before "
                    "the account is usable. Nothing below is verified.")
                log(f"[x] resume in a normal browser: {page.url}")
                return 6
            elif stage == "captcha":
                page.screenshot(path="x_captcha.png")
                if args.headless:
                    log(f"[x] stopped: {state['why']} (rerun without --headless to solve it)")
                    return 7
                log("[x] captcha on screen — solve it in the window, then press Enter")
                try:
                    input()
                except EOFError:
                    pass
                pass  # continue  # re-classify after the human solved the captcha
            elif stage == "cookie_notice":
                _click(page, "Accept all cookies", timeout=6000)
                page.wait_for_timeout(2000)
                pass  # continue  # re-classify; falling through would land on "unknown"
            elif stage == "code":
                code = args.code or wait_for_code(email_addr, client_id, msa,
                                                  baseline=baseline,
                                                  timeout=args.code_timeout)
                if not code:
                    log("[x] no code arrived; pass --code XXXXXX once you have it")
                    page.screenshot(path="x_stuck.png")
                    return 3
                if not fill_code(page, code):
                    log("[x] verification-code field not found")
                    page.screenshot(path="x_stuck.png")
                    return 4
            elif stage == "password_create":
                if not fill_password(page, password):
                    log("[x] password field not found")
                    return 4
            elif stage == "dob":
                if not fill_dob(page, dob=dob):
                    log("[x] birthday selects not found")
                    page.screenshot(path="x_stuck.png")
                    return 4
            elif stage == "setup":
                # pass the pinned handle: regenerating it each pass makes X reject
                # the retry as "already taken" and the pane never advances
                used = fill_setup(page, email_addr, password, handle=handle)
                if not used:
                    log("[x] name/username/password fields not all present")
                    page.screenshot(path="x_stuck.png")
                    return 4
                if not submit(page, "Continue"):
                    log("[x] setup pane did not accept the submit")
                    page.screenshot(path="x_stuck.png")
                    return 4
            elif stage == "preferences":
                if not fill_preferences(page):
                    log("[x] privacy preferences did not accept the submit")
                    page.screenshot(path="x_stuck.png")
                    return 4
            elif stage == "optional":
                # X chains several optional panes; each one is just a Skip. The
                # retry budget in the loop covers the repeats, so a fresh URL
                # keeps going and a pane that will not advance is caught.
                if not skip_optional(page):
                    page.screenshot(path="x_stuck.png")
                    return 4
            elif stage == "username":
                if not fill_username(page, handle):
                    log("[x] username field not found")
                    return 4
                submit(page, "Continue")
            else:  # unknown: keep the DOM so the next slice is obvious
                page.screenshot(path="x_stuck.png")
                dump = Path("x_unknown.html")
                try:
                    dump.write_text(page.content(), encoding="utf-8")
                except Exception:
                    pass
                log(f"[x] unknown pane dumped to {dump} — url={page.url}")
                if not args.headless:
                    log("[x] unexpected pane — finish it in the visible browser, then press Enter")
                    try:
                        input()
                    except EOFError:
                        pass
                else:
                    log(f"[x] unexpected pane at {page.url}")
                    return 4
        log("[x] ran out of steps without finishing")
        return 5
    except PWError as e:
        log(f"[x] browser died: {str(e).splitlines()[0]}")
        return 8
    except (OSError, ValueError) as e:
        # mail/network/cache trouble during the poll must surface as a failure,
        # not vanish under the Playwright-only handler
        log(f"[x] network/mail failure: {type(e).__name__}: {e}")
        return 3
    finally:
        try:
            close()
        except Exception as e:
            log(f"[cleanup] {type(e).__name__} while closing the browser")


# ── CodeBuddy connect: X session -> 9Router provider connection ───────────────
# 9Router owns the provider record, and its device-code API performs the whole
# OAuth exchange. We drive only the pane that needs a human (codebuddy.ai/login)
# and then poll. Writing providerConnections directly — what next-step/ did —
# breaks as soon as the schema gains a NOT NULL column, and races the server.
NINEROUTER = os.environ.get("NINEROUTER_URL", "http://localhost:20128").rstrip("/")
NINEROUTER_PW = os.environ.get("NINEROUTER_PASSWORD", "123456")  # 9Router's local default
CB_PROVIDER = "codebuddy-intl"
NR_PW_FIELD = "password"
# The login pane is an iframe and its ids change between builds (next-step/ still
# names #social-google, which is gone), so match the accessible name instead.
# Longest first: "Sign up with X" contains "X".
CB_X_LABELS = ("Sign up with X", "Sign in with X", "Continue with X", "Twitter", "X")
# The login pane is a Keycloak realm: the provider rows are <a id="social-*"> and
# the visible text is an icon, so the id is the only stable handle.
CB_X_SELECTORS = ("a#social-twitter", "a#social-x", "#social-twitter", "#social-x")
CB_POLICY_SELECTOR = 'input[id^="agree-policy"]'
# After the provider row, Keycloak raises an in-frame dialog whose buttons are
# literally "Confirm" / "Cancel"; X adds its own wording on the OAuth page.
CB_CONSENT_LABELS = ("Confirm", "Authorize app", "Authorize", "Continue as", "Continue")
# x.com/i/oauth2/authorize renders a landing page while signed out: the
# login form only exists after its own "Log in" button is clicked.
X_LOGIN_LABELS = ("Log in", "Sign in", "Login")

def nr_request(path, cookie=None, payload=None, timeout=25):
    """JSON call against the local 9Router API."""
    headers = {"Accept": "application/json"}
    data = None
    if payload is not None:
        data = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    if cookie:
        headers["Cookie"] = cookie
    url = f"{NINEROUTER}{path}"
    with urllib.request.urlopen(urllib.request.Request(url, data, headers),
                                timeout=timeout) as r:
        return json.load(r)

def nr_login():
    """Authenticate to 9Router; returns a Cookie header value."""
    creds = {NR_PW_FIELD: NINEROUTER_PW}
    body = json.dumps(creds).encode()
    req = urllib.request.Request(
        f"{NINEROUTER}/api/auth/login", body,
        {"Content-Type": "application/json", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        reply = json.load(r)
        jars = r.headers.get_all("Set-Cookie") or []
    if not reply.get("success"):
        raise RuntimeError(reply.get("error") or "9Router rejected the credentials")
    pairs = [c.split(";", 1)[0] for c in jars if "=" in c]
    if not pairs:
        raise RuntimeError("9Router returned no auth cookie")
    return "; ".join(pairs)

def cb_device_code(cookie):
    d = nr_request(f"/api/oauth/{CB_PROVIDER}/device-code", cookie)
    if not d.get("device_code") or not d.get("verification_uri"):
        raise ValueError(f"device-code payload missing fields: {sorted(d)}")
    return d

def cb_connected(cookie):
    """Emails already connected under this provider — poll always adds a row."""
    try:
        d = nr_request("/api/providers", cookie)
    except Exception:
        return set()
    return {(c.get("email") or "").lower() for c in d.get("connections") or []
            if c.get("provider") == CB_PROVIDER and c.get("email")}

def cb_pick_label(labels, available):
    """First wanted label present in `available`, longest first so specific
    wording beats its own substring ("Sign up with X" before "X")."""
    exact = {a.strip().lower(): a for a in available if a.strip()}
    for want in sorted(labels, key=len, reverse=True):
        if want.lower() in exact:
            return exact[want.lower()]
    for want in sorted(labels, key=len, reverse=True):
        for name in available:
            if want.lower() in name.lower():
                return name
    return None

def _cb_root(page):
    """The Keycloak login pane is an iframe; X's own pages are not."""
    try:
        if not page.query_selector("iframe"):
            return page
        return page.frame_locator("iframe").first
    except Exception:
        return page

def cb_buttons(page):
    """Accessible names of the clickable rows; the provider rows are links."""
    names = []
    for root in (page, _cb_root(page)):
        for role in ("button", "link"):
            try:
                names += [t.strip() for t in root.get_by_role(role).all_inner_texts()]
            except Exception:
                continue
    return [n for n in names if n]

def cb_click(page, name, timeout=15000):
    for root in (page, _cb_root(page)):
        for role in ("button", "link"):
            try:
                loc = root.get_by_role(role, name=name, exact=False).first
                loc.wait_for(state="visible", timeout=timeout)
                loc.click()
                return True
            except Exception:
                continue
    return False

def cb_tick_policy(page):
    """The sign-in rows stay gated until the agreement boxes are ticked.

    Keycloak hides the real inputs behind styled labels, so force the check
    instead of asking whether the input is visible.
    """
    ticked = 0
    for fr in page.frames:
        try:
            boxes = fr.locator(CB_POLICY_SELECTOR)
            for i in range(boxes.count()):
                box = boxes.nth(i)
                try:
                    if not box.is_checked():
                        box.check(force=True)
                        ticked += 1
                except Exception:
                    continue
        except Exception:
            continue
    return ticked

def cb_click_x(page):
    """Click the X provider row: stable Keycloak id first, then accessible name."""
    for fr in page.frames:
        for sel in CB_X_SELECTORS:
            try:
                loc = fr.locator(sel).first
                if loc.count() and loc.is_visible():
                    loc.click()
                    return sel
            except Exception:
                continue
    name = cb_pick_label(CB_X_LABELS, cb_buttons(page))
    return name if name and cb_click(page, name) else None

def x_login_inline(page, args, email_addr, password, msa, client_id):
    """Sign in to X inside the tab the OAuth flow opened (--fresh-x).

    The reversed flow: CodeBuddy's authorize page comes first, "Sign up with X"
    starts X's OAuth, and only then does X ask for credentials. On
    x.com/i/oauth2/authorize while signed out there is a landing page, so the
    login form appears only after its own "Log in" button is clicked.
    """
    try:
        baseline = uid_baseline(email_addr, client_id, msa)
    except SystemExit:
        raise
    except Exception:
        baseline = set()

    def on_x(p):
        return "x.com" in p.url or "twitter.com" in p.url

    submitted = asked_pw = 0
    for _ in range(24):
        for p in reversed(page.context.pages):
            if on_x(p) and p is not page:
                log(f"[cb] following the X window: {p.url[:70]}")
                page = p
                break

        # Redirected back to CodeBuddy/Keycloak: X's part is done.
        if page.url and not on_x(page):
            log(f"[cb] X finished, back at {page.url[:60]}")
            return True

        state = settle(page)
        stage = state["stage"]
        log(f"[cb] x stage={stage}" + (f" ({state['why']})" if state.get("why") else ""))

        if stage == "liveness":
            log("[cb] X wants a face scan here; not usable")
            return False
        if stage == "blocked":
            log(f"[cb] X blocked the login: {state.get('why')}")
            return False
        if stage == "captcha":
            log("[cb] captcha on the X pane — solve it in the window")
            try:
                input()
            except EOFError:
                pass
            continue

        if _visible(page, 'input[name="username_or_email"]', 1500):
            if submitted:
                log("[cb] X rejected the address")
                return False
            submitted += 1
            page.fill('input[name="username_or_email"]', email_addr)
            page.keyboard.press("Enter")
            page.wait_for_timeout(3000)
            continue
        if stage == "knowledge_check":
            if not fill_knowledge_check(page, email_addr, password, handle=args.handle):
                return False
            continue
        if stage == "code":
            code = args.code or wait_for_code(email_addr, client_id, msa,
                                              baseline=baseline, timeout=args.code_timeout)
            if not code:
                log("[cb] no X code arrived")
                return False
            fill_code(page, code)
            continue
        if stage == "password_create":
            if asked_pw:
                log("[cb] X rejected the password")
                return False
            asked_pw += 1
            login_password(page, password)
            continue

        # Landing page / consent: whatever this pane offers, take it.
        names = cb_buttons(page)
        for label in CB_CONSENT_LABELS + X_LOGIN_LABELS:
            if label in names and cb_click(page, label, timeout=4000):
                log(f"[cb] clicked {label!r} on X")
                page.wait_for_timeout(3000)
                break
        else:
            page.wait_for_timeout(3000)
    log("[cb] X login did not settle")
    return False

def connect_codebuddy(args, page, email_addr, password, msa, client_id):
    """Authorize codebuddy-intl in a new tab of the existing context."""
    try:
        cookie = nr_login()
    except Exception as e:
        log(f"[cb] 9Router login failed at {NINEROUTER}: {type(e).__name__}: {e}")
        return 10

    if email_addr.lower() in cb_connected(cookie) and not args.force:
        log(f"[cb] already connected: {email_addr} (--force adds another)")
        return 0

    try:
        dev = cb_device_code(cookie)
    except Exception as e:
        log(f"[cb] device-code request failed: {type(e).__name__}: {e}")
        return 10

    tab = page.context.new_page()  # same context => the X cookies come along

    def dump():
        try:
            Path("x_unknown.html").write_text(tab.content(), encoding="utf-8")
            tab.screenshot(path="x_stuck.png")
        except Exception:
            pass

    try:
        log(f"[cb] authorize at {dev['verification_uri']}")
        tab.goto(dev["verification_uri"], wait_until="domcontentloaded", timeout=90000)
        tab.wait_for_timeout(5000)
        ticked = cb_tick_policy(tab)
        if ticked:
            log(f"[cb] accepted {ticked} policy box(es)")
            tab.wait_for_timeout(800)

        before = len(page.context.pages)
        name = cb_click_x(tab)
        if not name:
            dump()
            log("[cb] no X sign-in row found — dumped to x_unknown.html")
            return 4
        log(f"[cb] clicked the X sign-in row ({name})")

        if getattr(args, "fresh_x", False):
            # Keycloak raises a Confirm dialog before it hands off to X.
            tab.wait_for_timeout(3000)
            for label in CB_CONSENT_LABELS:
                if cb_click(tab, label, timeout=2500):
                    log(f"[cb] confirmed {label!r}")
                    break
            tab.wait_for_timeout(5000)
            pages = page.context.pages
            xpage = pages[-1] if len(pages) > before else tab
            log(f"[cb] X login in {xpage.url[:80]}")
            if not x_login_inline(xpage, args, email_addr, password, msa, client_id):
                dump()
                return 4
            log("[cb] X signed in inside the OAuth flow")

        deadline = time.time() + args.cb_timeout
        asked_password = False
        last_note = 0
        while time.time() < deadline:
            try:
                r = nr_request(f"/api/oauth/{CB_PROVIDER}/poll", cookie,
                               {"deviceCode": dev["device_code"]})
            except Exception:
                r = {}
            if r.get("success"):
                conn = r.get("connection") or {}
                log(f"[cb] connected {CB_PROVIDER}: {conn.get('email') or email_addr}")
                return 0
            if r.get("fatal"):
                dump()
                log(f"[cb] authorize failed: {r.get('errorDescription') or r.get('error')}")
                return 11
            # The tab may close itself once CodeBuddy finishes the redirect, and
            # the Keycloak pane has its own password field — only touch the tab
            # while it is open, and only fill a password on X's own page.
            try:
                if not tab.is_closed() and time.time() - last_note > 10:
                    last_note = time.time()
                    cur = tab.url
                    log(f"[cb] waiting: {cur[:90]}")
                    # CodeBuddy's account picker needs one more click before it
                    # returns to the plugin redirect that authorizes the state.
                    if "/login/select" in cur or "/started" in cur:
                        for label in ("Continue", "Confirm", "Select", "Authorize", "Allow"):
                            if cb_click(tab, label, timeout=2500):
                                log(f"[cb] advanced the CodeBuddy picker ({label!r})")
                                break
                if not tab.is_closed():
                    on_x = "x.com" in tab.url or "twitter.com" in tab.url
                    if not asked_password and on_x and login_password(tab, password):
                        asked_password = True
                        log("[cb] X asked for the password again; supplied it")
                    else:
                        for label in CB_CONSENT_LABELS:
                            if cb_click(tab, label, timeout=2000):
                                log(f"[cb] confirmed {label!r}")
                                break
            except Exception as e:
                log(f"[cb] tab interaction stopped: {type(e).__name__}")
            time.sleep(5)

        dump()
        log(f"[cb] authorization never completed after {args.cb_timeout}s")
        return 11
    finally:
        try:
            tab.close()
        except Exception:
            pass

def run_codebuddy(args):
    """--login-codebuddy: connect the signed-in X account to codebuddy-intl."""
    email_addr, password, msa, client_id = load_account(args.account)
    log(f"[acct] {email_addr}")

    # --fresh-x reverses the order: the OAuth page comes first and X is signed in
    # inside that tab, so no saved cookie jar is needed.
    path = None if getattr(args, "fresh_x", False) else \
        SESSION_DIR / f"{email_addr.split('@')[0]}.json"
    if path is not None and not path.exists():
        log(f"[cb] no saved X session at {path}")
        log("[cb] run first: xreg.py login --save-session --proxy ''")
        return 3
    state = {}
    if path is not None:
        try:
            state = json.loads(path.read_text())
        except Exception as e:
            log(f"[cb] unreadable session file: {type(e).__name__}")
            return 3

    proxy = pick_proxy(args.proxy)
    try:
        ctx, close = launch(args, proxy)
    except Exception as e:
        log(f"[engine] launch failed: {type(e).__name__}: {e}")
        return 2
    try:
        page = ctx.pages[0] if getattr(ctx, "pages", None) else None
        page = page or ctx.new_page()
        if state:
            page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=90000)
            try:
                page.context.add_cookies(state.get("cookies") or [])
            except Exception as e:
                log(f"[cb] could not restore the X session: {type(e).__name__}")
                return 3
            page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(4000)
            if not logged_in(page):
                log("[cb] saved session is not signed in — rerun login --save-session")
                return 9
            log("[cb] X session restored")
        return connect_codebuddy(args, page, email_addr, password, msa, client_id)
    finally:
        try:
            close()
        except Exception:
            pass

def _self_check():
    """Offline checks for the two pure helpers."""
    assert extract_code("Your X confirmation code is 832104. Enter it.") == "832104"
    assert extract_code("Hi! Use 998812 to confirm your email.") == "998812"
    assert extract_code("Confirm your email address\nCode 771234\nThis code expires") == "771234"
    assert extract_code("no code here") is None
    assert extract_code("Verification code: 4451 2288") is None  # X codes are unspaced
    assert derive_handle("OgdenIntrieri341@outlook.com") == "OgdenIntrieri34"  # 15-char cap
    assert derive_handle("a.b-c@x.com") == "abc"

    # 3: multipart HTML fallback must survive the exhausted walk() iterator
    import email as _email
    raw = ("MIME-Version: 1.0\r\n"
           "Content-Type: multipart/alternative; boundary=B\r\n\r\n"
           "--B\r\nContent-Type: text/plain\r\n\r\n\r\n"
           "--B\r\nContent-Type: text/html\r\n\r\n"
           "<p>Your code is <b>778899</b></p>\r\n--B--\r\n")
    msg = _email.message_from_string(raw)
    body = body_text(msg)
    assert "778899" in body, f"html fallback lost the body: {body!r}"

    # 4: only X signup mail qualifies
    assert is_signup_mail({"from": "X <info@x.com>", "to": "a@outlook.com",
                           "subject": "1234 is your X confirmation code"}, "a@outlook.com")
    assert not is_signup_mail({"from": "spam@evil.com", "to": "a@outlook.com",
                               "subject": "123456"}, "a@outlook.com")
    assert not is_signup_mail({"from": "X <info@x.com>", "to": "other@x.com",
                               "subject": "code 123456"}, "a@outlook.com")

    # 7: exact DOB comparison, "1" must not match "21"
    assert _same_value("21", "1") is False
    assert _same_value("March", "March") is True
    assert _same_value("March v", "March") is True
    assert _same_value("16", "16") is True

    # 7: --dob rejects impossible dates and out-of-range years
    assert parse_dob_arg("1997-03-29") == (1997, 3, 29)
    for bad in ("1997-02-30", "1997-13-01", "1850-01-01", "1997-3", "not-a-date"):
        try:
            parse_dob_arg(bad)
        except Exception:
            pass
        else:
            raise AssertionError(f"parse_dob_arg accepted {bad!r}")
    h = derive_handle_from_name("OgdenIntrieri341@outlook.com")
    assert h.startswith("Ogden") and len(h) == 9 and h[5:].isdigit(), h
    # name must be digit-free: X rejects digits in the display-name field
    for addr in ("OgdenIntrieri341@outlook.com", "briyatylie3292@outlook.com",
                 "NevelMulderig544@outlook.com", "x_99@outlook.com",
                 "12345@x.com", "jane.doe@x.com"):
        nm = derive_name(addr)
        assert not any(ch.isdigit() for ch in nm), f"digits in name for {addr}: {nm!r}"
        assert 2 <= len(nm) <= NAME_MAX, f"bad name length for {addr}: {nm!r}"
        assert nm.isascii() and " " not in nm[:0], f"non-ascii name for {addr}: {nm!r}"
    assert derive_name("OgdenIntrieri341@outlook.com") == "Ogden Intrieri"
    assert derive_name("briyatylie3292@outlook.com") == "Briyatylie"
    assert derive_name("12345@x.com") == NAME_FALLBACK

    # submit()/_enabled_control: a gated control must not be picked as actionable
    class _El:
        def __init__(self, text, gated=False):
            self._t, self._g = text, gated
        def inner_text(self):
            return self._t
        def is_visible(self):
            return True
        def get_attribute(self, name):
            return "true" if (name == "aria-disabled" and self._g) else None

    class _Page:
        def __init__(self, els):
            self.els = els
        def query_selector_all(self, _sel):
            return self.els

    assert _enabled_control(_Page([_El("Continue", gated=True)]), "Continue") is None
    assert _enabled_control(_Page([_El("Continue", gated=False)]), "Continue") is not None
    assert _enabled_control(_Page([_El("Sign up", gated=False)]), "Continue") is None
    assert _control_exists(_Page([_El("Continue", gated=True)]), "Continue") is True

    # skip_label: the specific label must beat its own substring, or "Skip for
    # now" would be dismissed by the bare-"Skip" branch.
    assert skip_label(_Page([_El("Skip for now")])) == "Skip for now"
    assert skip_label(_Page([_El("Not now")])) == "Not now"
    assert skip_label(_Page([_El("Skip")])) == "Skip"
    assert skip_label(_Page([_El("Continue")])) is None
    # a gated longer label must fall through to the enabled shorter one
    assert skip_label(_Page([_El("Skip for now", gated=True), _El("Skip")])) == "Skip"
    for _ in range(500):  # 500 random dates must always be legal
        y, m, d = random_dob()
        assert 1990 <= y <= 2005 and 1 <= m <= 12 and 1 <= d <= 31
        import datetime
        datetime.date(y, m, d)  # raises on 30 Feb / 31 Apr

    # cb_pick_label: specific wording must beat its own substring, and a pane
    # that only offers Google must not be mistaken for an X button.
    assert cb_pick_label(CB_X_LABELS, ["Google", "X"]) == "X"
    assert cb_pick_label(CB_X_LABELS, ["Sign up with X", "X"]) == "Sign up with X"
    assert cb_pick_label(CB_X_LABELS, ["Continue with Google"]) is None
    assert cb_pick_label(CB_CONSENT_LABELS, ["Authorize app"]) == "Authorize app"
    assert cb_pick_label(CB_CONSENT_LABELS, ["Cancel", "Confirm"]) == "Confirm"
    assert cb_pick_label(CB_CONSENT_LABELS, ["Cancel"]) is None
    print("self-check OK")


def main():
    ap = argparse.ArgumentParser(
        prog="xreg.py",
        description="Register or sign in to an X account from account.txt.",
        epilog="No subcommand registers a new account (the default flow).")
    ap.add_argument("command", nargs="?", choices=("signup", "login", "check"),
                    default="signup",
                    help="signup (default) | login | check")
    ap.add_argument("--account", default="account.txt")
    ap.add_argument("--headless", action="store_true",
                    help="hide the browser (default visible: X's edge 403s headless more often)")
    ap.add_argument("--dry-run", action="store_true", help="open the signup page only")
    ap.add_argument("--mail-check", action="store_true", help="only prove the mailbox is readable")
    ap.add_argument("--mail", action="store_true", help="dump the newest messages and exit")
    ap.add_argument("--json", action="store_true", help="with --mail/--mail-check: JSON output")
    ap.add_argument("-n", "--count", type=int, default=1, help="messages to read with --mail")
    ap.add_argument("--mail-login", action="store_true",
                    help="force the OAuth device-code flow and rewrite the token cache")
    ap.add_argument("--code", help="verification code, if you already have it")
    ap.add_argument("--handle", help="desired X handle")
    ap.add_argument("--dob", type=parse_dob_arg,
                    help="override birthday as YYYY-MM-DD (default: random 1990-2005)")
    ap.add_argument("--code-timeout", type=int, default=180)
    ap.add_argument("--self-check", action="store_true", help="run the offline helper checks")
    ap.add_argument("--engine", choices=("camoufox", "chromium"), default="camoufox",
                    help="camoufox (patched Firefox, primary) or chromium (fallback)")
    ap.add_argument("--proxy", default=None,
                    help="proxy url; omit for auto-pick from PROXIES.txt, "
                         "pass --proxy '' to go direct")
    ap.add_argument("--retries", type=int, default=3,
                    help="full-flow attempts when blocked by liveness/edge/browser death")
    ap.add_argument("--retry-delay", type=float, default=8.0,
                    help="seconds to wait between full-flow attempts")
    ap.add_argument("--rotate", type=int, default=3,
                    help="how many spare proxies to try after a block (0 disables)")
    ap.add_argument("--no-rotate", action="store_true", help="never switch proxy mid-run")
    ap.add_argument("--save-session", action="store_true",
                    help="login: persist the cookie jar for reuse")
    ap.add_argument("--login-codebuddy", action="store_true",
                    help="connect this X account to 9Router's codebuddy-intl provider")
    ap.add_argument("--fresh-x", action="store_true",
                    help="with --login-codebuddy: log in to X inside the OAuth tab")
    ap.add_argument("--force", action="store_true",
                    help="with --login-codebuddy: connect even if already connected")
    ap.add_argument("--cb-timeout", type=int, default=300,
                    help="seconds to wait for the CodeBuddy authorization")
    a = ap.parse_args()
    if a.self_check:
        return _self_check()
    if a.login_codebuddy:
        return run_codebuddy(a)
    if a.command == "login":
        return run_login(a)
    if a.command == "check":
        return run_check(a)
    return run_signup(a)


if __name__ == "__main__":
    raise SystemExit(main())
