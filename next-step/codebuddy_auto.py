"""
=============================================================================
 9Router CodeBuddy (codebuddy-intl) Auto-Connector & Bulk Google Register
 Standalone Distribution Edition
=============================================================================
"""

from __future__ import annotations

import base64
import json
import os
import re
import ssl
import sys
import threading
import time
import urllib.parse
import urllib.request
import uuid
from datetime import datetime, timezone
from pathlib import Path
import sqlite3

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Determine base directory dynamically
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
GOOGLE_TXT_PATH = BASE_DIR / "Google.txt"
LOGS_DIR = BASE_DIR / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)


def load_config() -> dict:
    default_config = {
        "webshare_api_key": "",
        "proxy_rotate_every_accounts": 3,
        "delay_between_accounts_seconds": 3,
        "default_headless": True
    }
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                default_config.update(data)
        except Exception:
            pass
    return default_config


def get_9router_db_path() -> Path:
    """Discover 9Router SQLite database path automatically."""
    candidates = []
    if sys.platform == "win32":
        appdata = os.getenv("APPDATA")
        if appdata:
            candidates.append(Path(appdata) / "9router" / "db" / "data.sqlite")
        localappdata = os.getenv("LOCALAPPDATA")
        if localappdata:
            candidates.append(Path(localappdata) / "9router" / "db" / "data.sqlite")
        userprofile = os.getenv("USERPROFILE")
        if userprofile:
            candidates.append(Path(userprofile) / ".9router" / "db" / "data.sqlite")
    else:
        home = Path.home()
        candidates.append(home / ".config" / "9router" / "db" / "data.sqlite")
        candidates.append(home / ".9router" / "db" / "data.sqlite")

    for p in candidates:
        if p.exists():
            return p

    fallback = Path.home() / "AppData" / "Roaming" / "9router" / "db" / "data.sqlite"
    fallback.parent.mkdir(parents=True, exist_ok=True)
    return fallback


def request_codebuddy_state() -> tuple[str, str]:
    """Request official OAuth state from CodeBuddy."""
    url = (
        "https://www.codebuddy.ai/v2/plugin/auth/state"
        "?ideType=IDE&ideName=IDE&ideVersion=2.108.1&pluginVersion=2.108.1&platform=win32"
    )
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    }
    req = urllib.request.Request(url, data=b"", headers=headers, method="POST")
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, context=ctx, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        res_data = data.get("data", {})
        state = res_data.get("state")
        auth_url = res_data.get("authUrl")
        if not state or not auth_url:
            raise ValueError(f"Respon tidak valid dari CodeBuddy: {data}")
        return state, auth_url


def poll_codebuddy_token(state: str, timeout: int = 180, interval: int = 2) -> dict:
    """Poll CodeBuddy for authorization token completion."""
    url = f"https://www.codebuddy.ai/v2/plugin/auth/token?state={state}"
    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    }
    deadline = time.time() + timeout
    ctx = ssl.create_default_context()

    while time.time() < deadline:
        try:
            req = urllib.request.Request(url, headers=headers, method="GET")
            with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                token_data = data.get("data", {})
                if token_data and token_data.get("accessToken"):
                    return token_data
        except urllib.error.HTTPError as e:
            if e.code not in [404, 400, 429]:
                pass
        except Exception:
            pass
        time.sleep(interval)
    return {}


def get_user_resource(access_token: str) -> dict:
    url = "https://www.codebuddy.ai/v2/billing/meter/get-user-resource"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
    }
    ctx = ssl.create_default_context()
    try:
        req = urllib.request.Request(url, headers=headers, method="GET")
        with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return {}


def insert_codebuddy_to_9router(token_data: dict, email_hint: str = "") -> bool:
    db_path = get_9router_db_path()
    if not db_path.exists():
        print(f"[!] Database 9Router tidak ditemukan di: {db_path}")
        return False

    access_token = token_data.get("accessToken", "")
    refresh_token = token_data.get("refreshToken", "")
    expires_in = token_data.get("expiresIn", 86400 * 365)

    res_info = get_user_resource(access_token)
    u_data = res_info.get("data", {}) if isinstance(res_info, dict) else {}
    email = u_data.get("email") or u_data.get("userEmail") or email_hint or "CodeBuddy User"
    user_id = u_data.get("userId") or u_data.get("user_id") or ""

    now_utc = datetime.now(timezone.utc)
    expires_at = datetime.fromtimestamp(now_utc.timestamp() + expires_in, tz=timezone.utc).isoformat()
    now_iso = now_utc.isoformat().replace("+00:00", "Z")

    conn_data = {
        "accessToken": access_token,
        "refreshToken": refresh_token,
        "idToken": None,
        "tokenType": "Bearer",
        "expiresIn": expires_in,
        "expiresAt": expires_at,
        "scope": "",
        "email": email,
        "displayName": f"CodeBuddy ({email})",
        "testStatus": "active",
        "lastError": None,
        "lastErrorAt": None,
        "errorCode": None,
        "backoffLevel": 0,
        "providerSpecificData": {
            "userId": user_id,
            "email": email,
            "connectionProxyEnabled": False,
            "connectionProxyUrl": "",
            "connectionNoProxy": ""
        }
    }

    try:
        conn = sqlite3.connect(str(db_path))
        cur = conn.cursor()
        cur.execute("SELECT id, data FROM providerConnections WHERE provider = 'codebuddy-intl'")
        rows = cur.fetchall()

        matched_id = None
        for cid, d_str in rows:
            try:
                d = json.loads(d_str)
                em = d.get("email") or d.get("displayName")
                if em and email.lower() in em.lower():
                    matched_id = cid
                    break
            except Exception:
                pass

        if matched_id:
            cur.execute(
                "UPDATE providerConnections SET data = ?, updatedAt = ? WHERE id = ?",
                (json.dumps(conn_data), now_iso, matched_id)
            )
            print(f"[9Router DB] Koneksi diperbarui untuk: {email}")
        else:
            new_id = str(uuid.uuid4())
            cur.execute(
                "INSERT INTO providerConnections (id, provider, data, createdAt, updatedAt) VALUES (?, ?, ?, ?, ?)",
                (new_id, "codebuddy-intl", json.dumps(conn_data), now_iso, now_iso)
            )
            cur.execute("SELECT COUNT(*) FROM providerConnections WHERE provider = 'codebuddy-intl'")
            total = cur.fetchone()[0]
            print(f"[9Router DB] Berhasil menambahkan koneksi CodeBuddy #{total}: {email}")

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"[!] Error simpan ke database 9Router: {e}")
        return False


def extract_google_2fa_code(page) -> str | None:
    """Extract 2-digit 2FA prompt code from Google login page."""
    try:
        el = page.locator("[data-sign-in-code]").first
        if el.count() > 0:
            code = el.get_attribute("data-sign-in-code")
            if code:
                return code.strip()
    except Exception:
        pass

    try:
        elems = page.locator("div[role='presentation'] div, span.x7hBbf, div.x7hBbf, div.c3Hknd, div[jsname='bN97Pc'], div.l07pSd").all_inner_texts()
        for t in elems:
            tc = t.strip()
            if tc.isdigit() and 1 <= len(tc) <= 3:
                return tc
    except Exception:
        pass

    try:
        text = page.locator("#view_container, body").inner_text()
        m = re.search(r'(?:tap|ketuk|pilih|match|press|number|angka|nomor)\s*[:\n\s]*(\d{1,3})', text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        if any(w in text.lower() for w in ["check your phone", "periksa ponsel", "notification", "notifikasi", "tap yes", "ketuk ya"]):
            digits = re.findall(r'\b\d{2}\b', text)
            if digits:
                return digits[0]
            return "Cek notifikasi di HP Anda (Ketuk Ya / Setuju)"
    except Exception:
        pass

    return None


def fetch_webshare_proxies(token: str) -> list[dict]:
    """Fetch active proxies list from Webshare API and test connectivity."""
    if not token or len(token.strip()) < 10:
        return []

    print("[*] Mengambil daftar proxy dari Webshare API...", flush=True)
    url = "https://proxy.webshare.io/api/v2/proxy/list/?mode=direct&page=1&page_size=100"
    req = urllib.request.Request(url, headers={"Authorization": f"Token {token.strip()}"})
    ctx = ssl.create_default_context()
    
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=12) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            results = data.get("results", [])
            raw_proxies = []
            for r in results:
                raw_proxies.append({
                    "server": f"http://{r.get('proxy_address')}:{r.get('port')}",
                    "username": r.get("username"),
                    "password": r.get("password"),
                    "ip": r.get("proxy_address"),
                    "country": r.get("country_code", "US")
                })
            
            print(f"[*] Memvalidasi {len(raw_proxies)} proxy Webshare ke CodeBuddy...", flush=True)
            healthy = []
            for p in raw_proxies:
                proxy_url = f"http://{p['username']}:{p['password']}@{p['ip']}:{p['server'].split(':')[-1]}"
                try:
                    handler = urllib.request.ProxyHandler({'http': proxy_url, 'https': proxy_url})
                    opener = urllib.request.build_opener(handler)
                    t0 = time.time()
                    res = opener.open('https://www.codebuddy.ai', timeout=4)
                    dt = time.time() - t0
                    print(f"  [+] Proxy OK: {p['ip']} ({p['country']}) - {dt:.2f}s", flush=True)
                    healthy.append(p)
                except Exception:
                    pass

            print(f"[V] {len(healthy)}/{len(raw_proxies)} proxy aktif dan siap dipakai.\n", flush=True)
            return healthy if healthy else raw_proxies
    except Exception as e:
        print(f"[!] Gagal mengambil proxy Webshare: {e}", flush=True)
        return []


def get_existing_connected_emails() -> set[str]:
    db_path = get_9router_db_path()
    if not db_path.exists():
        return set()
    try:
        conn = sqlite3.connect(str(db_path))
        cur = conn.cursor()
        cur.execute("SELECT data FROM providerConnections WHERE provider='codebuddy-intl'")
        rows = cur.fetchall()
        emails = set()
        for (d_str,) in rows:
            try:
                d = json.loads(d_str)
                em = d.get("email") or d.get("displayName")
                if em:
                    if "(" in em and ")" in em:
                        em = em.split("(")[-1].split(")")[0]
                    emails.add(em.strip().lower())
            except Exception:
                pass
        conn.close()
        return emails
    except Exception:
        return set()


def register_single_google(email: str, password: str, proxy_dict: dict | None, index: int = 1, total: int = 1) -> bool:
    from camoufox.sync_api import Camoufox

    print("\n" + "=" * 65, flush=True)
    proxy_info = f"Proxy: {proxy_dict.get('ip')} ({proxy_dict.get('country')})" if proxy_dict else "Direct (No Proxy)"
    print(f" [{index}/{total}] MEMPROSES: {email} | {proxy_info}", flush=True)
    print("=" * 65, flush=True)

    # 1. State Request
    print("[1] Meminta OAuth State CodeBuddy...", flush=True)
    try:
        state, auth_url = request_codebuddy_state()
    except Exception as e:
        print(f"[!] Error request state: {e}", flush=True)
        return False

    token_box = {}
    stop_event = threading.Event()

    def _poller():
        try:
            tok = poll_codebuddy_token(state, timeout=180, interval=2)
            if tok and tok.get("accessToken"):
                token_box["token"] = tok
                stop_event.set()
        except Exception:
            stop_event.set()

    poller_thread = threading.Thread(target=_poller, daemon=True)
    poller_thread.start()

    # 2. Camoufox Stealth Automation
    camoufox_args = {
        "headless": True,
        "locale": "en-US",
        "geoip": True
    }
    if proxy_dict:
        camoufox_args["proxy"] = {
            "server": proxy_dict["server"],
            "username": proxy_dict["username"],
            "password": proxy_dict["password"]
        }

    print(f"[2] Meluncurkan browser Camoufox ({proxy_info})...", flush=True)
    with Camoufox(**camoufox_args) as browser:
        page = browser.new_page(viewport={"width": 1280, "height": 800})
        page.set_default_timeout(45000)

        # 2a. Open Auth URL
        loaded = False
        for attempt in range(2):
            try:
                page.goto(auth_url, timeout=30000)
                loaded = True
                break
            except Exception as e:
                print(f"[*] Percobaan {attempt+1} goto gagal: {e}", flush=True)
                time.sleep(2)

        if not loaded:
            print("[!] Halaman login gagal dimuat lewat proxy ini.", flush=True)
            return False

        # 2b. Agreement Checkbox
        print("[3] Menyetujui kebijakan & memilih Google...", flush=True)
        frame = page.frame_locator('iframe').first
        try:
            frame.locator("#agree-policy-account").evaluate("""el => {
                el.checked = true;
                if (typeof handlePolicyChange === 'function') handlePolicyChange(el);
                el.dispatchEvent(new Event('change', { bubbles: true }));
                document.querySelectorAll('input[type=checkbox]').forEach(c => {
                    c.checked = true;
                    if (typeof handlePolicyChange === 'function') handlePolicyChange(c);
                    c.dispatchEvent(new Event('change', { bubbles: true }));
                });
            }""")
        except Exception as e:
            print(f"[*] Policy notice: {e}", flush=True)

        time.sleep(1)

        # 2c. Click Google
        try:
            frame.locator("#social-google").first.click(timeout=15000)
        except Exception as e:
            print(f"[!] Error klik tombol google: {e}", flush=True)

        time.sleep(4)

        # 2d. Fill Email
        print(f"[4] Mengisi email Google: {email}...", flush=True)
        try:
            email_input = page.locator("input[type='email'], input#identifierId, input[name='identifier']").first
            email_input.wait_for(state="visible", timeout=20000)
            email_input.fill(email)
            time.sleep(0.5)

            next_btn = page.locator("#identifierNext, button:has-text('Next'), button:has-text('Berikutnya')").first
            next_btn.click()
            time.sleep(3.5)
        except Exception as e:
            print(f"[!] Error input email: {e}", flush=True)

        # 2e. Fill Password
        print("[5] Mengisi password...", flush=True)
        time.sleep(1)
        pwd_input = page.locator("input[type='password'], input[name='Passwd'], input[name='password']").first
        if pwd_input.count() == 0 or not pwd_input.is_visible():
            pwd_opt = page.locator("div:has-text('Enter your password'), li:has-text('Enter your password'), div[data-challengeindex]").first
            if pwd_opt.count() > 0 and pwd_opt.is_visible():
                pwd_opt.click()
                time.sleep(2)
            else:
                try_btn = page.locator("button:has-text('Try another way'), a:has-text('Try another way')").first
                if try_btn.count() > 0 and try_btn.is_visible():
                    try_btn.click()
                    time.sleep(2)
                    pwd_opt2 = page.locator("div:has-text('Enter your password'), li:has-text('Enter your password')").first
                    if pwd_opt2.count() > 0:
                        pwd_opt2.click()
                        time.sleep(2)

        try:
            pwd_input = page.locator("input[type='password'], input[name='Passwd'], input[name='password']").first
            pwd_input.wait_for(state="visible", timeout=15000)
            pwd_input.fill(password)
            time.sleep(0.5)

            pwd_next = page.locator("#passwordNext, button:has-text('Next'), button:has-text('Berikutnya')").first
            pwd_next.click()
            time.sleep(4)
            print("[V] Password berhasil disubmit!", flush=True)
        except Exception as e:
            print(f"[!] Error input password: {e}", flush=True)

        # 2f. 2FA Detection & Code Display + Callback Handshake
        print("[6] Memantau verifikasi Google & handshake CodeBuddy...", flush=True)
        deadline = time.time() + 120
        last_printed_2fa = None

        while time.time() < deadline and not stop_event.is_set():
            curr_url = page.url

            # Detect Google 2FA Code
            code_2fa = extract_google_2fa_code(page)
            if code_2fa and code_2fa != last_printed_2fa:
                last_printed_2fa = code_2fa
                print("\n" + "*" * 65, flush=True)
                print(f" [🔔 VERIFIKASI HP / 2FA GOOGLE TERDETEKSI!]", flush=True)
                print(f" 👉 AKUN: {email}", flush=True)
                if code_2fa.isdigit():
                    print(f" 👉 KETUK ANGKA INI DI NOTIFIKASI HP ANDA:  >>> [ {code_2fa} ] <<<", flush=True)
                else:
                    print(f" 👉 INSTRUKSI HP:  {code_2fa}", flush=True)
                print(f" (Menunggu Anda mengonfirmasi di HP...)", flush=True)
                print("*" * 65 + "\n", flush=True)

            # Check if redirected back to CodeBuddy
            if "codebuddy" in curr_url and ("complete" in curr_url or "select" in curr_url or "login" in curr_url or "started" in curr_url):
                print(f"[V] Berhasil landing di CodeBuddy: {curr_url[:80]}...", flush=True)
                try:
                    select_btn = page.locator("button, a.btn, div[role='button']").first
                    if select_btn.count() > 0 and select_btn.is_visible():
                        select_btn.click(timeout=1500)
                except Exception:
                    pass

                print("[*] Menunggu CodeBuddy menyelesaikan handshake token...", flush=True)
                t_token_wait = time.time() + 25
                while time.time() < t_token_wait and not stop_event.is_set():
                    time.sleep(1)
                break

            # Consent button (Continue / Allow / Next)
            try:
                cont_btn = page.locator("button:has-text('Continue'), button:has-text('Lanjutkan'), button:has-text('Allow'), button:has-text('Izinkan')").first
                if cont_btn.count() > 0 and cont_btn.is_visible():
                    cont_btn.click(timeout=2000)
                    time.sleep(3)
            except Exception:
                pass

            time.sleep(1.5)

    # 3. Harvest Token
    poller_thread.join(timeout=15)

    if "token" in token_box:
        token_data = token_box["token"]
        access_token = token_data.get("accessToken", "")
        res_info = get_user_resource(access_token)
        u_data = res_info.get("data", {}) if isinstance(res_info, dict) else {}
        acc_email = u_data.get("email") or u_data.get("userEmail") or email

        ok = insert_codebuddy_to_9router(token_data, email_hint=acc_email)
        if ok:
            print(f" [V] SUKSES! Akun ({acc_email}) terdaftar & aktif di 9Router!\n", flush=True)
            return True
    else:
        print(f" [!] Gagal mendapatkan token untuk {email}\n", flush=True)
        return False


def main():
    print("=" * 65)
    print(" [*] 9ROUTER CODEBUDDY BULK GOOGLE AUTO-REGISTER")
    print("=" * 65)

    config = load_config()
    webshare_key = config.get("webshare_api_key", "")
    rotate_step = int(config.get("proxy_rotate_every_accounts", 3))
    delay_sec = int(config.get("delay_between_accounts_seconds", 3))

    proxies = fetch_webshare_proxies(webshare_key)
    if not proxies:
        print("[*] Mode: Direct / Tanpa Proxy (Bisa masukkan Webshare Key di config.json)")

    if not GOOGLE_TXT_PATH.exists():
        print(f"[!] File {GOOGLE_TXT_PATH} tidak ditemukan! Silakan buat file Google.txt berisi email|password.")
        return

    with open(GOOGLE_TXT_PATH, "r", encoding="utf-8", errors="ignore") as f:
        lines = [l.strip() for l in f if l.strip() and not l.strip().startswith("#") and "|" in l]

    if not lines:
        print("[!] Tidak ada akun di Google.txt. Silakan isi format: email@gmail.com|password")
        return

    existing = get_existing_connected_emails()
    print(f"[*] Total akun di Google.txt: {len(lines)} ({len(existing)} sudah terhubung di 9Router).")

    queue = []
    for line in lines:
        parts = line.split("|", 1)
        em = parts[0].strip()
        pw = parts[1].strip()

        if em.lower() in existing:
            continue
        queue.append((em, pw))

    print(f"[*] Total akun yang siap diproses: {len(queue)}\n")
    if not queue:
        print("[V] Semua akun di Google.txt sudah terhubung ke 9Router!")
        return

    success_count = 0
    fail_count = 0

    for idx, (em, pw) in enumerate(queue, 1):
        if proxies:
            proxy_idx = ((idx - 1) // rotate_step) % len(proxies)
            active_proxy = proxies[proxy_idx]
        else:
            active_proxy = None

        ok = register_single_google(em, pw, active_proxy, index=idx, total=len(queue))
        if ok:
            success_count += 1
        else:
            fail_count += 1
        time.sleep(delay_sec)

    print("\n" + "=" * 65)
    print(f" [★] PROSES SELESAI! Sukses: {success_count} | Gagal: {fail_count}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    main()
