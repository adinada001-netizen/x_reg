"""Manual Interactive Browser Login for CodeBuddy -> 9Router."""

import json
import os
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path
from codebuddy_auto import (
    request_codebuddy_state,
    poll_codebuddy_token,
    insert_codebuddy_to_9router,
    get_user_resource
)

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

def main():
    print("=" * 65)
    print(" [*] CODEBUDDY MANUAL / INTERACTIVE LOGIN KE 9ROUTER")
    print("=" * 65)

    print("\n[1] Meminta OAuth State resmi CodeBuddy...")
    try:
        state, auth_url = request_codebuddy_state()
    except Exception as e:
        print(f"[!] Gagal meminta state: {e}")
        return

    print(f"[V] State ID : {state}")
    print(f"[V] Auth URL : {auth_url}")

    print("\n" + "-" * 65)
    print(" [👉] PETUNJUK:")
    print(" 1. Halaman login CodeBuddy akan terbuka otomatis di browser Anda.")
    print(" 2. Centang persetujuan kebijakan.")
    print(" 3. Login menggunakan Google, X (Twitter), GitHub, atau Apple.")
    print(" 4. Script ini otomatis menangkap token dan menghubungkannya ke 9Router!")
    print("-" * 65 + "\n")

    try:
        webbrowser.open(auth_url)
    except Exception:
        pass

    print("[*] Menunggu konfirmasi login dari CodeBuddy (timeout 10 menit)...")
    token_data = poll_codebuddy_token(state, timeout=600, interval=2)

    if not token_data or not token_data.get("accessToken"):
        print("[!] Timeout atau otorisasi belum selesai. Silakan coba lagi.")
        return

    print("[V] Token resmi CodeBuddy berhasil didapatkan!")
    access_token = token_data.get("accessToken", "")
    res_info = get_user_resource(access_token)
    u_data = res_info.get("data", {}) if isinstance(res_info, dict) else {}
    email = u_data.get("email") or u_data.get("userEmail") or "CodeBuddy Account"

    ok = insert_codebuddy_to_9router(token_data, email_hint=email)
    if ok:
        print("\n" + "=" * 65)
        print(f" [V] SUKSES! Akun ({email}) berhasil terhubung aktif ke 9Router!")
        print("=" * 65 + "\n")

if __name__ == "__main__":
    main()
