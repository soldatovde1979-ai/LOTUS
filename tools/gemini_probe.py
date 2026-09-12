"""
Диагностика доступа к Gemini API (generativelanguage.googleapis.com).

Печатает, на какие IP резолвится домен (системный DNS), и делает тестовый
запрос generateContent. Ключ берётся из первого аргумента командной строки
или переменной окружения GEMINI_API_KEY.

Запуск:
    python tools\\gemini_probe.py                      # резолв + тест без ключа
    python tools\\gemini_probe.py <API_KEY>            # тест с ключом
    set GEMINI_API_KEY=... && python tools\\gemini_probe.py

Только чтение и печать, ничего не меняет. Зависимости: stdlib.
"""

import os
import sys
import json
import ssl
import socket
import configparser
import urllib.request
import urllib.error
from urllib.parse import quote

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

HOST = "generativelanguage.googleapis.com"
URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"


def main():
    print("=== DNS ===")
    try:
        infos = socket.getaddrinfo(HOST, 443)
        ips = sorted(set(i[4][0] for i in infos))
        print(f"{HOST} -> {ips}")
    except OSError as e:
        print(f"resolv FAIL: {e}")

    api_key = sys.argv[1].strip() if len(sys.argv) > 1 else os.environ.get("GEMINI_API_KEY", "")
    key_source = "argv" if len(sys.argv) > 1 else ("env" if api_key else "")
    if not api_key:
        try:
            ini = configparser.ConfigParser(interpolation=None)
            ini.read(os.path.join(BASE_DIR, "set.ini"), encoding="utf-8-sig")
            api_key = ini.get("GOOGLE", "api_key", fallback="").strip()
            key_source = "set.ini [GOOGLE]" if api_key else ""
        except Exception:
            pass
    print("=== API ===")
    print(f"url: {URL}")
    print(f"key: {'set' if api_key else 'NOT SET'}" + (f" ({key_source})" if key_source else ""))

    payload = json.dumps({"contents": [{"parts": [{"text": "ping"}]}]}).encode("utf-8")
    req = urllib.request.Request(
        URL + ("?key=" + quote(api_key, safe="") if api_key else ""),
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try:
        with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
            print("HTTP:", resp.status)
            print(resp.read().decode("utf-8")[:500])
    except urllib.error.HTTPError as e:
        print("HTTP:", e.code)
        print(e.read().decode("utf-8")[:500])
    except urllib.error.URLError as e:
        print("URLError:", e)


if __name__ == "__main__":
    main()
