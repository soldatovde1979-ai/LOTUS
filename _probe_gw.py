# Временный диагностический скрипт: проверка gw-proxy 8787 (/v1/chat/completions)
import json
import urllib.request
import urllib.error

req = urllib.request.Request(
    "http://127.0.0.1:8787/v1/chat/completions",
    data=json.dumps({
        "model": "pro",
        "messages": [{"role": "user", "content": "reply with the single word: ok"}],
        "temperature": 0,
    }).encode(),
    headers={"Content-Type": "application/json", "Authorization": "Bearer local"},
)
try:
    r = urllib.request.urlopen(req, timeout=90)
    d = json.loads(r.read())
    print("HTTP", r.status)
    print("content:", str(d["choices"][0]["message"]["content"])[:300])
    print("model:", d.get("model"))
except urllib.error.HTTPError as e:
    print("HTTPError", e.code)
    print(e.read()[:800].decode(errors="replace"))
except Exception as e:
    print("ERR", type(e).__name__, e)
