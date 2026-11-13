# Временный диагностический скрипт: повторная отправка упавшего пакета (10 писем)
# для проверки — 502 воспроизводится стабильно или был разовым сбоем шлюза.
import json
import time
import notes_gemini as ng

data = json.load(open("dashboard_data.json", encoding="utf-8"))
fails = [e for e in data["emails"] if not e.get("kind")][:10]
print(f"Пакет из {len(fails)} писем (id: {[e['id'] for e in fails]})")

for attempt in (1, 2, 3):
    print(f"--- Попытка {attempt} ---")
    try:
        with open("promt-get-post.txt", encoding="utf-8") as f:
            tpl = f.read()
        payload = [{"id": e["id"], "sender": e["senderName"],
                    "email": e.get("senderEmail", ""),
                    "to": (e.get("to") or [])[:12], "cc": (e.get("cc") or [])[:12],
                    "subject": e["subject"], "body": e["body"][:400]} for e in fails]
        prompt = (tpl
                  .replace("[[TODAY]]", "2026-10-20")
                  .replace("[[E_TOP_LIST]]", "")
                  .replace("[[E_DIR_LIST]]", "")
                  .replace("[[MY_EMAIL]]", "test@svo.su")
                  .replace("[[E_TOP_EMAILS]]", "")
                  .replace("[[E_DIR_EMAILS]]", "")
                  .replace("[[EMAIL_COUNT]]", str(len(fails)))
                  .replace("[[EMAILS_PAYLOAD]]", json.dumps(payload, ensure_ascii=False, indent=2)))
        messages = [
            {"role": "system", "content": "Ты — корпоративный AI Triage ассистент. Отвечай только валидным JSON без markdown."},
            {"role": "user", "content": prompt},
        ]
        t0 = time.time()
        result, err = ng.ask_model(messages, temperature=0.1, timeout=120.0)
        dt = time.time() - t0
        if err:
            print(f"ошибка за {dt:.1f}s: {err}")
        else:
            items = result if isinstance(result, list) else (result.get("emails") if isinstance(result, dict) else None)
            n = len(items) if items is not None else -1
            print(f"успех за {dt:.1f}s, items={n}")
            break
    except Exception as e:
        print("EXC", type(e).__name__, e)
    time.sleep(3)
