# Временный диагностический скрипт: воспроизведение AI-пакета из 15 писем,
# которые остались неразобранными в последнем сборе (kind пустой).
import json
import notes_gemini as ng

data = json.load(open("dashboard_data.json", encoding="utf-8"))
fails = [e for e in data["emails"] if not e.get("kind")][:15]
print(f"Берём {len(fails)} неразобранных писем (id: {[e['id'] for e in fails]})")

before = [(e["id"], e.get("kind", ""), e.get("priority")) for e in fails]
ng.call_ai_triage(fails)

changed = sum(1 for e, b in zip(fails, before) if e.get("kind") != b[1])
print(f"Итог: разобрано {changed} из {len(fails)}")
for e in fails:
    print(f"  id={e['id']} kind={e.get('kind')!r} priority={e.get('priority')} "
          f"summary={str(e.get('summary'))[:60]!r}")
