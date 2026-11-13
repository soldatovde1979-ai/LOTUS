# Временный диагностический скрипт: один реальный AI-пакет (2 письма) через call_ai_triage
import json
import notes_gemini as ng

probe = [
    {
        "id": "9991",
        "senderName": "Anna I Svalova",
        "senderEmail": "ai.svalova@svo.su",
        "to": ["CN=Dmitriy E Soldatov/O=SVO@SVO"],
        "cc": [],
        "subject": "Accepted: Планерка 1С",
        "body": "Планерка 1С: принята приглашение, время 15:00.",
        "priority": 3,
    },
    {
        "id": "9992",
        "senderName": "Ivan I Ivanov",
        "senderEmail": "ii.ivanov@svo.su",
        "to": ["CN=Dmitriy E Soldatov/O=SVO@SVO"],
        "cc": [],
        "subject": "Вопрос по отчету",
        "body": "Добрый день, подскажи когда будет готов отчет по продажам? Нужен до пятницы.",
        "priority": 3,
    },
]

ng.call_ai_triage(probe)
for e in probe:
    print("---", e["id"])
    print("kind:", e.get("kind"), "| priority:", e.get("priority"),
          "| ask:", e.get("ask"), "| summary:", str(e.get("summary"))[:80])
