"""
Диагностика извлечения ссылок на задачи 1С:ДО.

Отвечает на вопрос, почему у письма документооборота появляется (или нет)
кнопка «Перейти в ДО»: показывает, что видит каждый источник ссылок —
обычный текст письма, MIME-часть и DXL-выгрузка — и какой адрес в итоге
попадает в do_url.

Запуск (из корня проекта, рядом с notes_gemini.py):
    python scripts/probe_do_links.py                # свежие письма с упоминанием 1С:ДО
    python scripts/probe_do_links.py <UNID>         # конкретный документ

Скрипт только читает и печатает — ничего не меняет. Вывод можно копировать
целиком, если фикс не сработал.
"""

import os
import re
import sys
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import win32com.client  # noqa: E402

from notes_gemini import (  # noqa: E402
    get_mail_database,
    extract_do_url,
    mentions_do_task,
    _url_pairs_from_text,
    _urls_from_mime,
    _urls_from_dxl,
)

SEP = "=" * 80


def dump(doc, session, subject):
    """Печатает, что увидел каждый источник ссылок для одного документа."""
    print(SEP)
    print("Тема :", str(subject)[:100])
    body_text = ""
    if doc.HasItem("Body"):
        item = doc.GetFirstItem("Body")
        if item is not None and hasattr(item, "Text"):
            body_text = str(item.Text)
    print("ДО?  :", "да" if mentions_do_task(str(subject), body_text) else "нет")

    text_urls = [u for u, _ in _url_pairs_from_text(body_text)]
    mime_urls = [u for u, _ in _urls_from_mime(doc)]
    dxl_urls = [u for u, _ in _urls_from_dxl(doc, session)]
    print("URL в тексте :", text_urls or "— не найдено")
    print("URL из MIME  :", mime_urls or "— не найдено")
    print("URL из DXL   :", dxl_urls or "— не найдено")

    do = extract_do_url(doc, body_text, session)
    print("do_url итог  :", do or "— пусто (кнопки «Перейти в ДО» не будет)")


def main():
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize()
    db = get_mail_database(session)

    # По UNID — проверяем конкретное письмо.
    if len(sys.argv) > 1:
        unid = sys.argv[1].strip()
        doc = db.GetDocumentByUNID(unid)
        if doc is None:
            print(f"[!] Документ с UNID {unid} не найден.")
            return
        subject = doc.GetItemValue("Subject")[0] if doc.HasItem("Subject") else "(Без темы)"
        dump(doc, session, subject)
        return

    # Без аргумента — ищем свежие письма с упоминанием 1С:ДО.
    start = (datetime.now() - timedelta(days=3)).strftime("%d.%m.%Y")
    query = f'@IsAvailable(DeliveredDate) & DeliveredDate >= [{start}]'
    coll = db.Search(query, None, 0)
    print(f"Найдено писем за 3 дня: {coll.Count}")
    doc = coll.GetFirstDocument()
    shown = 0
    while doc is not None and shown < 10:
        try:
            subject = doc.GetItemValue("Subject")[0] if doc.HasItem("Subject") else ""
            body_text = ""
            if doc.HasItem("Body"):
                item = doc.GetFirstItem("Body")
                if item is not None and hasattr(item, "Text"):
                    body_text = str(item.Text)
            if mentions_do_task(subject, body_text):
                dump(doc, session, subject)
                shown += 1
        except Exception as e:
            print(f"[!] Ошибка чтения документа: {e}")
        try:
            doc = coll.GetNextDocument(doc)
        except Exception:
            break
    if not shown:
        print("Писем с упоминанием 1С:ДО за последние 3 дня не найдено.")


if __name__ == "__main__":
    main()
