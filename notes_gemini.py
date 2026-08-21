import os
import sys
import re
import json
import configparser
import urllib.request
import urllib.error
import ssl
from http.server import HTTPServer, SimpleHTTPRequestHandler
from datetime import datetime, timedelta
import win32com.client
import warnings
import threading

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

warnings.filterwarnings("ignore")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "set.ini")
PROMPT_FILE = os.path.join(BASE_DIR, "prompt_template.txt")
ANALYSIS_LOCK = threading.Lock()

def load_config():
    config = configparser.ConfigParser(interpolation=None)
    if not os.path.exists(CONFIG_FILE):
        print(f"[!] Файл конфигурации не найден: {CONFIG_FILE}")
        sys.exit(1)
    config.read(CONFIG_FILE, encoding="utf-8-sig")
    return config

def extract_email_from_text(text):
    """Извлекает email из строки (From) — ищет адрес в угловых скобках или просто email."""
    if not text:
        return None
    # Попытка 1: адрес в угловых скобках <...>
    m = re.search(r'<([^<>]+@[^<>]+)>', text)
    if m:
        return m.group(1).strip()
    # Попытка 2: обычный email в строке
    m = re.search(r'[\w\.\-\+]+@[\w\-]+(?:\.[\w\-]+)+', text)
    if m:
        return m.group(0).strip()
    return None


def extract_sender_email(doc):
    """Извлекает email отправителя из документа Notes."""
    # Поля, которые могут содержать интернет-адрес
    for field in ("INetFrom", "InternetFrom", "InternetAddr", "InternetAddress", "Originator"):
        if doc.HasItem(field):
            try:
                val = doc.GetItemValue(field)[0]
                if val:
                    email = extract_email_from_text(str(val))
                    if email:
                        return email
            except Exception:
                pass
    # Последняя попытка: поле From
    if doc.HasItem("From"):
        try:
            val = doc.GetItemValue("From")[0]
            email = extract_email_from_text(str(val))
            if email:
                return email
        except Exception:
            pass
    return None


def build_name_email_mapping(emails_data):
    """Строит два словаря: полное имя -> email и фамилия -> email."""
    mapping_full = {}
    mapping_last = {}
    for em in emails_data:
        name = (em.get("senderName") or "").strip()
        email = (em.get("senderEmail") or "").strip().lower()
        if not name or not email:
            continue
        name_l = name.lower()
        if name_l not in mapping_full:
            mapping_full[name_l] = email
        parts = name.split()
        if parts:
            last_name = parts[0].lower()
            if last_name not in mapping_last:
                mapping_last[last_name] = email
    return mapping_full, mapping_last


def find_email_for_person(full_name, mapping_full, mapping_last):
    """Ищет email для ФИО по точному совпадению, фамилии или вхождению."""
    name_l = full_name.strip().lower()
    if not name_l:
        return None
    if name_l in mapping_full:
        return mapping_full[name_l]
    parts = name_l.split()
    if parts and parts[0] in mapping_last:
        return mapping_last[parts[0]]
    # Вхождение: если фамилия встречается в каком-то имени
    for key, email in mapping_full.items():
        if parts and parts[0] and parts[0] in key:
            return email
        if name_l in key or key in name_l:
            return email
    return None


def resolve_names_to_emails(config, emails_data, persons_key="e_top", emails_key="e_top_emails"):
    """Сопоставляет ФИО из конфига с почтами из писем.
       Обновляет emails_key в конфиге (список уже найденных почт).
       Возвращает (новый_список_почт, список_ненайденных_ФИО)."""
    persons = [x.strip() for x in config.get("TAGS", persons_key, fallback="").split(",") if x.strip()]
    emails_cfg = [x.strip().lower() for x in config.get("TAGS", emails_key, fallback="").split(",") if x.strip()]
    mapping_full, mapping_last = build_name_email_mapping(emails_data)

    new_emails = list(emails_cfg)
    unresolved = []

    for person in persons:
        email = find_email_for_person(person, mapping_full, mapping_last)
        if email:
            if email not in new_emails:
                new_emails.append(email)
        else:
            # Проверяем: а вдруг для этого ФИО почта уже была найдена ранее?
            already_have = any(person.lower() in e for e in emails_cfg)
            if not already_have:
                unresolved.append(person)

    return ",".join(new_emails), unresolved


def load_prompt_template():
    if os.path.exists(PROMPT_FILE):
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return "[[EMAILS_PAYLOAD]]"

def extract_statuses(doc, session):
    is_unread = False
    try:
        # Проверяем статус прочтения текущим пользователем
        read_flag = False
        try:
            read_flag = doc.GetRead(session.CommonUserName)
        except Exception:
            try:
                read_flag = doc.GetRead()
            except Exception:
                read_flag = doc.GetRead(session.UserName)
        is_unread = not read_flag
    except Exception:
        is_unread = False

    is_replied = False
    try:
        if doc.HasItem("$Replied") or doc.HasItem("_Replied"):
            is_replied = True
        elif doc.HasItem("$ActionFlags"):
            flags = doc.GetItemValue("$ActionFlags")[0]
            if isinstance(flags, (int, float)) and (int(flags) & 1024):
                is_replied = True
    except Exception:
        pass

    return is_unread, is_replied

def get_mail_database(session):
    mail_server = session.GetEnvironmentString("MailServer", True)
    mail_file = session.GetEnvironmentString("MailFile", True)
    
    if mail_file:
        db = session.GetDatabase(mail_server, mail_file)
        if db and db.IsOpen:
            return db
        elif db:
            db.Open()
            return db

    try:
        db_dir = session.GetDbDirectory(mail_server)
        db = db_dir.OpenMailDatabase()
        if db:
            return db
    except Exception:
        pass

    raise RuntimeError("Не удалось открыть почтовую базу Lotus Notes.")

def create_notes_reply(unid, reply_text):
    """Создаёт черновик ответного письма в Lotus Notes.
       Возвращает (ссылка_на_черновик, ошибка)."""
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize()
        db = get_mail_database(session)
        original = db.GetDocumentByUNID(unid)
        if not original:
            return None, "Оригинальное письмо не найдено по UNID"

        memo = db.CreateDocument()
        memo.ReplaceItemValue("Form", "Memo")

        # Получатель — отправитель исходного письма
        try:
            recipients = original.GetItemValue("From")[0]
            memo.ReplaceItemValue("SendTo", recipients)
        except Exception:
            pass

        # Тема
        try:
            subject = original.GetItemValue("Subject")[0]
            memo.ReplaceItemValue("Subject", "RE: " + str(subject))
        except Exception:
            memo.ReplaceItemValue("Subject", "RE: Письмо")

        # Тело: рекомендованный AI-ответ
        body = memo.CreateRichTextItem("Body")
        body.AppendText(reply_text)

        # Сохраняем черновик (без отправки)
        memo.Save(True, True)

        new_unid = str(memo.UniversalID).strip()
        replica_id = str(db.ReplicaID).replace(":", "").strip()
        link = f"notes:///{replica_id}/0/{new_unid}?OpenDocument&Edit"
        return link, None
    except Exception as e:
        return None, str(e)


def fetch_notes_emails(days=2, max_emails=40, max_chars=800):
    print(f"\n[*] Чтение почты Notes за последние {days} дн...")
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize()
    except Exception as e:
        print(f"[!] Сбой COM: {e}")
        return []

    try:
        db = get_mail_database(session)
    except Exception as db_err:
        print(f"[!] Ошибка БД: {db_err}")
        return []

    replica_id = str(db.ReplicaID).replace(":", "").strip()
    start_date = (datetime.now() - timedelta(days=days)).strftime("%d.%m.%Y")
    search_query = f'@IsAvailable(DeliveredDate) & DeliveredDate >= [{start_date}]'
    
    collection = db.Search(search_query, None, 0)
    count = collection.Count
    print(f"[+] Найдено писем в базе: {count}")
    if count == 0:
        return []

    emails = []
    doc = collection.GetFirstDocument()
    idx = 1

    while doc is not None:
        try:
            subject = doc.GetItemValue("Subject")[0] if doc.HasItem("Subject") else "(Без темы)"
            sender = doc.GetItemValue("From")[0] if doc.HasItem("From") else "(Неизвестный)"
            sender_email = extract_sender_email(doc)
            delivered = doc.GetItemValue("DeliveredDate")[0] if doc.HasItem("DeliveredDate") else doc.Created
            is_unread, is_replied = extract_statuses(doc, session)
            unid = str(doc.UniversalID).strip()

            body_text = ""
            if doc.HasItem("Body"):
                body_item = doc.GetFirstItem("Body")
                if body_item is not None and hasattr(body_item, "Text"):
                    body_text = body_item.Text.strip()

            clean_body = re.sub(r'\s+', ' ', body_text)[:max_chars]
            date_display = str(delivered)[:16]
            
            # Чистая ссылка протокола Lotus Notes
            lotus_link = f"notes:///{replica_id}/0/{unid}"

            unread_tag = "[НОВОЕ]" if is_unread else "[ПРОЧИТАНО]"
            print(f"  #{idx:02d} {unread_tag} {str(sender)[:22]} | {str(subject)[:35]}")

            emails.append({
                "id": str(idx),
                "unid": unid,
                "lotus_url": lotus_link,
                "date": date_display,
                "senderName": str(sender).replace("CN=", "").split("/")[0],
                "senderEmail": sender_email or "",
                "subject": str(subject),
                "isRead": not is_unread,
                "isReplied": is_replied,
                "needsReply": False,
                "priority": 3,
                "body": clean_body,
                "threadMessages": [{"text": clean_body}]
            })
            idx += 1
            if len(emails) >= max_emails:
                break
        except Exception as doc_err:
            print(f"[!] Пропуск: {doc_err}")

        doc = collection.GetNextDocument(doc)

    return emails

def set_doc_read_status(unid, read_state=True):
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize()
        db = get_mail_database(session)
        doc = db.GetDocumentByUNID(unid)
        if doc:
            if read_state:
                try: doc.MarkRead(session.CommonUserName)
                except: doc.MarkRead()
            else:
                try: doc.MarkUnread(session.CommonUserName)
                except: doc.MarkUnread()
            return True
    except Exception as e:
        print(f"[!] Ошибка смены статуса UNID {unid}: {e}")
    return False

def call_ai_triage(emails):
    cfg = load_config()
    e_top = cfg.get("TAGS", "e_top", fallback="")
    e_dir = cfg.get("TAGS", "e_dir", fallback="")
    api_key = cfg.get("DEEPSEEK", "api_key", fallback="").strip()
    base_url = cfg.get("DEEPSEEK", "base_url", fallback="https://api.deepseek.com").strip().rstrip("/")
    model = cfg.get("DEEPSEEK", "model", fallback="deepseek-chat").strip()
    batch_size = cfg.getint("NOTES", "batch_size", fallback=15)
    raw_template = load_prompt_template()

    print(f"[*] Отправка {len(emails)} писем в DeepSeek AI пакетами по {batch_size}...")

    # Разбиваем письма на пакеты
    batches = [emails[i:i + batch_size] for i in range(0, len(emails), batch_size)]
    total_batches = len(batches)

    for batch_idx, batch in enumerate(batches, 1):
        print(f"[*] Обработка пакета {batch_idx}/{total_batches} ({len(batch)} писем)...")

        prompt_payload = [{
            "id": e["id"],
            "sender": e["senderName"],
            "email": e.get("senderEmail", ""),
            "subject": e["subject"],
            "body": e["body"][:400]
        } for e in batch]

        prompt = (raw_template
            .replace("[[E_TOP_LIST]]", str(e_top))
            .replace("[[E_DIR_LIST]]", str(e_dir))
            .replace("[[EMAIL_COUNT]]", str(len(batch)))
            .replace("[[EMAILS_PAYLOAD]]", json.dumps(prompt_payload, ensure_ascii=False, indent=2))
        )

        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": "Ты — корпоративный AI Triage ассистент. Отвечай только валидным JSON без markdown."},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.1,
            "stream": False
        }

        try:
            req_data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
            req = urllib.request.Request(
                f"{base_url}/chat/completions",
                data=req_data,
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                method="POST"
                        )
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

            try:
                with urllib.request.urlopen(req, context=ctx, timeout=30.0) as resp:
                    res_json = json.loads(resp.read().decode('utf-8'))
            except urllib.error.URLError as url_err:
                print(f"[!] Сетевая ошибка при обращении к {base_url}: {url_err}")
                raise
            except Exception as http_err:
                print(f"[!] HTTP ошибка: {http_err}")
                raise

            raw_text = res_json["choices"][0]["message"]["content"].strip()

            # Убираем markdown код-блоки если они есть
            if raw_text.startswith("```"):
                raw_text = re.sub(r'^```(?:json)?\s*', '', raw_text)
                raw_text = re.sub(r'\s*```$', '', raw_text)

            match = re.search(r'(\{[\s\S]*\}|\[[\s\S]*\])', raw_text)
            if match:
                raw_text = match.group(1)
            
            try:
                                parsed = json.loads(raw_text)
            except json.JSONDecodeError as json_err:
                print(f"[!] Ошибка парсинга JSON из ответа AI: {json_err}")
                print(f"[!] Полученный текст: {raw_text[:500]}...")
                continue

            items = parsed if isinstance(parsed, list) else parsed.get("emails", [])

            ai_map = {str(it.get("id")): it for it in items}
            for e in batch:
                info = ai_map.get(str(e["id"]), {})
                if info:
                    # Приоритет: поддерживаем priority (P1/P2/P3), criticality (число),
                    # строки типа "High priority", "срочно", "высокий" и т.д.
                    raw_priority = info.get("priority", info.get("criticality", e["priority"]))
                    if isinstance(raw_priority, str):
                        try:
                            # Попытка извлечь число из "P1", "1", "High priority"
                            match_num = re.search(r'(\d)', raw_priority)
                            if match_num:
                                e["priority"] = int(match_num.group(1))
                            elif "высок" in raw_priority.lower() or "срочн" in raw_priority.lower():
                                e["priority"] = 1
                            elif "средн" in raw_priority.lower():
                                e["priority"] = 2
                            else:
                                e["priority"] = 3
                        except Exception:
                            pass
                    else:
                        try:
                            e["priority"] = int(raw_priority)
                        except Exception:
                            pass

                    # Ограничиваем приоритет диапазоном 1-3
                    if e["priority"] < 1: e["priority"] = 1
                    if e["priority"] > 3: e["priority"] = 3

                                        # needsReply: из булева поля или из action_type
                    raw_need = info.get("needsReply", info.get("action_type", ""))
                    if isinstance(raw_need, bool):
                        e["needsReply"] = raw_need
                    elif isinstance(raw_need, str):
                        e["needsReply"] = raw_need.lower() in ("нужен ответ", "yes", "true", "да", "требует ответа")

                    # Summary
                    if "summary" in info:
                        e["threadMessages"] = [{"text": info["summary"]}]

                    # Рекомендованный ответ AI
                    if info.get("suggested_reply"):
                        e["suggestedReply"] = info["suggested_reply"]
                    elif info.get("suggestedReply"):
                        e["suggestedReply"] = info["suggestedReply"]

                    # Отвечено по action_type
                    if info.get("action_type", "") == "Отвечено":
                        e["isReplied"] = True
            print(f"[✓] Пакет {batch_idx}/{total_batches} обработан успешно!")
        except Exception as e:
            print(f"[!] AI Triage для пакета {batch_idx}/{total_batches} пропущен из-за ошибки: {e}")

    print(f"[✓] AI анализ всех {len(emails)} писем завершен!")
    return emails

class NotesWebHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        if self.path == "/api/settings":
            cfg = load_config()
            data = {
                "e_top": cfg.get("TAGS", "e_top", fallback=""),
                "e_dir": cfg.get("TAGS", "e_dir", fallback=""),
                "e_top_emails": cfg.get("TAGS", "e_top_emails", fallback=""),
                "e_dir_emails": cfg.get("TAGS", "e_dir_emails", fallback="")
            }
            self.send_json(200, data)
        else:
            super().do_GET()

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length).decode('utf-8')
        req_data = json.loads(body) if body else {}

        if self.path == "/api/analyze":
            if not ANALYSIS_LOCK.acquire(blocking=False):
                self.send_json(429, {"error": "Анализ уже выполняется..."})
                return

            try:
                cfg = load_config()
                days = req_data.get("days", cfg.getint("NOTES", "default_days", fallback=2))
                max_emails = cfg.getint("NOTES", "max_emails", fallback=40)
                max_chars = cfg.getint("NOTES", "max_body_chars", fallback=800)

                emails = fetch_notes_emails(days=days, max_emails=max_emails, max_chars=max_chars)
                if emails:
                    # Сопоставляем ФИО из настроек с почтами из прочитанных писем
                    try:
                        e_top_emails, unresolved_top = resolve_names_to_emails(cfg, emails, "e_top", "e_top_emails")
                        e_dir_emails, unresolved_dir = resolve_names_to_emails(cfg, emails, "e_dir", "e_dir_emails")
                        cfg.set("TAGS", "e_top_emails", e_top_emails)
                        cfg.set("TAGS", "e_dir_emails", e_dir_emails)
                        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                            cfg.write(f)
                        if unresolved_top:
                            print(f"[i] Не найдены почты для E_TOP (повторим позже): {', '.join(unresolved_top)}")
                        if unresolved_dir:
                            print(f"[i] Не найдены почты для E_DIR (повторим позже): {', '.join(unresolved_dir)}")
                    except Exception as resolve_err:
                        print(f"[!] Ошибка сопоставления ФИО->почты: {resolve_err}")

                    emails = call_ai_triage(emails)

                with open(os.path.join(BASE_DIR, "dashboard_data.json"), "w", encoding="utf-8") as f:
                    json.dump({"emails": emails}, f, ensure_ascii=False, indent=2)

                self.send_json(200, {"emails": emails})
            except Exception as e:
                print(f"[!] Ошибка: {e}")
                self.send_json(500, {"error": str(e)})
            finally:
                ANALYSIS_LOCK.release()

        elif self.path == "/api/mark-read":
            try:
                unid = req_data.get("unid", "")
                is_read = req_data.get("is_read", True)
                success = set_doc_read_status(unid, is_read)
                self.send_json(200, {"status": "ok" if success else "failed"})
            except Exception as e:
                self.send_json(500, {"error": str(e)})

        elif self.path == "/api/save-settings":
            try:
                cfg = load_config()
                new_e_top = req_data.get("e_top", "")
                new_e_dir = req_data.get("e_dir", "")
                cfg.set("TAGS", "e_top", new_e_top)
                cfg.set("TAGS", "e_dir", new_e_dir)

                # Пересчитываем email-списки по данным из кэша
                cached_emails = []
                try:
                    with open(os.path.join(BASE_DIR, "dashboard_data.json"), "r", encoding="utf-8") as f:
                        cache = json.load(f)
                    cached_emails = cache.get("emails", [])
                except Exception:
                    pass

                if cached_emails:
                    e_top_emails, _ = resolve_names_to_emails(cfg, cached_emails, "e_top", "e_top_emails")
                    e_dir_emails, _ = resolve_names_to_emails(cfg, cached_emails, "e_dir", "e_dir_emails")
                    cfg.set("TAGS", "e_top_emails", e_top_emails)
                    cfg.set("TAGS", "e_dir_emails", e_dir_emails)

                with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                    cfg.write(f)

                self.send_json(200, {
                    "status": "ok",
                    "e_top_emails": cfg.get("TAGS", "e_top_emails", fallback=""),
                    "e_dir_emails": cfg.get("TAGS", "e_dir_emails", fallback="")
                })
            except Exception as e:
                self.send_json(500, {"error": str(e)})

        elif self.path == "/api/reply":
            try:
                unid = req_data.get("unid", "")
                text = req_data.get("text", "")
                if not unid:
                    self.send_json(400, {"error": "unid обязателен"})
                    return
                link, err = create_notes_reply(unid, text)
                if link:
                    self.send_json(200, {"status": "ok", "lotus_url": link})
                else:
                    self.send_json(500, {"error": err or "Ошибка создания ответа в Lotus Notes"})
            except Exception as e:
                self.send_json(500, {"error": str(e)})

    def send_json(self, status, payload):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(payload, ensure_ascii=False).encode('utf-8'))

if __name__ == "__main__":
    PORT = 8765
    server = HTTPServer(("127.0.0.1", PORT), NotesWebHandler)
    print(f"[*] Smart Mail Dashboard запущен: http://127.0.0.1:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()
