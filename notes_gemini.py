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
DATA_FILE = os.path.join(BASE_DIR, "dashboard_data.json")
STATE_FILE = os.path.join(BASE_DIR, "work_state.json")
ANALYSIS_LOCK = threading.Lock()

# Как именно удалось определить прочитанность в последнем сборе (уходит в UI)
READ_DETECTION = {"method": "none", "reliable": False, "detail": "", "candidates": {}}


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
    m = re.search(r'<([^<>]+@[^<>]+)>', text)
    if m:
        return m.group(1).strip()
    m = re.search(r'[\w\.\-\+]+@[\w\-]+(?:\.[\w\-]+)+', text)
    if m:
        return m.group(0).strip()
    return None


def extract_sender_email(doc):
    """Извлекает email отправителя из документа Notes."""
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
    """Ищет email для ФИО. Сравнение идёт по транслитерированной фамилии,
    поэтому «Столяров Егор Валерьевич» находит «Egor V Stolyarov»."""
    name_l = str(full_name or "").strip().lower()
    if not name_l:
        return None
    if name_l in mapping_full:
        return mapping_full[name_l]
    for key, email in mapping_full.items():
        if names_match(full_name, key):
            return email
    return None


def resolve_names_to_emails(config, emails_data, persons_key="e_top", emails_key="e_top_emails"):
    """Сопоставляет ФИО из конфига с почтами из писем."""
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
            already_have = any(person.lower() in e for e in emails_cfg)
            if not already_have:
                unresolved.append(person)

    return ",".join(new_emails), unresolved


def normalize_email_list(raw):
    """Приводит введённый руками список адресов к единому виду без дублей."""
    seen, out = set(), []
    for part in re.split(r'[,;\s]+', str(raw or "")):
        part = part.strip().lower()
        if part and part not in seen:
            seen.add(part)
            out.append(part)
    return ",".join(out)


def load_prompt_template():
    if os.path.exists(PROMPT_FILE):
        with open(PROMPT_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return "[[EMAILS_PAYLOAD]]"


# ============================================================================
#  ОПРЕДЕЛЕНИЕ ПРОЧИТАННОСТИ
#
#  Что было сломано:
#   1) db.CreateNoteCollection(True) — параметр selectAllFlag означает
#      "включить ВСЕ типы заметок", а не "только непрочитанные". Коллекция
#      возвращала каждый документ базы, и множество unread_unids содержало
#      вообще все письма => все письма считались новыми.
#   2) session.Evaluate("@Unread", doc) — такой @-функции в Notes нет,
#      Evaluate падал, и код уходил на сломанный путь (1).
#
#  Через COM надёжно работает NotesDocument.GetRead() (Notes 8.5+).
#  Всё остальное ниже — резерв.
# ============================================================================

def _call_get_read(doc, username):
    """Вызывает doc.GetRead(), перебирая допустимые сигнатуры COM.
       Возвращает True/False или бросает исключение."""
    try:
        return bool(doc.GetRead())
    except Exception:
        pass
    if username:
        return bool(doc.GetRead(username))
    raise RuntimeError("GetRead недоступен")


def _session_username(session):
    for prop in ("EffectiveUserName", "CommonUserName", "UserName"):
        try:
            val = getattr(session, prop, None)
            if val:
                return str(val)
        except Exception:
            pass
    return ""


def probe_get_read(db, session):
    """Проверяет, работает ли doc.GetRead() в этой сборке Notes.
       Возвращает (работает: bool, username: str)."""
    username = _session_username(session)

    sample = None
    for view_name in ("($Inbox)", "($All)", "Inbox"):
        try:
            view = db.GetView(view_name)
            if view is not None:
                sample = view.GetFirstDocument()
                if sample is not None:
                    break
        except Exception:
            pass
    if sample is None:
        try:
            sample = db.AllDocuments.GetFirstDocument()
        except Exception:
            pass
    if sample is None:
        print("[!] Не нашёл ни одного документа для проверки GetRead()")
        return False, username

    try:
        _call_get_read(sample, username)
        print(f"[+] Метки прочтения: doc.GetRead() (пользователь: {username or 'текущий ID'})")
        return True, username
    except Exception as e:
        print(f"[!] doc.GetRead() недоступен: {e}")
        return False, username


def _unid_of(entry_or_doc):
    try:
        doc = getattr(entry_or_doc, "Document", None) or entry_or_doc
        return str(doc.UniversalID).strip()
    except Exception:
        return None


def unread_via_get_all_unread_entries(view):
    """NotesView.GetAllUnreadEntries() — есть с Notes 8, в COM бывает не открыт."""
    col = view.GetAllUnreadEntries()
    if col is None:
        return None
    out = set()
    entry = col.GetFirstEntry()
    while entry is not None:
        u = _unid_of(entry)
        if u:
            out.add(u)
        entry = col.GetNextEntry(entry)
    return out


def unread_via_nav_all_unread(view):
    """NotesView.CreateViewNavFromAllUnread() — быстрее, но в ряде сборок сломан."""
    nav = view.CreateViewNavFromAllUnread()
    if nav is None:
        return None
    out = set()
    entry = nav.GetFirst()
    guard = 0
    while entry is not None and guard < 5000:
        u = _unid_of(entry)
        if u:
            out.add(u)
        entry = nav.GetNext(entry)
        guard += 1
    return out


def unread_via_entry_scan(view, limit=3000):
    """Полный обход вида с опросом ViewEntry.IsUnread."""
    view.AutoUpdate = False
    nav = view.CreateViewNav()
    entry = nav.GetFirst()
    out, scanned, answered = set(), 0, 0
    while entry is not None and scanned < limit:
        try:
            if entry.IsDocument:
                scanned += 1
                verdict = None
                try:
                    verdict = not bool(entry.GetRead())
                    answered += 1
                except Exception:
                    try:
                        verdict = bool(entry.IsUnread)
                        answered += 1
                    except Exception:
                        verdict = None
                if verdict:
                    u = _unid_of(entry)
                    if u:
                        out.add(u)
        except Exception:
            pass
        entry = nav.GetNext(entry)
    if not scanned or not answered:
        # вид обошли, но на вопрос «прочитано?» никто не ответил — данных нет
        return None
    return out


def fetch_all_unread(db):
    """NotesDatabase.GetAllUnreadDocuments() — в COM официально не поддержан."""
    nc = db.GetAllUnreadDocuments()
    if nc is None:
        return None
    out = set()
    nid = nc.GetFirstNoteId()
    while nid:
        try:
            doc = db.GetDocumentByID(nid)
            if doc and doc.IsValid:
                out.add(str(doc.UniversalID).strip())
        except Exception:
            pass
        nid = nc.GetNextNoteId(nid)
    return out


def gather_unread_sources(db, tag=""):
    """Опрашивает ВСЕ доступные способы получить непрочитанные.
    Возвращает {название источника: множество UNID}."""
    found = {}
    suffix = f" [{tag}]" if tag else ""

    try:
        s = fetch_all_unread(db)
        if s is not None:
            found["GetAllUnreadDocuments" + suffix] = s
    except Exception as e:
        print(f"[i] GetAllUnreadDocuments{suffix}: недоступен ({e})")

    for vname in ("($Inbox)", "($All)"):
        try:
            view = db.GetView(vname)
        except Exception:
            view = None
        if view is None:
            continue
        for label, fn in (("GetAllUnreadEntries", unread_via_get_all_unread_entries),
                          ("CreateViewNavFromAllUnread", unread_via_nav_all_unread),
                          ("обход IsUnread", unread_via_entry_scan)):
            name = f"{label} {vname}{suffix}"
            try:
                s = fn(view)
                if s is not None:
                    found[name] = s
            except Exception as e:
                print(f"[i] {name}: недоступен ({e})")
    return found


def open_local_replica(session, db_server):
    """Метки прочтения живут в той реплике, где человек читает почту.
    Если работа идёт в локальной реплике, на сервере всё будет «прочитано»."""
    try:
        mail_file = session.GetEnvironmentString("MailFile", True)
        if not mail_file:
            return None
        local = session.GetDatabase("", mail_file)
        if local is None:
            return None
        if not local.IsOpen:
            local.Open()
        if not local.IsOpen:
            return None
        if not db_server:          # уже открывали локальную — второй раз не нужно
            return None
        return local
    except Exception as e:
        print(f"[i] Локальная реплика недоступна: {e}")
        return None


DEFAULT_SUBJECT_BLOCK = "принято, отклонено"
VALID_STATES = ("new", "wip", "done")
STATE_LIMIT = 3000


def ensure_section(cfg, name):
    if not cfg.has_section(name):
        cfg.add_section(name)


def window_start(days, skip_weekends=True):
    """Начало окна выборки. Суббота и воскресенье не расходуют «дни»,
    поэтому в понедельник «за 2 дня» достаёт письма с четверга."""
    d = datetime.now()
    remaining = max(1, int(days))
    while remaining > 0:
        d -= timedelta(days=1)
        if skip_weekends and d.weekday() >= 5:
            continue
        remaining -= 1
    return d


def parse_subject_filters(raw):
    return [p.strip().lower() for p in str(raw or "").split(",") if p.strip()]


def subject_blocked(subject, patterns):
    """Отсекает служебные ответы («Принято», «Отклонено») по слову целиком,
    чтобы не задеть темы вида «Принято решение по бюджету»."""
    s = str(subject or "").lower()
    for pat in patterns:
        if re.search(r'(?<![0-9a-zа-яё])' + re.escape(pat) + r'(?![0-9a-zа-яё])', s):
            return pat
    return None


def to_iso(val):
    """Приводит дату Notes к ISO-строке, чтобы фронтенд не угадывал формат."""
    try:
        return datetime(val.year, val.month, val.day,
                        val.hour, val.minute, val.second).isoformat(timespec="seconds")
    except Exception:
        pass
    try:
        return datetime.fromtimestamp(float(val)).isoformat(timespec="seconds")
    except Exception:
        pass
    try:
        return datetime.strptime(str(val)[:19], "%Y-%m-%d %H:%M:%S").isoformat(timespec="seconds")
    except Exception:
        return ""


# ============================================================================
#  СОСТОЯНИЕ ОТРАБОТКИ
#  Отдельная от Notes ось: new / wip / done. Живёт в work_state.json и
#  переживает пересбор — ключ UNID стабилен, порядковый id меняется.
# ============================================================================

def load_work_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_work_state(st):
    if len(st) > STATE_LIMIT:
        ordered = sorted(st.items(), key=lambda kv: kv[1].get("updated", ""), reverse=True)
        st = dict(ordered[:STATE_LIMIT])
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"[!] Не удалось сохранить {STATE_FILE}: {e}")
        return False


def apply_work_state(emails):
    """Проставляет письмам сохранённое состояние отработки."""
    st = load_work_state()
    restored = 0
    for e in emails:
        rec = st.get(e.get("unid", ""))
        if rec:
            e["workState"] = rec.get("status", "new")
            e["stateAt"] = rec.get("updated", "")
            if e["workState"] != "new":
                restored += 1
        else:
            e["workState"] = "new"
            e["stateAt"] = ""
    if restored:
        print(f"[+] Восстановлено состояние отработки для {restored} писем")
    return emails


def set_work_state(unid, status):
    if status not in VALID_STATES:
        return False, f"недопустимое состояние: {status}"
    st = load_work_state()
    if status == "new":
        st.pop(unid, None)
    else:
        st[unid] = {"status": status,
                    "updated": datetime.now().isoformat(timespec="seconds")}
    return save_work_state(st), None


class ReadResolver:
    """Опрашивает все доступные источники меток прочтения и сверяет их.
    Выбор делается ПОСЛЕ сбора — когда видно, сколько насчитал каждый."""

    def __init__(self, db, session):
        self.getread_ok, self.username = probe_get_read(db, session)
        self.sources = gather_unread_sources(db, "сервер" if getattr(db, "Server", "") else "")

        # метки живут в той реплике, где читают почту
        try:
            local = open_local_replica(session, getattr(db, "Server", ""))
        except Exception:
            local = None
        if local is not None:
            print("[i] Проверяю ещё и локальную реплику почтовой базы...")
            self.sources.update(gather_unread_sources(local, "локальная реплика"))

        self.source = None
        self.tally = {}
        names = (["doc.GetRead()"] if self.getread_ok else []) + list(self.sources)
        print(f"[i] Источников меток прочтения найдено: {len(names)}"
              + (f" — {', '.join(names)}" if names else ""))

    def verdicts(self, doc):
        """Мнение каждого источника: True = НЕпрочитано."""
        out = {}
        unid = str(doc.UniversalID).strip()
        if self.getread_ok:
            try:
                out["doc.GetRead()"] = not _call_get_read(doc, self.username)
            except Exception:
                pass
        for name, unids in self.sources.items():
            out[name] = unid in unids
        return out

    def decide(self, all_verdicts, total):
        self.tally = {}
        for v in all_verdicts:
            for name, unread in v.items():
                self.tally[name] = self.tally.get(name, 0) + (1 if unread else 0)

        if not self.tally:
            print("\n[!] Ни один источник меток прочтения недоступен.")
            READ_DETECTION.update(method="недоступно", reliable=False, candidates={},
                                  detail="Notes не отдаёт метки прочтения через COM")
            return None

        print("\n[i] Метки прочтения — сверка источников:")
        width = max(len(n) for n in self.tally)
        for name, cnt in sorted(self.tally.items(), key=lambda kv: -kv[1]):
            share = (cnt / total * 100) if total else 0
            print(f"      {name:<{width}}  непрочитанных {cnt:>4} из {total} ({share:.0f}%)")

        plausible = [n for n, c in self.tally.items() if 0 < c < total]
        if plausible:
            # самый узкий правдоподобный ответ — обычно самый точный
            self.source = min(plausible, key=lambda n: self.tally[n])
            reliable, detail = True, ""
        else:
            nonzero = [n for n, c in self.tally.items() if c > 0]
            self.source = nonzero[0] if nonzero else sorted(self.tally)[0]
            reliable = False
            if self.tally[self.source] >= total:
                detail = "все письма определились как непрочитанные"
            else:
                detail = ("ни один источник не нашёл непрочитанных писем. Обычно это значит, "
                          "что база не ведёт метки прочтения (Свойства базы → Дополнительно → "
                          "«Не поддерживать неотмеченные») либо почту читают в другой реплике")
            print(f"[!] ПОДОЗРЕНИЕ: {detail}")

        print(f"[+] Выбран источник: {self.source}\n")
        READ_DETECTION.update(method=self.source, reliable=reliable, detail=detail,
                              candidates=dict(self.tally))
        return self.source


# ============================================================================
#  ИМЕНА
#  Notes отдаёт «CN=Egor V Stolyarov/O=SVO», а в set.ini записано
#  «Столяров Егор Валерьевич». Без транслитерации списки E_TOP/E_DIR
#  не совпадали ни с одним письмом.
# ============================================================================

TRANSLIT = {
    'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'e','ж':'zh','з':'z',
    'и':'i','й':'y','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r',
    'с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh',
    'щ':'shch','ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya',
}


def translit(text):
    return "".join(TRANSLIT.get(ch, ch) for ch in str(text or "").lower())


def fuzzy_name(text):
    """Сводит разные варианты латиницы к одному виду: Stolyarov/Stoliarov,
    Rylskikh/Rylskih, Tsvetkov/Cvetkov — всё к общему знаменателю."""
    t = translit(text)
    t = re.sub(r'[^a-z]', '', t)
    for a, b in (('shch','sc'), ('sch','sc'), ('sh','s'), ('ch','c'), ('zh','j'),
                 ('kh','h'), ('ts','c'), ('ya','a'), ('ia','a'), ('yu','u'),
                 ('iu','u'), ('ye','e'), ('yo','e'), ('y','i'), ('ck','k')):
        t = t.replace(a, b)
    return re.sub(r'(.)\1+', r'\1', t)


def clean_sender_name(raw):
    """Достаёт человекочитаемое имя из поля From любого формата."""
    src = str(raw or "").strip()
    if not src:
        return "(Неизвестный)"
    m = re.search(r'CN=([^/]+)', src, re.I)          # CN=Egor V Stolyarov/O=SVO
    if m:
        return m.group(1).strip()
    m = re.match(r'^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$', src)   # "Имя" <адрес>
    if m:
        name = m.group(1).strip()
        return name if name else m.group(2).strip()
    return src.strip('"').strip() or "(Неизвестный)"


def name_tokens(name):
    """Значимые части имени: и кириллица, и латиница, в общем написании."""
    parts = re.split(r'[\s,]+', str(name or "").strip())
    return [fuzzy_name(p) for p in parts if len(re.sub(r'[^A-Za-zА-Яа-яЁё]', '', p)) > 2]


def names_match(config_name, sender_name):
    """Фамилия из конфига (обычно первая) против любой части имени в Notes."""
    cfg_tokens = name_tokens(config_name)
    snd_tokens = name_tokens(sender_name)
    if not cfg_tokens or not snd_tokens:
        return False
    surname = cfg_tokens[0]
    if len(surname) < 4:
        return False
    if surname in snd_tokens:
        return True
    # запас на усечённые формы: Abakumova / Abakumov
    for t in snd_tokens:
        if len(t) >= 4 and (t.startswith(surname) or surname.startswith(t)):
            return True
    return False


SYSTEM_HINTS = re.compile(
    r'noreply|no-reply|notification|techsupport|helpdesk|support_|sup\.|monitoring|'
    r'digest|automat|portal|jira|1с:|1c:|hrlink|news@|robot|admin@|mailer', re.I)


def is_system_sender(name, email):
    """Человек или автоматика. Имя человека — минимум два «словесных» токена
    (латиница или кириллица), допускаются инициалы: «Egor V Stolyarov»."""
    blob = f"{name or ''} {email or ''}"
    if SYSTEM_HINTS.search(blob):
        return True
    parts = [p for p in re.split(r'[\s,]+', str(name or "").strip()) if p]
    if len(parts) < 2:
        return True
    wordish = sum(1 for p in parts
                  if re.fullmatch(r'[A-ZА-ЯЁ][a-zа-яё\-]{1,}|[A-ZА-ЯЁ]\.?', p))
    return wordish < 2


def tag_emails_by_persons(cfg, emails):
    """Проставляет письмам E_TOP / E_DIR по спискам из set.ini."""
    def lst(key):
        return [x.strip() for x in cfg.get("TAGS", key, fallback="").split(",") if x.strip()]
    def mails(key):
        return {x.strip().lower() for x in cfg.get("TAGS", key, fallback="").split(",") if x.strip()}

    tops, dirs = lst("e_top"), lst("e_dir")
    top_mails, dir_mails = mails("e_top_emails"), mails("e_dir_emails")
    counts = {"E_TOP": 0, "E_DIR": 0}

    for e in emails:
        addr = (e.get("senderEmail") or "").lower()
        name = e.get("senderName") or ""
        tag = None
        if addr and addr in top_mails:
            tag = "E_TOP"
        elif addr and addr in dir_mails:
            tag = "E_DIR"
        elif any(names_match(p, name) for p in tops):
            tag = "E_TOP"
        elif any(names_match(p, name) for p in dirs):
            tag = "E_DIR"
        if tag:
            e["senderTag"] = tag
            counts[tag] += 1
    if counts["E_TOP"] or counts["E_DIR"]:
        print(f"[+] Отмечено писем: от высшего руководства {counts['E_TOP']}, от директоров {counts['E_DIR']}")
    else:
        print("[i] Писем от лиц из списков E_TOP/E_DIR не найдено")
    return emails


def detect_replied(doc):
    """Определяет, был ли дан ответ на письмо."""
    try:
        if doc.HasItem("$Replied") or doc.HasItem("_Replied"):
            return True
        if doc.HasItem("$ActionFlags"):
            flags = doc.GetItemValue("$ActionFlags")[0]
            if isinstance(flags, (int, float)) and (int(flags) & 1024):
                return True
    except Exception:
        pass
    return False


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
    """Создаёт черновик ответного письма в Lotus Notes."""
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize()
        db = get_mail_database(session)
        original = db.GetDocumentByUNID(unid)
        if not original:
            return None, "Оригинальное письмо не найдено по UNID"

        memo = db.CreateDocument()
        memo.ReplaceItemValue("Form", "Memo")

        try:
            recipients = original.GetItemValue("From")[0]
            memo.ReplaceItemValue("SendTo", recipients)
        except Exception:
            pass

        try:
            subject = str(original.GetItemValue("Subject")[0])
            if not subject.upper().startswith("RE:"):
                subject = "RE: " + subject
            memo.ReplaceItemValue("Subject", subject)
        except Exception:
            memo.ReplaceItemValue("Subject", "RE: Письмо")

        body = memo.CreateRichTextItem("Body")
        body.AppendText(reply_text)

        memo.Save(True, True)

        new_unid = str(memo.UniversalID).strip()
        replica_id = str(db.ReplicaID).replace(":", "").strip()
        link = f"notes:///{replica_id}/0/{new_unid}?OpenDocument&Edit"
        return link, None
    except Exception as e:
        return None, str(e)


def fetch_notes_emails(days=2, max_emails=40, max_chars=800, skip_weekends=True, subject_block=""):
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

    resolver = ReadResolver(db, session)

    replica_id = str(db.ReplicaID).replace(":", "").strip()
    start_dt = window_start(days, skip_weekends)
    start_date = start_dt.strftime("%d.%m.%Y")
    if skip_weekends:
        print(f"[i] Окно выборки: с {start_date} ({days} раб. дн., выходные не в счёт)")
    else:
        print(f"[i] Окно выборки: с {start_date} ({days} календ. дн.)")
    blocked_patterns = parse_subject_filters(subject_block)
    if blocked_patterns:
        print(f"[i] Фильтр тем: {', '.join(blocked_patterns)}")
    search_query = f'@IsAvailable(DeliveredDate) & DeliveredDate >= [{start_date}]'

    collection = db.Search(search_query, None, 0)
    count = collection.Count
    print(f"[+] Найдено писем в базе: {count}")
    if count == 0:
        return []

    emails = []
    verdict_list = []          # мнения источников по каждому письму
    doc = collection.GetFirstDocument()
    idx = 1
    skipped_by_subject = 0

    while doc is not None:
        try:
            subject = doc.GetItemValue("Subject")[0] if doc.HasItem("Subject") else "(Без темы)"
            sender = doc.GetItemValue("From")[0] if doc.HasItem("From") else "(Неизвестный)"
            hit = subject_blocked(subject, blocked_patterns)
            if hit:
                skipped_by_subject += 1
                doc = collection.GetNextDocument(doc)
                continue

            sender_email = extract_sender_email(doc)
            delivered = doc.GetItemValue("DeliveredDate")[0] if doc.HasItem("DeliveredDate") else doc.Created
            verdicts = resolver.verdicts(doc)
            is_replied = detect_replied(doc)
            unid = str(doc.UniversalID).strip()

            body_text = ""
            if doc.HasItem("Body"):
                body_item = doc.GetFirstItem("Body")
                if body_item is not None and hasattr(body_item, "Text"):
                    body_text = body_item.Text.strip()

            clean_body = re.sub(r'\s+', ' ', body_text)[:max_chars]
            date_display = str(delivered)[:16]
            lotus_link = f"notes:///{replica_id}/0/{unid}"

            print(f"  #{idx:02d} {str(sender)[:24]} | {str(subject)[:40]}")

            verdict_list.append(verdicts)
            emails.append({
                "id": str(idx),
                "unid": unid,
                "lotus_url": lotus_link,
                "date": date_display,
                "dateIso": to_iso(delivered),
                "senderName": clean_sender_name(sender),
                "senderEmail": sender_email or "",
                "subject": str(subject),
                "isRead": True,
                "readKnown": True,
                "isReplied": is_replied,
                "needsReply": False,
                "priority": 3,
                "body": clean_body,
                "summary": "",
                "threadMessages": [{"text": clean_body}]
            })
            idx += 1
            if len(emails) >= max_emails:
                break
        except Exception as doc_err:
            print(f"[!] Пропуск: {doc_err}")

        doc = collection.GetNextDocument(doc)

    total = len(emails)
    if skipped_by_subject:
        print(f"[i] Отсеяно по фильтру тем: {skipped_by_subject}")

    # Источник меток выбираем ПОСЛЕ сбора — когда видно, что каждый из них насчитал
    source = resolver.decide(verdict_list, total)
    unread_count = 0
    for e, v in zip(emails, verdict_list):
        if source and source in v:
            e["isRead"] = not v[source]
            e["readKnown"] = True
        else:
            e["isRead"] = True
            e["readKnown"] = False
        if not e["isRead"]:
            unread_count += 1

    print(f"[+] Итог: {total} писем, непрочитанных {unread_count}")
    READ_DETECTION["unread"] = unread_count
    READ_DETECTION["total"] = total
    return apply_work_state(emails)


def set_doc_read_status(unid, read_state=True):
    """Меняет статус прочтения. Возвращает (успех, фактический_статус_прочитано|None)."""
    try:
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize()
        db = get_mail_database(session)
        doc = db.GetDocumentByUNID(unid)
        if not doc:
            return False, None

        username = _session_username(session)

        if read_state:
            try:
                doc.MarkRead(username) if username else doc.MarkRead()
            except Exception:
                doc.MarkRead()
        else:
            try:
                doc.MarkUnread(username) if username else doc.MarkUnread()
            except Exception:
                doc.MarkUnread()

        # Проверяем, что Notes действительно применил метку
        actual = None
        try:
            actual = _call_get_read(doc, username)
        except Exception:
            pass
        return True, actual
    except Exception as e:
        print(f"[!] Ошибка смены статуса UNID {unid}: {e}")
    return False, None


def call_ai_triage(emails, skip_system=True):
    cfg = load_config()
    e_top = cfg.get("TAGS", "e_top", fallback="")
    e_dir = cfg.get("TAGS", "e_dir", fallback="")
    api_key = cfg.get("DEEPSEEK", "api_key", fallback="").strip()
    base_url = cfg.get("DEEPSEEK", "base_url", fallback="https://api.deepseek.com").strip().rstrip("/")
    model = cfg.get("DEEPSEEK", "model", fallback="deepseek-chat").strip()
    batch_size = cfg.getint("NOTES", "batch_size", fallback=15)
    raw_template = load_prompt_template()

    targets = emails
    if skip_system:
        auto = [e for e in emails
                if not e.get("senderTag") and is_system_sender(e.get("senderName"), e.get("senderEmail"))]
        for e in auto:
            e["priority"] = 3
            e["summary"] = e["summary"] or "Автоматическое уведомление — AI-разбор не требуется"
            e["threadMessages"] = [{"text": e["summary"]}]
        targets = [e for e in emails if e not in auto]
        if auto:
            print(f"[i] Пропущено мимо AI как автоматика: {len(auto)} писем "
                  f"(экономия ~{max(0, (len(auto) + batch_size - 1) // batch_size)} запросов)")

    if not targets:
        print("[i] Живых писем для AI-разбора нет")
        return emails

    print(f"[*] Отправка {len(targets)} писем в DeepSeek AI пакетами по {batch_size}...")

    batches = [targets[i:i + batch_size] for i in range(0, len(targets), batch_size)]
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
                with urllib.request.urlopen(req, context=ctx, timeout=120.0) as resp:
                    res_json = json.loads(resp.read().decode('utf-8'))
            except urllib.error.URLError as url_err:
                print(f"[!] Сетевая ошибка при обращении к {base_url}: {url_err}")
                raise
            except Exception as http_err:
                print(f"[!] HTTP ошибка: {http_err}")
                raise

            raw_text = res_json["choices"][0]["message"]["content"].strip()

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
                if not info:
                    continue

                raw_priority = info.get("priority", info.get("criticality", e["priority"]))
                if isinstance(raw_priority, str):
                    try:
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

                if e["priority"] < 1: e["priority"] = 1
                if e["priority"] > 3: e["priority"] = 3

                raw_need = info.get("needsReply", info.get("action_type", ""))
                if isinstance(raw_need, bool):
                    e["needsReply"] = raw_need
                elif isinstance(raw_need, str):
                    e["needsReply"] = raw_need.lower() in ("нужен ответ", "yes", "true", "да", "требует ответа")

                if info.get("sender_tag"):
                    e["senderTag"] = info["sender_tag"]

                if info.get("summary"):
                    e["summary"] = info["summary"]
                    e["threadMessages"] = [{"text": info["summary"]}]

                if info.get("suggested_reply"):
                    e["suggestedReply"] = info["suggested_reply"]
                elif info.get("suggestedReply"):
                    e["suggestedReply"] = info["suggestedReply"]

                if info.get("action_type", "") == "Отвечено":
                    e["isReplied"] = True

            print(f"[✓] Пакет {batch_idx}/{total_batches} обработан успешно!")
        except Exception as e:
            print(f"[!] AI Triage для пакета {batch_idx}/{total_batches} пропущен из-за ошибки: {e}")

    print(f"[✓] AI-разбор завершён: обработано {len(targets)} из {len(emails)} писем")
    return emails


class NotesWebHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def log_message(self, fmt, *args):
        pass  # не засоряем консоль обращениями к статике

    def do_GET(self):
        if self.path == "/api/settings":
            cfg = load_config()
            api_key = cfg.get("DEEPSEEK", "api_key", fallback="").strip()
            data = {
                "e_top": cfg.get("TAGS", "e_top", fallback=""),
                "e_dir": cfg.get("TAGS", "e_dir", fallback=""),
                "e_top_emails": cfg.get("TAGS", "e_top_emails", fallback=""),
                "e_dir_emails": cfg.get("TAGS", "e_dir_emails", fallback=""),
                "default_days": cfg.getint("NOTES", "default_days", fallback=2),
                "max_emails": cfg.getint("NOTES", "max_emails", fallback=40),
                "max_body_chars": cfg.getint("NOTES", "max_body_chars", fallback=800),
                "batch_size": cfg.getint("NOTES", "batch_size", fallback=15),
                "base_url": cfg.get("DEEPSEEK", "base_url", fallback="https://api.deepseek.com"),
                "model": cfg.get("DEEPSEEK", "model", fallback="deepseek-chat"),
                "api_key_set": bool(api_key),
                "api_key_hint": (api_key[:4] + "…" + api_key[-4:]) if len(api_key) > 10 else "",
                "mark_read_on_done": cfg.getboolean("UI", "mark_read_on_done", fallback=True),
                "skip_weekends": cfg.getboolean("UI", "skip_weekends", fallback=True),
                "skip_ai_for_system": cfg.getboolean("UI", "skip_ai_for_system", fallback=True),
                "subject_block": cfg.get("UI", "subject_block", fallback=DEFAULT_SUBJECT_BLOCK),
                "read_detection": READ_DETECTION
            }
            self.send_json(200, data)

        elif self.path.split("?")[0] == "/api/data":
            # кэш последнего сбора, но состояние отработки — актуальное с диска
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    cache = json.load(f)
                cache["emails"] = apply_work_state(cache.get("emails", []))
                self.send_json(200, cache)
            except Exception:
                self.send_json(200, {"emails": []})
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
                skip_weekends = cfg.getboolean("UI", "skip_weekends", fallback=True)
                subject_block = cfg.get("UI", "subject_block", fallback=DEFAULT_SUBJECT_BLOCK)
                skip_ai = bool(req_data.get("skip_ai", False))

                emails = fetch_notes_emails(days=days, max_emails=max_emails, max_chars=max_chars,
                                            skip_weekends=skip_weekends, subject_block=subject_block)
                if emails:
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

                    emails = tag_emails_by_persons(cfg, emails)

                    if not skip_ai:
                        emails = call_ai_triage(
                            emails,
                            skip_system=cfg.getboolean("UI", "skip_ai_for_system", fallback=True))
                    else:
                        print("[i] AI-анализ пропущен по запросу (быстрый сбор)")

                result = {
                    "emails": emails,
                    "read_detection": READ_DETECTION,
                    "synced_at": datetime.now().isoformat(timespec="seconds"),
                    "days": days
                }
                with open(DATA_FILE, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)

                self.send_json(200, result)
            except Exception as e:
                print(f"[!] Ошибка: {e}")
                self.send_json(500, {"error": str(e)})
            finally:
                ANALYSIS_LOCK.release()

        elif self.path == "/api/mark-read":
            try:
                unid = req_data.get("unid", "")
                is_read = req_data.get("is_read", True)
                success, actual = set_doc_read_status(unid, is_read)
                self.send_json(200, {
                    "status": "ok" if success else "failed",
                    "is_read": actual if actual is not None else is_read,
                    "verified": actual is not None
                })
            except Exception as e:
                self.send_json(500, {"error": str(e)})

        elif self.path == "/api/set-state":
            try:
                unid = req_data.get("unid", "")
                status = req_data.get("status", "new")
                if not unid:
                    self.send_json(400, {"error": "unid обязателен"})
                    return
                saved, err = set_work_state(unid, status)
                if err:
                    self.send_json(400, {"error": err})
                    return
                # отработка может заодно проставлять метку прочтения в Notes
                if req_data.get("mark_read"):
                    try:
                        set_doc_read_status(unid, True)
                    except Exception as mark_err:
                        print(f"[!] Отработка: метку прочтения выставить не удалось: {mark_err}")
                self.send_json(200, {"status": "ok" if saved else "failed"})
            except Exception as e:
                self.send_json(500, {"error": str(e)})

        elif self.path == "/api/save-settings":
            try:
                cfg = load_config()
                for sec in ("TAGS", "NOTES", "DEEPSEEK", "UI"):
                    ensure_section(cfg, sec)

                def put(section, opt, val):
                    if val is None:
                        return
                    cfg.set(section, opt, str(val))

                put("TAGS", "e_top", req_data.get("e_top", ""))
                put("TAGS", "e_dir", req_data.get("e_dir", ""))

                # адреса можно править руками — сохраняем как есть,
                # автопоиск ниже только дополнит список, ничего не затирая
                if "e_top_emails" in req_data:
                    put("TAGS", "e_top_emails", normalize_email_list(req_data["e_top_emails"]))
                if "e_dir_emails" in req_data:
                    put("TAGS", "e_dir_emails", normalize_email_list(req_data["e_dir_emails"]))

                for key, section, opt in (("default_days", "NOTES", "default_days"),
                                          ("max_emails", "NOTES", "max_emails"),
                                          ("max_body_chars", "NOTES", "max_body_chars"),
                                          ("batch_size", "NOTES", "batch_size")):
                    if key in req_data:
                        try:
                            put(section, opt, int(req_data[key]))
                        except (TypeError, ValueError):
                            self.send_json(400, {"error": f"поле «{key}» должно быть числом"})
                            return

                if "base_url" in req_data:
                    put("DEEPSEEK", "base_url", str(req_data["base_url"]).strip().rstrip("/"))
                if "model" in req_data:
                    put("DEEPSEEK", "model", str(req_data["model"]).strip())
                # пустой ключ означает «не менять», а не «стереть»
                if str(req_data.get("api_key", "")).strip():
                    put("DEEPSEEK", "api_key", str(req_data["api_key"]).strip())

                if "mark_read_on_done" in req_data:
                    put("UI", "mark_read_on_done", "True" if req_data["mark_read_on_done"] else "False")
                if "skip_weekends" in req_data:
                    put("UI", "skip_weekends", "True" if req_data["skip_weekends"] else "False")
                if "skip_ai_for_system" in req_data:
                    put("UI", "skip_ai_for_system", "True" if req_data["skip_ai_for_system"] else "False")
                if "subject_block" in req_data:
                    put("UI", "subject_block", str(req_data["subject_block"]).strip())

                cached_emails = []
                try:
                    with open(DATA_FILE, "r", encoding="utf-8") as f:
                        cached_emails = json.load(f).get("emails", [])
                except Exception:
                    pass

                if cached_emails:
                    try:
                        e_top_emails, _ = resolve_names_to_emails(cfg, cached_emails, "e_top", "e_top_emails")
                        e_dir_emails, _ = resolve_names_to_emails(cfg, cached_emails, "e_dir", "e_dir_emails")
                        cfg.set("TAGS", "e_top_emails", e_top_emails)
                        cfg.set("TAGS", "e_dir_emails", e_dir_emails)
                    except Exception as res_err:
                        print(f"[!] Автопоиск адресов пропущен: {res_err}")

                try:
                    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                        cfg.write(f)
                except Exception as write_err:
                    self.send_json(500, {"error": f"не удалось записать set.ini: {write_err}"})
                    return

                self.send_json(200, {
                    "status": "ok",
                    "e_top_emails": cfg.get("TAGS", "e_top_emails", fallback=""),
                    "e_dir_emails": cfg.get("TAGS", "e_dir_emails", fallback="")
                })
            except Exception as e:
                import traceback; traceback.print_exc()
                self.send_json(500, {"error": f"{type(e).__name__}: {e}"})

        elif self.path == "/api/clear":
            try:
                what = req_data.get("what", "all")
                removed = []
                if what in ("all", "emails"):
                    try:
                        if os.path.exists(DATA_FILE):
                            os.remove(DATA_FILE); removed.append("собранные письма")
                    except Exception as e:
                        self.send_json(500, {"error": f"не удалось удалить кэш писем: {e}"})
                        return
                if what in ("all", "state"):
                    try:
                        if os.path.exists(STATE_FILE):
                            os.remove(STATE_FILE); removed.append("отметки отработки")
                    except Exception as e:
                        self.send_json(500, {"error": f"не удалось удалить отметки: {e}"})
                        return
                print(f"[i] Очистка данных: {', '.join(removed) if removed else 'нечего удалять'}")
                self.send_json(200, {"status": "ok", "removed": removed})
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

        else:
            self.send_json(404, {"error": "Неизвестный метод"})

    def send_json(self, status, payload):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-store')
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
