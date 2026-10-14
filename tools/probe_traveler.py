# -*- coding: utf-8 -*-
"""
Диагностика доступа к почте Domino по HTTPS (гипотеза: traveler.svo.aero:443).

Отвечает на три вопроса:
  1. Доступен ли сервер по HTTPS и что за TLS-сертификат.
  2. Какие интерфейсы на нём открыты: iNotes (веб-почта), Domino Data Service
     (REST+JSON, путь /api/notes), адресная книга names.nsf.
  3. Работает ли HTTP-аутентификация данными пользователя и можно ли прочитать
     список последних писем из папки ($Inbox).

Дополнительно: TCP-пробы типовых портов Domino (1352 NRPC, 143/993 IMAP,
110 POP3, 63148 DIIOP) до самого traveler и до почтового сервера из
окружения Notes (MailServer), если клиент установлен и разблокирован.

Скрипт ТОЛЬКО читает и печатает — ничего не меняет ни на сервере, ни локально.
Пароль НЕ вшивается в код: берётся из переменной окружения TRAVELER_PASSWORD
или вводится с клавиатуры (в терминал не отображается). Для сетевых проб без
аутентификации можно запустить с флагом --no-auth.

Запуск:
  py tools/probe_traveler.py                       # полный прогон, пароль спросит
  py tools/probe_traveler.py --no-auth             # только сетевые пробы
  py tools/probe_traveler.py login mail_file       # логин и имя почтового файла
                                                   # (например: mail/dsoldatov.nsf)
"""

import base64
import getpass
import os
import re
import socket
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

HOST = "traveler.svo.aero"
DNS_SUFFIX = "svo.aero"
DEFAULT_LOGIN = "dmitriy soldatov"
DEFAULT_MAIL_FILE_CANDIDATES = [
    "mail/dsoldatov.nsf",   # эвристика: инициал + фамилия — только проба GET, без записи
    "mail/dsoldatov",
]
TCP_PORTS = (1352, 143, 993, 110, 63148)
TIMEOUT = 10

SEP = "=" * 72


def head(title):
    print(f"\n{SEP}\n  {title}\n{SEP}")


# ----------------------------------------------------------------- сеть/TLS

def dns_and_tcp(host, port):
    """TCP-проба. Возвращает ('open'|'closed'|'error', подробность)."""
    try:
        with socket.create_connection((host, port), timeout=5):
            return "open", ""
    except socket.timeout:
        return "closed", "таймаут"
    except ConnectionRefusedError:
        return "closed", "соединение отклонено"
    except OSError as e:
        return "error", str(e)


def tls_info(host):
    """Сертификат сервера. Сначала строгая проверка, при ошибке — без неё."""
    ctx_verify = ssl.create_default_context()
    try:
        with socket.create_connection((host, 443), timeout=TIMEOUT) as sock:
            with ctx_verify.wrap_socket(sock, server_hostname=host) as tls:
                return tls.getpeercert(), "проверка TLS пройдена"
    except ssl.SSLCertVerificationError as e:
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            with socket.create_connection((host, 443), timeout=TIMEOUT) as sock:
                with ctx.wrap_socket(sock, server_hostname=host) as tls:
                    cert = tls.getpeercert()
            return cert, f"СЕРТИФИКАТ НЕ ПРОШЁЛ ПРОВЕРКУ: {e.verify_message}"
        except OSError as e2:
            return None, f"TLS-рукопожатие не удалось: {e2}"
    except OSError as e:
        return None, f"соединение не установлено: {e}"


# -------------------------------------------------------------------- HTTP

class Resp:
    def __init__(self, status, reason, headers, body, final_url=""):
        self.status = status
        self.reason = reason
        self.headers = headers
        self.body = body  # bytes, не более 4000 байт
        self.final_url = final_url


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """Запрещает urllib следовать за 302 — Location читаем сами."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def http_get(url, login=None, password=None, timeout=TIMEOUT, follow=True, big=False):
    """GET без/с Basic-аутентификацией. TLS-проверка отключена
    (в проекте принято [NETWORK] ssl_verify=False; сертификат печатаем отдельно).
    follow=False — не переходить по 302 (нужен заголовок Location).
    big=True — читать до 200000 байт (для XML-выгрузок)."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    limit = 200000 if big else 4000
    req = urllib.request.Request(url, method="GET")
    req.add_header("User-Agent", "LOTUS-probe/1.0")
    if login is not None and password is not None:
        token = base64.b64encode(f"{login}:{password}".encode("utf-8")).decode("ascii")
        req.add_header("Authorization", f"Basic {token}")
    try:
        if follow:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return Resp(r.status, r.reason, dict(r.headers), r.read(limit), r.geturl())
        opener = urllib.request.build_opener(
            urllib.request.HTTPSHandler(context=ctx), _NoRedirect())
        with opener.open(req, timeout=timeout) as r:
            return Resp(r.status, r.reason, dict(r.headers), r.read(limit), r.geturl())
    except urllib.error.HTTPError as e:
        return Resp(e.code, e.reason, dict(e.headers), e.read(limit), getattr(e, "url", ""))
    except (urllib.error.URLError, OSError) as e:
        return None


def short(resp, limit=160):
    if resp is None:
        return "нет ответа (сеть/TLS)"
    body = resp.body.decode("utf-8", "replace").replace("\n", " ")
    return f"{resp.status} {resp.reason} | {body[:limit]}"


# --------------------------------------------------------------- inbox XML

def parse_read_view_entries(xml_bytes):
    """Разбор ответа Domino ?ReadViewEntries (XML): первые N записей."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        return [f"<не XML: {e}>"]
    rows = []
    for viewentry in root.iter("viewentry"):
        texts = []
        for col in viewentry.iter("entrydata"):
            name = col.get("name") or f"col{col.get('columnnumber')}"
            val = " ".join(t.text or "" for t in col.iter("text")).strip()
            if val:
                texts.append(f"{name}={val}")
        rows.append(" | ".join(texts)[:300])
    return rows


def person_doc_hints(host, login, password):
    """Ищет person-документ в names.nsf (StartKey по фамилии) и вытаскивает
    из него MailFile и MailServer. Возвращает (mail_file, mail_server)."""
    parts = login.split()
    last = parts[-1] if parts else login
    url = (f"https://{host}/names.nsf/People?ReadViewEntries"
           f"&StartKey={urllib.parse.quote(last)}&Count=5")
    r = http_get(url, login, password, big=True)
    print(f"  names.nsf StartKey «{last}» -> {short(r, 60)}")
    if r is None or r.status != 200:
        return "", ""
    try:
        root = ET.fromstring(r.body)
    except ET.ParseError as e:
        print(f"  <не XML: {e}>")
        return "", ""
    # 1-й шаг: UNID и отображаемое имя из вида People
    entries = []
    for ve in root.iter("viewentry"):
        unid = ve.get("unid")
        texts = []
        for col in ve.iter("entrydata"):
            texts.append(" ".join(t.text or "" for t in col.iter("text")).strip())
        entries.append((unid, " | ".join(texts)))
    print(f"  найдено записей в People: {len(entries)}"
          + (f"; первая: «{entries[0][1][:50]}» (unid={entries[0][0]})" if entries else ""))
    # 2-й шаг: открываем документ (таймаут увеличен — сервер отвечал 0 байт),
    # ищем MailFile/MailServer в HTML
    for unid, label in entries:
        if not unid:
            continue
        d = http_get(f"https://{host}/names.nsf/People/{unid}?OpenDocument",
                     login, password, big=True, timeout=30)
        body = d.body.decode("utf-8", "replace") if d else ""
        m1 = re.search(r"([\w\-]+[/\\][\w\-\. ]+\.nsf)", body, re.I)
        m2 = re.search(r"CN=([\w\-]+)", body)
        if m1:
            mf = m1.group(1).replace("\\", "/")
            msrv = m2.group(1) if m2 else ""
            print(f"  person-документ «{label[:40]}» -> MailFile={mf} MailServer={msrv or '?'}")
            return mf, msrv
        bsize = len(d.body) if d else 0
        bstat = f"{d.status}" if d else "нет ответа"
        print(f"  person-документ «{label[:40]}»: {bstat}, {bsize} байт")
        if d is not None and d.status == 200:
            # извлекаем значения полей mailfile_fld / mailserver_fld
            flat = re.sub(r"\s+", " ", body)
            def field(fid):
                m = re.search(
                    rf'id="{fid}"[^>]*>.*?<font[^>]*>([^<]+)</font>', flat, re.I)
                if not m:
                    m = re.search(rf'id="{fid}"[^>]*>\s*([^<]+)', flat, re.I)
                return m.group(1).strip() if m else ""
            mf_val = field("mailfile_fld")
            msrv_val = field("mailserver_fld")
            if mf_val:
                mf_norm = mf_val.replace("\\", "/")
                if not mf_norm.lower().endswith(".nsf"):
                    mf_norm += ".nsf"
                print(f"  person-документ -> MailFile={mf_val} MailServer={msrv_val or '?'} (норм.: {mf_norm})")
                return mf_norm, msrv_val
            # диагностика: все вхождения 'mail' в HTML
            hits = [m.start() for m in re.finditer(r"mail", flat, re.I)]
            for pos in hits[:8]:
                print(f"    …{flat[max(0, pos - 60):pos + 120]}…")
            return "", ""  # документ открылся — дальше person-документы не перебираем
    return "", ""


# ------------------------------------------------------------------- COM

def mail_server_from_com():
    """MailServer/MailFile/ReplicaID из окружения клиента Notes (если доступен)."""
    try:
        import win32com.client  # noqa
        session = win32com.client.Dispatch("Lotus.NotesSession")
        session.Initialize("")
        srv = session.GetEnvironmentString("MailServer", True)
        mf = session.GetEnvironmentString("MailFile", True)
        replica = ""
        try:
            db = session.GetDatabase(srv, mf)
            replica = db.ReplicaID or ""
        except Exception:
            pass
        return srv, mf, replica
    except Exception as e:
        return None, f"COM-сессия не открыта: {e}", ""


def cn_to_hostname(cn):
    """Эвристика: 'CN=MAIL01/OU=SVO/O=AERO' -> 'MAIL01'."""
    part = cn.split("/")[0]
    if part.upper().startswith("CN="):
        part = part[3:]
    return part.strip()


# ------------------------------------------------------------------ main

def main():
    no_auth = "--no-auth" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]

    login = args[0] if args else DEFAULT_LOGIN
    mail_file = args[1] if len(args) > 1 else os.environ.get("TRAVELER_MAIL_FILE", "")
    password = os.environ.get("TRAVELER_PASSWORD", "")

    results = {}

    # 1. DNS + TCP
    head("1. СЕТЬ: DNS И TCP")
    try:
        ips = socket.gethostbyname_ex(HOST)[2]
        print(f"  {HOST} -> IP: {', '.join(ips)}")
    except OSError as e:
        print(f"  [!] DNS не разрешил {HOST}: {e}")
        ips = []
    for port in (443,):
        state, detail = dns_and_tcp(HOST, port)
        print(f"  TCP {HOST}:{port:<6} -> {state} {detail}")

    # 2. TLS
    head("2. TLS-СЕРТИФИКАТ")
    cert, note = tls_info(HOST)
    print(f"  {note}")
    if cert:
        print(f"  Subject   : {cert.get('subject')}")
        print(f"  Issuer    : {cert.get('issuer')}")
        print(f"  Действует до: {cert.get('notAfter')}")

    # 3. Интерфейсы без аутентификации
    head("3. ИНТЕРФЕЙСЫ БЕЗ АУТЕНТИФИКАЦИИ")
    probes = [
        ("главная сервера", f"https://{HOST}/"),
        ("names.nsf (адресная книга)", f"https://{HOST}/names.nsf?OpenDatabase"),
        ("DDS: /api/notes (REST+JSON)", f"https://{HOST}/api/notes"),
        ("каталог /mail/", f"https://{HOST}/mail/"),
        ("Traveler", f"https://{HOST}/traveler"),
        ("iNotes вход", f"https://{HOST}/mail/WebMailLite.nsf"),
        ("iNotes вход (mail/w.nsf)", f"https://{HOST}/mail/w.nsf"),
        ("веб-почта /webmail", f"https://{HOST}/webmail"),
        ("веб-почта /notes", f"https://{HOST}/notes"),
    ]
    for name, url in probes:
        r = http_get(url)
        www = r.headers.get("WWW-Authenticate", "") if r else ""
        print(f"  {name:<32} {short(r)}{(' | WWW-Authenticate: ' + www) if www and r and r.status == 401 else ''}")

    # 4. Аутентификация
    auth_ok = False
    if not no_auth:
        head("4. HTTP-АУТЕНТИФИКАЦИЯ (Basic)")
        if not password:
            password = getpass.getpass(f"  Пароль для '{login}' (Enter — пропустить): ")
        if password:
            r = http_get(f"https://{HOST}/names.nsf?OpenDatabase", login, password)
            auth_ok = r is not None and r.status == 200
            print(f"  names.nsf с Basic-парой -> {short(r, 80)}")
            print(f"  Вывод: {'аутентификация РАБОТАЕТ' if auth_ok else 'аутентификация НЕ прошла (пароль/тип пароля/метод)'}")
        else:
            print("  Пропущено: пароль не введён.")
    else:
        head("4. HTTP-АУТЕНТИФИКАЦИЯ")
        print("  Пропущено (флаг --no-auth).")

    # 5. Почтовый файл и ($Inbox)
    inbox_rows = []
    if auth_ok:
        head("5. ПОЧТОВЫЙ ФАЙЛ И ($Inbox)")
        chosen = None
        # --- Способ A: person-документ в names.nsf (StartKey по фамилии)
        mf_hint, msrv_hint = person_doc_hints(HOST, login, password)
        if mf_hint:
            chosen = mf_hint
            print(f"  База из names.nsf          : {chosen}")
        # --- Способ B: iNotes Redirect — сервер сам называет базу вошедшего
        if not chosen:
            r = http_get(f"https://{HOST}/iwaredir.nsf?Redirect",
                         login, password, follow=False)
            loc = (r.headers.get("Location", "") if r else "")
            print(f"  iwaredir.nsf (iNotes)      -> {short(r, 80)}")
            if r is not None and r.status == 500:
                print("  Тело ошибки 500 (первые 300):",
                      r.body.decode("utf-8", "replace").replace("\n", " ")[:300])
            if loc:
                print(f"  Location                   : {loc}")
                m = re.search(r"/((?:[\w\-\.~/]+)\.nsf)", loc, re.I)
                if m:
                    chosen = m.group(1)
                    print(f"  База из Location           : {chosen}")
        # --- Способ C: открытие по ReplicaID из COM
        srv, mf_com, replica = mail_server_from_com()
        if not chosen and replica:
            url = f"https://{HOST}/{replica}?OpenDatabase"
            r = http_get(url, login, password)
            print(f"  По ReplicaID {replica} -> {short(r, 80)}")
            if r is not None and r.status == 200:
                chosen = replica
        # --- Способ D: прямые кандидаты пути
        candidates = [mail_file] if mail_file else DEFAULT_MAIL_FILE_CANDIDATES
        for mf in candidates:
            if chosen:
                break
            if not mf:
                continue
            url = f"https://{HOST}/{mf.rstrip('/')}?OpenDatabase"
            r = http_get(url, login, password)
            print(f"  {mf:<28} -> {short(r, 80)}")
            if r is not None and r.status == 200:
                chosen = mf
        # --- Способ E: прямой HTTPS на почтовый сервер из COM (Mail8.svo.aero)
        srv, mf_com, replica = mail_server_from_com()
        mail_hosts = []
        if srv:
            g = cn_to_hostname(srv)
            if g:
                mail_hosts.append(g)
                mail_hosts.append(f"{g.lower()}.{DNS_SUFFIX}")
        if not chosen:
            for mh in mail_hosts:
                for mf in ([mf_com] if mf_com else []) + DEFAULT_MAIL_FILE_CANDIDATES:
                    if not mf:
                        continue
                    r = http_get(f"https://{mh}/{mf.rstrip('/')}?OpenDatabase",
                                 login, password)
                    print(f"  {mh}/{mf} -> {short(r, 80)}")
                    if r is not None and r.status == 200:
                        chosen = mf.replace("\\", "/")
                        print(f"  РАБОЧАЯ СВЯЗКА: https://{mh}/{chosen}")
                        break
                if chosen:
                    break
        # --- Способ F: iNotes Redirect по правильному синтаксису (?Redirect=1&Mode=...)
        redir_base = [
            f"https://{HOST}/iwaredir.nsf?Redirect=1&Mode=2",
            f"https://{HOST}/iwaredir.nsf?Redirect=1",
            f"https://{HOST}/iwaredir.nsf?Login&RedirectTo=/$$iNotes.nsf?Inbox",
        ]
        for u in redir_base:
            r = http_get(u, login, password, follow=False, timeout=30)
            loc = (r.headers.get("Location", "") if r else "")
            print(f"  {u.split('?')[1]:<28} -> {short(r, 60)}")
            if loc:
                print(f"    Location: {loc}")
            if r is not None and r.status == 200 and not loc:
                body = r.body.decode("utf-8", "replace").replace("\n", " ")
                m = re.search(r"((?:[\w\-\.~/\\]+)\.nsf)", body, re.I)
                if m and not chosen:
                    cand = m.group(1).replace("\\", "/")
                    print(f"    База из тела ответа: {cand}")
                    chosen = cand
        # --- Способ G: мобильный Traveler UI (тоже умеет листать inbox)
        for mu in (
            f"https://{HOST}/servlet/traveler/TravelerServlet",
            f"https://{HOST}/traveler/appinfo",
        ):
            r = http_get(mu, login, password, timeout=30)
            print(f"  Traveler {mu.rsplit('/', 1)[-1]:<16} -> {short(r, 60)}")
        # --- Читаем ($Inbox) только если база реально открылась по https
        if chosen:
            inbox_url = (
                f"https://{HOST}/{chosen.rstrip('/')}/"
                f"({urllib.parse.quote('($Inbox)')})?ReadViewEntries&Count=10"
            )
            r = http_get(inbox_url, login, password, big=True, timeout=30)
            print(f"  ReadViewEntries ($Inbox)   -> {short(r, 80)}")
            if r is not None and r.status == 200:
                inbox_rows = parse_read_view_entries(r.body)
            # DDS (REST+JSON) по той же базе
            dds_url = f"https://{HOST}/api/notes/{chosen}/views/($Inbox)?count=5"
            r2 = http_get(dds_url, login, password)
            print(f"  DDS views/($Inbox)         -> {short(r2, 80)}")
        else:
            print("  [!] База не найдена ни через iwaredir, ни по ReplicaID, ни по пути.")
            print("      Уточните имя и запустите:")
            print(f"      py tools/probe_traveler.py '{login}' mail/<файл>.nsf")
    else:
        head("5. ПОЧТОВЫЙ ФАЙЛ И ($Inbox)")
        print("  Пропущено (нет аутентификации).")

    # 6. Порты Domino
    head("6. TCP-ПОРТЫ DOMINO")
    targets = [HOST]
    if not auth_ok:
        srv, mf_com, replica = mail_server_from_com()
    if srv:
        print(f"  MailServer (из Notes) : {srv}")
        print(f"  MailFile   (из Notes) : {mf_com}")
        if replica:
            print(f"  ReplicaID  (из Notes) : {replica}")
        host_guess = cn_to_hostname(srv)
        print(f"  Проба hostname: {host_guess}")
        try:
            socket.gethostbyname(host_guess)
            targets.append(host_guess)
        except OSError:
            print("  (hostname по CN не резолвится — пробуем только traveler)")
    else:
        print(f"  {mf_com if mf_com else 'COM-сессия не открыта (клиент закрыт/заблокирован).'}")
    for target in targets:
        for port in TCP_PORTS:
            state, detail = dns_and_tcp(target, port)
            print(f"  TCP {target}:{port:<6} -> {state} {detail}")

    # СВОДКА
    head("СВОДКА (скопируйте этот блок)")
    print(f"  Хост            : {HOST} (IP: {', '.join(ips) if ips else '—'})")
    print(f"  TLS             : {note}")
    if cert:
        print(f"  TLS до          : {cert.get('notAfter')}")
    if not no_auth:
        print(f"  Basic-аутентиф. : {'успех' if auth_ok else 'НЕ прошла'}")
    for row in inbox_rows[:5]:
        print(f"  Письмо          : {row}")
    print(f"  MailServer      : {srv if srv else '—'}")

    print("\nГотово. Скрипт ничего не менял — только читал.")


if __name__ == "__main__":
    main()
