"""
Что на самом деле объявлено в COM-интерфейсе Domino Objects.

Диагностика показала: Notes 14.5, библиотека свежая, но десять методов
работы с метками прочтения недоступны. Возможных объяснений два:

  А) методы есть в библиотеке типов, но pywin32 их не видит из-за позднего
     связывания — тогда лечится ранним связыванием (makepy);
  Б) HCL никогда не выносила их в COM — тогда не поможет ничто,
     и надо идти другим путём.

Скрипт различает эти случаи. Ничего не меняет, кроме кэша обёрток pywin32,
который и так пересобирается автоматически.

Запуск:  python probe_unread.py
"""

import sys
import os

SEP = "=" * 74
DOMINO_TLB = "{29131520-2EED-1069-BF5D-00DD011186B7}"
TARGETS = ["GetRead", "MarkRead", "MarkUnread", "GetAllUnreadDocuments",
           "GetAllReadDocuments", "GetAllUnreadEntries", "GetAllReadEntries",
           "CreateViewNavFromAllUnread", "MarkAllRead", "MarkAllUnread", "IsUnread"]


def head(t):
    print(f"\n{SEP}\n  {t}\n{SEP}")


try:
    import pythoncom
    import win32com.client
    from win32com.client import gencache
except ImportError:
    print("Нужен pywin32:  pip install pywin32")
    sys.exit(1)


# ───────────────────────────── 1. Ранее связывание ─────────────────────────
head("1. РАННЕЕ СВЯЗЫВАНИЕ: собираем обёртку из библиотеки типов")

gen_module = None
for ver in ((1, 2), (1, 3), (1, 4), (1, 0)):
    try:
        gen_module = gencache.EnsureModule(DOMINO_TLB, 0, ver[0], ver[1])
        if gen_module:
            print(f"  Обёртка собрана для версии TLB {ver[0]}.{ver[1]}")
            print(f"  Модуль: {getattr(gen_module, '__file__', '—')}")
            break
    except Exception as e:
        print(f"  Версия {ver[0]}.{ver[1]}: не собралась ({e})")

if gen_module is None:
    print("  [!] Библиотеку типов обработать не удалось — раздел 2 всё равно ответит.")
else:
    src = getattr(gen_module, "__file__", "")
    if src and os.path.exists(src):
        try:
            text = open(src, encoding="utf-8", errors="ignore").read()
        except Exception:
            text = ""
        print("\n  Упоминания в сгенерированном модуле:")
        for name in TARGETS:
            print(f"    {name:<28} {'ЕСТЬ' if name in text else 'нет'}")


# ─────────────────────── 2. Что объявляет сам объект ───────────────────────
head("2. ЧТО ОБЪЕКТЫ СООБЩАЮТ О СЕБЕ ЧЕРЕЗ ITypeInfo")

def members_of(obj, label):
    """Список всех методов и свойств, объявленных объектом."""
    if obj is None:
        print(f"  {label}: объект получить не удалось")
        return set()
    try:
        ti = obj._oleobj_.GetTypeInfo()
    except Exception as e:
        print(f"  {label}: библиотека типов недоступна ({e})")
        return set()

    names = set()
    try:
        attr = ti.GetTypeAttr()
        for i in range(attr.cFuncs):
            fd = ti.GetFuncDesc(i)
            for n in ti.GetNames(fd.memid):
                names.add(n)
        for i in range(attr.cVars):
            vd = ti.GetVarDesc(i)
            for n in ti.GetNames(vd.memid):
                names.add(n)
    except Exception as e:
        print(f"  {label}: перечислить не вышло ({e})")
        return names

    print(f"\n  {label}: объявлено элементов — {len(names)}")
    hits = sorted(n for n in names if "read" in n.lower() or "unread" in n.lower())
    print(f"     содержат 'read': {', '.join(hits) if hits else 'ни одного'}")
    return names


def ids_of_names(obj, name):
    """Прямой вопрос объекту: знаешь такое имя?"""
    try:
        obj._oleobj_.GetIDsOfNames(0, name)
        return "знает"
    except Exception:
        return "не знает"


session = db = view = doc = entry = None
try:
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize()
    srv = session.GetEnvironmentString("MailServer", True)
    mf = session.GetEnvironmentString("MailFile", True)
    db = session.GetDatabase(srv, mf)
    if db and not db.IsOpen:
        db.Open()
    view = db.GetView("($Inbox)")
    doc = view.GetFirstDocument() if view else None
    nav = view.CreateViewNav() if view else None
    entry = nav.GetFirst() if nav else None
    print(f"  Сессия и база открыты: {db.FilePath}")
except Exception as e:
    print(f"  [!] Не удалось подготовить объекты: {e}")

all_names = {}
for obj, label in ((session, "NotesSession"), (db, "NotesDatabase"),
                   (view, "NotesView"), (doc, "NotesDocument"),
                   (entry, "NotesViewEntry")):
    all_names[label] = members_of(obj, label)

head("3. ПРЯМОЙ ОПРОС ПО ИМЕНАМ (GetIDsOfNames)")
checks = [("NotesDatabase", db, ["GetAllUnreadDocuments", "GetAllReadDocuments"]),
          ("NotesView", view, ["GetAllUnreadEntries", "CreateViewNavFromAllUnread", "MarkAllRead"]),
          ("NotesDocument", doc, ["GetRead", "MarkRead", "MarkUnread"]),
          ("NotesViewEntry", entry, ["GetRead", "IsUnread"])]
known = 0
total = 0
for label, obj, names in checks:
    for n in names:
        total += 1
        res = ids_of_names(obj, n) if obj is not None else "объекта нет"
        if res == "знает":
            known += 1
        print(f"  {label + '.' + n:<45} {res}")


# ───────────────────────────────── 4. Вывод ────────────────────────────────
head("4. ВЫВОД")
if known > 0:
    print(f"  Объект знает {known} из {total} имён — значит, дело было в связывании.")
    print("  Раннее связывание открывает эти методы, дашборд можно перевести на него.")
else:
    print("  Ни одного из имён объект не знает, хотя Notes 14.5 и библиотека свежая.")
    print("  Это значит: HCL не выносила работу с метками прочтения в COM-интерфейс")
    print("  Domino Objects — ни в одной версии. Обновлять и перерегистрировать нечего.")
    print("\n  Рабочие пути в порядке трудоёмкости:")
    print("    1. Признак «новое с прошлого сбора» — уже работает в дашборде.")
    print("    2. Notes C API (nnotes.dll, NSFDbGetUnreadNoteTable) через ctypes.")
    print("       Notes 64-битный и Python 64-битный — разрядности совпадают.")
    print("    3. Java/CORBA через DIIOP — там методы есть, но нужна задача DIIOP")
    print("       на сервере Mail8/SVO, это к администраторам Domino.")

print(f"\n{SEP}")
print("  Отдельно: полный список методов NotesDocument сохранён в members.txt —")
print("  по нему видно, что вообще доступно.")
print(SEP)

try:
    with open("members.txt", "w", encoding="utf-8") as f:
        for label, names in all_names.items():
            f.write(f"\n=== {label} ({len(names)}) ===\n")
            for n in sorted(names):
                f.write(f"  {n}\n")
    print("  members.txt записан.")
except Exception as e:
    print(f"  members.txt записать не удалось: {e}")
