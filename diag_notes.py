"""
Диагностика COM-интерфейса Lotus Notes.

Отвечает на один вопрос: какая библиотека Domino Objects зарегистрирована
в системе и почему в ней нет функций работы с метками прочтения.

Запуск:  python diag_notes.py
Ничего не меняет — только читает и печатает. Вывод можно копировать целиком.
"""

import sys
import os

SEP = "=" * 72


def head(title):
    print(f"\n{SEP}\n  {title}\n{SEP}")


def safe(fn, default="— не удалось получить"):
    try:
        val = fn()
        return val if val not in (None, "") else default
    except Exception as e:
        return f"— ошибка: {e}"


# ---------------------------------------------------------------- 1. Питон
head("1. ОКРУЖЕНИЕ")
print(f"  Python           : {sys.version.split()[0]}")
print(f"  Разрядность      : {'64-bit' if sys.maxsize > 2**32 else '32-bit'}")
try:
    import win32api  # noqa
    import win32com
    print(f"  pywin32          : установлен ({os.path.dirname(win32com.__file__)})")
except ImportError:
    print("  pywin32          : НЕ УСТАНОВЛЕН — дальше смысла нет")
    sys.exit(1)

if sys.maxsize > 2**32:
    print("\n  [!] Python 64-битный. Lotus Notes — 32-битное приложение, и его COM-сервер")
    print("      обычно тоже 32-битный. Если объекты вообще создаются, значит связь идёт")
    print("      через суррогатный процесс, и часть интерфейса может быть недоступна.")
    print("      Стоит проверить тот же скрипт на 32-битном Python.")


# ------------------------------------------------------------- 2. Реестр
head("2. ЧТО ЗАРЕГИСТРИРОВАНО ПОД ИМЕНЕМ Lotus.NotesSession")

import winreg


def reg_get(root, path, name=""):
    with winreg.OpenKey(root, path) as k:
        return winreg.QueryValueEx(k, name)[0]


clsid = None
for progid in ("Lotus.NotesSession", "Notes.NotesSession"):
    try:
        clsid = reg_get(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID")
        print(f"  ProgID           : {progid}")
        print(f"  CLSID            : {clsid}")
        break
    except OSError:
        continue

if not clsid:
    print("  [!] ProgID Lotus.NotesSession в реестре не найден.")
    print("      Клиент Notes не установлен либо COM-сервер не зарегистрирован.")
else:
    server_path = None
    for kind in ("InprocServer32", "LocalServer32"):
        try:
            server_path = reg_get(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\{kind}")
            print(f"  Тип сервера      : {kind}")
            print(f"  Файл             : {server_path}")
            break
        except OSError:
            continue

    if server_path:
        exe = server_path.strip('"').split('" ')[0].strip('"')
        print(f"  Файл существует  : {'да' if os.path.exists(exe) else 'НЕТ — ссылка битая'}")
        if os.path.exists(exe):
            try:
                import win32api
                info = win32api.GetFileVersionInfo(exe, "\\")
                ms, ls = info["FileVersionMS"], info["FileVersionLS"]
                ver = f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
                print(f"  Версия файла     : {ver}")
                print(f"  Размер           : {os.path.getsize(exe):,} байт")
            except Exception as e:
                print(f"  Версия файла     : не прочиталась ({e})")
            print(f"  Каталог Notes    : {os.path.dirname(exe)}")


# --------------------------------------------------- 3. Библиотека типов
head("3. БИБЛИОТЕКИ ТИПОВ DOMINO OBJECTS")
found_tlb = False
try:
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "TypeLib") as tl:
        for i in range(winreg.QueryInfoKey(tl)[0]):
            guid = winreg.EnumKey(tl, i)
            try:
                with winreg.OpenKey(tl, guid) as gk:
                    for j in range(winreg.QueryInfoKey(gk)[0]):
                        ver = winreg.EnumKey(gk, j)
                        try:
                            name = reg_get(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{guid}\{ver}")
                        except OSError:
                            continue
                        if name and ("domino" in name.lower() or "lotus" in name.lower()
                                     or "notes" in name.lower()):
                            found_tlb = True
                            print(f"  {name}")
                            print(f"     версия TLB    : {ver}")
                            print(f"     GUID          : {guid}")
                            for sub in ("win32", "win64"):
                                try:
                                    p = reg_get(winreg.HKEY_CLASSES_ROOT,
                                                rf"TypeLib\{guid}\{ver}\0\{sub}")
                                    print(f"     файл ({sub})  : {p}")
                                    print(f"     существует    : {'да' if os.path.exists(p) else 'НЕТ'}")
                                except OSError:
                                    pass
            except OSError:
                continue
except OSError as e:
    print(f"  Не удалось прочитать реестр: {e}")

if not found_tlb:
    print("  [!] Библиотека типов Domino Objects не зарегистрирована.")
    print("      Именно поэтому Python видит урезанный набор методов.")


# ------------------------------------------------ 4. Кэш pywin32 (makepy)
head("4. КЭШ ОБЁРТОК PYWIN32")
try:
    from win32com.client import gencache
    cache_dir = gencache.GetGeneratePath()
    print(f"  Каталог кэша     : {cache_dir}")
    if os.path.isdir(cache_dir):
        files = [f for f in os.listdir(cache_dir) if not f.startswith("__")]
        print(f"  Файлов в кэше    : {len(files)}")
        for f in files[:10]:
            print(f"     {f}")
        if files:
            print("\n  [i] Если кэш собран для старой версии Notes, он может перекрывать")
            print("      актуальную библиотеку. Каталог можно безопасно удалить —")
            print("      pywin32 соберёт его заново.")
    else:
        print("  Кэш пуст — используется позднее связывание.")
except Exception as e:
    print(f"  Не проверить: {e}")


# ------------------------------------------------------- 5. Живая сессия
head("5. ЧТО ГОВОРИТ САМА СЕССИЯ NOTES")
session = None
try:
    import win32com.client
    session = win32com.client.Dispatch("Lotus.NotesSession")
    session.Initialize()
    print(f"  Версия Notes     : {safe(lambda: session.NotesVersion)}")
    print(f"  Сборка           : {safe(lambda: session.NotesBuildVersion)}")
    print(f"  Пользователь     : {safe(lambda: session.UserName)}")
    print(f"  Платформа        : {safe(lambda: session.Platform)}")
    print(f"  Почтовый сервер  : {safe(lambda: session.GetEnvironmentString('MailServer', True))}")
    print(f"  Файл почты       : {safe(lambda: session.GetEnvironmentString('MailFile', True))}")
except Exception as e:
    print(f"  [!] Сессию создать не удалось: {e}")


# ------------------------------------------- 6. Наличие нужных методов
head("6. НАЛИЧИЕ МЕТОДОВ РАБОТЫ С МЕТКАМИ ПРОЧТЕНИЯ")
print("  Все они появились в Domino Objects 8.0.\n")

if session is None:
    print("  Пропущено — нет сессии.")
else:
    def has(obj, name):
        if obj is None:
            return "объекта нет"
        try:
            getattr(obj, name)
            return "ЕСТЬ"
        except Exception:
            return "нет"

    db = view = doc = entry = None
    try:
        srv = session.GetEnvironmentString("MailServer", True)
        mf = session.GetEnvironmentString("MailFile", True)
        db = session.GetDatabase(srv, mf)
        if db and not db.IsOpen:
            db.Open()
        print(f"  База открыта     : {db.Server or '(локально)'} / {db.FilePath}")
        print(f"  Документов       : {safe(lambda: db.AllDocuments.Count)}")
    except Exception as e:
        print(f"  [!] База не открылась: {e}")

    if db is not None:
        try:
            view = db.GetView("($Inbox)")
            doc = view.GetFirstDocument() if view else None
            nav = view.CreateViewNav() if view else None
            entry = nav.GetFirst() if nav else None
        except Exception:
            pass

        checks = [
            ("NotesDatabase.GetAllUnreadDocuments", db, "GetAllUnreadDocuments"),
            ("NotesDatabase.GetAllReadDocuments", db, "GetAllReadDocuments"),
            ("NotesView.GetAllUnreadEntries", view, "GetAllUnreadEntries"),
            ("NotesView.CreateViewNavFromAllUnread", view, "CreateViewNavFromAllUnread"),
            ("NotesView.MarkAllRead", view, "MarkAllRead"),
            ("NotesDocument.GetRead", doc, "GetRead"),
            ("NotesDocument.MarkRead", doc, "MarkRead"),
            ("NotesDocument.MarkUnread", doc, "MarkUnread"),
            ("NotesViewEntry.GetRead", entry, "GetRead"),
            ("NotesViewEntry.IsUnread", entry, "IsUnread"),
        ]
        width = max(len(c[0]) for c in checks)
        present = 0
        for label, obj, attr in checks:
            res = has(obj, attr)
            if res == "ЕСТЬ":
                present += 1
            print(f"    {label:<{width}}  {res}")

        print(f"\n  Доступно {present} из {len(checks)}.")
        if present == 0:
            print("  ВЫВОД: зарегистрирован COM-сервер Domino Objects старше версии 8.0.")
            print("         Метки прочтения из него получить нельзя ничем.")
        elif present < len(checks):
            print("  ВЫВОД: интерфейс неполный — вероятно, смешаны разные версии Notes.")
        else:
            print("  ВЫВОД: всё на месте, метки прочтения доступны.")

print(f"\n{SEP}\n  Готово. Скопируйте вывод целиком.\n{SEP}")