# -*- coding: utf-8 -*-
"""Контрольная проба NRPC: что реально отвечает сервер по COM/GetDatabase.

Только чтение. Запуск: py tools/probe_nrpc.py
"""
import sys

import win32com.client


def main():
    ses = win32com.client.Dispatch("Lotus.NotesSession")
    ses.Initialize()
    print("Session user:", ses.UserName)

    targets = [
        # (сервер, база)
        ("CN=Mail8/O=svo", "mail\\desoldatov.nsf"),
        ("traveler.svo.aero", "mail\\desoldatov.nsf"),
        ("traveler.svo.aero", "names.nsf"),
        ("traveler.svo.aero", "log.nsf"),
        ("traveler.svo.aero", "net_gde_net_takoy_bazy.nsf"),  # контроль отрицательного ответа
    ]
    for srv, fn in targets:
        try:
            db = ses.GetDatabase(srv, fn)
            opened = db.IsOpen
            title = db.Title if opened else ""
            print(f"{srv:20s} {fn:35s} IsOpen={opened} Title={title!r}")
        except Exception as e:  # noqa: BLE001
            print(f"{srv:20s} {fn:35s} ERROR: {e!r}"[:120])

    print()
    print("=== 5 последних писем из Inbox (Mail8) ===")
    db = ses.GetDatabase("CN=Mail8/O=svo", "mail\\desoldatov.nsf")
    view = db.GetView("($Inbox)")
    print("Всего входящих:", view.EntryCount)
    for i in range(1, 6):
        doc = view.GetNthDocument(i)
        if doc is None:
            break
        posted = (doc.GetItemValue("PostedDate") or [""])[0]
        frm = (doc.GetItemValue("From") or [""])[0]
        subj = (doc.GetItemValue("Subject") or [""])[0][:60]
        print(f"  {i} | {posted} | {frm} | {subj}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print("FATAL:", repr(e))
        sys.exit(1)
