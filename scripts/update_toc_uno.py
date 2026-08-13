#!/usr/bin/env python3
"""LibreOffice UNO で TOC / PAGE フィールドを更新して上書き保存する。

LibreOffice 同梱 Python で実行すること:
  /Applications/LibreOffice.app/Contents/Resources/python scripts/update_toc_uno.py 納品/*.docx
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
UNO_URL = "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"


def bootstrap():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    for _ in range(30):
        try:
            return resolver.resolve(UNO_URL)
        except NoConnectException:
            time.sleep(0.5)
    raise SystemExit("LibreOffice UNO に接続できません")


def ensure_soffice():
    ctx = None
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    try:
        ctx = resolver.resolve(UNO_URL)
        return ctx
    except NoConnectException:
        pass
    subprocess.Popen(
        [
            SOFFICE,
            "--headless",
            "--invisible",
            "--norestore",
            "--nolockcheck",
            "--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return bootstrap()


def file_url(path: str) -> str:
    return uno.systemPathToFileUrl(os.path.abspath(path))


def update_one(desktop, path: str) -> None:
    props = (
        PropertyValue(Name="Hidden", Value=True),
        PropertyValue(Name="ReadOnly", Value=False),
        PropertyValue(Name="MacroExecutionMode", Value=4),
    )
    doc = desktop.loadComponentFromURL(file_url(path), "_blank", 0, props)
    try:
        doc.reconnect()
    except Exception:
        pass
    try:
        indexes = doc.getDocumentIndexes()
        for i in range(indexes.getCount()):
            indexes.getByIndex(i).update()
    except Exception as exc:
        print("index update:", exc, file=sys.stderr)
    try:
        doc.refresh()
    except Exception:
        pass
    # PAGE フィールドも再計算
    try:
        doc.reformat()
    except Exception:
        pass
    doc.store()
    doc.close(True)
    print("updated:", os.path.basename(path))


def main() -> None:
    paths = [os.path.abspath(p) for p in sys.argv[1:]]
    if not paths:
        raise SystemExit("docx パスを指定してください")
    ctx = ensure_soffice()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    for p in paths:
        update_one(desktop, p)


if __name__ == "__main__":
    main()
