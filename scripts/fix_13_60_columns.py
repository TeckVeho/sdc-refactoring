#!/usr/bin/env python3
"""Align 1.3 / 6.0 tables to Ex生産情報一覧1 columns."""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

DOCS = Path("/Users/kohei/Projects/SDC/docs")

TUMIKAE_GAIYOU = {
    "Workbook_BeforeClose()": "保存ダイアログを抑止して閉じる",
    "Workbook_Open()": "起動時に画面クリアと積替品抽出を実行",
    "Open_oraconDB()": "ODBC で DB 接続を開く",
    "SQL_Exe()": "SQL 文を Execute で実行",
    "SQL_INSERT_UPDATE()": "キー存在チェック付き INSERT/UPDATE",
    "SQL_Delete()": "WHERE 条件で DELETE",
    "Disp_Sheet()": "SQL 結果をシートに転記",
    "Set_Array()": "SQL 結果を配列に格納",
    "画面クリア3詰替()": "画面データクリア・初期状態復帰",
    "modori()": "積替品シートに戻る",
    "詰替品データ更新()": "変更行を ExSeihinZ へ INSERT/UPDATE",
    "TumikaeHinn()": "DB から積替え対象品を抽出し一覧表示",
    "DataKakou()": "抽出データを加工してシートに表示",
    "積替製品TB表示()": "製品マスタと積替えフラグを積替TBに表示",
    "ExchengeDATE()": "日付フォーマット変換",
    "InsatuHanni()": "印刷範囲を設定",
    "Bookを閉じる()": "ブックを保存せず閉じる",
}

NYUUSHUKKA_GAIYOU = {
    ("ThisWorkbook", "Workbook_BeforeClose()"): "保存確認を抑止して閉じる",
    ("ThisWorkbook", "Workbook_Open()"): "シート保護設定後に画面クリアし入荷状況を表示",
    ("Sheet1", "Worksheet_Change()"): "出荷日変更時に出荷集計を起動",
    ("Sheet2", "Worksheet_Change()"): "入荷日変更時に入荷集計を起動",
    ("BD_Read入出荷", "出荷履歴データ()"): "出荷履歴を DB から取得",
    ("BD_Read入出荷", "Ric23HP_Zaiko()"): "入荷在庫を DB から取得",
    ("ExFunction", "ExchengeDay()"): "日付を mmdd / yyyymmdd に変換",
    ("ExFunction", "ExchengeDATE()"): "日付フォーマット変換",
    ("集計処理", "出荷集計()"): "出荷数量を集計してシート表示",
    ("集計処理", "入荷集計()"): "入荷数量を集計してシート表示",
    ("集計終了", "終了()"): "ブックを閉じる（最後の1冊なら Excel 終了）",
    ("表クリア", "画面クリア()"): "表示データをクリア",
    ("SQL_Execution", "Open_oraconDB()"): "ODBC で DB 接続を開く",
    ("SQL_Execution", "SQL_Exe()"): "SQL 文を Execute で実行",
    ("SQL_Execution", "SQL_INSERT_UPDATE()"): "キー存在チェック付き INSERT/UPDATE",
    ("SQL_Execution", "SQL_Delete()"): "WHERE 条件で DELETE",
    ("SQL_Execution", "Disp_Sheet()"): "SQL 結果をシートに転記",
    ("SQL_Execution", "Set_Array()"): "SQL 結果を配列に格納",
    ("装置稼働状況", "稼働状況()"): "装置稼働状況を DB 集計",
    ("装置稼働状況", "ExHenkan()"): "稼働状況用の値変換",
    ("運転時間", "線源登録表示()"): "指定期間の装置稼働時間を集計表示",
    ("運転時間", "ExHenkan()"): "稼働時間用の値変換",
    ("運転時間", "ChengTime()"): "時間値の変換",
}


def norm_mod(s: str) -> str:
    s = re.sub(r"\*+", "", s)
    s = re.sub(r"（.*?）", "", s)
    s = re.sub(r"\.(bas|cls|frm)$", "", s, flags=re.I)
    return s.strip()


def norm_kind(s: str) -> str:
    t = s.strip().lstrip(".")
    if t.lower() in {"bas", "cls", "frm"}:
        return f".{t.lower()}"
    return s.strip()


def split_row(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def find_table_after(text: str, heading: str) -> tuple[int, int, list[str], list[list[str]]] | None:
    m = re.search(rf"^### {re.escape(heading)}.*$", text, re.M)
    if not m:
        return None
    start = m.end()
    nxt = re.search(r"^### ", text[start:], re.M)
    section_end = start + (nxt.start() if nxt else len(text) - start)
    section = text[start:section_end]
    lines = section.splitlines(keepends=True)
    abs_pos = start
    for i, line in enumerate(lines):
        raw = line.rstrip("\n")
        if raw.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s\-:|]+\|$", lines[i + 1].rstrip("\n")):
            header = split_row(raw)
            rows = []
            j = i + 2
            while j < len(lines) and lines[j].startswith("|"):
                rows.append(split_row(lines[j].rstrip("\n")))
                j += 1
            table_start = abs_pos
            table_end = start + sum(len(x) for x in lines[:j])
            return table_start, table_end, header, rows
        abs_pos += len(line)
    return None


def render(header: list[str], rows: list[list[str]]) -> str:
    widths = [len(h) for h in header]
    for r in rows:
        for i, c in enumerate(r):
            widths[i] = max(widths[i], len(c))
    # keep compact markdown like the rest of the docs
    sep = "| " + " | ".join("---" for _ in header) + " |"
    h = "| " + " | ".join(header) + " |"
    body = ["| " + " | ".join(r) + " |" for r in rows]
    return "\n".join([h, sep, *body]) + "\n"


def counts_from_60(header: list[str], rows: list[list[str]]) -> Counter:
    mi = next(i for i, c in enumerate(header) if "モジュール" in c)
    c = Counter()
    for r in rows:
        if len(r) > mi:
            c[norm_mod(r[mi])] += 1
    return c


def col_index(header: list[str], *names: str) -> int | None:
    for i, c in enumerate(header):
        if c in names:
            return i
    return None


def fix_13(text: str, counts: Counter) -> str:
    found = find_table_after(text, "1.3 VBA モジュール一覧") or find_table_after(text, "1.3 VBAモジュール一覧")
    if not found:
        return text
    a, b, header, rows = found
    drop = {"行数", "コード行数", "OLE ストリーム", "行数（概算）"}
    keep_src = [c for c in header if c not in drop]
    # rebuild to canonical
    new_rows = []
    mi = col_index(header, "モジュール", "モジュール名")
    ki = col_index(header, "種別")
    ri = col_index(header, "主な役割")
    ci = col_index(header, "プロシージャ数")
    for r in rows:
        mod = r[mi]
        kind = norm_kind(r[ki]) if ki is not None else ""
        role = r[ri] if ri is not None else ""
        n = r[ci] if ci is not None else str(counts.get(norm_mod(mod), 0))
        if ci is None:
            n = str(counts.get(norm_mod(mod), 0))
        new_rows.append([r[0], r[1], mod, kind, n, role])
    new_header = ["✓", "No", "モジュール", "種別", "プロシージャ数", "主な役割"]
    return text[:a] + "\n" + render(new_header, new_rows) + text[b:]


def drop_60_cols(header, rows, drop_names, gaiyou_fn=None, rename_proc=True):
    keep_idx = [i for i, c in enumerate(header) if c not in drop_names]
    new_header = [header[i] for i in keep_idx]
    if rename_proc:
        new_header = ["プロシージャ" if c == "プロシージャ名" else c for c in new_header]
    new_rows = []
    for r in rows:
        nr = [r[i] if i < len(r) else "" for i in keep_idx]
        if gaiyou_fn is not None:
            nr.append(gaiyou_fn(header, r))
        new_rows.append(nr)
    if gaiyou_fn is not None:
        new_header.append("概要")
    return new_header, new_rows


def tumikae_gaiyou(header, r):
    pi = col_index(header, "プロシージャ", "プロシージャ名")
    proc = r[pi].strip("`")
    return TUMIKAE_GAIYOU.get(proc, r[-1].replace("📄", "").replace("🗄️", "").replace("📊", "").replace("🖥️", "").strip() or "—")


def nyuu_gaiyou(header, r):
    mi = col_index(header, "モジュール")
    pi = col_index(header, "プロシージャ", "プロシージャ名")
    proc = r[pi].strip("`")
    key = (norm_mod(r[mi]), proc)
    return NYUUSHUKKA_GAIYOU.get(key, "—")


FILES_13 = [
    "Ex1号機照射情報_仕様書.md",
    "ExR3ｵｰﾌﾞﾝ管理温度補正有_仕様書.md",
    "ExRic3詰替作業_仕様書.md",
    "Ex１号機作業指図書_仕様書.md",
    "Ex照射実績表示2_仕様書.md",
    "Ex顧客在庫報告_仕様書.md",
    "Ex入出荷集計_仕様書.md",
    "Ex線量検索_仕様書.md",
    "Exガンマ照射課実績集計_仕様書.md",
    "Ex1号機照射計画_仕様書.md",
    "ExJMM60φ記入用紙_仕様書.md",
]

# 6.0 transforms: filename -> drop cols, optional gaiyou
FILES_60 = {
    "Ex照射実績表示2_仕様書.md": (["行数"], None),
    "Ex顧客在庫報告_仕様書.md": (["行数"], None),
    "Exガンマ照射課実績集計_仕様書.md": (["行数"], None),
    "Ex線量検索_仕様書.md": (["行数", "場所"], None),
    "Ex1号機照射計画_仕様書.md": (["DB I/O", "呼出元"], None),
    "ExRic3詰替作業_仕様書.md": (["場所"], tumikae_gaiyou),
    "Ex入出荷集計_仕様書.md": (["行数", "DB", "呼出元"], nyuu_gaiyou),
}


def main():
    for fn in FILES_13:
        path = DOCS / fn
        text = path.read_text()
        found60 = find_table_after(text, "6.0 全プロシージャ一覧")
        counts = counts_from_60(found60[2], found60[3]) if found60 else Counter()
        new = fix_13(text, counts)
        if new != text:
            path.write_text(new)
            print("1.3 updated", fn, dict(counts))
        else:
            print("1.3 unchanged", fn)

    for fn, (drop, gaiyou) in FILES_60.items():
        path = DOCS / fn
        text = path.read_text()
        found = find_table_after(text, "6.0 全プロシージャ一覧")
        if not found:
            print("6.0 missing", fn)
            continue
        a, b, header, rows = found
        nh, nr = drop_60_cols(header, rows, set(drop), gaiyou)
        new = text[:a] + "\n" + render(nh, nr) + text[b:]
        path.write_text(new)
        print("6.0 updated", fn, "->", nh)


if __name__ == "__main__":
    main()
