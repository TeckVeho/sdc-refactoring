#!/usr/bin/env python3
"""Align 1.1 / 2.0 tables to Ex生産情報一覧1 columns."""
from __future__ import annotations

import re
from pathlib import Path

DOCS = Path("/Users/kohei/Projects/SDC/docs")

H11 = "| ✓ | No | シート名 | 最大行 | 最大列 | 保存時 Visible | VBA による動的切替 |\n| --- | --- | --- | --- | --- | --- | --- |\n"
H20 = "| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |\n| --- | --- | --- | --- | --- | --- |\n"


def replace_table_after(text: str, heading: str, new_table: str) -> str:
    m = re.search(rf"^### {re.escape(heading)}.*$", text, re.M)
    if not m:
        raise SystemExit(f"heading not found: {heading}")
    start = m.end()
    nxt = re.search(r"^### ", text[start:], re.M)
    section_end = start + (nxt.start() if nxt else len(text) - start)
    section = text[start:section_end]
    lines = section.splitlines(keepends=True)
    abs_pos = start
    for i, line in enumerate(lines):
        raw = line.rstrip("\n")
        if raw.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s\-:|]+\|$", lines[i + 1].rstrip("\n")):
            j = i + 2
            while j < len(lines) and lines[j].startswith("|"):
                j += 1
            a = abs_pos
            b = start + sum(len(x) for x in lines[:j])
            return text[:a] + "\n" + new_table + text[b:]
        abs_pos += len(line)
    raise SystemExit(f"table not found after {heading}")


T11 = {
    "Ex1号機照射情報_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 完了予定時間 | 50 | O (15) | visible | — |\n"
        + "|  | 2 | 照射データ | 1 | U (21) | visible | — |\n"
        + "|  | 3 | 線源情報 | 1 | D (4) | visible | — |\n"
        + "|  | 4 | 出荷日情報 | 4491 | J (10) | visible | — |\n",
    ),
    "ExR3ｵｰﾌﾞﾝ管理温度補正有_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 実績 | 133 | 22 (V) | visible | — |\n"
        + "| ✓ | 2 | 管理点線量 | 1748 | 35 (AI) | visible | — |\n"
        + "| ✓ | 3 | Oven | 60 | 28 (AB) | visible | — |\n"
        + "| ✓ | 4 | 線源状態 | 3343 | 30 (AD) | visible | — |\n"
        + "|  | 5 | 履歴記録 | 9 | 4 (D) | visible | — |\n"
        + "| ✓ | 6 | 作業用 | 210 | 14 (N) | visible | — |\n",
    ),
    "ExRic3詰替作業_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 積替品 | 235 | 15 | visible | — |\n"
        + "| ✓ | 2 | 積替TB | 10000 | 12 | visible | — |\n"
        + "|  | 3 | WorkTB | 133 | 9 | visible | — |\n",
    ),
    "Ex１号機作業指図書_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 指図書 | 58 | 27 | visible | — |\n"
        + "| ✓ | 2 | 線量率登録 | 21 | 15 | visible | — |\n"
        + "|  | 3 | 改訂履歴 | 105 | 4 | visible | — |\n"
        + "| ✓ | 4 | 設定値 | 50 | 18 | visible | — |\n"
        + "|  | 5 | ソフト変更履歴 | 20 | 4 | visible | — |\n",
    ),
    "Ex照射実績表示2_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 実績表示 | — | — | visible | — |\n"
        + "|  | 2 | 実績 | — | — | hidden | 保存時非表示。VBA が DB データを書き込む |\n"
        + "|  | 3 | 社員 | — | — | hidden | 保存時非表示。VBA が社員マスタを書き込む |\n",
    ),
    "Ex顧客在庫報告_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 標準 | — | — | visible | — |\n"
        + "| ✓ | 2 | 日機装 | — | — | visible | — |\n"
        + "| ✓ | 3 | オリエンタル | — | — | visible | — |\n"
        + "| ✓ | 4 | ITパック | — | — | visible | — |\n"
        + "|  | 5 | 抽出 | — | — | visible | — |\n",
    ),
    "Ex月末在庫集計_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 記録データ | 51 | S (19) | visible | — |\n"
        + "| ✓ | 2 | γ在庫 | 1051 | Z (26) | visible | — |\n"
        + "| ✓ | 3 | 1号機仕掛 | 136 | Y (25) | visible | — |\n"
        + "| ✓ | 4 | 23号機仕掛 | 408 | Y (25) | visible | — |\n"
        + "| ✓ | 5 | EB在庫 | 997 | Z (26) | visible | — |\n"
        + "| ✓ | 6 | LOCA | 917 | Q (17) | visible | — |\n"
        + "| ✓ | 7 | 集計表 | 29 | I (9) | visible | — |\n",
    ),
    "Ex入出荷集計_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 入荷状況 | 116 | AD | visible | — |\n"
        + "| ✓ | 2 | 出荷実績 | 801 | AZ | visible | — |\n"
        + "| ✓ | 3 | 稼働時間 | 77 | O | visible | — |\n",
    ),
    "Ex線量検索_仕様書.md": (
        "1.1 シート一覧",
        H11 + "| ✓ | 1 | 条件 | — | — | visible | — |\n",
    ),
    "Exガンマ照射課実績集計_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 集計 | 12 | 18 (R) | visible | — |\n"
        + "| ✓ | 2 | 1号機実績 | 520 | 20 (T) | visible | — |\n"
        + "| ✓ | 3 | 2号機実績 | 1 | 1 (A) | visible | — |\n"
        + "| ✓ | 4 | 3号機実績 | 1 | 1 (A) | visible | — |\n",
    ),
    "Ex1号機照射計画_仕様書.md": (
        "1.1 シート一覧",
        H11
        + "| ✓ | 1 | 照射計画 | 177 | 293 (KG) | visible | 行・列の表示/非表示（`画面クリア()`） |\n"
        + "| ✓ | 2 | 未計画品一覧 | 204 | 14 (N) | visible | — |\n"
        + "|  | 3 | 設定値 | 26 | 14 (N) | visible | — |\n"
        + "|  | 4 | 使用方法 | 27 | 4 (D) | visible | — |\n",
    ),
}

# 2.0: heading variants
T20 = {
    "Ex1号機照射情報_仕様書.md": (
        "2.0 可視性一覧",
        H20
        + "| 1 | 完了予定時間 | — | — | — | — |\n"
        + "| 2 | 照射データ | — | — | — | — |\n"
        + "| 3 | 線源情報 | — | — | — | — |\n"
        + "| 4 | 出荷日情報 | — | — | — | — |\n",
    ),
    "ExR3ｵｰﾌﾞﾝ管理温度補正有_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 実績 | — | — | — | — |\n"
        + "| 2 | 管理点線量 | — | — | — | — |\n"
        + "| 3 | Oven | — | — | — | — |\n"
        + "| 4 | 線源状態 | — | — | — | — |\n"
        + "| 5 | 履歴記録 | — | — | — | — |\n"
        + "| 6 | 作業用 | — | — | — | — |\n",
    ),
    "ExRic3詰替作業_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 積替品 | — | — | — | — |\n"
        + "| 2 | 積替TB | — | — | — | — |\n"
        + "| 3 | WorkTB | — | — | — | — |\n",
    ),
    "Ex１号機作業指図書_仕様書.md": (
        "2.0 可視性一覧",
        H20
        + "| 1 | 指図書 | — | — | — | — |\n"
        + "| 2 | 線量率登録 | — | — | — | — |\n"
        + "| 3 | 改訂履歴 | — | — | — | — |\n"
        + "| 4 | 設定値 | — | — | — | — |\n"
        + "| 5 | ソフト変更履歴 | — | — | — | — |\n",
    ),
    "Ex照射実績表示2_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 実績表示 | — | 常時表示 | — | — |\n"
        + "| 2 | 実績 | あり（保存時 Hidden） | VBA が在庫・実績を書き込むとき | 保存時は非表示のまま | `在庫実績データ()` |\n"
        + "| 3 | 社員 | あり（保存時 Hidden） | VBA が社員マスタを書き込むとき | 保存時は非表示のまま | `社員データ()` |\n",
    ),
    "Ex顧客在庫報告_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 標準 | — | — | — | — |\n"
        + "| 2 | 日機装 | — | — | — | — |\n"
        + "| 3 | オリエンタル | — | — | — | — |\n"
        + "| 4 | ITパック | — | — | — | — |\n"
        + "| 5 | 抽出 | — | — | — | — |\n",
    ),
    "Ex月末在庫集計_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 記録データ | — | — | — | — |\n"
        + "| 2 | γ在庫 | — | — | — | — |\n"
        + "| 3 | 1号機仕掛 | — | — | — | — |\n"
        + "| 4 | 23号機仕掛 | — | — | — | — |\n"
        + "| 5 | EB在庫 | — | — | — | — |\n"
        + "| 6 | LOCA | — | — | — | — |\n"
        + "| 7 | 集計表 | — | — | — | — |\n",
    ),
    "Ex入出荷集計_仕様書.md": (
        "2.0 可視性一覧",
        H20
        + "| 1 | 入荷状況 | — | — | — | — |\n"
        + "| 2 | 出荷実績 | — | — | — | — |\n"
        + "| 3 | 稼働時間 | — | — | — | — |\n",
    ),
    "Ex線量検索_仕様書.md": (
        "2.0 シート可視性一覧",
        H20 + "| 1 | 条件 | — | — | — | — |\n",
    ),
    "Exガンマ照射課実績集計_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 集計 | — | — | — | — |\n"
        + "| 2 | 1号機実績 | — | — | — | — |\n"
        + "| 3 | 2号機実績 | — | — | — | — |\n"
        + "| 4 | 3号機実績 | — | — | — | — |\n",
    ),
    "Ex1号機照射計画_仕様書.md": (
        "2.0 可視性一覧",
        H20
        + "| 1 | 照射計画 | 行・列単位 | DebugFlg 設定時に再表示 | `画面クリア()` で行 3-18・列 B-F・R を非表示 | `画面クリア()` |\n"
        + "| 2 | 未計画品一覧 | — | — | — | — |\n"
        + "| 3 | 設定値 | — | — | — | — |\n"
        + "| 4 | 使用方法 | — | — | — | — |\n",
    ),
    "ExDBファイル表示_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 使い方 | — | — | — | — |\n"
        + "| 2 | 項目TB | — | — | — | — |\n"
        + "| 3 | 抽出 | — | — | — | — |\n"
        + "| 4 | 抽出結果 | — | — | — | — |\n"
        + "| 5 | Table登録 | — | — | — | — |\n",
    ),
    "ExGM_EB会社ｺｰﾄﾞ変換TB_仕様書.md": (
        "2.0 シート可視性一覧",
        H20 + "| 1 | 登録 | — | — | — | — |\n",
    ),
    "ExRIC3線量不足報告書住電用_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 報告書 | — | — | — | — |\n"
        + "| 2 | 試験成績書用紙 | — | — | — | — |\n"
        + "| 3 | 改訂履歴 | — | — | — | — |\n",
    ),
    "線量率計算v2025-2_仕様書.md": (
        "2.0 シート可視性一覧",
        H20
        + "| 1 | 線量率RATE | — | — | — | — |\n"
        + "| 2 | 積算INTEG | — | — | — | — |\n"
        + "| 3 | 改訂履歴 | — | — | — | — |\n"
        + "| 4 | Sheet1 | — | — | — | — |\n",
    ),
}


def insert_calendar_20(text: str) -> str:
    old = "## 2. シート詳細\n\n### 2.1 カレンダー\n"
    new = (
        "## 2. シート詳細\n\n"
        "### 2.0 シート可視性一覧\n\n"
        + H20
        + "| 1 | カレンダー | — | — | — | — |\n\n"
        + "> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル\n\n"
        + "### 2.1 カレンダー\n"
    )
    if old not in text:
        raise SystemExit("calendar 2.1 block not found")
    return text.replace(old, new, 1)


def rename_20_heading(text: str) -> str:
    return text.replace("### 2.0 可視性一覧", "### 2.0 シート可視性一覧")


def main():
    for fn, (heading, table) in T11.items():
        path = DOCS / fn
        text = replace_table_after(path.read_text(), heading, table)
        path.write_text(text)
        print("1.1", fn)

    for fn, (heading, table) in T20.items():
        path = DOCS / fn
        text = replace_table_after(path.read_text(), heading, table)
        text = rename_20_heading(text)
        path.write_text(text)
        print("2.0", fn)

    cal = DOCS / "Exカレンダー_仕様書.md"
    cal.write_text(insert_calendar_20(cal.read_text()))
    print("2.0 calendar inserted")


if __name__ == "__main__":
    main()
