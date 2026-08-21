#!/usr/bin/env python3
"""残り25件: 3章余分列・4章見出し・レイアウト・コントロール・一点もの。"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path("/Users/kohei/Projects/SDC")
DOCS = ROOT / "docs"


def read(rel: str) -> str:
    return (DOCS / rel).read_text(encoding="utf-8")


def write(rel: str, text: str) -> None:
    (DOCS / rel).write_text(text, encoding="utf-8")


def drop_last_col_in_table(text: str, header_marker: str) -> str:
    """指定ヘッダを持つ表から最終列を落とす。"""
    lines = text.splitlines(keepends=True)
    out: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if header_marker in line and line.strip().startswith("|"):
            # header
            cells = split_row(line)
            out.append(join_row(cells[:-1]) + ("\n" if line.endswith("\n") else ""))
            i += 1
            if i < len(lines) and re.match(r"^\|[\s\-:|]+\|\s*$", lines[i]):
                sep = split_row(lines[i])
                out.append(join_row(sep[:-1]) + "\n")
                i += 1
            while i < len(lines) and lines[i].startswith("|"):
                cells = split_row(lines[i])
                out.append(join_row(cells[:-1]) + ("\n" if lines[i].endswith("\n") else ""))
                i += 1
            continue
        out.append(line)
        i += 1
    return "".join(out)


def split_row(line: str) -> list[str]:
    s = line.rstrip("\n")
    if s.endswith("\r"):
        s = s[:-1]
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def join_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def infer_kind(name: str, content: str) -> str:
    if name and name != "—":
        if "入力" in content:
            return "入力"
        if "数式" in content:
            return "数式"
        return "VBA代入"
    if any(k in content for k in ("ラベル", "タイトル", "ヘッダー", "ヘッダ")):
        return "ラベル"
    return "ラベル"


def convert_3col_layout(text: str) -> str:
    """| ✓ | セル範囲 | 名前付き範囲 | 内容 | → 基準7列。"""
    header = "| ✓ | セル範囲 | 名前付き範囲 | 内容 |"
    new_header = "| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |"
    new_sep = "| --- | --- | --- | --- | --- | --- | --- |"
    lines = text.splitlines(keepends=True)
    out: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i].strip() == header:
            out.append(new_header + "\n")
            i += 1
            if i < len(lines) and lines[i].startswith("| ---"):
                out.append(new_sep + "\n")
                i += 1
            n = 0
            while i < len(lines) and lines[i].startswith("|"):
                cells = split_row(lines[i])
                if len(cells) >= 4:
                    n += 1
                    mark, cell, name, content = cells[0], cells[1], cells[2], cells[3]
                    kind = infer_kind(name, content)
                    out.append(join_row([mark, str(n), cell, name, kind, content, content]) + "\n")
                else:
                    out.append(lines[i])
                i += 1
            continue
        out.append(lines[i])
        i += 1
    return "".join(out)


def drop_caption_parent(text: str) -> str:
    old = "| ✓ | No | コントロール | 種別 | キャプション | 親フレーム | 用途 |"
    new = "| ✓ | No | コントロール | 種別 | 用途 |"
    new_sep = "| --- | --- | --- | --- | --- |"
    lines = text.splitlines(keepends=True)
    out: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i].strip() == old:
            out.append(new + "\n")
            i += 1
            if i < len(lines) and lines[i].startswith("| ---"):
                out.append(new_sep + "\n")
                i += 1
            while i < len(lines) and lines[i].startswith("|"):
                cells = split_row(lines[i])
                if len(cells) >= 7:
                    # ✓ No コントロール 種別 キャプション 親フレーム 用途
                    out.append(join_row([cells[0], cells[1], cells[2], cells[3], cells[6]]) + "\n")
                else:
                    out.append(lines[i])
                i += 1
            continue
        out.append(lines[i])
        i += 1
    return "".join(out)


def fix_ch3() -> None:
    p = "Ex照射実績表示2_仕様書.md"
    write(p, drop_last_col_in_table(read(p), "業務的意味 | VBA参照"))

    p = "Exガンマ照射課実績集計_仕様書.md"
    write(p, drop_last_col_in_table(read(p), "業務的意味 | VBA からの参照"))

    # 入出荷: スコープ列を落とし、名前をコード体に
    p = "Ex入出荷集計_仕様書.md"
    text = read(p)
    old = "| ✓ | No | 名前 | 参照先 | スコープ | 業務的意味 |"
    new = "| ✓ | No | 名前 | 参照先 | 業務的意味 |"
    new_sep = "| --- | --- | --- | --- | --- |"
    lines = text.splitlines(keepends=True)
    out: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i].strip() == old:
            out.append(new + "\n")
            i += 1
            if i < len(lines) and lines[i].startswith("| ---"):
                out.append(new_sep + "\n")
                i += 1
            while i < len(lines) and lines[i].startswith("|"):
                cells = split_row(lines[i])
                if len(cells) >= 6:
                    mark, no, name, ref, _scope, meaning = cells[:6]
                    if not name.startswith("`"):
                        name = f"`{name}`"
                    out.append(join_row([mark, no, name, ref, meaning]) + "\n")
                else:
                    out.append(lines[i])
                i += 1
            continue
        out.append(lines[i])
        i += 1
    write(p, "".join(out))


def fix_getsumatsu_ch4() -> None:
    p = "Ex月末在庫集計_仕様書.md"
    text = read(p)
    text = text.replace("### 4.31 号機仕掛", "### 4.3 1号機仕掛")
    text = text.replace("### 4.423 号機仕掛", "### 4.4 23号機仕掛")
    write(p, text)


def fix_kokyaku_ch4() -> None:
    p = "Ex顧客在庫報告_仕様書.md"
    text = read(p)
    text = text.replace("### 標準\n", "### 4.1 標準\n", 1)
    text = text.replace("### 日機装\n", "### 4.2 日機装\n", 1)
    text = text.replace("### オリエンタル\n", "### 4.3 オリエンタル\n", 1)
    text = text.replace("### 抽出\n", "### 4.4 抽出\n", 1)
    write(p, text)


def fix_shosha_joho_purpose() -> None:
    p = "Ex1号機照射情報_仕様書.md"
    text = read(p)
    text = text.replace(
        "### 2.1 完了予定時間\n\n#### 非表示行・列\n\nなし。\n\nメイン画面。1号機固定照射中の製品一覧と完了予想時刻を表示する。\n",
        "### 2.1 完了予定時間\n\n**目的**: メイン画面。1号機固定照射中の製品一覧と完了予想時刻を表示する。\n\n#### 非表示行・列\n\nなし。\n",
        1,
    )
    text = text.replace(
        "### 2.2 照射データ\n\n#### 非表示行・列\n\nなし。\n\nDB テーブル `SYOUK1` から取得した照射管理データの中間格納シート。VBA の `Yomikomi()` により書き込まれる。\n",
        "### 2.2 照射データ\n\n**目的**: DB テーブル `SYOUK1` から取得した照射管理データの中間格納シート。VBA の `Yomikomi()` により書き込まれる。\n\n#### 非表示行・列\n\nなし。\n",
        1,
    )
    text = text.replace(
        "### 2.3 線源情報\n\n#### 非表示行・列\n\nなし。\n\nDB テーブル `SENGNR1` から取得した線源のタイマー情報の中間格納シート。VBA の `SenGenn()` により書き込まれる。\n",
        "### 2.3 線源情報\n\n**目的**: DB テーブル `SENGNR1` から取得した線源のタイマー情報の中間格納シート。VBA の `SenGenn()` により書き込まれる。\n\n#### 非表示行・列\n\nなし。\n",
        1,
    )
    text = text.replace(
        "### 2.4 出荷日情報\n\n#### 非表示行・列\n\nなし。\n\nDB テーブル `ExKeikakuX` から取得した出荷日マスタ。VBA の `SyukabiRead()` により書き込まれる。\n",
        "### 2.4 出荷日情報\n\n**目的**: DB テーブル `ExKeikakuX` から取得した出荷日マスタ。VBA の `SyukabiRead()` により書き込まれる。\n\n#### 非表示行・列\n\nなし。\n",
        1,
    )
    write(p, text)


def fix_tsumikae_5x() -> None:
    p = "ExRic3詰替作業_仕様書.md"
    text = read(p)
    old = """### 5.2 ユーザーフォーム上のボタン（サマリ）

なし。

### 5.3 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `画面クリア3詰替()` | **Ctrl+E** | 画面データクリア・初期状態復帰 |
"""
    new = """### 5.2 ショートカットキー

| No | マクロ名 | ショートカット | 処理概要 |
| --- | --- | --- | --- |
| 1 | `画面クリア3詰替()` | **Ctrl+E** | 画面データクリア・初期状態復帰 |

### 5.3 ユーザーフォーム上のボタン（サマリ）

なし。
"""
    if old not in text:
        raise SystemExit("詰替 5.2/5.3 ブロックが見つからない")
    write(p, text.replace(old, new, 1))


def fix_senryoritsu() -> None:
    p = "線量率計算v2025-2_仕様書.md"
    text = read(p)
    old_toc = """## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
9. [データフロー](#9-データフロー)
"""
    new_toc = """## 目次

1. [ファイル構成](#1-ファイル構成)
2. [シート詳細](#2-シート詳細)
3. [名前付き範囲一覧](#3-名前付き範囲一覧)
4. [数式一覧](#4-数式一覧)
5. [ボタン・マクロ対応](#5-ボタンマクロ対応)
6. [VBA モジュール仕様](#6-vba-モジュール仕様)
7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)
8. [DB 接続・外部連携](#8-db-接続外部連携)
9. [データフロー](#9-データフロー)
10. [セキュリティ注意事項](#10-セキュリティ注意事項)
11. [スコープ外（本仕様書に含まないもの）](#スコープ外本仕様書に含まないもの)
12. [付録: 機器一覧](#付録-機器一覧)
"""
    if old_toc not in text:
        raise SystemExit("線量率 TOC が見つからない")
    text = text.replace(old_toc, new_toc, 1)

    # 4章末尾の孤立 8.4 を外し、章を 5→10→スコープ外→付録 の順に並べる
    m4 = re.search(r"(## 4\. 数式一覧\n)", text)
    if not m4:
        raise SystemExit("線量率 4章が見つからない")
    start4 = m4.start()
    after4 = text[start4:]

    def take(src: str, heading: str, next_headings: list[str]) -> tuple[str, str]:
        m = re.search(rf"(^{re.escape(heading)}\n)", src, re.M)
        if not m:
            raise SystemExit(f"見出しなし: {heading}")
        start = m.start()
        end = len(src)
        for nh in next_headings:
            m2 = re.search(rf"(^{re.escape(nh)}\n)", src[start + 1 :], re.M)
            if m2:
                cand = start + 1 + m2.start()
                if cand < end:
                    end = cand
        return src[start:end], src[:start] + src[end:]

    nexts = [
        "### 8.4 外部ファイル連携",
        "## 9. データフロー",
        "## スコープ外（本仕様書に含まないもの）",
        "## 付録: 機器一覧",
        "## 5. ボタン・マクロ対応",
        "## 6. VBA モジュール仕様",
        "## 7. ユーザーフォーム仕様",
        "## 8. DB 接続・外部連携",
        "## 10. セキュリティ注意事項",
    ]
    ch4, rest = take(after4, "## 4. 数式一覧", nexts)
    orphan84, rest = take(rest, "### 8.4 外部ファイル連携", nexts)
    ch9, rest = take(rest, "## 9. データフロー", nexts)
    scope, rest = take(rest, "## スコープ外（本仕様書に含まないもの）", nexts)
    appendix, rest = take(rest, "## 付録: 機器一覧", nexts)
    ch5, rest = take(rest, "## 5. ボタン・マクロ対応", nexts)
    ch6, rest = take(rest, "## 6. VBA モジュール仕様", nexts)
    ch7, rest = take(rest, "## 7. ユーザーフォーム仕様", nexts)
    ch8, rest = take(rest, "## 8. DB 接続・外部連携", nexts)
    ch10, rest = take(rest, "## 10. セキュリティ注意事項", nexts)
    if rest.strip():
        raise SystemExit(f"線量率 並べ替え残り: {rest[:200]!r}")

    # 8.4 を 8章末へ
    ch8 = ch8.rstrip() + "\n\n" + orphan84.lstrip()
    if not ch8.endswith("\n"):
        ch8 += "\n"

    prefix = text[:start4]
    new_body = "".join([ch4, ch5, ch6, ch7, ch8, ch9, ch10, scope, appendix])
    write(p, prefix + new_body)


def fix_sumiden() -> None:
    p = "ExRIC3線量不足報告書住電用_仕様書.md"
    text = read(p)
    old = """**非表示行・列:**

| 種別 | 対象 | 内容 |
| --- | --- | --- |
| 非表示行 | 34 | VLOOKUP 用カラム番号（1〜50） |
| 非表示行 | 35 | DB カラム名ヘッダ（SYKNO, SESDATE, SOKUTCD 等） |
| 非表示行 | 36 | VLOOKUP 数式行（`Data` 範囲から列番号で抽出） |
| 非表示列 | — | なし |

**データ入力規則:**

| セル範囲 | マージ範囲 | 名前付き範囲 | 入力規則 | ソース |
| --- | --- | --- | --- | --- |
| `B7:F7` | `B7:F7` | `SenkNo` | リスト | VBA 動的生成（線量不足の照射管理番号） |
| `O7:R7` | `O7:R7` | `SaiSoku` | リスト | `SyainTB`（有資格社員名） |
| `G16:J16` | `G16:J16` | `SEnKind` | リスト | `SenSyu`（線量計種類） |
| `AG7:AI7` | `AG7:AI7` | `Sosi` | リスト | 固定値 "無し,有" |

#### 非表示行・列

なし。
"""
    new = """#### 非表示行・列

| 種類 | 対象 | 備考 |
| --- | --- | --- |
| 非表示行 | 34 | VLOOKUP 用カラム番号（1〜50） |
| 非表示行 | 35 | DB カラム名ヘッダ（SYKNO, SESDATE, SOKUTCD 等） |
| 非表示行 | 36 | VLOOKUP 数式行（`Data` 範囲から列番号で抽出） |
"""
    if old not in text:
        raise SystemExit("住電 2.1 レイアウトブロックが見つからない")
    text = text.replace(old, new, 1)
    text = text.replace("**非表示行・列:** なし\n\n#### 非表示行・列\n\nなし。\n", "#### 非表示行・列\n\nなし。\n")
    write(p, text)


def fix_calendar_layout() -> None:
    p = "Exカレンダー_仕様書.md"
    text = read(p)
    insert = """
#### 主要セル

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `$G$1` | `HyoujiBi` | 入力 | 表示開始月（YYYY/M/1 に正規化） | カレンダー表示の起点月 |
|  | 2 | `$AQ$2` | `Honnjitu` | VBA代入 | 本日日付 | 当日判定・休日更新の基準日 |
|  | 3 | `$A$1` | `Debug` | 手動設定 | 空=通常 / 非空=右クリック処理スキップ | デバッグモード切替 |
| ✓ | 4 | `$AQ$7:$AR$372` | `YasumiTB` | VBA代入 | `休日読込()` が DB から書込み | 休日テーブル（VLOOKUP 範囲） |
| ✓ | 5 | E7:F37 ほか12か月分 | `YasumiDay` | VBA代入 / 入力 | 休日フラグと更新前値 | 変更検出と一括クリア |
| ✓ | 6 | `$D$6` | — | 数式 | `=HyoujiBi` | 1か月目の月初日 |

"""
    anchor = "#### DB読み込み領域\n"
    if insert.strip() in text:
        return
    if anchor not in text:
        raise SystemExit("カレンダー DB読み込み領域 が見つからない")
    write(p, text.replace(anchor, insert + anchor, 1))


def fix_db_hidden_and_layout() -> None:
    p = "ExDBファイル表示_仕様書.md"
    text = read(p)
    old_20b = """### 2.0 b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
| --- | --- | --- |
| 項目TB | 2 | C〜D |
| 抽出 | 2〜3 | L〜M |
| Table登録 | 2 | I〜J, O〜P |
"""
    new_20b = """### 2.0 b 非表示行・列一覧

| 種類 | 対象 | 備考 |
| --- | --- | --- |
| 非表示行 | 項目TB 行2 | ヘッダー補助行 |
| 非表示列 | 項目TB C〜D | 内部列 |
| 非表示行 | 抽出 行2〜3 | 設定補助行 |
| 非表示列 | 抽出 L〜M | 内部列 |
| 非表示行 | Table登録 行2 | ヘッダー補助行 |
| 非表示列 | Table登録 I〜J, O〜P | 変更前値・内部列 |
"""
    if old_20b not in text:
        raise SystemExit("DB 2.0b が見つからない")
    text = text.replace(old_20b, new_20b, 1)

    def set_hidden(text: str, heading: str, block: str) -> str:
        pat = rf"(### {re.escape(heading)}\n.*?#### 非表示行・列\n\n)なし。\n"
        m = re.search(pat, text, re.S)
        if not m:
            raise SystemExit(f"DB {heading} 非表示 が見つからない")
        return text[: m.start()] + m.group(1) + block + text[m.end() :]

    text = set_hidden(
        text,
        "2.2 項目TB",
        "| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示行 | 2 | ヘッダー補助行 |\n| 非表示列 | C〜D | 内部列 |\n",
    )
    text = set_hidden(
        text,
        "2.3 抽出",
        "| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示行 | 2〜3 | 設定補助行 |\n| 非表示列 | L〜M | 内部列 |\n",
    )
    text = set_hidden(
        text,
        "2.5 Table登録",
        "| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示行 | 2 | ヘッダー補助行 |\n| 非表示列 | I〜J, O〜P | 変更前値・内部列 |\n",
    )

    text = convert_3col_layout(text)

    # 抽出結果ヘッダー部を基準7列へ
    old_res = """#### レイアウト構造（ヘッダー部）

| 行 | 列A | 列B以降 |
| --- | --- | --- |
| 1 | 「テーブル名称：」 | テーブル名 |
| 2 | 「テーブル内容：」 | テーブルの説明 |
| 3 | 「テーブル名　：」 | DBテーブル名 |
| 4 | 「データベース：」 | DB名（UID） |
| 5 | 「抽出条件　　：」 | SQL WHERE 条件文字列 |
| 6 | 「抽出件数　　：」 | 抽出されたレコード件数 |
| 7 | 「DB項目名→」 | DB フィールド名（B列〜） |
| 8 | 「備考→」 | 項目説明（B列〜） |
| 9 | 「文字長→」 | フィールド長（B列〜） |
| 11 | 「No」 | 項目名（B列〜） |
| 12〜30000 | データ行 | DB から取得したデータ（B列〜） |
"""
    new_res = """#### レイアウト構造（ヘッダー部）

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
|  | 1 | `$A$1` | — | ラベル | テーブル名称： | ヘッダーラベル |
|  | 2 | `$B$1` | — | VBA代入 | テーブル名 | 抽出対象テーブル名 |
|  | 3 | `$A$2` | — | ラベル | テーブル内容： | ヘッダーラベル |
|  | 4 | `$B$2` | — | VBA代入 | テーブルの説明 | テーブル説明 |
|  | 5 | `$A$3` | — | ラベル | テーブル名　： | ヘッダーラベル |
|  | 6 | `$B$3` | — | VBA代入 | DBテーブル名 | DB 上のテーブル名 |
|  | 7 | `$A$4` | — | ラベル | データベース： | ヘッダーラベル |
|  | 8 | `$B$4` | — | VBA代入 | DB名（UID） | 接続先 DB |
|  | 9 | `$A$5` | — | ラベル | 抽出条件　　： | ヘッダーラベル |
|  | 10 | `$B$5` | — | VBA代入 | SQL WHERE 条件文字列 | 抽出条件 |
|  | 11 | `$A$6` | — | ラベル | 抽出件数　　： | ヘッダーラベル |
|  | 12 | `$B$6` | — | VBA代入 | 抽出されたレコード件数 | 件数 |
| ✓ | 13 | `$B$12:$…$30000` | — | VBA代入 | DB から取得したデータ | 抽出結果本体（最大30,000行） |
"""
    if old_res not in text:
        raise SystemExit("DB 抽出結果レイアウトが見つからない")
    text = text.replace(old_res, new_res, 1)
    write(p, text)


def fix_gm_layout() -> None:
    p = "ExGM_EB会社ｺｰﾄﾞ変換TB_仕様書.md"
    text = read(p)
    text = text.replace(
        "#### 非表示行・列\n\nなし。\n",
        "#### 非表示行・列\n\n| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示行 | 4, 8 | 入力行と一覧の間の空行 |\n",
        1,
    )
    write(p, convert_3col_layout(text))


def fix_jmm60() -> None:
    p = "ExJMM60φ記入用紙_仕様書.md"
    text = read(p)
    old_20 = """### 2.0 シート可視性一覧

| No | シート名 | codeName | sheetId | 保存時 Visible |
| --- | --- | --- | --- | --- |
| 1 | 使用方法 | Sheet1 | 7 | visible |
| 2 | 測定値 | Sheet2 | 3 | visible |
| 3 | 報告書 | Sheet3 | 1 | visible |
"""
    new_20 = """### 2.0 シート可視性一覧

| No | シート | VBA による非表示化 | 表示するタイミング | 非表示にするタイミング | 制御プロシージャ |
| --- | --- | --- | --- | --- | --- |
| 1 | 使用方法 | — | — | — | — |
| 2 | 測定値 | — | — | — | — |
| 3 | 報告書 | — | — | — | — |

> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル
"""
    if old_20 not in text:
        raise SystemExit("JMM60 2.0 が見つからない")
    text = text.replace(old_20, new_20, 1)

    old_20b = """### 2.0 b 非表示行・列一覧

| シート | 非表示行 | 非表示列 |
| --- | --- | --- |
| 測定値 | なし | Q〜X |
| 報告書 | 1 | なし |
"""
    new_20b = """### 2.0 b 非表示行・列一覧

| 種類 | 対象 | 備考 |
| --- | --- | --- |
| 非表示列 | 測定値 Q〜X | DB 参照データ領域 |
| 非表示行 | 報告書 行1 | ヘッダー補助行 |
"""
    if old_20b not in text:
        raise SystemExit("JMM60 2.0b が見つからない")
    text = text.replace(old_20b, new_20b, 1)

    text = text.replace(
        "### 2.2 測定値\n\n**目的**: 線量測定データの入力・計算を行うメインシート。最大90バッチ（6本×30回照射×上中下3段）の測定値を記録する。\n\n#### 非表示行・列\n\nなし。\n",
        "### 2.2 測定値\n\n**目的**: 線量測定データの入力・計算を行うメインシート。最大90バッチ（6本×30回照射×上中下3段）の測定値を記録する。\n\n#### 非表示行・列\n\n| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示列 | Q〜X | DB から取得した参照データ領域 |\n",
        1,
    )
    text = text.replace(
        "### 2.3 報告書\n\n**目的**: 顧客向け線量測定結果報告書。「測定値」シートの計算結果を数式で参照し印刷する。\n\n#### 非表示行・列\n\nなし。\n",
        "### 2.3 報告書\n\n**目的**: 顧客向け線量測定結果報告書。「測定値」シートの計算結果を数式で参照し印刷する。\n\n#### 非表示行・列\n\n| 種類 | 対象 | 備考 |\n| --- | --- | --- |\n| 非表示行 | 1 | ヘッダー補助行 |\n",
        1,
    )

    insert = """
#### 主要セル

| ✓ | No | セル | 名前付き範囲 | 種別 | 実態（値/数式/VBA代入） | 業務的意味 |
| --- | --- | --- | --- | --- | --- | --- |
| ✓ | 1 | `$B$3` | `Uno` | VBA代入 | StartForm が在庫DBから代入 | 受付番号 |
| ✓ | 2 | `$D$3` | `Honnsuu` | VBA代入 | 在庫DBの入荷数 | 受付本数（印刷範囲・保存チェック） |
| ✓ | 3 | `$E$3` | `Dose` | VBA代入 | 在庫DBの指定線量 | 指定線量 |
| ✓ | 4 | `$G$3` | `SenKind` | 入力 | 線量計種類 | 計算前チェック対象 |
| ✓ | 5 | `$H$3` | `keicord` | 入力 | 計算式コード | 現在選択中の計算式 |
| ✓ | 6 | `$I$3` | `Kagenn` | VBA代入 | 在庫DBの下限線量 | 下限線量 |
| ✓ | 7 | `$J$3` | `Jyogen` | VBA代入 | 在庫DBの上限線量 | 上限線量 |
| ✓ | 8 | `$G$5:$K$544` | `DataTB` | 入力 / 数式 | 厚さ・ABS入力、測定線量は計算 | 測定データ領域 |
| ✓ | 9 | `$O$1` | `KirokuFolder` | 設定値 | `D:\\JMM線量\\` | 記録ファイル保存先 |
|  | 10 | `$A$1` | `Debug` | 手動設定 | 空=通常 / 値あり=デバッグ | デバッグモード |

"""
    anchor = "**繰返し構造（行5以降）**:\n"
    if anchor not in text:
        raise SystemExit("JMM60 繰返し構造 が見つからない")
    text = text.replace(anchor, insert + anchor, 1)
    write(p, text)


def main() -> None:
    fix_ch3()
    fix_getsumatsu_ch4()
    fix_kokyaku_ch4()
    fix_shosha_joho_purpose()
    fix_tsumikae_5x()
    fix_senryoritsu()
    fix_sumiden()
    fix_calendar_layout()
    fix_db_hidden_and_layout()
    fix_gm_layout()
    fix_jmm60()
    write("ExMenu_仕様書.md", drop_caption_parent(read("ExMenu_仕様書.md")))
    write("ExJMM90φ記入用紙_仕様書.md", drop_caption_parent(read("ExJMM90φ記入用紙_仕様書.md")))
    print("ok")


if __name__ == "__main__":
    main()
