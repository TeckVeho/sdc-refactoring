#!/usr/bin/env python3
"""Word仕様書レビュー対応: Markdown 横断の機械置換。

適用内容:
  - 9章構成ファイルへ 7章「ユーザーフォーム仕様」を挿入して 10章に揃える
  - 見出しの件数表記を削除
  - 1.3 / 8.3 見出し名を統一
  - 5.x を連番に振り直し
  - ショートカット「割り当てなし」行を削除
  - 表: #→No、✓を先頭列、一覧表へ No 列を挿入
  - 6.0 にスコープ列が無ければ追加（ヒューリスティック）
  - 3.1 に用途列が無ければ追加
  - 末尾の *以上* を削除
  - スコープ外セクションが無ければ追加
  - 冒頭メタキー名の正規化
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"

SKIP = {"_仕様書共通ルール.md"}

NINE_CHAPTER = {
    "ExDBファイル表示_仕様書.md",
    "Exカレンダー_仕様書.md",
    "ExGM_EB会社ｺｰﾄﾞ変換TB_仕様書.md",
}

H13_CANON = "1.3 VBA モジュール一覧"
H83_CANON = "8.3 SQL 一覧"

SQL_HEAD_RE = re.compile(
    r"^(#{3})\s*(\d+)\.3\s+(主要\s*)?SQL(\s*文)?(\s*一覧)?(\s*[（(](抜粋|全件|全\s*\d+\s*件)[）)])?\s*$"
)
COUNT_HEAD_RE = re.compile(r"（全\s*\d+\s*件）")
COUNT_BODY_RE = re.compile(r"^全\s+\*\*\d+\*\*\s+件。\s*$")
IJO_RE = re.compile(r"^\*?以\s*上\*?\s*$")
ASSIGN_NONE_RE = re.compile(r"割り当てなし")
HEADING_RE = re.compile(r"^(#{2,4})\s+(.*)$")
META_KEY_MAP = {
    "VBA プロジェクトサイズ": "VBA プロジェクト",
    "外部連携ファイル": "外部連携",
}

SCOPE_OUT = """\
## スコープ外（本仕様書に含まないもの）

- セル書式（色・罫線・フォント）
- 条件付き書式、グラフ・画像、印刷設定

必要な場合は Excel 画面のスクリーンショットで補完してください。
"""

USERFORM_CH7 = """\
## 7. ユーザーフォーム仕様

該当なし（ユーザーフォームなし）。

"""

LIST_HEAD_HINTS = (
    "シート一覧",
    "ユーザーフォーム一覧",
    "モジュール一覧",
    "名前付き範囲",
    "全プロシージャ一覧",
    "テーブル一覧",
    "シート上のボタン",
    "ユーザーフォーム上のボタン",
    "CommandBar",
    "コントロール一覧",
    "ショートカット",
    "データフロー",
    "セキュリティ注意事項",
    "シート可視性",
)


def split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def join_table_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def is_sep_row(cells: list[str]) -> bool:
    return all(re.fullmatch(r":?-{3,}:?", c.replace(" ", "") or "---") or set(c) <= set("-: ") for c in cells)


def parse_blocks(text: str) -> list[tuple[str, list[str]]]:
    lines = text.split("\n")
    blocks: list[tuple[str, list[str]]] = []
    i = 0
    while i < len(lines):
        if lines[i].strip().startswith("|"):
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i])
                i += 1
            blocks.append(("table", tbl))
        else:
            blocks.append(("line", [lines[i]]))
            i += 1
    return blocks


def current_heading_context(prev_headings: list[str]) -> str:
    return " / ".join(prev_headings[-3:])


def should_have_no(ctx: str, header: list[str]) -> bool:
    if any(h in header for h in ("No", "#", "番号", "項番")):
        return False
    if "項目" in header and "内容" in header and "✓" not in header:
        return False
    if header == ["種別", "表記", "例"] or header == ["用語", "意味"]:
        return False
    if header[:3] == ["章", "対象", "✓ の判定基準"] or "判定基準" in "".join(header):
        return False
    if "アイコン" in header:
        return False
    if any(x in header for x in ("数式", "セル")) and "数式" in ctx:
        return False
    if "制約" in header:  # 3.1 入力規則
        return False
    if any(k in ctx for k in LIST_HEAD_HINTS):
        return True
    if "✓" in header:
        return True
    return False


def infer_scope(module: str, proc: str, kind: str) -> str:
    blob = f"{module} {proc} {kind}"
    if "宣言" in blob or proc in ("—", "-", ""):
        return "—"
    if kind.strip() == "Event" or "Click" in proc or proc.startswith("Worksheet_") or proc.startswith("Workbook_"):
        return "Private"
    if ".frm" in module or "Form" in module:
        return "Private"
    if "ThisWorkbook" in module or re.search(r"Sheet\d", module):
        if kind.strip() == "Event":
            return "Private"
    return "Public"


def transform_table(tbl_lines: list[str], ctx: str) -> list[str]:
    rows = [split_table_row(ln) for ln in tbl_lines]
    if len(rows) < 2:
        return tbl_lines
    header, sep, data = rows[0], rows[1], rows[2:]
    ncols = len(header)
    data = [r + [""] * (ncols - len(r)) if len(r) < ncols else r[:ncols] for r in data]

    # ショートカットの割り当てなし行を削除
    if "ショートカット" in ctx or "ショートカット" in "".join(header):
        kept = []
        for r in data:
            if any(ASSIGN_NONE_RE.search(c) for c in r):
                continue
            kept.append(r)
        data = kept

    # # → No（既存番号は維持）
    header = ["No" if h in ("#", "番号", "項番", "No.") else h for h in header]

    # ✓ を先頭へ
    if "✓" in header and header.index("✓") != 0:
        i = header.index("✓")
        header = [header[i]] + header[:i] + header[i + 1 :]
        data = [[r[i]] + r[:i] + r[i + 1 :] for r in data]

    # 6.0 スコープ列
    if "全プロシージャ" in ctx and "スコープ" not in header:
        insert_at = None
        for name in ("プロシージャ名", "プロシージャ"):
            if name in header:
                insert_at = header.index(name) + 1
                break
        if insert_at is None:
            insert_at = 2 if header and header[0] == "✓" else 1
        header = header[:insert_at] + ["スコープ"] + header[insert_at:]
        mod_i = next((i for i, h in enumerate(header) if "モジュール" in h), None)
        proc_i = next((i for i, h in enumerate(header) if "プロシージャ" in h), None)
        kind_i = next((i for i, h in enumerate(header) if h in ("種別", "種類")), None)
        new_data = []
        for r in data:
            r = r[:insert_at] + [""] + r[insert_at:]
            mod = r[mod_i] if mod_i is not None and mod_i < len(r) else ""
            proc = r[proc_i] if proc_i is not None and proc_i < len(r) else ""
            kind = r[kind_i] if kind_i is not None and kind_i < len(r) else ""
            r[insert_at] = infer_scope(mod, proc, kind)
            new_data.append(r)
        data = new_data

    # 3.1 用途列
    if "入力規則" in ctx and "用途" not in header:
        header = header + ["用途"]
        data = [r + [""] for r in data]

    # No 列の挿入
    if should_have_no(ctx, header) and "No" not in header:
        insert_at = 1 if header and header[0] == "✓" else 0
        header = header[:insert_at] + ["No"] + header[insert_at:]
        data = [r[:insert_at] + [str(n)] + r[insert_at:] for n, r in enumerate(data, 1)]
    elif "No" in header:
        ni = header.index("No")
        for n, r in enumerate(data, 1):
            if not r[ni] or r[ni] in ("", "—", "-"):
                r[ni] = str(n)

    if not data and ("ショートカット" in ctx):
        return ["割り当て済みのショートカットキーなし。"]

    sep = ["---"] * len(header)
    out = [join_table_row(header), join_table_row(sep)]
    out.extend(join_table_row(r) for r in data)
    return out


def unify_heading(line: str) -> str:
    m = HEADING_RE.match(line)
    if not m:
        return line
    marks, title = m.group(1), m.group(2).strip()
    title = COUNT_HEAD_RE.sub("", title).strip()

    if re.match(r"^1\.3\s", title) or title in ("モジュール一覧", "VBAモジュール一覧", "VBA モジュール一覧"):
        return f"{marks} {H13_CANON}"
    if SQL_HEAD_RE.match(f"{marks} {title}") or re.match(r"^\d+\.3\s+.*SQL", title):
        num = re.match(r"^(\d+)\.3", title)
        n = num.group(1) if num else "8"
        return f"{marks} {n}.3 SQL 一覧"
    return f"{marks} {title}" if marks else line


def renumber_section5(lines: list[str]) -> list[str]:
    idxs = []
    for i, ln in enumerate(lines):
        m = re.match(r"^(###)\s+5\.(x|\d+)\s+(.*)$", ln)
        if m:
            idxs.append((i, m.group(3).strip()))
    if not idxs:
        return lines
    out = list(lines)
    for n, (i, title) in enumerate(idxs, 1):
        out[i] = f"### 5.{n} {title}"
    return out


def insert_chapter7(text: str) -> str:
    if re.search(r"^## 7\. ユーザーフォーム", text, re.M):
        return text
    # 7. DB → 8. DB, 8. データフロー → 9, 9. セキュリティ → 10
    # 見出し番号の繰り下げは後ろから
    def bump_head(m):
        marks, num, rest = m.group(1), int(m.group(2)), m.group(3)
        if num >= 7:
            return f"{marks} {num + 1}.{rest}"
        return m.group(0)

    text = re.sub(r"^(#{2,4})\s+(\d+)\.(.*)$", bump_head, text, flags=re.M)
    text = re.sub(
        r"^## 8\. DB 接続",
        USERFORM_CH7 + "## 8. DB 接続",
        text,
        count=1,
        flags=re.M,
    )
    text = re.sub(
        r"^7\. \[DB 接続・外部連携\]\(#7-db-接続外部連携\)\s*$",
        "7. [ユーザーフォーム仕様](#7-ユーザーフォーム仕様)\n"
        "8. [DB 接続・外部連携](#8-db-接続外部連携)",
        text,
        count=1,
        flags=re.M,
    )
    text = re.sub(
        r"^8\. \[データフロー\]\(#8-データフロー\)\s*$",
        "9. [データフロー](#9-データフロー)",
        text,
        count=1,
        flags=re.M,
    )
    text = re.sub(
        r"^9\. \[セキュリティ注意事項\]\(#9-セキュリティ注意事項\)\s*$",
        "10. [セキュリティ注意事項](#10-セキュリティ注意事項)",
        text,
        count=1,
        flags=re.M,
    )
    return text


def normalize_meta(text: str) -> str:
    def repl(m):
        key = META_KEY_MAP.get(m.group(1), m.group(1))
        return f"> **{key}**:{m.group(2)}"

    return re.sub(r"^> \*\*(.+?)\*\*:(.*)$", repl, text, flags=re.M)


def add_scope_out(text: str) -> str:
    if re.search(r"^##+ *スコープ外", text, re.M):
        return text
    return text.rstrip() + "\n\n" + SCOPE_OUT


def process_file(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    if path.name in NINE_CHAPTER:
        text = insert_chapter7(text)
    text = normalize_meta(text)

    lines = text.split("\n")
    # 見出し統一・件数行削除・以上削除
    new_lines = []
    for ln in lines:
        if COUNT_BODY_RE.match(ln.strip()):
            continue
        if IJO_RE.match(ln.strip()):
            continue
        if HEADING_RE.match(ln):
            ln = unify_heading(ln)
        new_lines.append(ln)
    lines = renumber_section5(new_lines)

    # 表変換（見出しコンテキスト付き）
    heading_stack: list[str] = []
    out: list[str] = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        hm = HEADING_RE.match(ln)
        if hm:
            level = len(hm.group(1))
            title = hm.group(2).strip()
            heading_stack = [h for h in heading_stack if h.count("#") < level]
            heading_stack.append("#" * level + " " + title)
            out.append(ln)
            i += 1
            continue
        if ln.strip().startswith("|"):
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i])
                i += 1
            ctx = " ".join(h.split(" ", 1)[-1] for h in heading_stack)
            out.extend(transform_table(tbl, ctx))
            continue
        out.append(ln)
        i += 1

    text = "\n".join(out)
    text = add_scope_out(text)
    # 連続空行の抑制（3行以上→2行）
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    if not text.endswith("\n"):
        text += "\n"
    return text


def main() -> None:
    targets = sys.argv[1:]
    if targets:
        files = [Path(t) for t in targets]
    else:
        files = sorted(p for p in DOCS.glob("*_仕様書.md") if p.name not in SKIP)
    for f in files:
        new = process_file(f)
        f.write_text(new, encoding="utf-8")
        print(f"normalized: {f.name}")


if __name__ == "__main__":
    main()
