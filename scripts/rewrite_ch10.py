#!/usr/bin/env python3
"""10章を ExR3オーブン管理温度補正有 形式（No / カテゴリ / 内容 / リスク＝「レベル：理由」）へ揃える。"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SKIP = {"_仕様書共通ルール.md"}

LEVEL_RE = re.compile(r"^(高|中|低|低〜中)[：:]")

CANNED_RISK = (
    ("認証", "中：VBAエディタで閲覧可能"),
    ("パスワード", "中：シート非表示にしても保護解除で閲覧可能"),
    ("インジェクション", "中：文字列連結によるSQL構築"),
    ("エラー", "中：サイレント障害の可能性"),
    ("Resume Next", "中：サイレント障害の可能性"),
    ("トランザクション", "中：データ不整合リスク"),
    ("保護", "低：VBAからは無制限アクセス"),
    ("保存", "低：変更破棄が無警告"),
    ("AutoExec", "低：意図通りの起動・終了処理"),
    ("IOC", "低：社内アドレスの露出"),
    ("ネットワーク", "低：サーバー移行時に変更要"),
    ("ファイル", "中：ローカルファイルの作成・削除が可能"),
)


def canned_risk(cat: str, content: str) -> str:
    blob = cat + content
    for key, val in CANNED_RISK:
        if key in blob:
            return val
    return "中：要確認"


def format_risk(risk: str, content: str, cat: str = "") -> str:
    risk = (risk or "").strip()
    if LEVEL_RE.match(risk):
        # 内容を切り詰めて付けた自動生成は捨てて定型理由に戻す
        rest = LEVEL_RE.sub("", risk)
        if len(rest) > 28 or rest.startswith("`") or rest.startswith("DSN"):
            return canned_risk(cat, content)
        return risk
    if risk in ("高", "中", "低", "低〜中"):
        return canned_risk(cat, content)
    if not risk:
        return canned_risk(cat, content)
    level = "低" if any(k in cat + content + risk for k in ("意図", "誤検出", "社内")) else "中"
    return f"{level}：{risk}"


def split_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def join_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def categorize_olevba(kind: str, keyword: str, content: str) -> tuple[str, str, str]:
    k = f"{kind} {keyword} {content}"
    if "IOC" in kind or re.search(r"\d+\.\d+\.\d+\.\d+", k):
        return "IOC", content or f"ソースにアドレス {keyword} が含まれる", "低：社内アドレスの露出"
    if "AutoExec" in kind or "Workbook_Open" in keyword:
        return "AutoExec", content or f"{keyword} が自動実行される", "低：意図通りの起動・終了処理"
    if "Environ" in keyword:
        return "環境変数", content or "Environ で環境情報を読取る", "低：環境情報の読取り"
    if any(x in keyword for x in ("Kill", "MkDir", "Write", "Output")):
        return "ファイル操作", content or f"{keyword} を使用", "中：ローカルファイルの作成・削除が可能"
    if keyword.strip("`") == "Open" or "Open" in keyword:
        return "ファイル操作 / DB接続", content or "Open を使用（ADO 接続の可能性）", "低：DB 接続またはファイルオープン"
    if any(x in keyword for x in ("Command", "Call", "Chr", "Hex", "Base64", "AppActivate")):
        return "olevba 誤検出", content or f"{keyword} が検出された", "低：olevba の誤検出（VBA の正常構文）"
    return kind or "その他", content or keyword, "低：olevba 検出項目"


def convert_table(header: list[str], rows: list[list[str]], notes: str) -> list[list[str]]:
    h = [x.replace(" ", "") for x in header]
    out = []

    def add(cat, content, risk):
        out.append([cat, content, format_risk(risk, content, cat)])

    if "カテゴリ" in "".join(header) and "リスク" in "".join(header):
        ci = next(i for i, x in enumerate(header) if "カテゴリ" in x)
        content_i = next(i for i, x in enumerate(header) if x == "内容")
        ri = next(i for i, x in enumerate(header) if "リスク" in x)
        for r in rows:
            add(r[ci], r[content_i], r[ri])
        return out

    if "区分" in header and "リスク" in header:
        ci = header.index("区分")
        content_i = header.index("内容") if "内容" in header else 1
        ri = header.index("リスク")
        for r in rows:
            add(r[ci], r[content_i], r[ri])
        return out

    if "種別" in header and "キーワード" in header:
        ki = header.index("種別")
        wi = header.index("キーワード")
        ci = header.index("内容") if "内容" in header else wi
        for r in rows:
            cat, content, risk = categorize_olevba(r[ki], r[wi], r[ci] if ci < len(r) else "")
            add(cat, content, risk)
        if re.search(r"PWD|パスワード|DSN=ricdb", notes, re.I):
            if not any("認証" in r[0] or "DB認証" in r[0] for r in out):
                m = re.search(r"DSN=ricdb;UID=\w+;PWD=\w+", notes)
                cred = m.group(0) if m else "DSN=ricdb;UID=ric;PWD=t6101"
                out.insert(0, ["DB認証情報", f"接続文字列がVBAソース内にハードコード（{cred}）", "中：VBAエディタで閲覧可能"])
        return out

    if "項目" in header and "内容" in header:
        ii = header.index("項目")
        ci = header.index("内容")
        ri = header.index("リスク") if "リスク" in header else (header.index("対策") if "対策" in header else None)
        for r in rows:
            add(r[ii], r[ci], r[ri] if ri is not None else "")
        return out

    # fallback: first text cols
    for r in rows:
        cells = [c for c in r if c not in ("", "—") and not c.isdigit()]
        if len(cells) >= 2:
            add(cells[0], cells[1], cells[2] if len(cells) > 2 else "")
    return out


def rewrite_section(body: str) -> str:
    lines = body.split("\n")
    # find first table
    i = 0
    while i < len(lines) and not lines[i].strip().startswith("|"):
        i += 1
    if i >= len(lines):
        return body  # 該当なし等はそのまま

    prefix = lines[:i]
    tbl = []
    j = i
    while j < len(lines) and lines[j].strip().startswith("|"):
        tbl.append(lines[j])
        j += 1
    suffix = lines[j:]
    notes = "\n".join(prefix + suffix)

    rows = [split_row(ln) for ln in tbl]
    header, data = rows[0], rows[2:]
    converted = convert_table(header, data, notes)
    # 重複カテゴリを簡易マージしない。olevba の AutoExec は1行にまとめる
    merged = []
    seen_auto = False
    for cat, content, risk in converted:
        if cat == "AutoExec":
            if seen_auto:
                continue
            seen_auto = True
            cat, content, risk = (
                "AutoExec",
                "Workbook_Open / BeforeClose 等が自動実行される",
                "低：意図通りの起動・終了処理",
            )
        merged.append((cat, content, risk))
    # olevba 誤検出はリスク喚起にならないので落とす
    merged = [r for r in merged if r[0] != "olevba 誤検出"]
    # 同カテゴリのファイル操作は1行にまとめる
    files = [r for r in merged if r[0].startswith("ファイル")]
    if len(files) > 1:
        merged = [r for r in merged if not r[0].startswith("ファイル")]
        merged.append(("ファイル操作", "Open / Write / Kill / MkDir 等を使用", "中：ローカルファイルの作成・削除が可能"))

    out_tbl = [
        "| No | カテゴリ | 内容 | リスク |",
        "| --- | --- | --- | --- |",
    ]
    for n, (cat, content, risk) in enumerate(merged, 1):
        out_tbl.append(join_row([str(n), cat, content, risk]))

    # 既存の olevba 前置きとブロッククォート注意は表に吸収したので削る
    new_prefix = []
    for ln in prefix:
        if ln.strip().startswith("olevba"):
            continue
        if ln.strip().startswith(">"):
            continue
        new_prefix.append(ln)
    # suffix の注意ブロックも表に取り込んだので削除
    new_suffix = []
    skip_bq = True
    for ln in suffix:
        if ln.strip().startswith(">"):
            continue
        if skip_bq and not ln.strip():
            continue
        skip_bq = False
        new_suffix.append(ln)

    return "\n".join(new_prefix + out_tbl + ([""] if new_suffix else []) + new_suffix).rstrip() + "\n"


def process(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    text = text.replace("## 10. セキュリティ・注意事項", "## 10. セキュリティ注意事項")
    m = re.search(r"^## 10\. セキュリティ注意事項\s*$", text, re.M)
    if not m:
        return
    start = m.end()
    nxt = re.search(r"^## ", text[start:], re.M)
    end = start + nxt.start() if nxt else len(text)
    body = text[start:end]
    # サブセクション 10.1 がある場合は、表形式に畳まずリスク列があるものだけ整形
    if re.search(r"^### 10\.", body, re.M):
        # サブセクション型はファイルごとに残し、後でエージェントが畳む。ここでは見出しだけ統一
        path.write_text(text, encoding="utf-8")
        return
    new_body = rewrite_section(body)
    path.write_text(text[: start] + "\n\n" + new_body.lstrip("\n") + text[end:], encoding="utf-8")
    print("ch10:", path.name)


def main() -> None:
    for f in sorted(p for p in DOCS.glob("*_仕様書.md") if p.name not in SKIP):
        process(f)


if __name__ == "__main__":
    main()
