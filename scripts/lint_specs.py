#!/usr/bin/env python3
"""仕様書 Markdown の記載ルール lint。機械判定できる項目のみ。"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SKIP = {"_仕様書共通ルール.md"}


def tables(md: str):
    lines = md.split("\n")
    i = 0
    while i < len(lines):
        if lines[i].strip().startswith("|") and i + 1 < len(lines) and re.match(
            r"^\|[\s:|-]+\|$", lines[i + 1].strip()
        ):
            hdr = [c.strip() for c in lines[i].strip().split("|")[1:-1]]
            yield hdr
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                i += 1
        else:
            i += 1


def lint_file(path: Path) -> list[str]:
    t = path.read_text(encoding="utf-8")
    issues = []
    if re.search(r"（全\s*\d+\s*件）", t):
        issues.append("件数表記（全 N 件）が残っている")
    if re.search(r"^#{3,4}\s+\d+\.x ", t, re.M):
        issues.append("見出し番号が 5.x のまま")
    if "割り当てなし" in t:
        issues.append("ショートカット「割り当てなし」が残っている")
    if re.search(r"^\*?以\s*上\*?\s*$", t, re.M):
        issues.append("本文末尾の「以上」が残っている（docx 側で付与）")
    if not re.search(r"^##+ *スコープ外", t, re.M):
        issues.append("スコープ外セクションなし")
    if not re.search(r"^## 10\. セキュリティ注意事項", t, re.M):
        issues.append("10章見出しが「10. セキュリティ注意事項」ではない")
    m13 = re.search(r"^### 1\.3 (.+)$", t, re.M)
    if m13 and m13.group(1).strip() != "VBA モジュール一覧":
        issues.append(f"1.3 見出しが不統一: {m13.group(1).strip()}")
    m83 = re.search(r"^### 8\.3 (.+)$", t, re.M)
    if m83 and m83.group(1).strip() != "SQL 一覧":
        issues.append(f"8.3 見出しが不統一: {m83.group(1).strip()}")
    chs = [int(x) for x in re.findall(r"^## (\d+)\.", t, re.M)]
    if chs and max(chs) != 10:
        issues.append(f"最大章番号が 10 ではない: {max(chs)}")
    if re.search(r"^## 7\. DB", t, re.M):
        issues.append("7章が DB 接続のまま（ユーザーフォーム章が欠落）")
    for hdr in tables(t):
        if "#" in hdr:
            issues.append(f"表見出しに # が残っている: {hdr}")
            break
        if "✓" in hdr and hdr.index("✓") != 0:
            issues.append(f"✓列が先頭ではない: {hdr}")
            break
    return issues


def main() -> int:
    files = sorted(p for p in DOCS.glob("*_仕様書.md") if p.name not in SKIP)
    failed = 0
    for f in files:
        issues = lint_file(f)
        if issues:
            failed += 1
            print(f"✗ {f.name}")
            for iss in issues:
                print(f"    - {iss}")
        else:
            print(f"✓ {f.name}")
    print(f"\n{len(files) - failed}/{len(files)} files clean")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
