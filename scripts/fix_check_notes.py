#!/usr/bin/env python3
"""✓説明を基準位置・引用ブロックに揃える（横断ルール A/B/C）。"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SKIP = {"_仕様書共通ルール.md", "feedback.md"}

NOTE = {
    "1.1": "> ✓ = ユーザーが直接操作する、または VBA が動的に表示/非表示を切り替えるシート",
    "1.2": "> ✓ = ユーザー入力を受け付ける、または業務フローの起点となるフォーム",
    "1.3": "> ✓ = ユーザー操作の起点 / DB I/O を含む / 他モジュールから呼び出される / コード行数上位 25%",
    "2": "> 以下の各シートのレイアウト構造表における ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結するセル",
    "3": "> ✓ = VBA から `Range()` で代入または参照され、業務ロジックに直結する名前付き範囲",
    "5": "> ✓ = DB 更新・画面遷移・計算実行など副作用のある操作を起動するボタン",
    "6.0": "> ✓ = ユーザー操作の起点（Click イベント等） / DB I/O を実行 / 他モジュールから呼び出される Public",
    "7": "> ✓ = ユーザー入力を受け付ける、またはイベントで業務処理を起動するコントロール",
    "8.2": "> ✓ = INSERT / UPDATE / DELETE の対象テーブル（参照のみのテーブルは ✓ なし）",
}

HEADING_RE = re.compile(r"^(#{2,4})\s+(.+)$")
BQ_CHECK_RE = re.compile(r"^>\s*.*✓")
BARE_CHECK_RE = re.compile(r"^✓\s*(判定基準)?\s*[=:：]")


def is_check_header(line: str) -> bool:
    s = line.strip()
    if not s.startswith("|"):
        return False
    cells = [c.strip() for c in s.strip("|").split("|")]
    return bool(cells) and cells[0] == "✓"


def is_check_note(line: str) -> bool:
    return bool(BQ_CHECK_RE.match(line) or BARE_CHECK_RE.match(line))


def heading_key(title: str) -> str | None:
    m = re.match(r"^(\d+)(?:\.(\d+[a-z]?))?", title.strip())
    if not m:
        return None
    return f"{m.group(1)}.{m.group(2)}" if m.group(2) else m.group(1)


def parse_headings(lines: list[str]) -> list[tuple[int, int, str | None, str]]:
    out = []
    for i, line in enumerate(lines):
        m = HEADING_RE.match(line)
        if m:
            out.append((i, len(m.group(1)), heading_key(m.group(2)), m.group(2).strip()))
    return out


def section_end(headings, idx: int, n: int) -> int:
    level = headings[idx][1]
    for j in range(idx + 1, len(headings)):
        if headings[j][1] <= level:
            return headings[j][0]
    return n


def find_heading(headings, key: str, level: int | None = None):
    for i, h in enumerate(headings):
        if h[2] == key and (level is None or h[1] == level):
            return i, h
    return None, None


def range_has_check_table(lines, start: int, end: int) -> bool:
    return any(is_check_header(lines[i]) for i in range(start, end))


def first_table(lines, start: int, end: int) -> int | None:
    for i in range(start, end):
        if lines[i].strip().startswith("|"):
            return i
    return None


def notes_in_range(lines, start: int, end: int) -> list[int]:
    return [i for i in range(start, end) if is_check_note(lines[i])]


def window_after_heading(lines, h_idx: int, end: int) -> tuple[int, int]:
    """見出し直後〜最初の表または次見出し直前。"""
    i = h_idx + 1
    while i < end and not lines[i].strip():
        i += 1
    return h_idx + 1, i


def has_note_before_table(lines, h_idx: int, end: int) -> bool:
    t = first_table(lines, h_idx + 1, end)
    lim = t if t is not None else end
    return any(is_check_note(lines[i]) for i in range(h_idx + 1, lim))


def insert_note(lines: list[str], after: int, note: str) -> None:
    """after は見出し行 index。直後に空行＋note＋空行を入れる（重複空行は抑制）。"""
    insert_at = after + 1
    block = []
    if insert_at < len(lines) and lines[insert_at].strip():
        block.append("")
    block.append(note)
    nxt = insert_at
    # 既に空行があれば追加しない
    if nxt >= len(lines) or lines[nxt].strip():
        block.append("")
    lines[insert_at:insert_at] = block


def delete_indices(lines: list[str], idxs: list[int]) -> None:
    for i in sorted(set(idxs), reverse=True):
        del lines[i]
        # 連続空行を1つに
        if i < len(lines) and i > 0 and not lines[i].strip() and not lines[i - 1].strip():
            del lines[i]


def convert_bare_notes(lines: list[str], note_map_line: dict[int, str] | None = None) -> int:
    n = 0
    for i, line in enumerate(lines):
        if BARE_CHECK_RE.match(line):
            lines[i] = note_map_line[i] if note_map_line and i in note_map_line else f"> {line.strip()}"
            n += 1
    return n


def process(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()
    actions: list[str] = []

    # 凡例内の ✓ = は表セルなので触らない。本文の裸 ✓説明だけ引用化。
    legend_end = next(
        (i for i, l in enumerate(lines) if l.startswith("## ") and "凡例" not in l and i > 0),
        0,
    )

    def refresh():
        return parse_headings(lines), len(lines)

    headings, n = refresh()

    # --- 5章: 説明を章直下へ ---
    hi, h = find_heading(headings, "5", 2)
    if h:
        end = section_end(headings, hi, n)
        if range_has_check_table(lines, h[0], end):
            notes = notes_in_range(lines, h[0] + 1, end)
            # 5.x 配下（次の ### 以降）にある説明を除去
            sub = next((hh[0] for hh in headings if hh[0] > h[0] and hh[0] < end and hh[1] >= 3), end)
            misplaced = [i for i in notes if i >= sub]
            if misplaced:
                delete_indices(lines, misplaced)
                actions.append("5: 5.x 下の✓説明を削除して章直下へ")
                headings, n = refresh()
                hi, h = find_heading(headings, "5", 2)
                end = section_end(headings, hi, n)
            if h and not has_note_before_table(lines, h[0], end):
                insert_note(lines, h[0], NOTE["5"])
                actions.append("5: 章直下に✓説明を追加")

    # --- 番号付き小見出し（表の直前） ---
    for key, level in (("1.1", 3), ("1.2", 3), ("1.3", 3), ("3", 2), ("6.0", 3), ("7", 2), ("8.2", 3)):
        headings, n = refresh()
        hi, h = find_heading(headings, key, level)
        if not h:
            continue
        end = section_end(headings, hi, n)
        if not range_has_check_table(lines, h[0], end):
            continue
        if has_note_before_table(lines, h[0], end):
            # 裸の説明があれば引用化（この範囲のみ）
            t = first_table(lines, h[0] + 1, end) or end
            for i in range(h[0] + 1, t):
                if BARE_CHECK_RE.match(lines[i]):
                    lines[i] = NOTE[key]
                    actions.append(f"{key}: 裸✓説明を引用化")
            continue
        insert_note(lines, h[0], NOTE[key])
        actions.append(f"{key}: ✓説明を追加")

    # --- 2章: 2.0 の後・2.1 の前 ---
    headings, n = refresh()
    hi2, h2 = find_heading(headings, "2", 2)
    if h2:
        end2 = section_end(headings, hi2, n)
        if range_has_check_table(lines, h2[0], end2):
            notes = notes_in_range(lines, h2[0] + 1, end2)
            # レイアウト用の既存説明（「以下の各シート」または 2.0 直後）
            hi20, h20 = find_heading(headings, "2.0", 3)
            hi21 = None
            for j, hh in enumerate(headings):
                if hh[2] and hh[2].startswith("2.") and hh[2] not in ("2", "2.0") and hh[0] > h2[0]:
                    hi21 = j
                    break
            # 正しい位置: 2.1 の直前、または 2.0 終端
            if hi21 is not None:
                target = headings[hi21][0]
            elif h20:
                target = section_end(headings, hi20, n)
            else:
                target = h2[0] + 1
            layout_notes = [
                i
                for i in notes
                if "レイアウト" in lines[i] or "各シート" in lines[i]
            ]
            already_ok = any(target - 4 <= i < target for i in layout_notes)
            if not already_ok:
                if layout_notes:
                    delete_indices(lines, layout_notes)
                    headings, n = refresh()
                    hi21 = None
                    for j, hh in enumerate(headings):
                        if hh[2] and hh[2].startswith("2.") and hh[2] not in ("2", "2.0") and hh[1] == 3:
                            # 最初の 2.x（2.0 以外）
                            if hh[2] != "2.0":
                                hi21 = j
                                break
                    if hi21 is not None:
                        target = headings[hi21][0]
                    else:
                        target = headings[find_heading(headings, "2", 2)[0]][0] + 1
                # 2.1 見出しの直前に挿入
                insert_at = target
                block = [NOTE["2"], ""]
                if insert_at > 0 and lines[insert_at - 1].strip():
                    block = [""] + block
                lines[insert_at:insert_at] = block
                actions.append("2: レイアウト✓説明を 2.0 後 / 2.1 前に配置")

    # --- 4章・9章: 既存の裸説明だけ引用化（新規追加しない） ---
    headings, n = refresh()
    for key in ("4", "9"):
        hi, h = find_heading(headings, key, 2)
        if not h:
            continue
        end = section_end(headings, hi, n)
        for i in range(h[0] + 1, end):
            if BARE_CHECK_RE.match(lines[i]):
                lines[i] = f"> {lines[i].strip()}"
                actions.append(f"{key}: 既存✓説明を引用化")

    # 凡例より後の残りの裸 ✓説明
    headings, n = refresh()
    legend_end = next(
        (h[0] for h in headings if h[3].startswith("目次") or (h[1] == 2 and h[2] == "1")),
        0,
    )
    for i in range(legend_end, n):
        if BARE_CHECK_RE.match(lines[i]):
            lines[i] = f"> {lines[i].strip()}"
            actions.append(f"L{i+1}: 裸✓説明を引用化")

    if actions:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return actions


def main() -> int:
    files = sorted(p for p in DOCS.glob("*_仕様書.md") if p.name not in SKIP)
    changed = 0
    for f in files:
        actions = process(f)
        if actions:
            changed += 1
            print(f"✓ {f.name}")
            for a in actions:
                print(f"    - {a}")
        else:
            print(f"· {f.name}（変更なし）")
    print(f"\n{changed}/{len(files)} files updated")
    return 0


if __name__ == "__main__":
    sys.exit(main())
