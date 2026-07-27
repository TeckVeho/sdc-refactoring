#!/usr/bin/env python3
"""Markdown仕様書 → docx変換スクリプト

✅️ Ex生産情報一覧1_仕様書.docx のレイアウト・書式を再現して
docs/配下の .md 仕様書を .docx に変換する。
"""

import re
import sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ── スタイル定義 ──────────────────────────────────────────
FONT_NAME = "Arial"
FONT_NAME_EA = "Arial"

COLOR_H1 = RGBColor(0x1F, 0x38, 0x64)  # #1f3864
COLOR_H2 = RGBColor(0x2E, 0x5F, 0xA3)  # #2e5fa3
COLOR_H3 = RGBColor(0x2E, 0x75, 0xB6)  # #2e75b6
COLOR_H4 = RGBColor(0x55, 0x55, 0x55)  # #555555
COLOR_NOTE = RGBColor(0x59, 0x59, 0x59)  # #595959
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SIZE_TITLE = Pt(30)
SIZE_SUBTITLE = Pt(20)
SIZE_H1 = Pt(16)
SIZE_H2 = Pt(14)
SIZE_H3 = Pt(12)
SIZE_H4 = Pt(11)
SIZE_BODY = Pt(10)
SIZE_NOTE = Pt(9)
SIZE_CODE = Pt(9)
SIZE_META = Pt(8)
SIZE_TOC_H2 = Pt(11)
SIZE_TOC_H3 = Pt(10)

PAGE_MARGIN = Twips(900)
PAGE_WIDTH = Twips(11906)
PAGE_HEIGHT = Twips(15120)

TABLE_HEADER_BG = "2e5fa3"
META_ROW_EVEN_BG = "ebf3fb"
META_ROW_ODD_BG = "ffffff"
CODE_BG = "f2f2f2"


def set_run_font(run, name=FONT_NAME, size=SIZE_BODY, color=COLOR_BLACK,
                 bold=False, font_ea=None):
    run.font.name = name
    run.font.size = size
    run.font.color.rgb = color
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rpr.append(parse_xml(
        f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_ea or name}"/>'
    ))


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run._element.append(parse_xml(f'<w:br {nsdecls("w")} w:type="page"/>'))


def set_cell_shading(cell, color_hex):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>'
    )
    tcPr.append(shading)


def set_table_borders(table, sz=4, color="000000"):
    tbl = table._element
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        tbl.insert(0, tblPr)
    borders_xml = f"""
    <w:tblBorders {nsdecls("w")}>
      <w:top w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
      <w:left w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
      <w:bottom w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
      <w:right w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
      <w:insideH w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
      <w:insideV w:val="single" w:sz="{sz}" w:color="{color}" w:space="0"/>
    </w:tblBorders>"""
    tblPr.append(parse_xml(borders_xml))


def add_styled_paragraph(doc, text, size=SIZE_BODY, color=COLOR_BLACK,
                         bold=False, alignment=None, font=FONT_NAME,
                         space_before=None, space_after=None, indent_left=None):
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment
    if space_before is not None:
        p.paragraph_format.space_before = space_before
    if space_after is not None:
        p.paragraph_format.space_after = space_after
    if indent_left is not None:
        p.paragraph_format.left_indent = indent_left

    parts = parse_inline_markdown(text)
    for part_text, part_bold, part_code in parts:
        run = p.add_run(part_text)
        is_bold = bold or part_bold
        f = "Courier New" if part_code else font
        set_run_font(run, name=f, size=size, color=color, bold=is_bold)

    return p


def parse_inline_markdown(text):
    """インライン Markdown（**太字** / `コード`）をパース"""
    parts = []
    pattern = re.compile(r'(\*\*(.+?)\*\*|`([^`]+)`)')
    last = 0
    for m in pattern.finditer(text):
        if m.start() > last:
            parts.append((text[last:m.start()], False, False))
        if m.group(2):
            parts.append((m.group(2), True, False))
        elif m.group(3):
            parts.append((m.group(3), False, True))
        last = m.end()
    if last < len(text):
        parts.append((text[last:], False, False))
    if not parts:
        parts.append((text, False, False))
    return parts


def add_code_block(doc, lines):
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        pPr = p._element.get_or_add_pPr()
        pPr.append(parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="{CODE_BG}" w:val="clear"/>'
        ))
        run = p.add_run(line)
        set_run_font(run, name="Courier New", size=SIZE_CODE, color=COLOR_BLACK)


def add_table_from_rows(doc, header, rows):
    ncols = len(header)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(table)

    for ci, h in enumerate(header):
        cell = table.rows[0].cells[ci]
        set_cell_shading(cell, TABLE_HEADER_BG)
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h.strip())
        set_run_font(run, size=SIZE_NOTE, color=COLOR_WHITE, bold=True)

    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = table.rows[ri + 1].cells[ci]
            val = row[ci].strip() if ci < len(row) else ""
            cell.text = ""
            p = cell.paragraphs[0]
            parts = parse_inline_markdown(val)
            for part_text, part_bold, part_code in parts:
                run = p.add_run(part_text)
                f = "Courier New" if part_code else FONT_NAME
                set_run_font(run, name=f, size=SIZE_NOTE, color=COLOR_BLACK,
                             bold=part_bold)


# ── Markdown パーサー ──────────────────────────────────────

def parse_md(md_path):
    """Markdown を構造化トークン列に変換"""
    text = Path(md_path).read_text(encoding="utf-8")
    lines = text.split("\n")
    tokens = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # コードブロック
        if line.startswith("```"):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].startswith("```"):
                code_lines.append(lines[i])
                i += 1
            tokens.append(("code", code_lines))
            i += 1
            continue

        # 見出し
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            tokens.append(("heading", level, m.group(2).strip()))
            i += 1
            continue

        # テーブル
        if line.strip().startswith("|") and "|" in line[1:]:
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            if len(table_lines) >= 2:
                header = [c.strip() for c in table_lines[0].split("|")[1:-1]]
                rows = []
                for tl in table_lines[2:]:
                    cols = [c.strip() for c in tl.split("|")[1:-1]]
                    rows.append(cols)
                tokens.append(("table", header, rows))
            continue

        # blockquote
        if line.startswith("> ") or line.strip() == ">":
            bq_lines = []
            while i < len(lines) and (lines[i].startswith("> ") or lines[i].strip() == ">"):
                content = lines[i][2:] if lines[i].startswith("> ") else ""
                bq_lines.append(content)
                i += 1
            tokens.append(("blockquote", bq_lines))
            continue

        # 水平線
        if re.match(r'^---+\s*$', line.strip()):
            tokens.append(("hr",))
            i += 1
            continue

        # 空行
        if line.strip() == "":
            i += 1
            continue

        # リスト項目
        m = re.match(r'^(\s*)[-*]\s+(.*)', line)
        if m:
            tokens.append(("list_item", m.group(2).strip(), len(m.group(1))))
            i += 1
            continue

        # 番号付きリスト
        m = re.match(r'^(\s*)\d+\.\s+(.*)', line)
        if m:
            tokens.append(("list_item", m.group(2).strip(), len(m.group(1))))
            i += 1
            continue

        # 通常の段落
        tokens.append(("paragraph", line.strip()))
        i += 1

    return tokens


# ── メタ情報抽出 ──────────────────────────────────────────

def extract_metadata(tokens):
    """先頭の blockquote からメタ情報を抽出"""
    meta = {}
    title = ""

    for tok in tokens:
        if tok[0] == "heading" and tok[1] == 1:
            title = tok[2]
            break

    for tok in tokens:
        if tok[0] == "blockquote":
            for line in tok[1]:
                m = re.match(r'\*\*(.+?)\*\*:\s*(.*)', line)
                if m:
                    meta[m.group(1).strip()] = m.group(2).strip()
            if meta:
                break

    return title, meta


# ── docx 生成 ─────────────────────────────────────────────

def build_docx(tokens, output_path):
    doc = Document()

    # ページ設定
    section = doc.sections[0]
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.left_margin = PAGE_MARGIN
    section.right_margin = PAGE_MARGIN
    section.top_margin = PAGE_MARGIN
    section.bottom_margin = PAGE_MARGIN

    # デフォルトフォント
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = SIZE_BODY
    style.paragraph_format.space_after = Pt(4)
    rpr = style.element.get_or_add_rPr()
    rpr.append(parse_xml(
        f'<w:rFonts {nsdecls("w")} w:eastAsia="{FONT_NAME}"/>'
    ))

    title, meta = extract_metadata(tokens)

    # ── 表紙 ──
    doc.add_paragraph()
    doc.add_paragraph()
    add_styled_paragraph(doc, title.replace(" 仕様書", ""),
                         size=SIZE_TITLE, color=COLOR_H1, bold=True,
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_styled_paragraph(doc, "システム仕様書",
                         size=SIZE_SUBTITLE, color=COLOR_H2, bold=True,
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    if meta:
        rows_data = list(meta.items())
        table = doc.add_table(rows=len(rows_data), cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)
        for ri, (key, val) in enumerate(rows_data):
            bg = META_ROW_EVEN_BG if ri % 2 == 0 else META_ROW_ODD_BG
            for ci, text in enumerate([key, val]):
                cell = table.rows[ri].cells[ci]
                set_cell_shading(cell, bg)
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(text)
                set_run_font(run, size=SIZE_META, bold=True)

    add_page_break(doc)

    # ── 目次（見出しから自動生成）──
    add_styled_paragraph(doc, "目次", size=SIZE_H1, color=COLOR_H1, bold=True,
                         space_after=Pt(12))

    headings_for_toc = []
    for tok in tokens:
        if tok[0] == "heading" and tok[1] in (2, 3):
            if tok[1] == 2 and tok[2].strip() == "目次":
                continue
            headings_for_toc.append((tok[1], tok[2]))

    for level, text in headings_for_toc:
        if level == 2:
            add_styled_paragraph(doc, text,
                                 size=SIZE_TOC_H2, color=COLOR_H1, bold=True,
                                 space_before=Pt(2), space_after=Pt(2))
        else:
            add_styled_paragraph(doc, text,
                                 size=SIZE_TOC_H3, color=COLOR_H2, bold=True,
                                 indent_left=Twips(360),
                                 space_before=Pt(1), space_after=Pt(1))

    add_page_break(doc)

    # ── 本文 ──
    skip_first_heading = True
    skip_first_blockquote = True
    first_h2 = True
    skip_toc_heading = True
    skip_toc_list = False

    for tok in tokens:
        if tok[0] == "heading":
            level, text = tok[1], tok[2]
            if skip_first_heading and level == 1:
                skip_first_heading = False
                continue

            if level == 2 and skip_toc_heading and text.strip() == "目次":
                skip_toc_heading = False
                skip_toc_list = True
                continue

            if level == 2:
                if not first_h2:
                    add_page_break(doc)
                first_h2 = False
                add_styled_paragraph(doc, text,
                                     size=SIZE_H1, color=COLOR_H1, bold=True,
                                     space_before=Pt(12), space_after=Pt(8))
            elif level == 3:
                add_styled_paragraph(doc, text,
                                     size=SIZE_H2, color=COLOR_H2, bold=True,
                                     space_before=Pt(10), space_after=Pt(6))
            elif level == 4:
                add_styled_paragraph(doc, text,
                                     size=SIZE_H3, color=COLOR_H3, bold=True,
                                     space_before=Pt(8), space_after=Pt(4))
            elif level >= 5:
                add_styled_paragraph(doc, text,
                                     size=SIZE_H4, color=COLOR_H4, bold=True,
                                     space_before=Pt(6), space_after=Pt(3))

        elif tok[0] == "blockquote":
            if skip_first_blockquote:
                skip_first_blockquote = False
                continue
            bq_text = " ".join(tok[1]).strip()
            if bq_text:
                add_styled_paragraph(doc, f"📌 {bq_text}",
                                     size=SIZE_NOTE, color=COLOR_NOTE)

        elif tok[0] == "paragraph":
            text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', tok[1])
            add_styled_paragraph(doc, text, size=SIZE_BODY)

        elif tok[0] == "list_item":
            text = tok[1]
            if skip_toc_list and re.match(r'\[.+\]\(#', text):
                continue
            skip_toc_list = False
            indent = Twips(360) if tok[2] == 0 else Twips(360 + tok[2] * 180)
            text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
            p = add_styled_paragraph(doc, f"• {text}", size=SIZE_BODY,
                                     indent_left=indent)

        elif tok[0] == "table":
            header, rows = tok[1], tok[2]
            add_table_from_rows(doc, header, rows)
            doc.add_paragraph()

        elif tok[0] == "code":
            add_code_block(doc, tok[1])
            doc.add_paragraph()

        elif tok[0] == "hr":
            pass

    doc.save(str(output_path))
    print(f"✅ 生成完了: {output_path}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 md_to_docx.py <input.md> [output.docx]")
        sys.exit(1)

    md_path = Path(sys.argv[1])
    if len(sys.argv) >= 3:
        out_path = Path(sys.argv[2])
    else:
        out_path = md_path.with_suffix(".docx")

    tokens = parse_md(md_path)
    build_docx(tokens, out_path)


if __name__ == "__main__":
    main()
