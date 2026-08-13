#!/usr/bin/env python3
"""docs/*.md を 納品/*.docx に一括変換する。

目次は静的見出し一覧（TOC フィールドなし）のため、LibreOffice による
フィールド更新は不要。残す場合は --update-fields を指定する。
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
OUT = ROOT / "納品"
SOFFICE = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
SKIP = {"_仕様書共通ルール.md"}


def md_files() -> list[Path]:
    return sorted(p for p in DOCS.glob("*_仕様書.md") if p.name not in SKIP)


def generate_docx(md: Path, dest: Path, cfg: dict) -> None:
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "md_to_docx.py"),
        str(md),
        "--out",
        str(dest),
        "--rev",
        cfg["rev"],
        "--date",
        cfg["date"],
    ]
    if cfg.get("issuer"):
        cmd += ["--issuer", cfg["issuer"]]
    if cfg.get("client"):
        cmd += ["--client", cfg["client"]]
    subprocess.check_call(cmd)


def update_toc_libreoffice(docx_paths: list[Path]) -> None:
    """docx→docx 再保存で TOC / PAGE フィールドを更新する。"""
    if not SOFFICE.exists():
        print("WARNING: LibreOffice が見つかりません。目次ページ番号は未確定です。", file=sys.stderr)
        print("  brew install --cask libreoffice", file=sys.stderr)
        return
    with tempfile.TemporaryDirectory(prefix="lo-toc-") as tmp:
        tmp_dir = Path(tmp)
        profile = tmp_dir / "profile"
        outdir = tmp_dir / "out"
        outdir.mkdir()
        cmd = [
            str(SOFFICE),
            "--headless",
            "--norestore",
            "--nolockcheck",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to",
            "docx",
            "--outdir",
            str(outdir),
        ] + [str(p) for p in docx_paths]
        subprocess.check_call(cmd)
        for src in docx_paths:
            converted = outdir / src.name
            if not converted.exists():
                print(f"WARNING: LibreOffice 出力なし: {src.name}", file=sys.stderr)
                continue
            shutil.copy2(converted, src)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--rev", default="1.1")
    ap.add_argument("--date", default="2026-08-13")
    ap.add_argument("--issuer", default="")
    ap.add_argument("--client", default="")
    ap.add_argument(
        "--update-fields",
        action="store_true",
        help="LibreOffice でフィールドを再保存する（既定: しない。目次は静的）",
    )
    ap.add_argument("files", nargs="*", help="対象 md（省略時は docs 全件）")
    args = ap.parse_args()

    if args.files:
        mds = [Path(f) for f in args.files]
    else:
        mds = md_files()

    OUT.mkdir(parents=True, exist_ok=True)
    cfg = {"rev": args.rev, "date": args.date, "issuer": args.issuer, "client": args.client}
    dests = []
    for md in mds:
        dest = OUT / (md.stem + ".docx")
        print(f"generate: {md.name} → {dest.name}")
        generate_docx(md, dest, cfg)
        dests.append(dest)

    if args.update_fields:
        print("LibreOffice でフィールドを再保存中...")
        update_toc_libreoffice(dests)
        print("field update done")


if __name__ == "__main__":
    main()
