#!/usr/bin/env python3
"""納品フォルダを Windows エクスプローラーで文字化けしない ZIP にパックする。

macOS の Finder / 標準 zip は日本語ファイル名を UTF-8 バイト列のまま書き、
ZIP の Language Encoding Flag (UTF-8 / EFS, general purpose bit 11) を
立てないことがある。日本語 Windows は CP932 として解釈し「納品」→「邂榊刀」
のように化ける。

本スクリプトは UTF-8 フラグ付きで再パックする。__MACOSX / .DS_Store は除外。

使い方:
  python3 scripts/pack_nohin_zip.py
  python3 scripts/pack_nohin_zip.py --out ~/Desktop/納品.zip
"""

from __future__ import annotations

import argparse
import os
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SOURCE = REPO_ROOT / "納品"
DEFAULT_OUTPUT = Path.home() / "Desktop" / "納品.zip"

ZIP_UTF8_FLAG = 0x800
SKIP_FILE_NAMES = frozenset({".DS_Store"})


def _arcname(source_dir: Path, file_path: Path) -> str:
    rel = file_path.relative_to(source_dir.parent)
    return rel.as_posix()


def _should_skip(file_path: Path) -> bool:
    if file_path.name in SKIP_FILE_NAMES:
        return True
    return "__MACOSX" in file_path.parts


def pack_delivery_zip(source_dir: Path, output_zip: Path) -> int:
    source_dir = source_dir.resolve()
    if not source_dir.is_dir():
        raise SystemExit(f"ソースがフォルダではありません: {source_dir}")

    files = sorted(
        p for p in source_dir.rglob("*") if p.is_file() and not _should_skip(p)
    )
    if not files:
        raise SystemExit(f"パックするファイルがありません: {source_dir}")

    output_zip.parent.mkdir(parents=True, exist_ok=True)
    if output_zip.exists():
        output_zip.unlink()

    with zipfile.ZipFile(
        output_zip, mode="w", compression=zipfile.ZIP_DEFLATED
    ) as zf:
        for file_path in files:
            name = _arcname(source_dir, file_path)
            zinfo = zipfile.ZipInfo.from_file(file_path, name)
            zinfo.flag_bits |= ZIP_UTF8_FLAG
            zinfo.compress_type = zipfile.ZIP_DEFLATED
            with file_path.open("rb") as src:
                zf.writestr(zinfo, src.read())

    return len(files)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
        help=f"納品 docx フォルダ (既定: {DEFAULT_SOURCE})",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"出力 ZIP パス (既定: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()
    count = pack_delivery_zip(args.source, args.out.expanduser().resolve())
    print(f"Wrote {args.out.expanduser().resolve()} ({count} files)")


if __name__ == "__main__":
    main()
