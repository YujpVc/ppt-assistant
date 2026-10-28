#!/usr/bin/env python3
"""Render a PPTX to PDF and PNG pages for visual review.

Uses a system ``soffice`` when available, then the user-scoped official
LibreOffice bundle installed by the setup script. The output stays beside the
input deck so foreground/background jobs see the same files.
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def find_soffice() -> Path | None:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    candidates = [
        Path.home()
        / ".local/share/dsh-tools/libreoffice-official/opt/libreoffice26.8/opt/libreoffice26.8/program/soffice",
        Path.home() / ".local/share/dsh-tools/libreoffice/usr/lib/libreoffice/program/soffice",
    ]
    if found:
        candidates.insert(0, Path(found))
    return next((path for path in candidates if path.is_file() and os.access(path, os.X_OK)), None)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--outdir", type=Path)
    parser.add_argument("--dpi", type=int, default=110)
    args = parser.parse_args()
    pptx = args.pptx.resolve()
    if not pptx.is_file():
        print(f"render_ppt: file not found: {pptx}", file=sys.stderr)
        return 2
    soffice = find_soffice()
    if soffice is None:
        print("render_ppt: no usable soffice/libreoffice found", file=sys.stderr)
        return 3
    outdir = (args.outdir or pptx.parent / ".render").resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    for old_page in outdir.glob("page-*.png"):
        old_page.unlink()
    profile = outdir / ".lo-profile"
    profile.mkdir(exist_ok=True)
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    command = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(outdir),
        str(pptx),
    ]
    result = subprocess.run(command, env=env, text=True, capture_output=True)
    if result.returncode:
        print(result.stdout, end="")
        print(result.stderr, end="", file=sys.stderr)
        return result.returncode
    pdf = outdir / f"{pptx.stem}.pdf"
    if not pdf.is_file():
        print(f"render_ppt: conversion completed without PDF: {pdf}", file=sys.stderr)
        return 4
    prefix = outdir / "page"
    subprocess.run(["pdftoppm", "-png", "-r", str(args.dpi), str(pdf), str(prefix)], check=True)
    raw_pages = sorted(outdir.glob("page-[0-9]*.png"))
    staging = []
    for index, raw_page in enumerate(raw_pages, 1):
        temporary = outdir / f".render-page-{index}.png"
        raw_page.rename(temporary)
        staging.append(temporary)
    for index, temporary in enumerate(staging, 1):
        temporary.rename(outdir / f"page-{index:02d}.png")
    print(f"renderer={soffice}")
    print(f"pdf={pdf}")
    print(f"pages={len(staging)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
