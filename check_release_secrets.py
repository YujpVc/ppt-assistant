#!/usr/bin/env python3
"""Refuse a release that contains a non-empty image API key or secret file."""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FORBIDDEN = {"image-config.json", ".env", "secrets.json", "credentials.json"}
PATTERNS = [
    re.compile(r'["\']api_key["\']\s*:\s*["\']([^"\']+)["\']'),
    re.compile(r'["\']key["\']\s*:\s*["\'](sk-[^"\']+)["\']'),
    re.compile(r'UUAPI_API_KEY\s*=\s*["\']?([^\s"\']+)'),
]

def main() -> int:
    problems = []
    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts or "node_modules" in path.parts:
            continue
        if path.name in FORBIDDEN:
            problems.append(f"forbidden secret file: {path.relative_to(ROOT)}")
            continue
        if path.suffix.lower() not in {".py", ".md", ".yml", ".yaml", ".json", ".sh", ".txt"}:
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        for pattern in PATTERNS:
            for match in pattern.finditer(text):
                value = match.group(1).strip()
                if value and value not in {"...", "YOUR_KEY", "your-key", "<key>"}:
                    problems.append(f"non-empty key in {path.relative_to(ROOT)}")
                    break
    if problems:
        print("release blocked:\n" + "\n".join(f"- {p}" for p in problems), file=sys.stderr)
        return 2
    print("release secret scan: ok")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
