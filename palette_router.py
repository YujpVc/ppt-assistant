#!/usr/bin/env python3
"""Pick a palette next to the last one used, not the same default every deck.

Usage:
  bash run_py.sh palette_router.py --seed "申报书标题"
  bash run_py.sh palette_router.py --seed "标题" --json
"""
from __future__ import annotations
import argparse, hashlib, json, os
from datetime import datetime, timezone
from pathlib import Path

PALETTES = {
    "ivory_terracotta": {"bg":"F7F3EC","panel":"FFFDF8","text":"1C1B19","muted":"6F675C","accent":"B04A2A","accent2":"A8763E","accent3":"5B6B7A","soft":"F2E7D9","tint":"E8EFE9"},
    "cool_teal": {"bg":"F3F5F3","panel":"FFFFFF","text":"182421","muted":"5F6D68","accent":"2F5D62","accent2":"A8763E","accent3":"6E8A99","soft":"DFE7E3","tint":"E8EFE9"},
    "chalk_indigo": {"bg":"F5F6F8","panel":"FFFFFF","text":"1B2430","muted":"5F6B7A","accent":"2E4A7D","accent2":"C08A4E","accent3":"6B7FA3","soft":"E4E8F0","tint":"E8EDF5"},
    "sage_charcoal": {"bg":"EFF2EC","panel":"FBFDF9","text":"20261F","muted":"5E6859","accent":"4F6B57","accent2":"A8763E","accent3":"7C8C7A","soft":"DFE7E3","tint":"E8EFE9"},
    "bone_aubergine": {"bg":"F6F1EE","panel":"FFFCFA","text":"241D26","muted":"6C5F6B","accent":"6A3F5E","accent2":"C08A5A","accent3":"7B6C86","soft":"EEE3E8","tint":"F0E8EC"},
    "mist_navy": {"bg":"F2F4F7","panel":"FFFFFF","text":"18202B","muted":"5D6877","accent":"1F3A5F","accent2":"B58A55","accent3":"6A7F99","soft":"E3E9EF","tint":"E9EEF3"},
    "paper_olive": {"bg":"F5F2E9","panel":"FFFDF6","text":"25251D","muted":"6A6758","accent":"5C6B3B","accent2":"B08A4F","accent3":"7A8465","soft":"ECE7D8","tint":"F0EDDF"},
    "cool_cobalt": {"bg":"F4F7FA","panel":"FFFFFF","text":"152033","muted":"5D6B7D","accent":"2255A4","accent2":"C08A4E","accent3":"7C8FA8","soft":"E1EAF4","tint":"E8F0F8"},
}
ORDER = list(PALETTES)
HISTORY = Path(os.environ.get("PPT_PALETTE_HISTORY", Path.home() / ".config/ppt-assistant/palette-history.json"))

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", default="ppt-assistant")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--no-history", action="store_true")
    args = ap.parse_args()
    h = int(hashlib.sha256(args.seed.encode("utf-8")).hexdigest()[:8], 16)
    last = None
    if not args.no_history:
        try:
            last = json.loads(HISTORY.read_text()).get("last")
        except Exception:
            pass
    start = h % len(ORDER)
    if last in ORDER:
        start = (ORDER.index(last) + 1 + (h % 3)) % len(ORDER)
    name = ORDER[start]; colors = PALETTES[name]
    if not args.no_history:
        try:
            HISTORY.parent.mkdir(parents=True, exist_ok=True)
            HISTORY.write_text(json.dumps({"last": name, "seed": args.seed, "at": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False, indent=1))
        except Exception:
            pass
    out = {"name": name, "colors": colors}
    if args.json:
        print(json.dumps(out, ensure_ascii=False, indent=1))
    else:
        print(f"palette: {name}")
        print("meta.accent:", colors["accent"])
        print("meta.colors:", json.dumps({k: colors[k] for k in ("bg","panel","text","muted","tint","soft","accent2","accent3")}, ensure_ascii=False))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
