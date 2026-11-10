#!/usr/bin/env python3
"""Vision QC for a generated image before it enters a deck.

Checks the things a prompt cannot guarantee: does the subject match the brief,
is there accidental text or a watermark, is the background clean, and which
side is quiet enough for copy. Returns JSON and exits 2 when the asset fails.

    python3 image_check.py assets/hero.png --intent "robot arm and amber bottle,
        negative space on the left" --min-score 7
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import urllib.request
from pathlib import Path

BASE = "https://api.deepseek.com"
MODEL = "deepseek-v4-flash-vision-exp"
PROMPT = """You are a strict image QA for presentation assets. Judge the image against the intent.
Return ONLY a JSON object with this shape:
{"subject_match": 0-10, "has_text": true/false, "has_watermark": true/false,
 "background_clean": 0-10, "negative_space": "left|right|top|bottom|none",
 "issues": ["short concrete issue"], "score": 0-10}
Rules: score reflects fitness as a full-bleed slide image; accidental text, logos,
watermarks, malformed hands or collages are automatic failures; "negative_space"
is the quietest side for overlaying copy. Intent: %s"""


def _key() -> str:
    key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if key:
        return key
    homes = [os.environ.get("DSH_HOME", "").strip(),
             os.path.expanduser("~/.dsh-linux"), os.path.expanduser("~/.dsh")]
    for home in homes:
        if not home:
            continue
        path = os.path.join(home, ".credentials.yaml")
        if not os.path.isfile(path):
            continue
        try:
            for line in open(path, encoding="utf-8"):
                stripped = line.strip()
                if stripped.startswith("DEEPSEEK_API_KEY:"):
                    value = stripped.split(":", 1)[1].strip().strip("'\"")
                    if value:
                        return value
        except OSError:
            continue
    return ""


def _mime(path: Path) -> str:
    return {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".webp": "image/webp"}.get(path.suffix.lower(), "image/png")


def _parse(text: str) -> dict:
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError(f"模型没有返回 JSON：{text[:200]}")
    data = json.loads(text[start:end + 1])
    for field in ("subject_match", "has_text", "has_watermark",
                  "background_clean", "negative_space", "issues", "score"):
        if field not in data:
            raise ValueError(f"模型 JSON 缺少 {field}")
    return data


def check(image: Path, intent: str, min_score: float, timeout: int = 120) -> dict:
    key = _key()
    if not key:
        raise RuntimeError("未找到 DEEPSEEK_API_KEY（env、$DSH_HOME、~/.dsh-linux、~/.dsh）")
    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": PROMPT % intent},
            {"type": "image_url", "image_url": {
                "url": f"data:{_mime(image)};base64," +
                       base64.b64encode(image.read_bytes()).decode()}},
        ]}],
        "temperature": 0,
    }
    data = json.dumps(payload).encode("utf-8")
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    try:
        import requests
    except ImportError:
        requests = None
    if requests is not None:
        proxy = (os.environ.get("PPT_IMAGE_PROXY") or os.environ.get("HTTPS_PROXY")
                 or os.environ.get("https_proxy"))
        proxies = {"http": proxy, "https": proxy} if proxy else None
        response = requests.post(BASE + "/chat/completions", data=data, headers=headers,
                                 timeout=timeout, proxies=proxies)
        if response.status_code >= 400:
            raise RuntimeError(f"视觉接口 {response.status_code}: {response.text[:300]}")
        body = response.json()
    else:
        request = urllib.request.Request(BASE + "/chat/completions", data=data,
                                         headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = json.loads(response.read().decode("utf-8"))
    text = body["choices"][0]["message"]["content"]
    report = _parse(text)
    report["pass"] = (float(report["score"]) >= min_score
                      and not report["has_text"] and not report["has_watermark"])
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("image", type=Path)
    parser.add_argument("--intent", required=True, help="what the image must show and leave room for")
    parser.add_argument("--min-score", type=float, default=7.0)
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--json", action="store_true", help="print the raw report only")
    parser.add_argument("--transparent", action="store_true",
                        help="asset is a transparent cutout; ignore the flattened white preview")
    args = parser.parse_args()
    intent = args.intent
    if args.transparent:
        intent += (" The asset is a transparent-background cutout; transparency is flattened "
                   "to white in this preview, so do not treat the white background as an error.")
    try:
        report = check(args.image, intent, args.min_score, args.timeout)
    except Exception as error:
        print(f"image_check: {error}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"score={report['score']} pass={report['pass']} "
              f"subject={report['subject_match']} clean={report['background_clean']} "
              f"text={report['has_text']} watermark={report['has_watermark']} "
              f"quiet={report['negative_space']}")
        for issue in report.get("issues") or []:
            print(f"- {issue}")
    return 0 if report["pass"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
