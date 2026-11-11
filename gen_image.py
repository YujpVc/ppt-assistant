#!/usr/bin/env python3
"""Generate deck imagery with an OpenAI-compatible Images API (gpt-image).

The PPT assistant owns its images: a hero/section picture is generated for the
deck's own topic instead of borrowing a stock photo that contradicts the copy.
The tool is a thin, dependency-free client (stdlib only) so it also runs before
``requirements.txt`` is installed.

Credentials live outside the repository. Resolution order:

  1. ``--config PATH``
  2. ``$PPT_IMAGE_CONFIG``
  3. ``~/.config/ppt-assistant/image.json``   (user-owned, never shipped)
  4. environment: ``UUAPI_API_KEY`` / ``UUAPI_BASE_URL``

No API key is bundled with this package. Copy ``image-config.example.json`` to
``~/.config/ppt-assistant/image.json`` and fill in your own credentials.

A config is ``{"api_base": "https://uuapi.io/v1", "api_key": "..."}``. The key is
never printed, copied into the deck, or written into the metadata sidecar.

Examples:
    python3 gen_image.py --prompt-file hero.txt --out assets/hero.png \\
        --size landscape --model gpt-image-2

    python3 gen_image.py --prompt "..." --out assets/lead.png \\
        --ref product.png --size 1024x1024
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import mimetypes
import os
import sys
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path
from typing import Any

ENGINE = Path(__file__).resolve().parent
MAX_REFERENCE_BYTES = 20 * 1024 * 1024
SIZE_ALIASES = {
    "landscape": "1536x1024",
    "16:9": "1536x1024",
    "wide": "1536x1024",
    "square": "1024x1024",
    "1:1": "1024x1024",
    "portrait": "1024x1536",
    "3:4": "1024x1536",
}
DEFAULT_MODELS = ("gpt-image-2", "gpt-image-2.5-flare")


class ImageError(RuntimeError):
    pass


def install_proxy() -> str | None:
    """Route this one process through a SOCKS proxy when asked.

    The workspace can lose direct egress to the image host while an SSH tunnel
    is available (``ssh -D``). PySocks is optional: when it is missing the
    command keeps using the normal route instead of failing unexpectedly.
    """
    proxy = (os.environ.get("PPT_IMAGE_PROXY") or os.environ.get("ALL_PROXY") or "")
    if not proxy.startswith("socks"):
        return None
    try:
        import socket
        import socks
        from urllib.parse import urlsplit
    except ImportError:
        return None
    parsed = urlsplit(proxy)
    if not parsed.hostname:
        return None
    socks.set_default_proxy(socks.SOCKS5, parsed.hostname, parsed.port or 1080,
                            rdns=proxy.startswith("socks5h"))
    socket.socket = socks.socksocket
    return proxy


def config_candidates() -> list[Path]:
    home = Path.home()
    return [
        home / ".config/ppt-assistant/image.json",
    ]


def load_config(explicit: str | None) -> tuple[str, str, Path | None]:
    if explicit:
        paths = [Path(explicit).expanduser()]
    elif os.environ.get("PPT_IMAGE_CONFIG"):
        paths = [Path(os.environ["PPT_IMAGE_CONFIG"]).expanduser()]
    else:
        paths = config_candidates()
    for path in paths:
        if path.is_file():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError) as error:
                raise ImageError(f"图片配置无法读取: {path}: {error}") from error
            base = str(data.get("api_base") or data.get("base_url") or "").strip()
            key = str(data.get("api_key") or data.get("key") or "").strip()
            if base and key:
                return base.rstrip("/"), key, path
            raise ImageError(f"图片配置缺少 api_base/api_key: {path}")
    base = os.environ.get("UUAPI_BASE_URL", "").strip()
    key = os.environ.get("UUAPI_API_KEY", "").strip()
    if base and key:
        return base.rstrip("/"), key, None
    raise ImageError(
        "未找到生图配置。请复制 image-config.example.json 到 "
        "~/.config/ppt-assistant/image.json 并填入自己的 api_base/api_key，"
        "或用 PPT_IMAGE_CONFIG 指向自己的配置文件，或导出 UUAPI_API_KEY。")


def ssl_context():
    # The upstream is HTTPS; keep verification on. An explicit opt-out exists for
    # the rare local TLS-inspection setup, matching the image-web tool.
    if os.environ.get("PPT_IMAGE_INSECURE_TLS") == "1":
        import ssl
        return ssl._create_unverified_context()
    return None


def _proxy_map():
    proxy = (os.environ.get("PPT_IMAGE_PROXY") or os.environ.get("HTTPS_PROXY")
             or os.environ.get("https_proxy"))
    if proxy and proxy.startswith(("socks", "http")):
        return {"http": proxy, "https": proxy}
    return None


def api_call(url: str, key: str, body: bytes, content_type: str,
             timeout: int) -> dict[str, Any]:
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": content_type,
        "User-Agent": "ppt-assistant-image/1.0",
    }
    try:
        import requests
    except ImportError:
        requests = None
    if requests is not None:
        try:
            response = requests.post(url, data=body, headers=headers,
                                     timeout=timeout, proxies=_proxy_map())
        except requests.RequestException as error:
            raise ImageError(f"无法连接图片服务: {error}") from error
        if response.status_code >= 400:
            detail = response.text
            try:
                detail = response.json().get("error", detail)
            except ValueError:
                pass
            raise ImageError(f"接口返回 {response.status_code}: {str(detail)[:600]}")
        try:
            return response.json()
        except ValueError as error:
            raise ImageError("接口返回的不是 JSON") from error

    request = urllib.request.Request(url, data=body, method="POST", headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=ssl_context()) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", "replace")
        try:
            detail = json.loads(detail).get("error", detail)
        except json.JSONDecodeError:
            pass
        raise ImageError(f"接口返回 {error.code}: {str(detail)[:600]}") from error
    except urllib.error.URLError as error:
        raise ImageError(f"无法连接图片服务: {error.reason}") from error
    except TimeoutError as error:
        raise ImageError(f"图片服务超时（{timeout}s）") from error


def download(url: str, timeout: int = 120) -> bytes:
    try:
        import requests
    except ImportError:
        requests = None
    if requests is not None:
        try:
            response = requests.get(url, timeout=timeout, proxies=_proxy_map(),
                                    headers={"User-Agent": "ppt-assistant-image/1.0"})
        except requests.RequestException as error:
            raise ImageError(f"图片下载失败: {error}") from error
        if response.status_code >= 400:
            raise ImageError(f"图片下载失败: HTTP {response.status_code}")
        return response.content
    request = urllib.request.Request(url, headers={"User-Agent": "ppt-assistant-image/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=ssl_context()) as response:
            return response.read()
    except (urllib.error.HTTPError, urllib.error.URLError) as error:
        raise ImageError(f"图片下载失败: {error}") from error


def multipart(fields: dict[str, str], files: list[tuple[str, str, str, bytes]]) -> tuple[bytes, str]:
    boundary = f"----ppt-image-{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend((
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            str(value).encode("utf-8"),
            b"\r\n",
        ))
    for field, filename, mime_type, data in files:
        chunks.extend((
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'.encode(),
            f"Content-Type: {mime_type}\r\n\r\n".encode(),
            data,
            b"\r\n",
        ))
    chunks.append(f"--{boundary}--\r\n".encode())
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def read_reference(path: Path) -> tuple[str, str, bytes]:
    if not path.is_file():
        raise ImageError(f"参考图不存在: {path}")
    data = path.read_bytes()
    if not data or len(data) > MAX_REFERENCE_BYTES:
        raise ImageError(f"参考图必须大于 0 且不超过 20 MB: {path}")
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    if not mime.startswith("image/"):
        raise ImageError(f"参考图必须是图片: {path}")
    return path.name, mime, data


def save_payload(payload: bytes, target: Path) -> str:
    if not payload:
        raise ImageError("接口返回了空图片")
    if payload[:8] == b"\x89PNG\r\n\x1a\n":
        suffix = ".png"
    elif payload[:3] == b"\xff\xd8\xff":
        suffix = ".jpg"
    else:
        suffix = target.suffix or ".png"
    final = target if target.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"} else target.with_suffix(suffix)
    final.parent.mkdir(parents=True, exist_ok=True)
    final.write_bytes(payload)
    return str(final)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    prompt_group = parser.add_mutually_exclusive_group(required=True)
    prompt_group.add_argument("--prompt", help="prompt text")
    prompt_group.add_argument("--prompt-file", help="read prompt text from a file")
    parser.add_argument("--out", required=True, help="output image path")
    parser.add_argument("--size", default="landscape",
                        help="landscape(1536x1024) | square(1024x1024) | portrait(1024x1536) | WxH")
    parser.add_argument("--model", default=os.environ.get("PPT_IMAGE_MODEL", DEFAULT_MODELS[0]))
    parser.add_argument("--ref", action="append", default=[], help="reference image (repeatable, max 4)")
    parser.add_argument("--mask", help="optional mask image for edits")
    parser.add_argument("--transparent", action="store_true",
                        help="request a transparent background (PNG cutout)")
    parser.add_argument("--n", type=int, default=1, help="number of images (1-4)")
    parser.add_argument("--timeout", type=int, default=480)
    parser.add_argument("--retries", type=int, default=1)
    parser.add_argument("--config", help="path to image config JSON")
    parser.add_argument("--no-meta", action="store_true", help="skip the .json sidecar")
    parser.add_argument("--dry-run", action="store_true", help="validate inputs and print the request only")
    args = parser.parse_args()
    install_proxy()

    if args.prompt_file:
        prompt = Path(args.prompt_file).expanduser().read_text(encoding="utf-8").strip()
    else:
        prompt = args.prompt.strip()
    if not prompt:
        raise ImageError("提示词不能为空")
    if not 1 <= args.n <= 4:
        raise ImageError("--n 必须在 1–4 之间")
    if len(args.ref) > 4:
        raise ImageError("参考图最多 4 张")

    size = SIZE_ALIASES.get(args.size.lower(), args.size)
    if size not in {"1024x1024", "1024x1536", "1536x1024"}:
        raise ImageError(f"当前账户支持 1024x1024 / 1024x1536 / 1536x1024，收到 {args.size}")

    base, key, config_path = load_config(args.config)
    mode = "edit" if args.ref else "generate"
    if args.mask and mode != "edit":
        raise ImageError("--mask 只能与 --ref 一起使用")

    if args.dry_run:
        print(json.dumps({
            "mode": mode, "model": args.model, "size": size, "n": args.n,
            "config": str(config_path) if config_path else None,
            "endpoint": f"{base}/images/{'edits' if mode == 'edit' else 'generations'}",
            "prompt_chars": len(prompt), "references": [str(Path(r).expanduser()) for r in args.ref],
            "out": str(Path(args.out).expanduser()),
        }, ensure_ascii=False, indent=2))
        return 0

    started = time.time()
    last_error: ImageError | None = None
    for attempt in range(args.retries + 1):
        try:
            if mode == "edit":
                files = [("image", *read_reference(Path(path).expanduser())) for path in args.ref]
                if args.mask:
                    files.append(("mask", *read_reference(Path(args.mask).expanduser())))
                fields = {"model": args.model, "prompt": prompt, "size": size,
                          "n": str(args.n)}
                if args.transparent:
                    fields.update({"background": "transparent", "output_format": "png"})
                body, content_type = multipart(fields, files)
                result = api_call(f"{base}/images/edits", key, body, content_type, args.timeout)
            else:
                payload = {"model": args.model, "prompt": prompt, "size": size,
                           "n": args.n, "response_format": "url"}
                if args.transparent:
                    payload.update({"background": "transparent", "output_format": "png"})
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                result = api_call(f"{base}/images/generations", key, body,
                                  "application/json", args.timeout)
            break
        except ImageError as error:
            last_error = error
            if attempt < args.retries:
                time.sleep(3 + attempt * 4)
                continue
            raise
    else:  # pragma: no cover - the loop always breaks or raises
        raise last_error or ImageError("图片生成失败")

    items = result.get("data") or []
    if not items:
        raise ImageError(f"接口未返回图片: {json.dumps(result, ensure_ascii=False)[:400]}")

    target = Path(args.out).expanduser()
    saved: list[str] = []
    for index, item in enumerate(items):
        suffix = "" if len(items) == 1 else f"-{index + 1}"
        item_target = target.with_name(f"{target.stem}{suffix}{target.suffix or '.png'}")
        if item.get("url"):
            saved.append(save_payload(download(item["url"]), item_target))
        elif item.get("b64_json"):
            try:
                payload = base64.b64decode(item["b64_json"], validate=True)
            except (ValueError, TypeError) as error:
                raise ImageError("接口返回的 base64 图片无效") from error
            saved.append(save_payload(payload, item_target))
        else:
            raise ImageError("接口返回项既没有 url 也没有 b64_json")

    elapsed = round(time.time() - started, 1)
    summary = {
        "ok": True,
        "mode": mode,
        "model": result.get("model") or args.model,
        "requested_model": args.model,
        "size": size,
        "files": saved,
        "seconds": elapsed,
        "prompt_sha256": hashlib.sha256(prompt.encode("utf-8")).hexdigest()[:16],
        "prompt_chars": len(prompt),
        "transparent": bool(args.transparent),
    }
    if not args.no_meta:
        meta_path = Path(saved[0]).with_suffix(".json")
        meta_path.write_text(json.dumps({
            **summary,
            "prompt": prompt,
            "references": [str(Path(r).expanduser()) for r in args.ref],
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        summary["meta"] = str(meta_path)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ImageError as error:
        print(f"gen_image: {error}", file=sys.stderr)
        raise SystemExit(2)
