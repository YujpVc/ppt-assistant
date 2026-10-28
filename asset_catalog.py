#!/usr/bin/env python3
"""Create and verify a deterministic inventory of presentation source assets."""
from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import tempfile
from pathlib import Path


KIND_BY_SUFFIX = {
    ".pptx": "presentation", ".ppt": "presentation", ".pdf": "document",
    ".docx": "document", ".xlsx": "data", ".csv": "data", ".json": "data",
    ".md": "text", ".txt": "text", ".html": "text", ".css": "text",
    ".png": "image", ".jpg": "image", ".jpeg": "image", ".gif": "image",
    ".webp": "image", ".mp4": "video", ".mov": "video", ".avi": "video",
    ".mkv": "video", ".wav": "audio", ".mp3": "audio",
}


def digest(path: Path) -> str:
    """Return the SHA256 of a regular file."""
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def relative_name(value) -> str:
    """Validate a canonical, non-hidden POSIX path from a JSON manifest."""
    if (not isinstance(value, str) or not value or "\\" in value
            or Path(value).is_absolute() or Path(value).as_posix() != value
            or any(part.startswith(".") for part in value.split("/"))):
        raise ValueError("manifest contains an invalid relative path")
    return value


def _entry(root: Path, path: Path) -> dict:
    relative = path.relative_to(root).as_posix()
    suffix = path.suffix.lower()
    return {
        "path": relative,
        "bytes": path.stat().st_size,
        "sha256": digest(path),
        "kind": KIND_BY_SUFFIX.get(suffix, "other"),
        "mime": mimetypes.guess_type(path.name)[0] or "application/octet-stream",
    }


def scan(root: Path, *, exclude: set[str] | None = None) -> dict:
    """Inventory regular files below ``root`` in stable relative-path order."""
    root = root.resolve()
    if not root.is_dir():
        raise ValueError("asset root must be a directory")
    excluded = {relative_name(item) for item in (exclude or set())}
    files = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.is_symlink():
            continue
        relative = path.relative_to(root).as_posix()
        if relative in excluded or any(part.startswith(".") for part in path.relative_to(root).parts):
            continue
        files.append(_entry(root, path))
    return {"version": 1, "files": files}


def write_manifest(root: Path, output: Path, markdown: Path | None = None) -> dict:
    """Write JSON and optional Markdown inventory without including the outputs."""
    root = root.resolve()
    output = output.resolve()
    if markdown and markdown.resolve() == output:
        raise ValueError("JSON and Markdown outputs must be separate")
    excluded = {output.relative_to(root).as_posix()} if output.is_relative_to(root) else set()
    if markdown and markdown.resolve().is_relative_to(root):
        excluded.add(markdown.resolve().relative_to(root).as_posix())
    manifest = scan(root, exclude=excluded)
    manifest["excluded"] = sorted(excluded)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=output.parent, delete=False) as stream:
        temporary = Path(stream.name)
        json.dump(manifest, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    temporary.replace(output)
    if markdown:
        groups: dict[str, int] = {}
        for item in manifest["files"]:
            groups[item["kind"]] = groups.get(item["kind"], 0) + 1
        lines = ["# Asset inventory", "", f"Files: {len(manifest['files'])}", "",
                 "| Kind | Count |", "| --- | ---: |"]
        lines.extend(f"| {kind} | {groups[kind]} |" for kind in sorted(groups))
        lines.extend(["", "| Path | Kind | Bytes | SHA256 |", "| --- | --- | ---: | --- |"])
        for item in manifest["files"]:
            name = item["path"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            name = name.replace("|", "&#124;").replace("`", "&#96;").replace("\n", "&#10;").replace("\r", "&#13;")
            lines.append(f"| {name} | {item['kind']} | {item['bytes']} | `{item['sha256']}` |")
        markdown.parent.mkdir(parents=True, exist_ok=True)
        markdown.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return manifest


def verify(root: Path, manifest_path: Path, *, strict: bool = False) -> dict:
    """Verify listed paths and hashes; strict mode also rejects unlisted files."""
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or manifest.get("version") != 1 or not isinstance(manifest.get("files"), list):
        raise ValueError("unsupported asset manifest")
    root = root.resolve()
    if not root.is_dir():
        raise ValueError("asset root must be a directory")
    excluded_values = manifest.get("excluded", [])
    if not isinstance(excluded_values, list):
        raise ValueError("excluded must be a list of relative paths")
    excluded = {relative_name(value) for value in excluded_values}
    listed = set()
    for item in manifest["files"]:
        if not isinstance(item, dict):
            raise ValueError("asset entry must be an object")
        path = relative_name(item.get("path"))
        if path in listed or path in excluded:
            raise ValueError("manifest contains an invalid or duplicate path")
        listed.add(path)
        unresolved = root / path
        candidate = unresolved.resolve()
        if (not candidate.is_relative_to(root) or not candidate.is_file()
                or any(parent.is_symlink() for parent in [unresolved, *unresolved.parents] if parent != root)):
            raise ValueError(f"asset is missing or leaves root: {path}")
        if digest(candidate) != item.get("sha256"):
            raise ValueError(f"SHA256 mismatch: {path}")
    if strict:
        if manifest_path.resolve().is_relative_to(root):
            excluded.add(manifest_path.resolve().relative_to(root).as_posix())
        actual = {item["path"] for item in scan(root, exclude=excluded)["files"]}
        if actual != listed:
            raise ValueError("strict inventory differs from the asset root")
    return {"verified": len(listed), "strict": strict}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    writer = commands.add_parser("scan")
    writer.add_argument("root", type=Path)
    writer.add_argument("--output", type=Path, required=True)
    writer.add_argument("--markdown", type=Path)
    checker = commands.add_parser("verify")
    checker.add_argument("root", type=Path)
    checker.add_argument("manifest", type=Path)
    checker.add_argument("--strict", action="store_true")
    args = parser.parse_args()
    result = write_manifest(args.root, args.output, args.markdown) if args.command == "scan" \
        else verify(args.root, args.manifest, strict=args.strict)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
