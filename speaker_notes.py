#!/usr/bin/env python3
"""Export speaker-script Markdown or embed it without rewriting slide content."""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import posixpath
import re
import tempfile
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from pptx import Presentation

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A}


def xml(data: bytes):
    """Parse local OOXML without resolving entities or fetching resources."""
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def dump(element) -> bytes:
    return ET.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def child(parent, namespace, tag, **attrs):
    return ET.SubElement(parent, f"{{{namespace}}}{tag}", **attrs)


def relationship(root, kind):
    return next((item for item in root if item.get("Type") == R + "/" + kind), None)


def add_relationship(root, kind, target):
    ids = {item.get("Id") for item in root}
    number = 1
    while f"rId{number}" in ids:
        number += 1
    return child(root, PR, "Relationship", Id=f"rId{number}", Type=R + "/" + kind, Target=target)


def slide_parts(entries: dict[str, bytes]) -> list[str]:
    """Resolve presentation order through relationships, including reordered slides."""
    targets = {item.get("Id"): item.get("Target") for item in xml(entries["ppt/_rels/presentation.xml.rels"])}
    return [resolve_part("ppt/presentation.xml", targets[item.get(f"{{{R}}}id")])
            for item in xml(entries["ppt/presentation.xml"]).find(f"{{{P}}}sldIdLst")]


def resolve_part(source: str, target: str) -> str:
    """Resolve relative or package-absolute OOXML relationship targets."""
    return posixpath.normpath(target.lstrip("/") if target.startswith("/")
                             else posixpath.join(posixpath.dirname(source), target))


def parse_script(path: Path, count: int) -> tuple[dict[int, str], int]:
    """Require one nonempty P-numbered section per slide; reject unfinished drafts."""
    raw = re.sub(r"<!--.*?-->", "", path.read_text(encoding="utf-8"), flags=re.S)
    sections, budget = {}, 0
    pattern = r"^## P(\d+)\s*[｜|:：]\s*([^\n]+)\n(.*?)(?=^## |\Z)"
    for match in re.finditer(pattern, raw, re.M | re.S):
        number = int(match[1])
        title = match[2].strip()
        body = re.sub(r"<!--.*?-->", "", match[3], flags=re.S).strip()
        if number in sections or not body or "TODO speaker script" in body:
            raise ValueError(f"P{number}: duplicate, empty or unfinished speaker script")
        timing = re.search(r"[（(](\d+)\s*(?:秒|s|seconds)[）)]", title)
        budget += int(timing[1]) if timing else 0
        sections[number] = f"P{number}｜{title}\n\n{body}"
    if set(sections) != set(range(1, count + 1)):
        raise ValueError(f"script must cover P1 through P{count} exactly")
    return sections, budget


def notes_document(text: str):
    root = ET.Element(f"{{{P}}}notes", nsmap={"p": P, "a": A, "r": R})
    tree = child(child(root, P, "cSld"), P, "spTree")
    nv = child(tree, P, "nvGrpSpPr")
    child(nv, P, "cNvPr", id="1", name="")
    child(nv, P, "cNvGrpSpPr")
    child(nv, P, "nvPr")
    child(tree, P, "grpSpPr")
    sp = child(tree, P, "sp")
    nv = child(sp, P, "nvSpPr")
    child(nv, P, "cNvPr", id="2", name="Notes")
    child(nv, P, "cNvSpPr")
    child(child(nv, P, "nvPr"), P, "ph", type="body", idx="1")
    child(sp, P, "spPr")
    tx = child(sp, P, "txBody")
    child(tx, A, "bodyPr")
    child(tx, A, "lstStyle")
    replace_text(tx, text)
    child(child(root, P, "clrMapOvr"), A, "masterClrMapping")
    return root


def replace_text(tx, text: str):
    for paragraph in list(tx):
        if paragraph.tag == f"{{{A}}}p":
            tx.remove(paragraph)
    for line in text.split("\n"):
        paragraph = child(tx, A, "p")
        run = child(paragraph, A, "r")
        child(run, A, "rPr", lang="zh-CN", sz="1100")
        child(run, A, "t").text = line


def semantic_part(data: bytes) -> bytes:
    """Remove only notes declarations for preservation comparisons."""
    root = xml(data)
    for item in list(root):
        if (item.tag == f"{{{P}}}notesMasterIdLst"
                or item.get("Type") in (R + "/notesSlide", R + "/notesMaster")
                or item.get("ContentType", "").endswith((".notesSlide+xml", ".notesMaster+xml"))):
            root.remove(item)
    return ET.tostring(root, method="c14n")


def preservation_changes(before: dict[str, bytes], after: dict[str, bytes]) -> list[str]:
    """Return changes outside notes parts and their required declarations."""
    problems = []
    for name in set(before) | set(after):
        if name.startswith("ppt/notesSlides/") or name.startswith("ppt/notesMasters/"):
            continue
        old, new = before.get(name), after.get(name)
        if old == new:
            continue
        declarations = name in ("[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels")
        slide_rels = name.startswith("ppt/slides/_rels/") and name.endswith(".rels")
        if declarations or slide_rels:
            empty = dump(ET.Element(f"{{{PR}}}Relationships"))
            if new is not None and semantic_part(old or empty) == semantic_part(new):
                continue
        problems.append(name)
    return sorted(problems)


def embed(pptx: Path, script: Path, output: Path) -> dict:
    """Atomically write complete notes after verifying all other content is preserved."""
    with ZipFile(pptx) as source:
        infos = source.infolist()
        if len({item.filename for item in infos}) != len(infos) or source.testzip():
            raise ValueError("invalid or duplicate ZIP entries")
        before = {item.filename: source.read(item) for item in infos}
    entries = before.copy()
    parts = slide_parts(entries)
    sections, budget = parse_script(script, len(parts))
    content_types = xml(entries["[Content_Types].xml"])
    for page, slide in enumerate(parts, 1):
        relfile = posixpath.join(posixpath.dirname(slide), "_rels", posixpath.basename(slide) + ".rels")
        rels = xml(entries[relfile]) if relfile in entries else ET.Element(f"{{{PR}}}Relationships", nsmap={None: PR})
        rel = relationship(rels, "notesSlide")
        if rel is None:
            number = 1
            while f"ppt/notesSlides/notesSlide{number}.xml" in entries:
                number += 1
            note_file = f"ppt/notesSlides/notesSlide{number}.xml"
            add_relationship(rels, "notesSlide", f"../notesSlides/notesSlide{number}.xml")
            entries[relfile] = dump(rels)
            root = notes_document(sections[page])
            note_rels = ET.Element(f"{{{PR}}}Relationships", nsmap={None: PR})
            add_relationship(note_rels, "slide", posixpath.relpath(slide, "ppt/notesSlides"))
            # A notes master is optional. Reuse it when the deck already has one.
            masters = sorted(n for n in entries if re.fullmatch(r"ppt/notesMasters/notesMaster\d+\.xml", n))
            if masters:
                add_relationship(note_rels, "notesMaster", posixpath.relpath(masters[0], "ppt/notesSlides"))
            entries[f"ppt/notesSlides/_rels/notesSlide{number}.xml.rels"] = dump(note_rels)
            child(content_types, CT, "Override", PartName="/" + note_file,
                  ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
        else:
            note_file = resolve_part(slide, rel.get("Target"))
            root = xml(entries[note_file])
            bodies = root.xpath('.//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]/p:txBody', namespaces=NS)
            if len(bodies) != 1:
                raise ValueError(f"P{page}: expected one notes body placeholder")
            replace_text(bodies[0], sections[page])
        entries[note_file] = dump(root)
    entries["[Content_Types].xml"] = dump(content_types)
    changes = preservation_changes(before, entries)
    if changes:
        raise ValueError(f"notes update changed presentation content: {changes}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".pptx", delete=False) as stream:
        temporary = Path(stream.name)
    try:
        with ZipFile(temporary, "w") as archive:
            for info in infos:
                archive.writestr(copy.copy(info), entries[info.filename])
            for name in entries.keys() - before.keys():
                archive.writestr(name, entries[name])
        result = Presentation(temporary)
        for page, slide in enumerate(result.slides, 1):
            if not slide.has_notes_slide or slide.notes_slide.notes_text_frame.text != sections[page]:
                raise ValueError(f"P{page}: notes readback mismatch")
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)
    return {"pages": len(parts), "notes_written": len(sections), "budget_seconds": budget,
            "content_preserved": True, "sha256": hashlib.sha256(output.read_bytes()).hexdigest()}


def export_script(pptx: Path, output: Path):
    """Create a draft from slide text and existing notes, with no invented achievements."""
    if pptx.resolve() == output.resolve():
        raise ValueError("Markdown output must be separate from the source deck")
    prs = Presentation(pptx)
    lines = ["# Speaker script", "", "<!-- Set per-page timing, rehearse and verify evidence before embedding. -->", ""]
    for page, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title.text if slide.shapes.title else f"Slide {page}"
        title = " ".join(title.split())
        body = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        text = "\n".join(s.text for s in slide.shapes if s.has_text_frame).replace("--", "- -")
        lines += [f"## P{page}｜{title}", "", f"<!-- Slide content:\n{text}\n-->", "", body.strip() or "[TODO speaker script]", ""]
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    export = commands.add_parser("export")
    export.add_argument("pptx", type=Path)
    export.add_argument("markdown", type=Path)
    write = commands.add_parser("embed")
    write.add_argument("pptx", type=Path)
    write.add_argument("markdown", type=Path)
    write.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.command == "export":
        export_script(args.pptx, args.markdown)
    else:
        print(json.dumps(embed(args.pptx, args.markdown, args.output), ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
