#!/usr/bin/env python3
"""
templafy_tool.py — Find and bind <<Placeholder>> tokens in .docx/.dotx files.

Usage:
  templafy_tool.py PATH -f [-r] [-d]
  templafy_tool.py PATH -b BINDING [-p PLACEHOLDER] [-r] [-d]
  templafy_tool.py PATH -f -b BINDING [-r] [-d]

Flags:
  PATH                  File or directory to process.
  -f, --find            List all unique <<...>> placeholders (bound and unbound).
  -b, --bind BINDING    Wrap every unbound <<...>> run in a Templafy SDT content
                        control with the given binding expression.
  -p, --placeholder     Limit -b to a specific placeholder name, e.g. ClientShortName.
  -r, --recursive       Walk directories recursively.
  -d, --dry-run         Show what would change without modifying any files.

Output file naming:
  Single file  → <stem>_bound<ext>  (written next to the source)
  Directory    → each file gets its own _bound copy
"""

import argparse
import copy
import os
import random
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ── Namespaces ────────────────────────────────────────────────────────────────

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

# Register namespaces we'll encounter so round-trips preserve prefix names.
_NAMESPACES = {
    "wpc":  "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx":   "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "mc":   "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o":    "urn:schemas-microsoft-com:office:office",
    "oel":  "http://schemas.microsoft.com/office/2019/extlst",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m":    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v":    "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp":   "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10":  "urn:schemas-microsoft-com:office:word",
    "w":    W,
    "w14":  W14,
    "w15":  "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex":"http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid":"http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16":  "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16sdtdh":"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16se":"http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg":  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi":  "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne":  "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps":  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
for _prefix, _uri in _NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)


def wtag(local: str) -> str:
    return f"{{{W}}}{local}"


# ── Parts of a docx that can hold visible text ────────────────────────────────

_TEXT_PARTS_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$"
)


def _text_parts(names: list[str]) -> list[str]:
    return [n for n in names if _TEXT_PARTS_RE.match(n)]


# ── Logical-text extraction (handles split runs) ──────────────────────────────

def _para_logical_text(para: ET.Element, include_sdt: bool = True) -> str:
    """
    Return the concatenated text of all <w:t> elements in direct run children
    of `para`, spanning across <w:proofErr> elements.

    include_sdt=True  → also include text inside child <w:sdt> sdtContent
                        (used by -f to list all placeholders including bound ones).
    include_sdt=False → skip child <w:sdt> elements entirely
                        (used by bind/dry-run counting to find only unbound text).
    """
    parts: list[str] = []
    for child in para:
        tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag_local == "r":
            for t in child:
                t_local = t.tag.split("}")[-1] if "}" in t.tag else t.tag
                if t_local == "t" and t.text:
                    parts.append(t.text)
        elif tag_local == "proofErr":
            pass  # transparent — don't add text, don't break the stream
        elif tag_local == "sdt" and include_sdt:
            for sc in child.iter(wtag("sdtContent")):
                for t in sc.iter(wtag("t")):
                    if t.text:
                        parts.append(t.text)
    return "".join(parts)


_PH_RE = re.compile(r"<<([^<>]+?)>>")


def _find_in_logical_text(text: str) -> set[str]:
    return set(_PH_RE.findall(text))


# ── Find placeholders in one XML part ────────────────────────────────────────

def _find_in_xml_bytes(data: bytes) -> set[str]:
    """
    Parse the XML and walk every <w:p> collecting placeholders from logical text.
    This handles both single-run tokens (<<Currency>>) and split-run tokens
    (<<, ClientShortName, >> in separate runs separated by proofErr elements).
    """
    found: set[str] = set()
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        # Fallback: regex on raw bytes for the simple un-escaped form
        raw = data.decode("utf-8", errors="replace")
        found |= _find_in_logical_text(raw)
        return found

    for para in root.iter(wtag("p")):
        text = _para_logical_text(para)
        found |= _find_in_logical_text(text)
    return found


def _get_binding_for_placeholder(path: Path, placeholder: str) -> str:
    """Return the Templafy binding expression for an already-bound placeholder."""
    import json
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for part in _text_parts(zf.namelist()):
                try:
                    root = ET.fromstring(zf.read(part))
                except ET.ParseError:
                    continue
                for sdt in root.iter(wtag("sdt")):
                    sdtPr = sdt.find(wtag("sdtPr"))
                    if sdtPr is None:
                        continue
                    tag_el = sdtPr.find(wtag("tag"))
                    if tag_el is None:
                        continue
                    tag_val = tag_el.get(wtag("val"), "")
                    if "templafy" not in tag_val:
                        continue
                    sdtContent = sdt.find(wtag("sdtContent"))
                    if sdtContent is None:
                        continue
                    text = "".join(t.text or "" for t in sdtContent.iter(wtag("t")))
                    if f"<<{placeholder}>>" in text:
                        try:
                            return json.loads(tag_val)["templafy"]["binding"]
                        except Exception:
                            return tag_val
    except zipfile.BadZipFile:
        pass
    return "(unknown)"


def _find_bound_in_xml_bytes(data: bytes) -> set[str]:
    """
    Return placeholder names that are already wrapped in a Templafy SDT,
    i.e. their text lives inside a <w:sdtContent> whose sibling <w:tag>
    contains a "templafy" JSON binding.
    """
    bound: set[str] = set()
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return bound
    for sdt in root.iter(wtag("sdt")):
        sdtPr = sdt.find(wtag("sdtPr"))
        if sdtPr is None:
            continue
        tag_el = sdtPr.find(wtag("tag"))
        if tag_el is None:
            continue
        tag_val = tag_el.get(wtag("val"), "")
        if "templafy" not in tag_val:
            continue
        # This SDT has a Templafy binding — collect placeholder names from its content
        sdtContent = sdt.find(wtag("sdtContent"))
        if sdtContent is None:
            continue
        text = "".join(t.text or "" for t in sdtContent.iter(wtag("t")))
        for m in _PH_RE.finditer(text):
            bound.add(m.group(1))
    return bound


def find_in_docx(path: Path) -> tuple[set[str], set[str]]:
    """
    Return (all_placeholders, bound_placeholders) found in a .docx.
    bound_placeholders is the subset that already live inside a Templafy SDT.
    """
    found: set[str] = set()
    bound: set[str] = set()
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for part in _text_parts(zf.namelist()):
                data = zf.read(part)
                found |= _find_in_xml_bytes(data)
                bound |= _find_bound_in_xml_bytes(data)
    except zipfile.BadZipFile as exc:
        print(f"  [WARN] Cannot read {path}: {exc}", file=sys.stderr)
    return found, bound


# ── SDT builder ───────────────────────────────────────────────────────────────

def _new_id() -> int:
    return random.randint(-(2**31), -1)


def _make_sdt(binding: str, runs: list[ET.Element]) -> ET.Element:
    """
    Build a Templafy <w:sdt> that wraps the given list of run elements.

    The tag JSON follows the exact structure Templafy generates:
      {"templafy":{"binding":"<binding>","visibility":"",
                   "disableUpdates":false,"removeAndKeepContent":false,"type":"text"}}
    """
    escaped = binding.replace('"', '\\"')
    tag_val = (
        '{"templafy":{"binding":"' + escaped + '",'
        '"visibility":"","disableUpdates":false,'
        '"removeAndKeepContent":false,"type":"text"}}'
    )

    sdt = ET.Element(wtag("sdt"))

    sdtPr = ET.SubElement(sdt, wtag("sdtPr"))
    alias = ET.SubElement(sdtPr, wtag("alias"))
    alias.set(wtag("val"), binding)
    tag_el = ET.SubElement(sdtPr, wtag("tag"))
    tag_el.set(wtag("val"), tag_val)
    id_el = ET.SubElement(sdtPr, wtag("id"))
    id_el.set(wtag("val"), str(_new_id()))

    ET.SubElement(sdt, wtag("sdtEndPr"))

    sc = ET.SubElement(sdt, wtag("sdtContent"))
    for r in runs:
        sc.append(copy.deepcopy(r))

    return sdt


# ── Paragraph-level binding ───────────────────────────────────────────────────

def _run_text(run: ET.Element) -> str:
    return "".join(t.text or "" for t in run.iter(wtag("t")))


def _is_sdt(elem: ET.Element) -> bool:
    return elem.tag == wtag("sdt")


def _split_streak(streak_text: str, streak_elems: list[ET.Element], binding: str,
                  target_ph: str | None) -> tuple[list[ET.Element], int]:
    """
    Given a streak of text (from consecutive runs/proofErr siblings) and the
    original elements, produce a new list of elements with every matching
    <<placeholder>> wrapped in its own SDT.  Handles multiple placeholders
    within the same streak (e.g. "<<StudyTitle>> | <<CRM>> | 6/25/2021").
    """
    # Get the run to use as a formatting template
    tmpl_run = next((e for e in streak_elems
                     if (e.tag.split("}")[-1] if "}" in e.tag else e.tag) == "r"), None)
    last_run = next((e for e in reversed(streak_elems)
                     if (e.tag.split("}")[-1] if "}" in e.tag else e.tag) == "r"), None)

    result: list[ET.Element] = []
    n_bound = 0
    cursor = 0

    for m in _PH_RE.finditer(streak_text):
        ph_name = m.group(1)
        ph_full = m.group(0)          # the full <<name>> token
        start, end = m.start(), m.end()

        # Skip if not our target
        if target_ph and ph_name != target_ph:
            continue

        # Emit any literal text before this placeholder
        if start > cursor:
            before = streak_text[cursor:start]
            br = ET.Element(wtag("r"))
            if tmpl_run is not None:
                rpr = tmpl_run.find(wtag("rPr"))
                if rpr is not None:
                    br.append(copy.deepcopy(rpr))
            t_el = ET.SubElement(br, wtag("t"))
            t_el.text = before
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            result.append(br)

        # Build the placeholder run, preserving formatting from the template
        ph_run = ET.Element(wtag("r"))
        if tmpl_run is not None:
            rpr = tmpl_run.find(wtag("rPr"))
            if rpr is not None:
                ph_run.append(copy.deepcopy(rpr))
        t_el = ET.SubElement(ph_run, wtag("t"))
        t_el.text = ph_full
        result.append(_make_sdt(binding, [ph_run]))
        n_bound += 1
        cursor = end

    if n_bound == 0:
        # Nothing matched — return original elements unchanged
        return streak_elems, 0

    # Emit any trailing literal text after the last placeholder
    if cursor < len(streak_text):
        after = streak_text[cursor:]
        ar = ET.Element(wtag("r"))
        ref = last_run if last_run is not None else (streak_elems[-1] if streak_elems else None)
        if ref is not None:
            rpr = ref.find(wtag("rPr"))
            if rpr is not None:
                ar.append(copy.deepcopy(rpr))
        t_el = ET.SubElement(ar, wtag("t"))
        t_el.text = after
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        result.append(ar)

    return result, n_bound


def bind_paragraph(
    para: ET.Element,
    binding: str,
    target_ph: str | None,
    dry_run: bool,
) -> int:
    """
    Scan a <w:p>'s direct children for unbound <<...>> tokens and wrap each
    in a Templafy SDT.  Handles:
      - tokens split across multiple runs by <w:proofErr> elements
      - multiple placeholders within the same run/streak
    Returns the number of SDT wraps applied (or that would be applied).
    """
    children = list(para)
    new_children: list[ET.Element] = []
    n_bound = 0
    i = 0

    while i < len(children):
        child = children[i]
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if local not in ("r", "proofErr"):
            new_children.append(child)
            i += 1
            continue

        # Collect a consecutive streak of run/proofErr siblings
        streak_elems: list[ET.Element] = []
        streak_text = ""
        j = i
        while j < len(children):
            c = children[j]
            cl = c.tag.split("}")[-1] if "}" in c.tag else c.tag
            if cl not in ("r", "proofErr"):
                break
            streak_elems.append(c)
            if cl == "r":
                streak_text += _run_text(c)
            j += 1

        if not _PH_RE.search(streak_text):
            new_children.extend(streak_elems)
            i = j
            continue

        expanded, count = _split_streak(streak_text, streak_elems, binding, target_ph)
        new_children.extend(expanded)
        n_bound += count
        i = j

    if n_bound > 0 and not dry_run:
        for child in list(para):
            para.remove(child)
        for child in new_children:
            para.append(child)

    return n_bound


# ── Process one XML part ──────────────────────────────────────────────────────

def bind_in_xml(
    data: bytes,
    binding: str,
    target_ph: str | None,
    dry_run: bool,
) -> tuple[bytes, int]:
    """
    Apply SDT bindings to all paragraphs in the XML part.
    Returns (new_bytes, count_bound).
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        print(f"  [WARN] XML parse error: {exc}", file=sys.stderr)
        return data, 0

    # Collect paragraphs that are NOT descendants of any <w:sdtContent>
    # (those are already inside bound SDTs and must not be touched)
    sdt_content_para_ids: set[int] = set()
    for sc in root.iter(wtag("sdtContent")):
        for p in sc.iter(wtag("p")):
            sdt_content_para_ids.add(id(p))

    total = 0
    for para in root.iter(wtag("p")):
        if id(para) in sdt_content_para_ids:
            continue
        total += bind_paragraph(para, binding, target_ph, dry_run)

    if total == 0 or dry_run:
        return data, total

    new_data = ET.tostring(root, encoding="unicode", xml_declaration=False).encode("utf-8")
    return new_data, total


# ── Process one docx ──────────────────────────────────────────────────────────

def bind_in_docx(
    src: Path,
    dst: Path,
    binding: str,
    target_ph: str | None,
    dry_run: bool,
) -> dict[str, int]:
    """
    Apply bindings across all text parts of a .docx.
    Returns {placeholder_name: count}.
    """
    totals: dict[str, int] = {}

    try:
        with zipfile.ZipFile(src, "r") as zf_in:
            names = zf_in.namelist()
            parts = _text_parts(names)

            if dry_run:
                for part in parts:
                    data = zf_in.read(part)
                    found = _find_in_xml_bytes(data)
                    targets = ({target_ph} & found) if target_ph else found
                    if not targets:
                        continue
                    root = ET.fromstring(data)
                    sdt_content_para_ids: set[int] = set()
                    for sc in root.iter(wtag("sdtContent")):
                        for p in sc.iter(wtag("p")):
                            sdt_content_para_ids.add(id(p))
                    for para in root.iter(wtag("p")):
                        if id(para) in sdt_content_para_ids:
                            continue
                        t = _para_logical_text(para, include_sdt=False)
                        for m in _PH_RE.finditer(t):
                            ph = m.group(1)
                            if ph in targets:
                                totals[ph] = totals.get(ph, 0) + 1
                return totals

            tmp = dst.with_suffix(".tmp")
            with zipfile.ZipFile(src, "r") as zf_in, \
                 zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for name in names:
                    data = zf_in.read(name)
                    if name in parts:
                        new_data, n = bind_in_xml(data, binding, target_ph, dry_run=False)
                        if n > 0:
                            # Attribute count to whichever placeholders we processed
                            found = _find_in_xml_bytes(data)
                            targets = ({target_ph} & found) if target_ph else found
                            for ph in targets:
                                totals[ph] = totals.get(ph, 0) + n
                            data = new_data
                    zf_out.writestr(name, data)
            tmp.replace(dst)

    except zipfile.BadZipFile as exc:
        print(f"  [ERROR] Cannot open {src}: {exc}", file=sys.stderr)

    return totals


# ── File / directory collection ───────────────────────────────────────────────

def collect_docx(path: Path, recursive: bool) -> list[Path]:
    if path.is_file():
        if path.suffix.lower() in (".docx", ".dotx"):
            return [path]
        print(f"[WARN] Not a .docx/.dotx file: {path}", file=sys.stderr)
        return []
    elif path.is_dir():
        globs = ["**/*.docx", "**/*.dotx"] if recursive else ["*.docx", "*.dotx"]
        files: list[Path] = []
        for g in globs:
            files.extend(path.glob(g))
        return sorted(set(files))
    else:
        print(f"[ERROR] Path does not exist: {path}", file=sys.stderr)
        sys.exit(1)


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        prog="templafy_tool.py",
        description=(
            "Inspect and add Templafy SDT bindings to <<Placeholder>> tokens "
            "in .docx / .dotx files."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("path", type=Path, help="File or directory to process.")
    parser.add_argument(
        "-f", "--find", action="store_true",
        help="List all unique <<...>> placeholders in the document(s).",
    )
    parser.add_argument(
        "-b", "--bind", metavar="BINDING",
        help="Templafy binding expression to apply, e.g. 'Form.MyField.Value'.",
    )
    parser.add_argument(
        "-p", "--placeholder", metavar="NAME",
        help=(
            "Limit --bind to a specific placeholder, e.g. 'ClientShortName'. "
            "If omitted, all placeholders receive the binding."
        ),
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="Process directories recursively.",
    )
    parser.add_argument(
        "-d", "--dry-run", action="store_true",
        help="Show what would change without writing any files.",
    )
    parser.add_argument(
        "-o", "--overwrite", action="store_true",
        help="Write the result back to the original file instead of creating a _bound copy.",
    )

    args = parser.parse_args()

    if not args.find and not args.bind:
        parser.error("Specify at least -f (find) or -b BINDING (bind).")

    files = collect_docx(args.path, args.recursive)
    if not files:
        print("No .docx/.dotx files found.")
        return

    dry_label = " [DRY RUN]" if args.dry_run else ""

    # ── Find ──────────────────────────────────────────────────────────────────
    if args.find:
        all_found: set[str] = set()
        for docx in files:
            found, bound = find_in_docx(docx)
            all_found |= found
            rel = docx.relative_to(args.path) if args.path.is_dir() else docx.name
            if found:
                print(f"\n{rel}{dry_label}")
                max_len = max(len(ph) for ph in found)
                for ph in sorted(found):
                    if ph in bound:
                        binding_val = _get_binding_for_placeholder(docx, ph)
                        padding = " " * (max_len - len(ph) + 2)
                        print(f"  <<{ph}>>{padding}→  {binding_val}")
                    else:
                        print(f"  <<{ph}>>")
            else:
                print(f"\n{rel}  (no placeholders found)")

        if len(files) > 1 and all_found:
            print(f"\n── Unique placeholders across all files ──")
            for ph in sorted(all_found):
                print(f"  <<{ph}>>")

    # ── Bind ──────────────────────────────────────────────────────────────────
    if args.bind:
        for docx in files:
            rel = docx.relative_to(args.path) if args.path.is_dir() else docx.name
            print(f"\n{rel}{dry_label}")

            if args.dry_run:
                totals = bind_in_docx(
                    docx, docx, args.bind, args.placeholder, dry_run=True
                )
                if totals:
                    for ph, n in sorted(totals.items()):
                        print(f"  Would bind <<{ph}>>  →  '{args.bind}'  "
                              f"({n} unbound occurrence(s))")
                else:
                    print("  No matching unbound placeholders found.")
            else:
                out = docx if args.overwrite else docx.with_name(docx.stem + "_bound" + docx.suffix)
                totals = bind_in_docx(
                    docx, out, args.bind, args.placeholder, dry_run=False
                )
                if totals:
                    for ph, n in sorted(totals.items()):
                        print(f"  Bound <<{ph}>>  →  '{args.bind}'  "
                              f"({n} occurrence(s))")
                    label = "(overwritten)" if args.overwrite else f"→  {out}"
                    print(f"  Written  {label}")
                else:
                    print("  No matching unbound placeholders found; no output written.")
                    if not args.overwrite and out.exists():
                        out.unlink()


if __name__ == "__main__":
    main()
