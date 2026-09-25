#!/usr/bin/env python3
"""Add OOXML custom colors to Word and PowerPoint packages.

The script opens a temporary text file using the editor named by EDITOR.
Enter one color per line, for example:

    #dd0000 a kind of red
    #00aa00 some green

Blank lines and lines beginning with '# ' are ignored. The custom-color list
is rounded up to a complete group of ten, with white entries named "blank"
used as fillers. At least ten entries are always written.

Supported package types:
    .docx, .docm, .dotx, .dotm
    .pptx, .pptm, .potx, .potm

By default the source package is replaced in place. Use --backup to create a
side-by-side backup before replacement.
"""

from __future__ import annotations

import argparse
import copy
import os
import re
import shlex
import shutil
import stat
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
QNAME = "{%s}%s" % (A_NS, "%s")

# The theme location is the same for the related Office package types.
THEME_BY_EXTENSION = {
    ".docx": "word/theme/theme1.xml",
    ".docm": "word/theme/theme1.xml",
    ".dotx": "word/theme/theme1.xml",
    ".dotm": "word/theme/theme1.xml",
    ".pptx": "ppt/theme/theme1.xml",
    ".pptm": "ppt/theme/theme1.xml",
    ".potx": "ppt/theme/theme1.xml",
    ".potm": "ppt/theme/theme1.xml",
}

COLOR_RE = re.compile(
    r"^\s*#(?P<hex>[0-9a-fA-F]{6})\s+(?P<name>.*?)\s*$"
)
MAX_CUSTOM_COLORS = 50
FILLER_COLOR = ("FFFFFF", "blank")
EDITOR_TEMPLATE = """# Enter one color per line.
# Format: #RRGGBB color name
# Example: #dd0000 a kind of red
# Blank lines and lines beginning with '# ' are ignored.

"""

# Register the namespace used by theme XML so new elements serialize as a:a.
ET.register_namespace("a", A_NS)

DEBUG = False


class ColorInputError(ValueError):
    """Raised when the editor contents do not contain a valid color list."""


def debug(message: str) -> None:
    if DEBUG:
        print(f"[debug] {message}")


def get_theme_zip_path(file_path: str | os.PathLike[str]) -> Optional[str]:
    """Return the internal theme1.xml path for a supported Office package."""
    extension = Path(file_path).suffix.lower()
    return THEME_BY_EXTENSION.get(extension)


def collect_office_files(input_path: Path, recursive: bool) -> List[Path]:
    """Collect supported Office files from a file or directory input."""
    if input_path.is_file():
        if get_theme_zip_path(input_path) is None:
            raise ValueError(
                f"Unsupported file extension: {input_path.suffix or '(none)'}"
            )
        return [input_path]

    if not input_path.is_dir():
        raise FileNotFoundError(f"Input path does not exist: {input_path}")

    files: List[Path] = []
    if recursive:
        for root, dirs, filenames in os.walk(input_path):
            # Match the behavior of the supplied extraction/replacement scripts.
            dirs[:] = sorted(d for d in dirs if not d.startswith("_"))
            for filename in sorted(filenames):
                candidate = Path(root) / filename
                if filename.startswith("~$"):
                    continue
                if get_theme_zip_path(candidate) is not None:
                    files.append(candidate)
    else:
        for candidate in sorted(input_path.iterdir()):
            if (
                candidate.is_file()
                and not candidate.name.startswith("~$")
                and get_theme_zip_path(candidate) is not None
            ):
                files.append(candidate)

    return files


def launch_editor() -> str:
    """Open the temporary color-list file using the user's EDITOR."""
    editor = os.environ.get("EDITOR")
    if not editor:
        raise RuntimeError(
            "The EDITOR environment variable is not set. "
            "For example: export EDITOR=nano"
        )

    try:
        editor_command = shlex.split(editor)
    except ValueError as exc:
        raise RuntimeError(f"Could not parse EDITOR={editor!r}: {exc}") from exc

    if not editor_command:
        raise RuntimeError("The EDITOR environment variable is empty.")

    fd, temp_path = tempfile.mkstemp(
        prefix="office-custom-colors-",
        suffix=".txt",
        text=True,
    )
    os.close(fd)

    try:
        with open(temp_path, "w", encoding="utf-8", newline="\n") as color_file:
            color_file.write(EDITOR_TEMPLATE)

        debug(f"Launching editor: {' '.join(editor_command)} {temp_path}")
        subprocess.run([*editor_command, temp_path], check=True)

        with open(temp_path, "r", encoding="utf-8-sig") as color_file:
            return color_file.read()
    except FileNotFoundError as exc:
        raise RuntimeError(
            f"Editor command not found: {editor_command[0]!r}. "
            "Check the EDITOR environment variable."
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            f"The editor exited with status {exc.returncode}; no Office files were changed."
        ) from exc
    finally:
        try:
            os.unlink(temp_path)
        except FileNotFoundError:
            pass


def parse_color_list(text: str) -> Tuple[List[Tuple[str, str]], int]:
    """Parse editor text and return (colors, number_of_fillers)."""
    colors: List[Tuple[str, str]] = []
    errors: List[str] = []

    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.lstrip("\ufeff") if line_number == 1 else raw_line
        stripped = line.strip()

        if not stripped or stripped.startswith("# "):
            continue

        match = COLOR_RE.match(line)
        if not match:
            errors.append(
                f"line {line_number}: expected '#RRGGBB color name', got {raw_line!r}"
            )
            continue

        color_hex = match.group("hex").upper()
        color_name = match.group("name").strip()
        if not color_name:
            errors.append(f"line {line_number}: the color name cannot be empty")
            continue

        colors.append((color_hex, color_name))

    if errors:
        raise ColorInputError("\n".join(errors))

    if len(colors) > MAX_CUSTOM_COLORS:
        raise ColorInputError(
            f"{len(colors)} custom colors were entered, but Office supports at most "
            f"{MAX_CUSTOM_COLORS} custom colors in this palette."
        )

    # Complete groups of ten: 1 -> 10, 14 -> 20, 20 -> 20.
    target_count = max(10, ((len(colors) + 9) // 10) * 10)
    filler_count = target_count - len(colors)
    colors.extend([FILLER_COLOR] * filler_count)
    return colors, filler_count


def qname(local_name: str) -> str:
    return QNAME % local_name


def build_custom_color_list(colors: Sequence[Tuple[str, str]]) -> ET.Element:
    """Create an a:custClrLst element from (hex RGB, display name) pairs."""
    custom_list = ET.Element(qname("custClrLst"))
    for color_hex, color_name in colors:
        custom_color = ET.SubElement(
            custom_list,
            qname("custClr"),
            {"name": color_name},
        )
        ET.SubElement(
            custom_color,
            qname("srgbClr"),
            {"val": color_hex},
        )
    return custom_list


def update_theme_xml(
    theme_xml: bytes,
    colors: Sequence[Tuple[str, str]],
) -> bytes:
    """Replace the theme's a:custClrLst with the supplied custom colors."""
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError as exc:
        raise ValueError(f"theme1.xml is not valid XML: {exc}") from exc

    if root.tag != qname("theme"):
        raise ValueError(
            "theme1.xml does not have the expected DrawingML a:theme root element"
        )

    custom_tag = qname("custClrLst")
    extra_tag = qname("extraClrSchemeLst")
    extension_tag = qname("extLst")

    # Remove any previous custom-color list so rerunning the script is
    # deterministic rather than appending duplicate palettes.
    for child in list(root):
        if child.tag == custom_tag:
            root.remove(child)

    new_custom_list = build_custom_color_list(colors)

    # Brandwares places custClrLst immediately after extraClrSchemeLst. If a
    # package omits extraClrSchemeLst, put it before extLst or at the end.
    insert_at = len(root)
    for index, child in enumerate(list(root)):
        if child.tag == extra_tag:
            insert_at = index + 1
            break
    else:
        for index, child in enumerate(list(root)):
            if child.tag == extension_tag:
                insert_at = index
                break

    root.insert(insert_at, new_custom_list)

    # Keep the output broadly compatible with Office's expected UTF-8 XML.
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def next_backup_path(office_path: Path) -> Path:
    """Return a non-destructive backup name beside the source file."""
    first = Path(str(office_path) + ".bak")
    if not first.exists():
        return first

    counter = 1
    while True:
        candidate = Path(str(office_path) + f".bak.{counter}")
        if not candidate.exists():
            return candidate
        counter += 1


def create_backup(office_path: Path) -> Path:
    backup_path = next_backup_path(office_path)
    shutil.copy2(office_path, backup_path)
    return backup_path


def replace_theme_xml(
    office_path: Path,
    colors: Sequence[Tuple[str, str]],
    make_backup: bool,
) -> bool:
    """Replace theme1.xml in place, atomically, for one Office package."""
    theme_zip_path = get_theme_zip_path(office_path)
    if theme_zip_path is None:
        print(f"[skip] {office_path}: unsupported extension")
        return False

    if not office_path.is_file():
        print(f"[skip] {office_path}: file not found")
        return False

    if office_path.stat().st_size == 0:
        print(f"[skip] {office_path}: file is empty")
        return False

    temp_path: Optional[Path] = None
    try:
        with zipfile.ZipFile(office_path, "r") as source_zip:
            bad_entry = source_zip.testzip()
            if bad_entry is not None:
                raise zipfile.BadZipFile(
                    f"CRC check failed for package entry {bad_entry!r}"
                )

            if theme_zip_path not in source_zip.namelist():
                print(f"[skip] {office_path}: {theme_zip_path} was not found")
                return False

            original_theme_info = next(
                info
                for info in source_zip.infolist()
                if info.filename == theme_zip_path
            )
            original_theme = source_zip.read(original_theme_info)
            updated_theme = update_theme_xml(original_theme, colors)

            debug(
                f"{office_path}: theme={theme_zip_path}, "
                f"original_xml={len(original_theme)} bytes, "
                f"updated_xml={len(updated_theme)} bytes"
            )

            fd, temp_name = tempfile.mkstemp(
                prefix=f".{office_path.name}.",
                suffix=office_path.suffix,
                dir=str(office_path.parent),
            )
            os.close(fd)
            temp_path = Path(temp_name)

            with zipfile.ZipFile(temp_path, "w", allowZip64=True) as output_zip:
                output_zip.comment = source_zip.comment
                for info in source_zip.infolist():
                    if info.filename == theme_zip_path:
                        continue
                    output_zip.writestr(info, source_zip.read(info))

                replacement_info = copy.copy(original_theme_info)
                output_zip.writestr(replacement_info, updated_theme)

            # Verify the newly built package before touching the source file.
            with zipfile.ZipFile(temp_path, "r") as check_zip:
                bad_entry = check_zip.testzip()
                if bad_entry is not None:
                    raise zipfile.BadZipFile(
                        f"CRC check failed in generated package entry {bad_entry!r}"
                    )
                ET.fromstring(check_zip.read(theme_zip_path))

            # Preserve the source file's permission bits when replacing it.
            shutil.copymode(office_path, temp_path)

            backup_path: Optional[Path] = None
            if make_backup:
                backup_path = create_backup(office_path)
                debug(f"Backup created: {backup_path}")

            os.replace(temp_path, office_path)
            temp_path = None

            backup_message = f"; backup: {backup_path}" if backup_path else ""
            print(f"[ok] Updated {office_path}{backup_message}")
            return True

    except (zipfile.BadZipFile, OSError, ValueError, ET.ParseError) as exc:
        print(f"[skip] {office_path}: {exc}", file=sys.stderr)
        return False
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Add OOXML custom colors to Word/PowerPoint files. "
            "The file is replaced in place unless --backup is used."
        ),
        epilog=(
            "Input format: one '#RRGGBB color name' per line. "
            "EDITOR must be set, for example 'export EDITOR=nano' or "
            "'export EDITOR=\"code --wait\"'."
        ),
    )
    parser.add_argument(
        "input_path",
        help="A supported Office file, or a directory of Office files",
    )
    parser.add_argument(
        "-d",
        "--debug",
        action="store_true",
        help="Print diagnostic information",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="When input_path is a directory, process supported files recursively",
    )
    parser.add_argument(
        "-b",
        "--backup",
        action="store_true",
        help="Create a non-destructive .bak backup before each replacement",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    global DEBUG

    parser = build_parser()
    args = parser.parse_args(argv)
    DEBUG = args.debug

    input_path = Path(args.input_path).expanduser()
    debug(f"Input path: {input_path}")
    debug(f"Recursive: {args.recursive}; backup: {args.backup}")

    try:
        office_files = collect_office_files(input_path, args.recursive)
    except (FileNotFoundError, ValueError) as exc:
        parser.error(str(exc))

    if not office_files:
        print(f"No supported Office files found in {input_path}", file=sys.stderr)
        return 1

    print(
        f"Preparing custom colors for {len(office_files)} file(s). "
        "Close the Office files before continuing."
    )

    try:
        editor_text = launch_editor()
        colors, filler_count = parse_color_list(editor_text)
    except (ColorInputError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    custom_count = len(colors) - filler_count
    print(
        f"Color list: {custom_count} custom color(s), "
        f"{filler_count} white filler color(s), "
        f"{len(colors)} total palette entries."
    )
    debug(f"Parsed colors: {colors}")

    success_count = 0
    failure_count = 0
    for office_file in office_files:
        if replace_theme_xml(office_file, colors, args.backup):
            success_count += 1
        else:
            failure_count += 1

    print(f"Finished: {success_count} updated, {failure_count} skipped/failed.")
    return 0 if failure_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
