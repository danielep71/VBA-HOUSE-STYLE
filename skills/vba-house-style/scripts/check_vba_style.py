#!/usr/bin/env python3
"""Check (and optionally fix) the mechanical parts of the VBA house style.

Usage:
    python check_vba_style.py [--fix] [--max-line N] PATH [PATH ...]

PATH may be a .bas/.cls/.frm file or a folder (searched recursively).

Report mode lists findings and exits 1 when any ERROR is found.
--fix rewrites only whitespace that cannot change behavior: line endings to
CRLF, a final CRLF, trailing whitespace, and leading tabs (4 spaces each).
The .frm designer block (VERSION ... up to the first Attribute line) is
never touched. Bytes are round-tripped as latin-1, so the Windows code page
of the export is preserved. Standard library only.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

SUFFIXES = {".bas", ".cls", ".frm"}
FIELDS = ["PURPOSE", "WHY THIS EXISTS", "INPUTS", "RETURNS", "BEHAVIOR",
          "ERROR POLICY", "DEPENDENCIES", "NOTES", "UPDATED"]
PROC_RE = re.compile(
    r"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I)
END_RE = re.compile(r"^End\s+(Sub|Function|Property)\b", re.I)
BANNER_RE = re.compile(r"^'[-=]{5,}\s*$")
LABEL_RE = re.compile(r"' [A-Z][A-Z /]+")
STRING_RE = re.compile(r'"(?:[^"]|"")*"')
DECL_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Global|Friend|Dim|Static)\s+)?(?:Const\s+)?"
    r"(?:WithEvents\s+)?\w+(?:\([^)]*\))?\s+As\s+", re.I)
COMPACT_MAX = 15        # body lines up to which a Private procedure may be compact
COMPACT_FIELDS = ["PURPOSE", "UPDATED"]


def code_start(lines: list[str]) -> int:
    """Index of the first line after a .frm designer block (0 otherwise)."""
    if lines and lines[0].startswith("VERSION ") and len(lines) > 1 \
            and lines[1].startswith("Begin "):
        for i, line in enumerate(lines):
            if line.startswith("Attribute "):
                return i
    return 0


def code_length(line: str) -> int:
    """Length of a line without string-literal contents and, for a
    declaration, without its trailing comment (tables and long literals are
    exempt from the line limit; the code around them is not)."""
    bare = STRING_RE.sub('""', line)
    if DECL_RE.match(bare) and "'" in bare:
        bare = bare[:bare.index("'")].rstrip()
    return len(bare)


def field_problems(labels: list[str], required: list[str]) -> list[str]:
    """Missing fields, standard fields out of order, and extra sections that
    are not between the standard fields and UPDATED."""
    problems = []
    seen = [f for f in labels if f in FIELDS]
    missing = [f for f in required if f not in seen]
    if missing:
        problems.append("header missing " + ", ".join(missing))
    if seen != sorted(seen, key=FIELDS.index):
        problems.append("header fields out of order")
    for n, label in enumerate(labels):
        if label in FIELDS:
            continue
        before = [f for f in labels[n + 1:] if f in FIELDS and f != "UPDATED"]
        if before:
            problems.append(f"extra section {label} must follow {before[-1]}")
            break
    return problems


def fix(raw: bytes) -> bytes:
    text = raw.decode("latin-1").replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    start = code_start(lines)
    for i in range(start, len(lines)):
        line = lines[i]
        lead = len(line) - len(line.lstrip(" \t"))
        lines[i] = (line[:lead].replace("\t", "    ") + line[lead:]).rstrip(" \t")
    return ("\r\n".join(lines) + "\r\n").encode("latin-1")


def check(path: Path, max_line: int) -> list[tuple[str, int, str]]:
    out: list[tuple[str, int, str]] = []
    raw = path.read_bytes()
    text = raw.decode("latin-1")

    def add(level: str, line_no: int, msg: str) -> None:
        out.append((level, line_no, msg))

    if text.count("\n") != text.count("\r\n") or "\r" in text.replace("\r\n", ""):
        add("ERROR", 0, "line endings are not all CRLF")
    if text and not text.endswith("\r\n"):
        add("ERROR", 0, "file does not end with CRLF")

    lines = text.replace("\r\n", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    start = code_start(lines)

    names = re.findall(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
    if names and names[0] != path.stem:
        add("ERROR", 0, f'VB_Name "{names[0]}" does not match file name')
    if not re.search(r"^[ \t]*Option Explicit\b", text, re.M | re.I):
        add("ERROR", 0, "missing Option Explicit")

    for n, line in enumerate(lines[start:], start + 1):
        if "\t" in line:
            add("ERROR", n, "tab character")
        if line != line.rstrip(" \t"):
            add("ERROR", n, "trailing whitespace")
        if BANNER_RE.match(line) and len(line.rstrip()) != 79:
            add("ERROR", n, f"banner is {len(line.rstrip())} chars, expected 79")
        if max_line and len(line) > max_line and code_length(line) > max_line:
            add("WARN", n, f"line is {len(line)} chars (> {max_line})")

    # Component header fields (first header after Option Explicit).
    top = next((n for n in range(start, min(start + 200, len(lines)))
                if re.match(r"' (MODULE|CLASS):", lines[n])), None)
    if top is None:
        add("WARN", 0, "no ' MODULE: / ' CLASS: component header")
    else:
        n = top + 1
        while n < len(lines) and lines[n].startswith("'") \
                and not (BANNER_RE.match(lines[n]) and "UPDATED" in
                         [ln[2:] for ln in lines[top:n] if LABEL_RE.fullmatch(ln)]):
            n += 1
        labels = [ln[2:].strip() for ln in lines[top:n] if LABEL_RE.fullmatch(ln)]
        for msg in field_problems(labels, FIELDS):
            add("WARN", top + 1, f"component {msg}")

    # Procedures: header fields in order, spacing between procedures.
    i = start
    while i < len(lines):
        m = PROC_RE.match(lines[i])
        if not m or lines[i].lstrip().startswith("'"):
            i += 1
            continue
        sig_line, name = i + 1, m.group(2)
        # Blank lines between the previous End ... and this signature.
        j, blanks = i - 1, 0
        while j >= start and lines[j].strip() == "":
            blanks += 1
            j -= 1
        if j >= start and END_RE.match(lines[j]) and blanks != 1:
            add("WARN", sig_line, f"{name}: {blanks} blank lines after previous End, expected 1")
        # Header block: from signature end to the first code line.
        k = i
        while k < len(lines) and lines[k].rstrip().endswith(" _"):
            k += 1
        end = k + 1
        while end < len(lines) and not END_RE.match(lines[end]):
            end += 1
        h = k + 1
        while h < end and lines[h].strip() == "":
            h += 1
        header = []
        while h < end and lines[h].startswith("'"):
            header.append(lines[h])
            h += 1
        labels = [ln[2:].strip() for ln in header if LABEL_RE.fullmatch(ln)]
        body = sum(1 for ln in lines[h:end] if ln.strip() and not ln.lstrip().startswith("'"))
        compact = lines[i].lstrip().lower().startswith("private ") and body <= COMPACT_MAX
        for msg in field_problems(labels, COMPACT_FIELDS if compact else FIELDS):
            add("WARN", sig_line, f"{name}: {msg}")
        i = end + 1
    return out


def collect(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for p in map(Path, paths):
        if p.is_dir():
            files += sorted(f for f in p.rglob("*") if f.suffix.lower() in SUFFIXES)
        elif p.suffix.lower() in SUFFIXES:
            files.append(p)
        else:
            print(f"skip (not .bas/.cls/.frm): {p}", file=sys.stderr)
    return files


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--fix", action="store_true", help="normalize whitespace and CRLF in place")
    ap.add_argument("--max-line", type=int, default=120, help="warn above this length (0 = off)")
    args = ap.parse_args(argv)

    errors = 0
    for path in collect(args.paths):
        if args.fix:
            raw = path.read_bytes()
            fixed = fix(raw)
            if fixed != raw:
                path.write_bytes(fixed)
                print(f"fixed: {path}")
        for level, line_no, msg in check(path, args.max_line):
            errors += level == "ERROR"
            where = f"{path}:{line_no}" if line_no else str(path)
            print(f"{level:5} {where}: {msg}")
    print(f"{errors} error(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
