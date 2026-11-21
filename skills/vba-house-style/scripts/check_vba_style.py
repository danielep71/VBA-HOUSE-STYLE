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
DECL_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Global|Friend|Dim|Static)\s+)?(?:Const\s+)?"
    r"(?:WithEvents\s+)?\w+(?:\([^)]*\))?\s+As\s+", re.I)
# A date/time literal holds only digits, date and time separators, an AM/PM
# marker or a month name. Anything else between two "#" is file-number syntax
# (Print #1, A: Print #2, B), not a literal.
DATE_RE = re.compile(
    r"#\s*(?:\d[\d/\-.:, ]*|[A-Za-z]{3,9}\.?\s+\d[\d,: ]*)"
    r"(?:\s*[AaPp][Mm])?\s*#")
REM_RE = re.compile(r"Rem\b", re.I)
COMPARE_CONST_RE = re.compile(r"(?:VBA\.)?vb(?:Binary|Text|Database)Compare$", re.I)
NAMED_COMPARE_RE = re.compile(r"Compare\s*:=", re.I)
# Attribute records are VBE-managed export metadata, not statements.
ATTRIBUTE_RE = re.compile(r"^Attribute\s+\w", re.I)
LABEL_LINE_RE = re.compile(r"^\s*\w+:\s*$")
COMPARE_CALL_RE = re.compile(r"(?<![\w.])(?:VBA\.)?(InStr|StrComp)\s*\(", re.I)
EMPTY_COMPARE_RE = re.compile(r'(?:<>|=)\s*""|""\s*(?:<>|=)')
NEXT_RE = re.compile(r"^\s*Next\s*$", re.I)
PARAM_RE = re.compile(
    r"^(?:Optional\s+)?(?:ByVal|ByRef)\s+\w+(?:\(\s*\))?\s+As\s+\S", re.I)
COMPACT_MAX = 15        # body lines up to which a Private procedure may be compact
COMPACT_FIELDS = ["PURPOSE", "UPDATED"]
TRIVIAL_MAX = 1         # body lines up to which a procedure is trivial (one-line getter)
# Body sections (SKILL.md section 6): standard ones have a fixed rank, task
# sections sit between INITIALIZE and EXIT PROCEDURE.
SECTION_RANK = {"DECLARE": 0, "INITIALIZE": 1, "EXIT PROCEDURE": 3, "ERROR HANDLER": 4}
TASK_RANK = 2
LOCAL_DIM_RE = re.compile(r"\s*(?:Dim|Static)\s+\w", re.I)
PROC_NAME_RE = re.compile(r'\s*Const\s+PROC_NAME\s+As\s+String\s*=\s*"[^"]*"\s*$', re.I)
ON_ERROR_GOTO_RE = re.compile(r"\s*On\s+Error\s+GoTo\s+([A-Za-z]\w*)\s*$", re.I)
LABEL_AT_RE = re.compile(r"(\s*)([A-Za-z]\w*):(?!=)")
EXIT_PROC_RE = re.compile(r"\s*Exit\s+(?:Sub|Function|Property)\s*$", re.I)
# VBA library functions that must be written VBA.<Name> (SKILL.md section 7).
# Kept short on purpose: a name that a project may define itself is a false
# positive. Left, Right and Mid appear only with "$": bare Left/Right are
# UserForm and control properties, and Mid is also a statement.
VBA_QUALIFIED = (
    "Len", "LenB", "Trim", "Trim$", "LTrim", "LTrim$", "RTrim", "RTrim$",
    "UCase", "UCase$", "LCase", "LCase$", "Left$", "Right$", "Mid$",
    "InStr", "InStrRev", "StrComp",
    "CStr", "CLng", "CInt", "CDbl", "CBool", "CDate",
)
VBA_NAME_RE = re.compile(
    r"(?<![\w$])(" + "|".join(re.escape(n) for n in
                              sorted(VBA_QUALIFIED, key=len, reverse=True))
    + r")(?![\w$])", re.I)
VBA_SPELLING = {n.lower(): n for n in VBA_QUALIFIED}
DECLARE_STMT_RE = re.compile(
    r"^(?:(?:Public|Private|Friend)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)", re.I)
EVENT_RE = re.compile(r"^(?:Public\s+)?Event\s+(\w+)", re.I)
VAR_STMT_RE = re.compile(
    r"^(?:(?:Public|Private|Global|Friend|Dim|Static)\s+|ReDim\s+(?:Preserve\s+)?)"
    r"(?:Const\s+)?(.*)$|^Const\s+(.*)$", re.I)
BLOCK_RE = re.compile(r"^(?:(?:Public|Private)\s+)?(Enum|Type)\s+(\w+)", re.I)
AS_TYPE_RE = re.compile(r"\bAs\s+(?:New\s+)?(\w+)", re.I)
SCALAR_TYPES = {"string", "variant", "long", "integer", "double", "single", "currency",
                "date", "boolean", "byte", "longlong", "longptr", "decimal"}


def code_start(lines: list[str]) -> int:
    """Index of the first line after a .frm designer block (0 otherwise)."""
    if lines and lines[0].startswith("VERSION ") and len(lines) > 1 \
            and lines[1].startswith("Begin "):
        for i, line in enumerate(lines):
            if line.startswith("Attribute "):
                return i
    return 0


CODE, STRING, DATE, COMMENT = "code", "string", "date", "comment"
Token = tuple[str, str]


def tokenize(line: str) -> list[Token]:
    """Split one physical line into (kind, text) tokens: code, string
    literal ("" escapes), date literal, comment (apostrophe, or Rem at the
    start of a statement). Joined, the texts give back the line. Every
    statement rule reads these tokens, never the raw text."""
    tokens: list[Token] = []
    code: list[str] = []
    stmt_start = True
    i, n = 0, len(line)

    def flush() -> None:
        if code:
            tokens.append((CODE, "".join(code)))
            code.clear()

    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            end = min(j + 1, n)
            flush()
            tokens.append((STRING, line[i:end]))
            i, stmt_start = end, False
            continue
        if ch == "'":
            flush()
            tokens.append((COMMENT, line[i:]))
            return tokens
        if ch == "#":
            m = DATE_RE.match(line, i)
            if m:
                flush()
                tokens.append((DATE, m.group()))
                i, stmt_start = m.end(), False
                continue
        if stmt_start and ch in "Rr" and REM_RE.match(line, i):
            flush()
            tokens.append((COMMENT, line[i:]))
            return tokens
        if ch == ":":
            stmt_start = True
        elif not ch.isspace():
            stmt_start = False
        code.append(ch)
        i += 1
    flush()
    return tokens


def code_view(tokens: list[Token]) -> str:
    """The code of a line for the statement rules: non-empty strings become
    "S", empty ones stay "", date literals become #D#, comments are dropped."""
    out = []
    for kind, text in tokens:
        if kind == CODE:
            out.append(text)
        elif kind == STRING:
            out.append('""' if text == '""' else '"S"')
        elif kind == DATE:
            out.append("#D#")
    return "".join(out).rstrip()


def code_length(line: str) -> int:
    """Length of a line without string-literal contents and, for a
    declaration, without its trailing comment (tables and long literals are
    exempt from the line limit; the code around them is not)."""
    tokens = tokenize(line)
    bare = "".join('""' if kind == STRING else text for kind, text in tokens)
    if DECL_RE.match(bare) and any(kind == COMMENT for kind, _ in tokens):
        bare = "".join('""' if kind == STRING else text for kind, text in tokens
                       if kind != COMMENT).rstrip()
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


def body_sections(lines: list[str], first: int, last: int) -> list[tuple[str, int]]:
    """(title, index of the title line) of each section banner in
    lines[first:last]: an UPPERCASE title after one space, between two banner
    lines. Header fields and the centered title are not sections."""
    out = []
    for n in range(max(first + 1, 1), last - 1):
        line = lines[n]
        title = line[2:].strip()
        if line.startswith("' ") and not line.startswith("'  ") and title \
                and title == title.upper() and re.search("[A-Z]", title) \
                and BANNER_RE.match(lines[n - 1]) and BANNER_RE.match(lines[n + 1]):
            out.append((title, n))
    return out


def body_problems(lines: list[str], stmts: list[tuple[int, str]], first: int,
                  last: int, proc_name: str, kind: str) -> list[tuple[int, str]]:
    """Layout problems of one procedure body (SKILL.md section 6), as
    (line number, message). stmts holds the (line index, code) of each
    statement in lines[first:last]."""
    problems: list[tuple[int, str]] = []
    sections = body_sections(lines, first, last)

    def section_at(idx: int) -> str | None:
        title = None
        for name, at in sections:
            if at >= idx:
                break
            title = name
        return title

    # Order: DECLARE, INITIALIZE, task sections, EXIT PROCEDURE, ERROR HANDLER.
    prev, prev_rank = "", -1
    for title, at in sections:
        rank = SECTION_RANK.get(title, TASK_RANK)
        if rank < prev_rank:
            problems.append((at + 1, f"section {title} is out of order (after {prev})"))
            break
        prev, prev_rank = title, rank

    # DECLARE starts with PROC_NAME; needed for local variables or an error handler.
    handlers = [(idx, m.group(1)) for idx, code in stmts
                if (m := ON_ERROR_GOTO_RE.match(code))]
    declare = next((at for title, at in sections if title == "DECLARE"), None)
    if declare is not None:
        idx, code = next(((i, c) for i, c in stmts if i > declare), (declare, ""))
        if section_at(idx) != "DECLARE" or not PROC_NAME_RE.match(code):
            problems.append((idx + 1, "DECLARE does not start with Const PROC_NAME As String"))
        else:
            n = idx
            while lines[n].rstrip().endswith(" _") and n + 1 < last:
                n += 1
            value = next((t[1:-1] for line in lines[idx:n + 1]
                          for k, t in tokenize(line) if k == STRING), proc_name)
            if value.lower() != proc_name.lower():
                problems.append((idx + 1, f'PROC_NAME is "{value}", expected "{proc_name}"'))
    else:
        need = next((i for i, c in stmts if LOCAL_DIM_RE.match(c)), None)
        if need is None and handlers:
            need = handlers[0][0]
        if need is not None:
            problems.append((need + 1, "no DECLARE section for local variables or error handler"))

    # On Error GoTo <label> in INITIALIZE; the label at column 0; the clean
    # exit in EXIT PROCEDURE before the label.
    if handlers:
        idx, label = handlers[0]
        if section_at(idx) != "INITIALIZE":
            problems.append((idx + 1, f"On Error GoTo {label} is not in INITIALIZE"))
        label_at = next((i for i, c in stmts if (m := LABEL_AT_RE.match(c))
                         and m.group(2).lower() == label.lower()), None)
        if label_at is None:
            problems.append((idx + 1, f"label {label}: not found"))
        else:
            if lines[label_at][:1].isspace():
                problems.append((label_at + 1, f"label {label}: is not at column 0"))
            if not any(EXIT_PROC_RE.match(c) and section_at(i) == "EXIT PROCEDURE"
                       for i, c in stmts if i < label_at):
                problems.append((label_at + 1,
                                 f"no Exit {kind} in EXIT PROCEDURE before {label}:"))
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

    # Procedures: header fields in order, spacing between procedures, body
    # sections.
    component = names[0] if names else path.stem
    statements = {n - 1: code for n, code in logical_lines(lines, start) if code.strip()}
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
        if not compact and body > TRIVIAL_MAX:
            stmts = [(n, statements[n]) for n in range(k + 1, end) if n in statements]
            kind = m.group(1).split()[0].capitalize()
            # Sections start after the banner that closes the header, so the
            # procedure title is never read as a section.
            first = k + 1
            updated = next((n for n in range(k + 1, end) if lines[n].rstrip() == "' UPDATED"), None)
            if updated is not None:
                first = next((n + 1 for n in range(updated, end) if BANNER_RE.match(lines[n])), first)
            for line_no, msg in body_problems(lines, stmts, first, end,
                                              f"{component}.{name}", kind):
                add("WARN", line_no, f"{name}: {msg}")
        i = end + 1
    out += statement_findings(lines, start)
    return out


def logical_lines(lines: list[str], start: int) -> list[tuple[int, str]]:
    """(first line number, code) per statement line, with continuation lines
    joined. Comment-only and blank lines give empty code."""
    out: list[tuple[int, str]] = []
    n = start
    while n < len(lines):
        first, code = n + 1, code_view(tokenize(lines[n]))
        while code.endswith(" _") and n + 1 < len(lines):
            n += 1
            code = code[:-1] + " " + code_view(tokenize(lines[n])).strip()
        out.append((first, code))
        n += 1
    return out


def split_args(text: str) -> list[str]:
    """Top-level comma-separated parts of text (parentheses nest)."""
    parts, depth, cur = [], 0, ""
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(cur.strip())
            cur = ""
        else:
            cur += ch
    parts.append(cur.strip())
    return parts


def call_args(code: str, open_at: int) -> tuple[list[str], int] | None:
    """Arguments of the call whose "(" is at open_at, and the index after its
    closing ")". None when the parenthesis is not closed on this statement."""
    depth = 0
    for i in range(open_at, len(code)):
        if code[i] == "(":
            depth += 1
        elif code[i] == ")":
            depth -= 1
            if depth == 0:
                return split_args(code[open_at + 1:i]), i + 1
    return None


def missing_compare(name: str, args: list[str]) -> bool:
    """True when the InStr/StrComp call certainly has no compare argument.
    The decision uses the argument count and named-argument syntax, never
    the text of the inputs; an ambiguous call (a variable that may be the
    compare mode) is not reported."""
    if any(NAMED_COMPARE_RE.match(a) for a in args):
        return False
    if name.lower() == "strcomp":
        return len(args) == 2
    if len(args) == 2:
        return True
    if len(args) == 3:
        if COMPARE_CONST_RE.fullmatch(args[2]):
            return False
        return re.fullmatch(r"\d+", args[0]) is not None
    return False


def signature_problems(code: str, m: re.Match[str]) -> list[str]:
    """ByVal/ByRef and As problems of a procedure signature."""
    problems: list[str] = []
    open_at = code.find("(", m.end())
    if open_at < 0:
        return problems
    parsed = call_args(code, open_at)
    if parsed is None:
        return problems
    params, after = parsed
    for param in params:
        if not param or re.match(r"ParamArray\b", param, re.I):
            continue
        param = re.sub(r"\s*=.*$", "", param)
        if PARAM_RE.match(param):
            continue
        name = re.sub(r"^(?:Optional\s+)?(?:(?:ByVal|ByRef)\s+)?", "", param,
                      flags=re.I).split()[0:1]
        label = name[0] if name else param
        if not re.search(r"\b(ByVal|ByRef)\b", param, re.I):
            problems.append(f"parameter {label} has no ByVal or ByRef")
        if not re.search(r"\bAs\s+\w", param, re.I):
            problems.append(f"parameter {label} has no As type")
    kind = m.group(1).lower()
    if (kind == "function" or kind == "property get") \
            and not re.match(r"\s+As\s+\w", code[after:], re.I):
        problems.append("no As return type")
    return problems


def first_word(text: str) -> str | None:
    """The identifier a declaration part starts with (after its modifiers)."""
    text = re.sub(r"^(?:(?:Optional|ByVal|ByRef|ParamArray|WithEvents|Const)\s+)+",
                  "", text.strip(), flags=re.I)
    m = re.match(r"[A-Za-z]\w*", text)
    return m.group().lower() if m else None


def declared_names(statements: list[tuple[int, str]]) -> tuple[set[str], set[str]]:
    """Lower-case names the file declares (procedures, Declare and Event
    names, their parameters, variables and constants of any scope, Enum and
    Type members), and the subset declared As a type that is not a VBA
    scalar (a user-defined type or an object). A library name the file
    declares is its own; Len of a user-defined type cannot be VBA.Len."""
    names: set[str] = set()
    records: set[str] = set()
    block = False

    def add(part: str) -> None:
        word = first_word(part)
        if not word:
            return
        names.add(word)
        m = AS_TYPE_RE.search(part)
        if m and m.group(1).lower() not in SCALAR_TYPES:
            records.add(word)

    for _, code in statements:
        code = code.strip()
        if not code or ATTRIBUTE_RE.match(code):
            continue
        if block:
            if re.match(r"End\s+(?:Enum|Type)\b", code, re.I):
                block = False
            else:
                add(code)
            continue
        m = BLOCK_RE.match(code)
        if m:
            block = True
            names.add(m.group(2).lower())
            continue
        m = PROC_RE.match(code) or DECLARE_STMT_RE.match(code) or EVENT_RE.match(code)
        if m:
            names.add(m.group(m.lastindex).lower())
            open_at = code.find("(", m.end())
            parsed = call_args(code, open_at) if open_at >= 0 else None
            for param in parsed[0] if parsed else []:
                add(param)
            continue
        m = VAR_STMT_RE.match(code)
        if m:
            for part in split_args(m.group(1) or m.group(2)):
                add(part)
    return names, records


def unqualified_calls(code: str, declared: set[str],
                      records: set[str]) -> list[tuple[str, str]]:
    """Library names of VBA_QUALIFIED used bare in one statement's code, as
    (listed spelling, written text). Not reported: a member (after "." or
    "!"), a named argument or an Open ... Len = clause (followed by "=" or
    ":="), a name the file declares, the Mid statement, and Len or LenB of a
    variable declared As a non-scalar type."""
    found: list[tuple[str, str]] = []
    for m in VBA_NAME_RE.finditer(code):
        before = code[:m.start()].rstrip()
        after = code[m.end():].lstrip()
        if before.endswith((".", "!")) or after.startswith((":=", "=")):
            continue
        name = m.group(1)
        if name.rstrip("$").lower() in declared:
            continue
        # The Mid statement (Mid$(Text, 1, 1) = "x") cannot be qualified.
        if name.lower() == "mid$" and re.search(r"(?:^|\b(?:Then|Else))$", before, re.I):
            continue
        if name.lower() in ("len", "lenb") and after.startswith("("):
            parsed = call_args(after, 0)
            arg = re.fullmatch(r"([A-Za-z]\w*)(?:\(.*\))?", parsed[0][0]) \
                if parsed and len(parsed[0]) == 1 else None
            if arg and arg.group(1).lower() in records:
                continue
        found.append((VBA_SPELLING[name.lower()], name))
    return found


def statement_findings(lines: list[str], start: int) -> list[tuple[str, int, str]]:
    """Statement-level conventions (SKILL.md section 7), as WARN findings."""
    out: list[tuple[str, int, str]] = []
    statements = logical_lines(lines, start)
    declared, records = declared_names(statements)
    for n, code in statements:
        if not code.strip() or ATTRIBUTE_RE.match(code):
            continue
        # VBA library calls qualified with VBA.
        for name, written in unqualified_calls(code, declared, records):
            out.append(("WARN", n, f"use VBA.{name} instead of {written}"))
        # Next names its counter.
        if NEXT_RE.match(code):
            out.append(("WARN", n, "Next without the counter name"))
        # One statement per line: ":" outside labels, named arguments (:=) excluded.
        plain = code.replace(":=", "  ")
        if ":" in plain and not LABEL_LINE_RE.match(plain):
            out.append(("WARN", n, "statement separator ':' (one statement per line)"))
        # Explicit compare mode in InStr / StrComp.
        for call in COMPARE_CALL_RE.finditer(code):
            parsed = call_args(code, call.end() - 1)
            if parsed is None:
                continue
            args, name = parsed[0], call.group(1)
            if missing_compare(name, args):
                out.append(("WARN", n, f"{name} without an explicit compare mode"))
        # vbNullString rather than "".
        if EMPTY_COMPARE_RE.search(code) and not re.match(r"\s*(?:\w+\s+)*Const\b", code, re.I):
            out.append(("WARN", n, 'use vbNullString instead of ""'))
        # Signatures: ByVal/ByRef and As.
        m = PROC_RE.match(code)
        if m:
            for problem in signature_problems(code, m):
                out.append(("WARN", n, f"{m.group(2)}: {problem}"))
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
