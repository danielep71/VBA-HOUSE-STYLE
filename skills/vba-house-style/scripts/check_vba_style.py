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
import datetime
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
# Header field text starts three spaces after the apostrophe (SKILL.md
# section 4); deeper lines continue it.
FIELD_TEXT_RE = re.compile(r"'   \S")
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
SELECT_RE = re.compile(r"Select\s+Case\b", re.I)
CASE_ELSE_RE = re.compile(r"Case\s+Else\b", re.I)
END_SELECT_RE = re.compile(r"End\s+Select\b", re.I)
LOCAL_VAR_RE = re.compile(r"(?:Dim|Static)\s+(.*)$", re.I)
# Hungarian prefix: a lower-case type tag followed by an upper-case letter
# (strName, wsData). Case-sensitive, so Strength and Intent are words.
HUNGARIAN_RE = re.compile(r"(str|lng|int|obj|bln|dbl|var|rng|ws|wb)[A-Z]")
PASCAL_RE = re.compile(r"[A-Z][A-Za-z0-9]*$")
VISIBILITY_RE = re.compile(r"(?:Public|Private|Friend)\s", re.I)
CONST_STMT_RE = re.compile(r"(?:(?:Public|Private|Global)\s+)?Const\s+(.*)$", re.I)
LABEL_COL0_RE = re.compile(r"' \S")
UPPER_SNAKE_RE = re.compile(r"[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*$")
# A one-line declaration whose "As" is aligned with its neighbours (group 1).
# The modifier is required, so Type members, Enum members and Declare
# statements never match.
ALIGN_DECL_RE = re.compile(
    r"\s*(?:(?:Public|Private|Global|Friend|Dim|Static)\s+(?:Const\s+)?|Const\s+)"
    r"(?:WithEvents\s+)?\w+(?:\([^)]*\))?\s+(As)\s+", re.I)
# Statement indentation (SKILL.md section 6): the intent comment at 4 spaces,
# the code at 8, 4 more per open block. Select Case opens two levels: Case
# sits one level in, its body two.
INTENT_INDENT, CODE_INDENT, INDENT_STEP = 4, 8, 4
BLOCK_OPEN_RE = {
    "For": re.compile(r"For\b", re.I),
    "Do": re.compile(r"Do\b", re.I),
    "While": re.compile(r"While\b", re.I),
    "With": re.compile(r"With\b", re.I),
    "Select": re.compile(r"Select\s+Case\b", re.I),
}
BLOCK_CLOSE_RE = {
    "If": re.compile(r"End\s*If\b", re.I),
    "For": re.compile(r"Next\b", re.I),
    "Do": re.compile(r"Loop\b", re.I),
    "While": re.compile(r"Wend\b", re.I),
    "With": re.compile(r"End\s+With\b", re.I),
    "Select": re.compile(r"End\s+Select\b", re.I),
}
BLOCK_LEVELS = {"Select": 2}
IF_RE = re.compile(r"If\b", re.I)
ELSEIF_RE = re.compile(r"ElseIf\b", re.I)
ELSE_RE = re.compile(r"Else\s*$", re.I)
CASE_RE = re.compile(r"Case\b", re.I)
THEN_END_RE = re.compile(r"\bThen$", re.I)
SEPARATOR_RE = re.compile(r":(?!=)")
CC_IF_RE = re.compile(r"#\s*If\b", re.I)
CC_END_IF_RE = re.compile(r"#\s*End\s*If\b", re.I)


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


FieldText = dict[str, tuple[int, list[tuple[int, str]]]]


def header_fields(lines: list[str], first: int, last: int) -> FieldText:
    """Standard fields of the header in lines[first:last]: label -> (label
    line index, [(line index, text line)]). Text lines are the comment lines
    after the label up to the next label or banner; lone "'" lines are left
    out."""
    fields: FieldText = {}
    current: list[tuple[int, str]] | None = None
    for n in range(first, last):
        line = lines[n]
        if BANNER_RE.match(line):
            current = None
        elif LABEL_RE.fullmatch(line):
            label = line[2:].strip()
            current = [] if label in FIELDS and label not in fields else None
            if current is not None:
                fields[label] = (n, current)
        elif current is not None and line.startswith("'") and line.strip() != "'":
            current.append((n, line))
    return fields


def param_names(code: str, m: re.Match[str]) -> list[str] | None:
    """Parameter names of a signature, or None when the list cannot be read."""
    open_at = code.find("(", m.end())
    if open_at < 0:
        return []
    parsed = call_args(code, open_at)
    if parsed is None:
        return None
    names = []
    for param in parsed[0]:
        if not param:
            continue
        param = re.sub(r"^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)*", "", param, flags=re.I)
        name = re.match(r"[A-Za-z]\w*", param)
        if name is None:
            return None
        names.append(name.group(0))
    return names


def field_value_problems(lines: list[str], first: int, last: int,
                         params: list[str] | None, kind: str,
                         component: str | None = None) -> list[tuple[int, str]]:
    """Header field values (SKILL.md section 4): field text at 3 spaces, a
    valid UPDATED date, a Sub that RETURNS Nothing, INPUTS that match the
    signature (procedures) or say None at module / class level (components).
    params is None when the signature could not be read; kind is the
    procedure kind in lower case. Fields that are absent are not checked."""
    out: list[tuple[int, str]] = []
    fields = header_fields(lines, first, last)
    for label, (at, text) in fields.items():
        if text and not FIELD_TEXT_RE.match(text[0][1]):
            out.append((text[0][0] + 1, f"{label} text not at 3 spaces"))
    if "UPDATED" in fields:
        at, text = fields["UPDATED"]
        value = text[0][1][1:].strip() if text else ""
        if not valid_date(value):
            out.append(((text[0][0] if text else at) + 1,
                        f"UPDATED {value or '(empty)'} is not a YYYY-MM-DD date"))
    if kind == "sub" and fields.get("RETURNS", (0, []))[1]:
        at, text = fields["RETURNS"]
        if not re.match(r"Nothing\b", text[0][1][1:].strip()):
            out.append((text[0][0] + 1, "RETURNS of a Sub is not Nothing"))
    if "INPUTS" not in fields or not fields["INPUTS"][1]:
        return out
    at, text = fields["INPUTS"]
    first_text = text[0][1][1:].strip()
    if component is not None:
        want = f"None at {component} level"
        if first_text.rstrip(".") != want:
            out.append((text[0][0] + 1, f"component INPUTS is not '{want}'"))
        return out
    if params is None:
        return out
    if re.match(r"None\b", first_text):
        if params:
            out.append((text[0][0] + 1, f"INPUTS is None but the signature "
                        f"declares {', '.join(params)}"))
        return out
    named = [line for _, line in text if FIELD_TEXT_RE.match(line)]
    if not all(re.fullmatch(r"'   [A-Za-z]\w*", line) for line in named):
        return out
    listed = [line[4:] for line in named]
    lower = {p.lower() for p in params}
    missing = [p for p in params if p.lower() not in {n.lower() for n in listed}]
    extra = [n for n in listed if n.lower() not in lower]
    if missing:
        out.append((at + 1, f"INPUTS does not list {', '.join(missing)}"))
    if extra:
        out.append((at + 1, f"INPUTS lists {', '.join(extra)}, not in the signature"))
    return out


def valid_date(value: str) -> bool:
    """True for a real calendar date written YYYY-MM-DD."""
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return False
    try:
        datetime.date.fromisoformat(value)
    except ValueError:
        return False
    return True


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


def blanks_above(lines: list[str], n: int) -> tuple[int, int]:
    """(blank lines right above line n, index of the line above them; -1 at
    the top of the file)."""
    j = n - 1
    while j >= 0 and lines[j].strip() == "":
        j -= 1
    return n - 1 - j, j


def section_spacing_problems(lines: list[str], first: int,
                             last: int) -> list[tuple[int, str]]:
    """Section banners in lines[first:last] without exactly one blank line
    above the opening banner line, or with a blank line right below the
    closing one (SKILL.md section 2). One finding per kind, with the count
    of the others."""
    bad: dict[str, list[tuple[int, str]]] = {}
    for title, n in body_sections(lines, first, last):
        blanks, _ = blanks_above(lines, n - 1)
        if blanks != 1:
            bad.setdefault("before", []).append(
                (n, f"{blanks} blank lines before section banner {title}, expected 1"))
        if n + 2 < last and lines[n + 2].strip() == "":
            bad.setdefault("after", []).append(
                (n + 3, f"blank line after section banner {title}, expected none"))
    out = []
    for found in bad.values():
        line_no, msg = found[0]
        more = f"; {len(found) - 1} more" if len(found) > 1 else ""
        out.append((line_no, msg + more))
    return sorted(out)


def group_banner_blocks(lines: list[str], first: int, last: int,
                        procs: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """(first, last) index of each group banner block in lines[first:last]
    outside procedures: a run of comment lines at column 0 that starts with
    a lone "'" and holds at least two banner lines."""
    inside = set()
    for a, b in procs:
        inside.update(range(a, b + 1))
    out, n = [], first
    while n < last:
        if n in inside or lines[n].rstrip() != "'":
            n += 1
            continue
        m = n
        while m + 1 < last and m + 1 not in inside and lines[m + 1].startswith("'"):
            m += 1
        if sum(1 for k in range(n, m + 1) if BANNER_RE.match(lines[k])) >= 2:
            out.append((n, m))
        n = m + 1
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


def is_intent_comment(line: str) -> bool:
    """A comment line whose text starts right after the apostrophe. Banner
    lines, section labels (' TITLE), lone apostrophes and commented-out code
    (an apostrophe followed by spaces) are not intent comments."""
    text = line.lstrip(" ")
    return text.startswith("'") and text[1:2] not in ("", " ") \
        and not BANNER_RE.match(text)


def block_step(code: str, stack: list[str]) -> int | None:
    """Apply one statement to the stack of open blocks and return the level
    at which the statement sits, or None when the statement cannot be
    classified (the caller then stays silent)."""
    depth = sum(BLOCK_LEVELS.get(k, 1) for k in stack)
    for kind, close in BLOCK_CLOSE_RE.items():
        if close.match(code):
            count = 1
            if kind == "For":
                counters = code[4:].strip()
                count = len(split_args(counters)) if counters else 1
            for _ in range(count):
                if not stack or stack[-1] != kind:
                    return None
                stack.pop()
            return sum(BLOCK_LEVELS.get(k, 1) for k in stack)
    if ELSEIF_RE.match(code):
        ok = stack and stack[-1] == "If" and THEN_END_RE.search(code)
        return depth - 1 if ok else None
    if ELSE_RE.match(code):
        return depth - 1 if stack and stack[-1] == "If" else None
    if re.match(r"Else\b", code, re.I):
        return None
    if CASE_RE.match(code):
        return depth - 1 if stack and stack[-1] == "Select" else None
    if IF_RE.match(code):
        if THEN_END_RE.search(code):
            stack.append("If")
        elif not re.search(r"\bThen\b", code, re.I):
            return None
        return depth
    for kind, opener in BLOCK_OPEN_RE.items():
        if opener.match(code):
            stack.append(kind)
            break
    return depth


def indent_problems(lines: list[str], first: int, last: int) -> list[tuple[int, str]]:
    """Statements and intent comments out of indentation in the body
    sections of lines[first:last], except DECLARE (SKILL.md section 6), as
    (line number, message). A statement sits at 8 spaces plus 4 per open
    block; an intent comment outside any block sits at 4. Continuation
    lines, labels, #If lines and comments inside a block are not checked.
    One finding per section and kind. When a line cannot be classified, the
    blocks do not balance, or a statement inside #If ... #End If opens or
    closes a block, the procedure is not reported."""
    sections = {at: title for title, at in body_sections(lines, first, last)}
    if not sections:
        return []
    stack: list[str] = []
    conditional = 0                 # open #If ... #End If regions
    title: str | None = None
    bad: dict[tuple[str, str], list[tuple[int, int, int]]] = {}
    n = first
    while n < last:
        title = sections.get(n, title)
        line = lines[n]
        lead = len(line) - len(line.lstrip(" "))
        code = code_view(tokenize(line))
        stop = n
        while code.endswith(" _") and stop + 1 < last:
            stop += 1
            code = code[:-1] + " " + code_view(tokenize(lines[stop])).strip()
        code = code.strip()
        checked = title is not None and title != "DECLARE"
        if line[lead:lead + 1] == "\t":
            return []
        if not code:
            if checked and not stack and is_intent_comment(line) and lead != INTENT_INDENT:
                bad.setdefault((title, "intent comment"), []).append(
                    (n + 1, lead, INTENT_INDENT))
        elif code.startswith("#"):
            if CC_IF_RE.match(code):
                conditional += 1
            elif CC_END_IF_RE.match(code):
                conditional -= 1
        elif not LABEL_LINE_RE.match(code):
            if SEPARATOR_RE.search(code) or code[0].isdigit():
                return []
            before = list(stack)
            level = block_step(code, stack)
            if level is None or level < 0:
                return []
            # Alternative #If branches may open or close blocks differently.
            if conditional > 0 and stack != before:
                return []
            want = CODE_INDENT + INDENT_STEP * level
            if checked and lead != want:
                bad.setdefault((title, "statement"), []).append((n + 1, lead, want))
        n = stop + 1
    if stack:
        return []
    out: list[tuple[int, str]] = []
    for (section, kind), found in bad.items():
        line_no, got, want = found[0]
        more = f"; {len(found) - 1} more in {section}" if len(found) > 1 else ""
        out.append((line_no, f"{kind} indented {got} spaces, expected {want}{more}"))
    return sorted(out)


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
    header_end = start
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
        header_end = n
        labels = [ln[2:].strip() for ln in lines[top:n] if LABEL_RE.fullmatch(ln)]
        for msg in field_problems(labels, FIELDS):
            add("WARN", top + 1, f"component {msg}")
        level = "class" if lines[top].startswith("' CLASS:") else "module"
        for line_no, msg in field_value_problems(lines, top, header_end, [], "",
                                                 component=level):
            add("WARN", line_no, msg if msg.startswith("component") else f"component {msg}")

    # Procedures: header fields in order, spacing between procedures, body
    # sections.
    component = names[0] if names else path.stem
    first_proc = next((n for n in range(start, len(lines))
                       if PROC_RE.match(lines[n]) and not lines[n].lstrip().startswith("'")),
                      len(lines))
    for line_no, msg in alignment_problems(lines, start, first_proc):
        add("WARN", line_no, msg)
    if top is not None:
        for line_no, msg in banner_problems(lines, header_end + 1, first_proc):
            add("WARN", line_no, msg)
        for line_no, msg in section_spacing_problems(lines, header_end + 1, first_proc):
            add("WARN", line_no, msg)
    statements = {n - 1: code for n, code in logical_lines(lines, start) if code.strip()}
    procs: list[tuple[int, int]] = []
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
        header, header_at = [], h
        while h < end and lines[h].startswith("'"):
            header.append(lines[h])
            h += 1
        for line_no, msg in title_problems(lines, header_at, h):
            add("WARN", line_no, f"{name}: {msg}")
        labels = [ln[2:].strip() for ln in header if LABEL_RE.fullmatch(ln)]
        body = sum(1 for ln in lines[h:end] if ln.strip() and not ln.lstrip().startswith("'"))
        compact = lines[i].lstrip().lower().startswith("private ") and body <= COMPACT_MAX
        for msg in field_problems(labels, COMPACT_FIELDS if compact else FIELDS):
            add("WARN", sig_line, f"{name}: {msg}")
        sig_code = statements.get(i, "")
        sig_m = PROC_RE.match(sig_code)
        params = param_names(sig_code, sig_m) if sig_m else None
        for line_no, msg in field_value_problems(lines, header_at, h, params,
                                                 m.group(1).split()[0].lower()):
            add("WARN", line_no, f"{name}: {msg}")
        # Sections start after the banner that closes the header, so the
        # procedure title is never read as a section.
        # Without UPDATED the header block ends where its comment lines do.
        first = k + 1
        updated = next((n for n in range(k + 1, end) if lines[n].rstrip() == "' UPDATED"), None)
        banners_from = h
        if updated is not None:
            first = next((n + 1 for n in range(updated, end) if BANNER_RE.match(lines[n])), first)
            banners_from = first
        for line_no, msg in banner_problems(lines, banners_from, end):
            add("WARN", line_no, f"{name}: {msg}")
        if header:
            for line_no, msg in section_spacing_problems(lines, banners_from, end):
                add("WARN", line_no, f"{name}: {msg}")
            # Blank lines around the header (SKILL.md section 6): one after a
            # one-line signature, none after a multi-line one, one before End.
            want = 0 if k > i else 1
            if header_at - k - 1 != want:
                add("WARN", sig_line, f"{name}: {header_at - k - 1} blank lines after "
                                      f"the signature, expected {want}")
            blanks, above = blanks_above(lines, end)
            if above >= h and blanks != 1:
                add("WARN", end + 1, f"{name}: {blanks} blank lines before "
                                     f"{lines[end].strip()}, expected 1")
        for line_no, msg in indent_problems(lines, first, end):
            add("WARN", line_no, f"{name}: {msg}")
        if not compact and body > TRIVIAL_MAX:
            stmts = [(n, statements[n]) for n in range(k + 1, end) if n in statements]
            kind = m.group(1).split()[0].capitalize()
            for line_no, msg in body_problems(lines, stmts, first, end,
                                              f"{component}.{name}", kind):
                add("WARN", line_no, f"{name}: {msg}")
            sections = body_sections(lines, first, end)
            for pos, (title, at) in enumerate(sections):
                if title == "DECLARE":
                    stop = sections[pos + 1][1] if pos + 1 < len(sections) else end
                    for line_no, msg in alignment_problems(lines, at, stop):
                        add("WARN", line_no, f"{name}: {msg}")
        procs.append((i, end))
        i = end + 1
    # Group banner blocks (SKILL.md section 6): one blank line after the
    # previous End, two after the declarations, one before the signature.
    groups_from = header_end + 1 if top is not None else start
    for first_line, last_line in group_banner_blocks(lines, groups_from, len(lines), procs):
        blanks, above = blanks_above(lines, first_line)
        if above >= 0 and END_RE.match(lines[above]):
            if blanks != 1:
                add("WARN", first_line + 1, f"group banner: {blanks} blank lines after "
                                            "previous End, expected 1")
        elif top is not None and above > header_end and blanks != 2:
            add("WARN", first_line + 1, f"group banner: {blanks} blank lines after "
                                        "the declarations, expected 2")
        below = last_line + 1
        while below < len(lines) and lines[below].strip() == "":
            below += 1
        sig = PROC_RE.match(lines[below]) if below < len(lines) else None
        if sig and below - last_line - 1 != 1:
            add("WARN", below + 1, f"{sig.group(2)}: {below - last_line - 1} blank lines "
                                   "after the group banner, expected 1")
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


def scoped_parts(statements: list[tuple[int, str]]):
    """(line number, statement, inside a procedure) per statement. A line is
    split at ":" (not ":="; strings are masked). The signature counts as
    inside the procedure, the End line as outside."""
    in_proc = False
    for n, code in statements:
        for part in re.split(r":(?!=)", code):
            part = part.strip()
            if not part:
                continue
            if PROC_RE.match(part):
                in_proc = True
            elif END_RE.match(part):
                in_proc = False
            yield n, part, in_proc


def block_findings(statements: list[tuple[int, str]]) -> list[tuple[int, str]]:
    """Select Case blocks without Case Else, and local variables (Dim or
    Static inside a procedure) with a Hungarian prefix or a name that is not
    PascalCase, as (line number, message). Module-level variables,
    constants and parameters are not local variables."""
    out: list[tuple[int, str]] = []
    selects: list[list] = []        # [line number, has Case Else] per open block
    for n, part, in_proc in scoped_parts(statements):
        if SELECT_RE.match(part):
            selects.append([n, False])
        elif CASE_ELSE_RE.match(part) and selects:
            selects[-1][1] = True
        elif END_SELECT_RE.match(part) and selects:
            line_no, has_else = selects.pop()
            if not has_else:
                out.append((line_no, "Select Case without Case Else"))
        m = LOCAL_VAR_RE.match(part) if in_proc else None
        if m:
            for decl in split_args(m.group(1)):
                name = re.match(r"\s*(?:WithEvents\s+)?([A-Za-z]\w*)", decl, re.I)
                if not name:
                    continue
                name = name.group(1)
                prefix = HUNGARIAN_RE.match(name)
                if prefix:
                    out.append((n, f"local variable {name} has a Hungarian prefix "
                                   f"({prefix.group(1)})"))
                elif not PASCAL_RE.match(name):
                    out.append((n, f"local variable {name} is not PascalCase"))
    return sorted(out)


def const_findings(statements: list[tuple[int, str]]) -> list[tuple[int, str]]:
    """Module-level constants whose name is not UPPER_SNAKE_CASE, as (line
    number, message). Constants inside a procedure are not checked."""
    out: list[tuple[int, str]] = []
    for n, part, in_proc in scoped_parts(statements):
        m = None if in_proc else CONST_STMT_RE.match(part)
        for decl in split_args(m.group(1)) if m else []:
            name = re.match(r"([A-Za-z]\w*)[%&^!#@$]?", decl)
            if name and not UPPER_SNAKE_RE.match(name.group(1)):
                out.append((n, f"constant {name.group(1)} is not UPPER_SNAKE_CASE"))
    return out


def banner_problems(lines: list[str], first: int, last: int) -> list[tuple[int, str]]:
    """Section banners in lines[first:last] that are not rule / label / rule
    with an UPPERCASE label one space after the apostrophe, as (line number,
    message). A label line sits between two banner lines; a column-0 comment
    with a banner on only one side is a broken block. Group banners (empty
    comment lines around a centered title) and intent comments (indented) are
    not section banners."""
    out: list[tuple[int, str]] = []
    for n in range(first, last):
        line = lines[n]
        before = n > first and BANNER_RE.match(lines[n - 1]) is not None
        after = n + 1 < last and BANNER_RE.match(lines[n + 1]) is not None
        if BANNER_RE.match(line) or not line.startswith("'") or line.rstrip() == "'":
            continue
        if before and after:
            text = line[1:].rstrip()
            label = text.strip()
            if not text.startswith(" ") or text.startswith("  "):
                out.append((n + 1, f'section label "{label}" is not left-aligned '
                                   "one space after the apostrophe"))
            elif label != label.upper() or not re.search("[A-Z]", label):
                out.append((n + 1, f'section label "{label}" is not UPPERCASE'))
        elif (before or after) and LABEL_COL0_RE.match(line):
            side = "closing" if before else "opening"
            out.append((n + 1, f'section label "{line[2:].strip()}" has no {side} '
                               "banner line"))
    return out


def title_problems(lines: list[str], first: int, last: int) -> list[tuple[int, str]]:
    """The procedure title banner in the header lines[first:last], if any: a
    banner, one title line and a banner. The title must be UPPERCASE and
    centered in the 78 columns after the apostrophe, with (78 - len(title))
    // 2 leading spaces (SKILL.md section 2)."""
    banner = next((n for n in range(first, last) if BANNER_RE.match(lines[n])), None)
    if banner is None or banner + 2 >= last or BANNER_RE.match(lines[banner + 1]) \
            or not BANNER_RE.match(lines[banner + 2]):
        return []
    text = lines[banner + 1][1:].rstrip()
    title = text.strip()
    if not title:
        return []
    if title != title.upper() or not re.search("[A-Z]", title):
        return [(banner + 2, f'procedure title "{title}" is not UPPERCASE')]
    lead, want = len(text) - len(text.lstrip(" ")), max(0, (78 - len(title)) // 2)
    if lead != want:
        return [(banner + 2, f'procedure title "{title}" is not centered '
                             f"({lead} leading spaces, expected {want})")]
    return []


def alignment_problems(lines: list[str], first: int, last: int) -> list[tuple[int, str]]:
    """Declarations in lines[first:last] whose "As" or trailing comment is not
    in the column of the first declaration of their group, as (line number,
    message). A group is a run of consecutive one-line declarations; a blank
    line, a comment line, a statement or a continued line ends it. One
    finding per group and kind."""
    groups: list[list[int]] = []
    run: list[int] = []
    continued = False
    for n in range(first, last):
        wraps = code_view(tokenize(lines[n])).endswith(" _")
        if not continued and not wraps and ALIGN_DECL_RE.match(lines[n]):
            run.append(n)
        elif run:
            groups.append(run)
            run = []
        continued = wraps
    if run:
        groups.append(run)
    out: list[tuple[int, str]] = []
    for run in groups:
        as_cols = [ALIGN_DECL_RE.match(lines[n]).start(1) for n in run]
        cmt_cols = {}
        for n in run:
            at = 0
            for kind, text in tokenize(lines[n]):
                if kind == COMMENT and text.startswith("'"):
                    cmt_cols[n] = at
                at += len(text)
        for label, cols in (("As", dict(zip(run, as_cols))),
                            ("trailing comment", cmt_cols)):
            if not cols:
                continue
            ref = cols[min(cols)]
            bad = next((n for n in sorted(cols) if cols[n] != ref), None)
            if bad is not None:
                out.append((bad + 1, f"{label} is not aligned with its group "
                                     f"(column {cols[bad] + 1}, expected {ref + 1})"))
    return sorted(out)


def statement_findings(lines: list[str], start: int) -> list[tuple[str, int, str]]:
    """Statement-level conventions (SKILL.md section 7), as WARN findings."""
    out: list[tuple[str, int, str]] = []
    statements = logical_lines(lines, start)
    for n, code in statements:
        if not code.strip() or ATTRIBUTE_RE.match(code):
            continue
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
            if not VISIBILITY_RE.match(code):
                out.append(("WARN", n, f"{m.group(2)}: no Public, Private or Friend"))
    out += [("WARN", n, msg) for n, msg in block_findings(statements)]
    out += [("WARN", n, msg) for n, msg in const_findings(statements)]
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
