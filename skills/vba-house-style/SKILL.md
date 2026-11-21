---
name: vba-house-style
description: Format, write or review exported VBA source (.bas, .cls, .frm) in the house style - 79-column comment banners, a fixed nine-field documentation header on every module and procedure, DECLARE/INITIALIZE/.../EXIT PROCEDURE/ERROR HANDLER section banners, two-level statement indentation under one-line intent comments, column-aligned declarations, VBA.-qualified built-ins, CRLF line endings. Use when the user asks to format, restyle, tidy, document or review VBA code, or writes new VBA in a repository that should follow this style.
---

# VBA house style

Apply this style to exported VBA components: standard modules (`.bas`),
class modules (`.cls`) and the code part of UserForms (`.frm`). The binary
`.frx` file and the designer block at the top of a `.frm` (from `VERSION` to
the first `Attribute` line) are written by the VBE: never edit them.

`references/M_Example.bas` is a complete worked example of every rule below.
Read it before formatting anything; copy its shapes rather than inventing new
ones.

## Workflow

1. Read the whole file first. Note the component kind (MODULE / CLASS /
   FORM), the project prefix already in use (for example `DP_`, `INV_`) and
   any existing comments worth keeping.
2. Run the checker in report mode to see the mechanical problems:
   `python scripts/check_vba_style.py <files or folders>`
3. Reformat by hand following the rules below. Formatting is **behavior
   preserving**: never rename public members, change signatures, reorder
   `Attribute` lines, alter string literals, or change logic. If a rule would
   require a behavioral change (for example adding an error handler that
   changes what the caller sees), propose it instead of doing it.
4. Run `python scripts/check_vba_style.py --fix <files>` to normalize
   whitespace and line endings, then run it again in report mode until it is
   clean (or every remaining warning is a deliberate exception you can name).
5. Summarize what changed. Write header text in the repository's language
   (English by default) and never mention an AI assistant in comments.

## 1. File-level rules

- Line endings CRLF, file ends with CRLF, no tabs, no trailing whitespace.
  Exception: the `.frm` designer block keeps the VBE's bytes exactly.
- Encoding: leave as exported (Windows code page). Do not convert to UTF-8.
- Indent unit: 4 spaces.
- Order at the top of the code: the `Attribute` lines, one blank line,
  `Option Explicit` (plus `Option Private Module` / `Option Compare` when the
  file already uses them), one blank line, then the component header.
- Keep `.editorconfig` / `.gitattributes` consistent with this when the repo
  has them (`*.bas/*.cls/*.frm`: `eol=crlf`, 4 spaces).

## 2. Banners

Every banner line is exactly 79 characters: an apostrophe and 78 dashes.

```
'------------------------------------------------------------------------------
```

A test-harness module may use `=` instead of `-` for its component header
only; keep whichever one the file already uses.

Three banner kinds:

- **Section banner** (declarations sections and procedure-body sections):
  title left-aligned after one space, UPPERCASE.

  ```
  '------------------------------------------------------------------------------
  ' PRIVATE CONSTANTS
  '------------------------------------------------------------------------------
  ```

- **Procedure title banner**: title UPPERCASE, centered in the 78 columns
  after the apostrophe (leading spaces = `(78 - len(title)) // 2`). The title
  is a short verb phrase describing the procedure, not its identifier.

- **Group banner** (separates families of procedures in a long module): an
  empty comment line above and below a centered title, all wrapped in two
  banner lines and surrounded by one empty `'` line.

  ```
  '
  '------------------------------------------------------------------------------
  '
  '                                   SETTINGS
  '
  '------------------------------------------------------------------------------
  '
  ```

## 3. Component header

Immediately after `Option Explicit` and one blank line: a lone `'` line, then
the banner with `' MODULE: <Name>` or `' CLASS: <Name>` (UserForms and
document modules such as `ThisWorkbook` use `MODULE:`), then the nine fields
below, then a closing banner.

## 4. The nine documentation fields

Module, class and procedure headers all use the same fields, in this order,
each a single-space-indented UPPERCASE label followed by text indented three
spaces, with a lone `'` line between fields:

1. `PURPOSE` - one or two sentences on what it does.
2. `WHY THIS EXISTS` - the problem it solves or why it is shaped this way.
3. `INPUTS` - `None`, or each parameter name on its own line (3 spaces) with
   its description on the next lines (5 spaces), parameters separated by `'`.
   Module/class level: `None at module level` / `None at class level`.
4. `RETURNS` - `Nothing` for a Sub, otherwise the value and its edge cases.
5. `BEHAVIOR` - what happens, in order; bullet lists use `'     - item`.
6. `ERROR POLICY` - what raises, what is suppressed, what is best-effort.
   `None` is allowed for trivial getters.
7. `DEPENDENCIES` - object models, other components, members it calls.
8. `NOTES` - design decisions, caveats, call examples (code examples indented
   7 spaces). `None` when there is nothing to say.
9. `UPDATED` - the date of the last substantive change, `YYYY-MM-DD`.

No field is ever omitted; write `None` instead.

**Compact header.** A `Private` procedure whose body has at most 15 code
lines (comments and blank lines not counted) may use a compact header:
`PURPOSE` and `UPDATED`, plus any of the other fields that say something
useful, still in the order above. `Public`, `Friend` and event procedures
always use all nine.

**Extra sections.** A module, class or procedure header may add its own
UPPERCASE sections (for example `PUBLIC SURFACE`) after the standard fields
and before `UPDATED`, which always comes last.

Do not invent facts: if you
cannot tell from the code why something exists, write what the code shows and
flag the gap to the user rather than guessing. When you reformat without
changing behavior, keep the existing `UPDATED` date; when you change behavior,
set it to today.

## 5. Declarations section (top of module)

Use section banners, in this order, omitting empty ones:
`PRIVATE TYPES`, `WINDOWS API DECLARATIONS`, `PRIVATE CONSTANTS`,
`PUBLIC CONSTANTS`, `PUBLIC ENUMS`, `PUBLIC TYPES`, `PUBLIC STATE`,
`PRIVATE STATE` (classes: `PRIVATE CONSTANTS`, `PRIVATE VARIABLES`, ...).
A file may add its own clearly named sections (for example
`TEST FAULT INJECTION STATE`).

- Everything under a section banner is indented 4 spaces.
- Align `As` in one column and the trailing `'comment` in another, within each
  block; columns fall on multiples of 4. Every declaration gets a trailing
  comment (no space after the apostrophe: `'Live clock interval`).
- Blank lines separate logical groups inside a section.
- Naming: constants `UPPER_SNAKE_CASE` with the project prefix
  (`<PFX>_ERR_NO_OWNER`); Win32 constants keep their SDK names (`GWL_STYLE`);
  module-private variables `m` + PascalCase (`mIsInitialized`, or
  `m<PFX>_Name` in shared modules); public globals `g<PFX>_Name`;
  enums and types `<PFX>_PascalCase` (Win32 structs keep SDK names).
- API declarations: `#If Mac Then` / `#Else` / `#If VBA7 Then` ... `PtrSafe`
  ... `#Else` ... legacy `#End If`, nested 4 spaces per level; one parameter
  per line after `( _`, each `ByVal`/`ByRef` explicit.
- Error-number constants: one block per procedure family, documented with a
  comment line above the block.

## 6. Procedure layout

```
<Signature>
                                  <- one blank line
'                                 <- lone apostrophe
'------------------------------------------------------------------------------
'                           CENTERED TITLE
'------------------------------------------------------------------------------
' PURPOSE ... (nine fields, or compact) ... ' UPDATED
'------------------------------------------------------------------------------
                                  <- one blank line
<body sections>
                                  <- one blank line
End Sub
```

- Signature: explicit visibility (`Public` / `Private` / `Friend`), explicit
  `ByVal` / `ByRef` on every parameter, explicit `As` type on parameters and
  functions. A signature that does not fit on one line puts each parameter on
  its own line, indented 4, ending in `, _`; the closing `)` follows the last
  parameter. When the signature is multi-line, the lone `'` follows it
  directly with no blank line.
- Exactly one blank line between `End Sub/Function/Property` and the next
  signature or group banner block; two blank lines between the declarations
  section and the first group banner block; one blank line between a group
  banner block and the signature below it.
- Naming: in standard modules, public and private procedures are
  `<ModulePrefix>_<Area>_<Verb><Object>` (for example
  `M_Settings_GetNamespace`); event handlers keep VBA's required names
  (`UserForm_Initialize`, `mLabel_Click`); class members are PascalCase.
  Do not rename existing public procedures while formatting.

### Body sections

The body is a sequence of section banners. Standard ones, in this order:

| Section | Contains |
|---|---|
| `DECLARE` | `Const PROC_NAME As String = "<Component>.<Procedure>"`, then a blank line, then every `Dim` (aligned `As`, aligned `'comment`) |
| `INITIALIZE` | `On Error GoTo ErrorHandler` and initial state |
| task-specific sections | `VALIDATE INPUTS`, `NORMALIZE INPUTS`, `LOAD SETTINGS`, `STORE REFERENCES`, `RETURN VALUE`... - a short UPPERCASE verb phrase per logical step |
| `EXIT PROCEDURE` | the clean exit (`Exit Sub` / `Exit Function`, preceded by cleanup) |
| `ERROR HANDLER` | the label at column 0, then capture / cleanup / re-raise |

Trivial procedures (a one-line getter) keep only the task sections they need;
`DECLARE` / `INITIALIZE` / `EXIT PROCEDURE` / `ERROR HANDLER` appear only
when there is something to put in them.

### Statements: intent comment + double indent

Inside a section, every statement or small group of statements is preceded by
a one-line intent comment at 4 spaces (no space after `'`), and the code sits
at 8 spaces. Nested blocks add 4 spaces per level from there.

```
'------------------------------------------------------------------------------
' VALIDATE REFERENCES
'------------------------------------------------------------------------------
    'Reject a missing owner form
        If OwnerForm Is Nothing Then
            Err.Raise vbObjectError + <PFX>_ERR_NO_OWNER, PROC_NAME, _
                "OwnerForm cannot be Nothing"
        End If
    'Store the owner form reference
        Set mOwnerForm = OwnerForm
```

- The intent comment says *why/what*, in sentence case, no final period
  (a multi-line intent comment continues at 4 spaces).
- One intent comment can cover several closely related statements.
- Labels (`ErrorHandler:`, `CleanExit:`, `SafeExit:`) are at column 0.

### Error handling pattern

```
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    '...cleanup...
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error
        Err.Raise ErrorNumber, PROC_NAME, "<What failed>: " & ErrorDescription
```

Capture `Err` before any other call; raise custom errors as
`vbObjectError + <constant>` with `PROC_NAME` as the source.

## 7. Statement-level conventions

- Qualify VBA library calls: `VBA.Len`, `VBA.Trim$`, `VBA.UCase$`,
  `VBA.InStr`, `VBA.CStr`, `VBA.vbNullString`... Prefer the `$` string
  variants. Use `VBA.LenB(x) = 0` or `VBA.Len(x) = 0` for empty checks, and
  `vbNullString` rather than `""`.
- Explicit compare mode in `InStr` / `StrComp` (`vbBinaryCompare` or
  `vbTextCompare`).
- `Next` always names its counter (`Next CharIndex`, `Next Ctl`).
- One statement per line; no `:` statement separators. Single-line
  `If ... Then <one statement>` is allowed for guards and `Exit`.
- Line continuation: break after an operator or comma, continuation lines
  indented 4 more than the statement; for `Err.Raise` keep number and source
  on the first line and the message on the next.
- Local variables: PascalCase, descriptive (`NormalizedHoverMode`,
  `ShapeIndex`), no Hungarian prefixes; object variables released with
  `Set x = Nothing` in the exit/cleanup path.
- `Select Case` with `Case Else` whenever unexpected values matter.
- Keep lines readable (aim for 100 characters, hard limit 120). String
  literal contents and the trailing comment of an aligned declaration do not
  count toward the limit; the code around them does.

## What not to do

- Do not touch `Attribute` lines, the `.frm` designer block, or `.frx` files.
- Do not change behavior, public names, signatures or string contents while
  formatting.
- Do not convert the file encoding or line endings to LF.
- Do not add a file to the target repository that mentions an AI assistant,
  and do not copy this skill into a public repository.
