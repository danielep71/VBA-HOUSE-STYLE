# Changelog

All notable changes to the `vba-house-style` plugin. Dates are in
Europe/Rome time. Versions follow `.claude-plugin/plugin.json`.

## [1.3.0] - 2026-11-21

### Added

- Checker: WARN findings for procedure body sections: local variables or an
  error handler without a `DECLARE` section, `DECLARE` not starting with
  `Const PROC_NAME As String = "<Component>.<Procedure>"` (or naming another
  procedure), the first `On Error GoTo` outside `INITIALIZE`, the error label
  not at column 0, sections out of order (`DECLARE`, `INITIALIZE`, task
  sections, `EXIT PROCEDURE`, `ERROR HANDLER`), and no clean exit in
  `EXIT PROCEDURE` before the error label (#7). One-line procedures and
  compact `Private` procedures are exempt. Exit code and `--fix` are
  unchanged.
- Checker: WARN for VBA library calls written without `VBA.` (`Len`, `Trim$`,
  `UCase$`, `InStr`, `CStr`...; the list is `VBA_QUALIFIED`) (#8). Members,
  names the file declares, the `Mid$` statement, the `Len =` clause of `Open`
  and `Len` / `LenB` of a user-defined type, an object or a `With` member
  (`Item.Inner`, `.Inner`) are not reported. No `--fix`.
- Checker: WARN for a `Select Case` without `Case Else` (nested blocks
  matched to their own `End Select`) and for a local variable (`Dim` or
  `Static` inside a procedure) with a Hungarian prefix (`strName`, `wsData`)
  or a name that is not PascalCase (#9). Module-level variables, constants
  and parameters are not checked; `Strength` and `Intent` are not reported.
  No `--fix`.

### Changed

- Checker: one tokenizer (code, string, date literal, comment) now feeds every
  statement rule, replacing the per-rule regular expressions. What the rules
  report on valid VBA is unchanged (#6). Two edge cases follow the new
  classification: a comment (apostrophe or `Rem`) counts in full toward the
  line-length limit, even when it quotes text, and a `Rem` after a
  declaration is exempt like an apostrophe comment.

## [1.2.0] - 2026-11-21

### Added

- Checker: WARN findings for statement-level conventions: `Next` without its
  counter, `:` between statements, `InStr` / `StrComp` without a compare mode,
  parameters without `ByVal` / `ByRef` or `As`, a function without `As`, and
  `= ""` instead of `vbNullString` (#3). Exit code and `--fix` are unchanged.
- Issue label catalogue (`.github/labels.json`, 20 labels) with the sync and
  drift workflows and a guide, `docs/ISSUE_LABELS.md`.
- `LICENSE` (MIT); the repository is now public.

### Removed

- `AGENTS.md` and `CLAUDE.md` (working instructions, kept out of the public
  repository).

## [1.1.1] - 2026-11-21

### Changed

- Raised the version to test the marketplace sync on claude.ai.

### Documentation

- README: the claude.ai steps after each release (**Check for updates**, then
  a new session) and the one-time setup, including pressing **Add** on the
  plugin so it gets updates.

## [1.1.0] - 2026-11-21

### Changed

- Calibrated the checker on real code: small `Private` helpers may use a
  compact `PURPOSE` + `UPDATED` header; extra header sections are allowed;
  string contents and declaration comments do not count toward line length.

### Added

- Unit tests for the checker (standard library only).
- The plugin build stops when the README badge version differs from
  `plugin.json`.
- CI workflow: runs the tests and builds the plugin ZIP as an artifact.
- `AGENTS.md` with the working instructions for agents.
- `.claude-plugin/marketplace.json`: claude.ai can sync the plugin from this
  repository.

## [1.0.0] - 2026-11-20

### Added

- The `vba-house-style` skill: rules, workflow, a worked example
  (`M_Example.bas`) and the style checker with `--fix`.
- Plugin layout and the ZIP build script.
