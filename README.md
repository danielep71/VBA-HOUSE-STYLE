<div align="center">

# 🎨 VBA-HOUSE-STYLE

**One skill, one style: exported VBA that reads the same in every repository.**

[![Plugin](https://img.shields.io/badge/plugin-v1.3.0-6f42c1?style=for-the-badge)](.claude-plugin/plugin.json)
[![VBA](https://img.shields.io/badge/VBA-.bas%20%C2%B7%20.cls%20%C2%B7%20.frm-217346?style=for-the-badge)](skills/vba-house-style/references/M_Example.bas)
[![Python](https://img.shields.io/badge/python-3%20%C2%B7%20stdlib%20only-3776AB?style=for-the-badge&logo=python&logoColor=white)](skills/vba-house-style/scripts/check_vba_style.py)
[![EOL](https://img.shields.io/badge/line%20endings-CRLF-555555?style=for-the-badge)](.gitattributes)
[![License](https://img.shields.io/badge/license-MIT-2ea44f?style=for-the-badge)](LICENSE)

[✨ Features](#-features) ·
[👀 At a glance](#-the-style-at-a-glance) ·
[📁 Layout](#-layout) ·
[🔍 Checker](#-the-checker) ·
[🚀 Release](#-release-a-new-version) ·
[💻 Local install](#-local-only-install) ·
[📝 Changelog](CHANGELOG.md) ·
[🏷️ Issue labels](docs/ISSUE_LABELS.md)

</div>

---

## ✨ Features

| | Rule | In short |
|---|---|---|
| 📏 | **Banners** | Every banner is exactly 79 characters: `'` plus 78 dashes; procedure titles UPPERCASE and centered, section labels UPPERCASE |
| 🧾 | **Nine-field header** | `PURPOSE` · `WHY THIS EXISTS` · `INPUTS` · `RETURNS` · `BEHAVIOR` · `ERROR POLICY` · `DEPENDENCIES` · `NOTES` · `UPDATED`; small `Private` helpers may use a compact `PURPOSE` + `UPDATED` header |
| 🧱 | **Body sections** | `DECLARE` → `INITIALIZE` → task sections → `EXIT PROCEDURE` → `ERROR HANDLER` |
| ↘️ | **Two-level indent** | One-line intent comment at 4 spaces, code at 8 |
| 📐 | **Aligned declarations** | `As` and the trailing `'comment` in fixed columns; explicit `Public` / `Private` / `Friend`; module constants in `UPPER_SNAKE_CASE` |
| 🔤 | **Optional `VBA.`** | `Len` and `VBA.Len` are both accepted; qualify when a reference may go missing or a name collides |
| ↩️ | **CRLF, no tabs** | No trailing whitespace; the `.frm` designer block is never touched |
| 🛡️ | **Behavior preserving** | Formatting never renames public members, changes signatures or alters logic |

## 👀 The style at a glance

An excerpt from [`M_Example.bas`](skills/vba-house-style/references/M_Example.bas):

```vb
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Example.M_Greeting_Build"

    Dim NormalizedName          As String       'Trimmed person name
    Dim ErrorNumber             As Long         'Captured error number

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a blank name
        If Len(NormalizedName) = 0 Then
            Err.Raise vbObjectError + EX_ERR_BLANK_NAME, PROC_NAME, _
                "PersonName cannot be blank"
        End If
```

## 📁 Layout

```text
vba-house-style/
├── .claude-plugin/
│   ├── marketplace.json            🛒 lets claude.ai sync the plugin from this repo
│   └── plugin.json                 📦 name, version, description
├── .github/
│   ├── labels.json                 🏷️ issue label catalogue
│   ├── scripts/                    🔄 label sync and drift scripts
│   └── workflows/                  ⚙️ CI, label sync, label drift
├── docs/
│   └── ISSUE_LABELS.md             📋 how the labels are used
├── skills/
│   └── vba-house-style/
│       ├── SKILL.md                📘 the rules and the workflow
│       ├── references/
│       │   └── M_Example.bas       🧪 a complete worked example
│       └── scripts/
│           └── check_vba_style.py  🔍 checker; --fix normalizes whitespace only
├── tests/
│   └── test_check_vba_style.py     🧪 checker tests
├── tools/
│   └── build_plugin.py             🛠️ builds the uploadable ZIP
├── CHANGELOG.md                    📝 what changed in each version
└── LICENSE                         ⚖️ MIT
```

## 🔍 The checker

```powershell
python skills\vba-house-style\scripts\check_vba_style.py src\          # report
python skills\vba-house-style\scripts\check_vba_style.py --fix src\    # normalize, then report
```

| Level | Flags |
|---|---|
| ❌ **ERROR** | not CRLF · missing final CRLF · `VB_Name` ≠ file name · no `Option Explicit` · tabs · trailing whitespace · banner ≠ 79 chars |
| ⚠️ **WARN** | line > `--max-line` (default 120; string contents and declaration comments not counted) · missing `MODULE:` / `CLASS:` header · header fields missing or out of order · blank lines between procedures ≠ 1 · `Next` without its counter · `:` between statements · `InStr` / `StrComp` without a compare mode · parameters without `ByVal`/`ByRef` or `As` · function without `As` · `= ""` instead of `vbNullString` · procedure without `Public` / `Private` / `Friend` · module-level constant not `UPPER_SNAKE_CASE` · `As` or trailing `'comment` not aligned within a group of declarations (module top and `DECLARE` sections) · procedure title not UPPERCASE or not centered in the 78 columns · section label not UPPERCASE, not one space after the apostrophe, or without its opening or closing banner line · `Select Case` without `Case Else` · local variable with a Hungarian prefix (`strName`) or not PascalCase · local variables or error handler without a `DECLARE` section · `DECLARE` not starting with `Const PROC_NAME` or a `PROC_NAME` other than `<Component>.<Procedure>` · first `On Error GoTo` outside `INITIALIZE` · error label not at column 0 · body sections out of order · no `Exit Sub/Function/Property` in `EXIT PROCEDURE` before the error label (one-line and compact `Private` procedures exempt) |

`--fix` touches whitespace and line endings only, never code, and skips the
`.frm` designer block. The exit code is `1` while any error remains.

Tests (standard library only); CI runs the same command, plus a style check of
`M_Example.bas`, which must stay clean:

```powershell
python -m unittest discover -s tests
python skills\vba-house-style\scripts\check_vba_style.py skills\vba-house-style\references
```

## 🚀 Release a new version

1. ✏️ Change the skill under `skills/vba-house-style/`.
2. 🔢 Raise `version` in `.claude-plugin/plugin.json` (for example 1.0.0 to
   1.0.1) and the **plugin** badge at the top of this file. The build stops
   if the two differ.
3. 📝 In `CHANGELOG.md`, move `[Unreleased]` under the new version with its
   date.
4. 🔀 Merge the task pull requests into `release/vX.Y.Z`, then integrate that
   branch into `main`.
5. 🔄 On claude.ai, **Manage marketplaces > Check for updates** to pull the
   new version now; only a raised version reaches users.
6. 🆕 Open a new session: a running session keeps the version it started
   with.

### ☁️ One-time setup on claude.ai

**Customize > Plugins > Add > Add marketplace**, enter
`danielep71/VBA-HOUSE-STYLE` and turn on **Sync automatically**. Then open the plugin
`vba-house-style` and press **Add**: until then it is only listed, not
installed, and gets no updates. After each release, **Check for updates** pulls the new
version.

### 📦 Fallback: upload a ZIP

```powershell
python tools\build_plugin.py
```

It writes `dist\vba-house-style-plugin-<version>.zip`; the latest **CI** run
on `main` has the same ZIP as the artifact `vba-house-style-plugin`. Upload it
under **Plugins > Add > Upload plugin**.

## 💻 Local-only install

Without claude.ai, copy the skill folder into your user skills:

```powershell
$dest = Join-Path $HOME '.claude\skills'
git clone https://github.com/danielep71/VBA-HOUSE-STYLE.git "$env:TEMP\vba-house-style"
Copy-Item "$env:TEMP\vba-house-style\skills\vba-house-style" $dest -Recurse -Force
```

> [!WARNING]
> Install one way only: the claude.ai plugin already syncs to a signed-in
> local Claude Code, and two copies (marketplace, ZIP upload, or folder) would
> show up as two skills.

---

<div align="center">
<sub>MIT licensed · see <a href="LICENSE">LICENSE</a></sub>
</div>
