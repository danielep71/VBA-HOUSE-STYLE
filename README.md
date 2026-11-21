<div align="center">

# 🎨 VBA-HOUSE-STYLE

**One skill, one style: exported VBA that reads the same in every repository.**

[![Plugin](https://img.shields.io/badge/plugin-v1.1.0-6f42c1?style=for-the-badge)](.claude-plugin/plugin.json)
[![VBA](https://img.shields.io/badge/VBA-.bas%20%C2%B7%20.cls%20%C2%B7%20.frm-217346?style=for-the-badge)](skills/vba-house-style/references/M_Example.bas)
[![Python](https://img.shields.io/badge/python-3%20%C2%B7%20stdlib%20only-3776AB?style=for-the-badge&logo=python&logoColor=white)](skills/vba-house-style/scripts/check_vba_style.py)
[![EOL](https://img.shields.io/badge/line%20endings-CRLF-555555?style=for-the-badge)](.gitattributes)
[![Private](https://img.shields.io/badge/repo-private-critical?style=for-the-badge&logo=github)](#)

[✨ Features](#-features) ·
[👀 At a glance](#-the-style-at-a-glance) ·
[📁 Layout](#-layout) ·
[🔍 Checker](#-the-checker) ·
[🚀 Release](#-release-a-new-version) ·
[💻 Local install](#-local-only-install)

</div>

---

## ✨ Features

| | Rule | In short |
|---|---|---|
| 📏 | **Banners** | Every banner is exactly 79 characters: `'` plus 78 dashes |
| 🧾 | **Nine-field header** | `PURPOSE` · `WHY THIS EXISTS` · `INPUTS` · `RETURNS` · `BEHAVIOR` · `ERROR POLICY` · `DEPENDENCIES` · `NOTES` · `UPDATED`; small `Private` helpers may use a compact `PURPOSE` + `UPDATED` header |
| 🧱 | **Body sections** | `DECLARE` → `INITIALIZE` → task sections → `EXIT PROCEDURE` → `ERROR HANDLER` |
| ↘️ | **Two-level indent** | One-line intent comment at 4 spaces, code at 8 |
| 📐 | **Aligned declarations** | `As` and the trailing `'comment` in fixed columns |
| 🔤 | **Qualified built-ins** | `VBA.Len`, `VBA.Trim$`, `VBA.CStr`… |
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
        If VBA.Len(NormalizedName) = 0 Then
            Err.Raise vbObjectError + EX_ERR_BLANK_NAME, PROC_NAME, _
                "PersonName cannot be blank"
        End If
```

## 📁 Layout

```text
vba-house-style/
├── .claude-plugin/
│   └── plugin.json                 📦 name, version, description
├── skills/
│   └── vba-house-style/
│       ├── SKILL.md                📘 the rules and the workflow
│       ├── references/
│       │   └── M_Example.bas       🧪 a complete worked example
│       └── scripts/
│           └── check_vba_style.py  🔍 checker; --fix normalizes whitespace only
├── tests/
│   └── test_check_vba_style.py     🧪 checker tests
└── tools/
    └── build_plugin.py             🛠️ builds the uploadable ZIP
```

## 🔍 The checker

```powershell
python skills\vba-house-style\scripts\check_vba_style.py src\          # report
python skills\vba-house-style\scripts\check_vba_style.py --fix src\    # normalize, then report
```

| Level | Flags |
|---|---|
| ❌ **ERROR** | not CRLF · missing final CRLF · `VB_Name` ≠ file name · no `Option Explicit` · tabs · trailing whitespace · banner ≠ 79 chars |
| ⚠️ **WARN** | line > `--max-line` (default 120; string contents and declaration comments not counted) · missing `MODULE:` / `CLASS:` header · header fields missing or out of order · blank lines between procedures ≠ 1 |

`--fix` touches whitespace and line endings only, never code, and skips the
`.frm` designer block. The exit code is `1` while any error remains.

Tests (standard library only):

```powershell
python -m unittest discover -s tests
```

## 🚀 Release a new version

1. ✏️ Change the skill under `skills/vba-house-style/`.
2. 🔢 Raise `version` in `.claude-plugin/plugin.json` (for example 1.0.0 to
   1.0.1) and the **plugin** badge at the top of this file. The build stops
   if the two differ.
3. 📦 Build the ZIP (needs Python 3 and Git):

   ```powershell
   python tools\build_plugin.py
   ```

   It writes `dist\vba-house-style-plugin-<version>.zip`.
   Without a local build, download the same ZIP from the latest
   **CI** run on `main` (GitHub > Actions > CI > artifact
   `vba-house-style-plugin`).
4. ☁️ On claude.ai open **Plugins**, open "Vba house style", and upload the new
   ZIP (or **Add > Upload plugin**). New sessions pick up the new version, in
   the cloud and in a signed-in local Claude Code.

## 💻 Local-only install

Without claude.ai, copy the skill folder into your user skills:

```powershell
$dest = Join-Path $HOME '.claude\skills'
git clone https://github.com/danielep71/VBA-HOUSE-STYLE.git "$env:TEMP\vba-house-style"
Copy-Item "$env:TEMP\vba-house-style\skills\vba-house-style" $dest -Recurse -Force
```

> [!WARNING]
> Do not install both ways at once: the claude.ai plugin already syncs to a
> signed-in local Claude Code, and two copies would show up as two skills.

---

<div align="center">
<sub>🔒 Private · keep this skill out of public repositories</sub>
</div>
