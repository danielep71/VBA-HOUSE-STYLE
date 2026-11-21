# VBA-HOUSE-STYLE

Plugin with one skill that formats exported VBA source (`.bas`, `.cls`,
`.frm`) in a fixed house style: 79-column banners, a nine-field documentation
header, section banners inside procedures, aligned declarations, CRLF line
endings.

## Layout

- `.claude-plugin/plugin.json` - plugin name, version, description
- `skills/vba-house-style/SKILL.md` - the rules and the workflow
- `skills/vba-house-style/references/M_Example.bas` - a complete worked example
- `skills/vba-house-style/scripts/check_vba_style.py` - checker; `--fix`
  normalizes whitespace only
- `tools/build_plugin.py` - builds the uploadable ZIP

## Release a new version

1. Change the skill under `skills/vba-house-style/`.
2. Raise `version` in `.claude-plugin/plugin.json` (for example 1.0.0 to 1.0.1).
3. Build the ZIP (needs Python 3 and Git):

   ```powershell
   python tools\build_plugin.py
   ```

   It writes `dist\vba-house-style-plugin-<version>.zip`.
4. On claude.ai open Plugins, open "Vba house style", and upload the new ZIP
   (or **Add > Upload plugin**). New sessions pick up the new version, in the
   cloud and in a signed-in local Claude Code.

## Local-only install (without claude.ai)

```powershell
$dest = Join-Path $HOME '.claude\skills'
git clone https://github.com/danielep71/VBA-HOUSE-STYLE.git "$env:TEMP\vba-house-style"
Copy-Item "$env:TEMP\vba-house-style\skills\vba-house-style" $dest -Recurse -Force
```

Do not install both ways at once: the claude.ai plugin already syncs to a
signed-in local Claude Code, and two copies would show up as two skills.
