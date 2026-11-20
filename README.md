# vba-house-style

Skill that formats exported VBA source (`.bas`, `.cls`, `.frm`) in a fixed
house style: 79-column banners, a nine-field documentation header, section
banners inside procedures, aligned declarations, CRLF line endings.

- `SKILL.md` - the rules and the workflow
- `references/M_Example.bas` - a complete worked example
- `scripts/check_vba_style.py` - checker; `--fix` normalizes whitespace only

## Install (Windows PowerShell 5.1)

```powershell
$dest = Join-Path $HOME '.claude\skills\vba-house-style'
git clone <this-repo-url> $dest
```

Update later with `git -C $dest pull`.
