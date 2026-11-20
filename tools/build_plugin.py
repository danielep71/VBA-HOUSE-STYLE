#!/usr/bin/env python3
"""Build the uploadable plugin ZIP from the tracked files.

Usage:
    python tools/build_plugin.py

Writes dist/vba-house-style-plugin-<version>.zip, where <version> comes from
.claude-plugin/plugin.json. The archive holds one top-level folder named after
the plugin, containing .claude-plugin/plugin.json and skills/. Only files
tracked by Git are included, so local scratch files never ship. Entry paths
use forward slashes and keep each file's bytes as committed in the working
tree (the VBA example stays CRLF). Standard library only.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INCLUDE = (".claude-plugin/", "skills/")


def main() -> int:
    manifest = json.loads((ROOT / ".claude-plugin/plugin.json").read_text(encoding="utf-8"))
    name, version = manifest["name"], manifest["version"]
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        print(f"plugin.json version '{version}' is not MAJOR.MINOR.PATCH", file=sys.stderr)
        return 1

    skill_md = ROOT / "skills" / name / "SKILL.md"
    front = re.match(r"---\n(.*?)\n---\n", skill_md.read_text(encoding="utf-8"), re.S)
    if not front or not re.search(rf"^name: {re.escape(name)}$", front.group(1), re.M):
        print(f"{skill_md}: frontmatter name must be '{name}'", file=sys.stderr)
        return 1

    tracked = subprocess.run(
        ["git", "ls-files", "-z"], cwd=ROOT, check=True, capture_output=True
    ).stdout.decode("utf-8").split("\0")
    files = sorted(p for p in tracked if p.startswith(INCLUDE))

    out = ROOT / "dist" / f"{name}-plugin-{version}.zip"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in files:
            zf.write(ROOT / rel, f"{name}/{rel}")

    for rel in files:
        print(f"  {name}/{rel}")
    print(f"built {out.relative_to(ROOT).as_posix()} ({len(files)} files)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
