# Changelog

All notable changes to the `vba-house-style` plugin. Dates are in
Europe/Rome time. Versions follow `.claude-plugin/plugin.json`.

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
