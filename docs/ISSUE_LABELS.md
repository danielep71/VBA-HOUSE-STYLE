# Issue labels

This repository defines 20 issue labels, each with a name, a description and a
colour. `.github/labels.json` is the source of truth: change a label through a
pull request, never in the GitHub UI. The catalogue follows the one of K-TEMPLATE;
only the `P1` description differs, because this repository has no calculations.

## Automation

| Workflow | Purpose |
| --- | --- |
| Sync issue labels | Brings the live labels to the catalogue after a change to it lands on `main`, or on a manual run on `main`. Refuses to delete a label that an issue or pull request still carries, and checks the exact result. |
| Detect issue-label drift | Compares the live labels with the catalogue every week and on demand. Read-only: it reports drift and never changes a label. |
| CI | Runs the self-tests of both scripts on every pull request. |

`prune: true` removes live labels that are outside the catalogue, but only when
nothing uses them. Preview a run with
`node .github/scripts/labels-sync.mjs --mode reconcile --dry-run`.

## Using the labels

Every issue has exactly one priority and the type labels that fit the work.

### Priorities

| Label | Use |
| --- | --- |
| `P1` | Only together with `bug`: the checker gives a wrong result on valid VBA, or `--fix` damages the source. Never for documentation, tooling or release work. |
| `P2` | Material correctness, robustness or maintainability defect. |
| `P3` | Non-blocking quality, testing, documentation, setup or technical-debt work. |

### Types

`bug`, `enhancement`, `documentation`, `behavior-change` (changes what the
checker reports or what `--fix` rewrites), `refactor`, `tests`, `ci`,
`repository`, `release` (version, changelog, tag, marketplace sync) and
`security`.

### Status

`blocked` while a named prerequisite is unresolved (remove it once resolved),
`question` when a decision or information is needed, `help wanted`,
`good first issue`, and the closing labels `duplicate`, `invalid` and
`wontfix`, which go with closing the issue and a one-line reason.

## Catalogue

| Label | Colour | Description |
| --- | --- | --- |
| `behavior-change` | `#C24E00` | Intentional externally observable behavior change |
| `blocked` | `#B60205` | Cannot proceed until a named dependency is complete |
| `bug` | `#D73A4A` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `ci` | `#24292F` | Automated quality checks, workflow maintenance, or dependency updates |
| `documentation` | `#0075CA` | Documentation for users or for my repository maintenance |
| `duplicate` | `#CFD3D7` | An existing issue or pull request already covers this work |
| `enhancement` | `#A2EEEF` | New capability or meaningful improvement to supported behavior |
| `good first issue` | `#7057FF` | Well-scoped task suitable for a first contribution |
| `help wanted` | `#008672` | I welcome external investigation, implementation, validation, or review |
| `invalid` | `#E4E669` | Report is incomplete, unsupported, or not actionable as submitted |
| `P1` | `#B60205` | Confirmed checker defect: wrong result on valid VBA, or --fix damages source (only with bug) |
| `P2` | `#D93F0B` | Material correctness, robustness, or maintainability defect |
| `P3` | `#FBCA04` | Non-blocking quality, testing, documentation, setup or technical-debt work |
| `question` | `#D876E3` | I need further information or need to make a decision |
| `refactor` | `#6F42C1` | Internal restructuring intended to preserve supported behavior |
| `release` | `#B60276` | Release preparation, certification, evidence, packaging, or publication |
| `repository` | `#1D76DB` | Repository structure, governance, metadata, tooling, or maintenance |
| `security` | `#8B0000` | Security boundary, vulnerability handling, supply chain, or hardening |
| `tests` | `#006B75` | Regression coverage, test harnesses, fixtures, or manual validation |
| `wontfix` | `#6E7781` | I have decided not to implement or continue this work |

## First run

Run **Sync issue labels** from `main` once Actions is enabled; it also runs by
itself whenever `.github/labels.json` changes on `main`.
