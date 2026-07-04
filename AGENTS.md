# Agent Guide

This repository is a Windows-only Microsoft Word COM CLI/API project. Treat
`msword_cli.py` as the main product surface and keep CLI behavior, API behavior,
tests, and docs aligned.

## Read These Files First

- `README.md` for product usage, installation, and supported workflows.
- `DEVNOTES.md` for architectural constraints and design decisions that must not
  be rediscovered during implementation.
- `plugins/README.md` for plugin architecture, discovery, install modes, and
  plugin development workflow.
- `tests/README.md` for the canonical test workflow, including sandbox-specific
  execution details.

`DEVNOTES.md` is a required project document, not optional background reading.
If a change touches CLI shape, command growth, plugin ergonomics, or other
design tradeoffs, review `DEVNOTES.md` before editing code.

## Change Rules

- Keep API logic out of thin Click wrappers where practical.
- Preserve lazy startup: `--help` and `--version` must not start Word.
- Raise `WordAPIError` in the API layer and translate it at the CLI layer.
- Keep plugin failures isolated so one broken entry point does not block later
  plugins.
- Keep documentation in sync with behavior changes. Update `README.md`,
  `DEVNOTES.md`, `plugins/README.md`, and `tests/README.md` when the change
  affects their scope.

## File Boundaries

Keep information clearly separated by file responsibility.

- `README.md` is for product and developer usage.
- `DEVNOTES.md` is for architectural constraints, design rationale, and
  implementation tradeoffs.
- `plugins/README.md` is for plugin-specific behavior, setup, discovery, and
  development procedure.
- `tests/README.md` is for test policy and execution procedure, including
  sandbox-specific guidance.
- `AGENTS.md` is for agent workflow rules and pointers to the canonical
  documents above.

Avoid duplicating the same operational guidance across multiple files when one
canonical source is enough. Prefer short references to the owning file over
copying instructions into another document. If ownership is unclear, resolve
that first instead of spreading the same guidance across the repo.

## Testing Is Mandatory

Any code or behavior change must be validated with tests before considering the
work complete. Do not skip test execution because a change looks small.

`tests/README.md` is the canonical source for:

- which test suites must run for a given change;
- how to run tests in normal development;
- how to run opt-in integration coverage against real Word;
- how to run tests correctly in restricted or sandboxed environments.

Do not duplicate or invent alternate test procedures in agent guidance. Follow
`tests/README.md`, and update that file if the test workflow changes.
