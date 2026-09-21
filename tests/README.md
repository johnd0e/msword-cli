# Tests

What is covered:
- CLI command behavior through the Click command group in `msword_cli.cli`
- High-level API behavior for `WordClient` and `Document`
- Bundled plugin behavior such as `save-as` and `compare-merge`, plus lazy
  client init, public `native` COM escape hatches, error conversion, and the
  newer review/content/property commands
  (`find`, `replace`, comments, revisions, fields, properties, summary, statistics)

Design notes:
- Tests use mocked COM / `pywin32` objects and do not require Microsoft Word to  be installed.
- The default suite is unit-level verification of command wiring and API behavior.
  Live Word checks are opt-in.

Run the suite from the project root:

```bash
./scripts/uv-run.ps1 sync --dev
./scripts/run-pytest.ps1
```

The required static check is:

```bash
uv run ruff check .
```

The project gate enables `E`, `F`, `I`, `B`, `UP`, `BLE001`, `DTZ001`, `FA100`, `FA102`, `PERF203`, `PERF401`, `PLR0911`, `PLR0912`, `PLR0913`, `PLR0917`, `PLR2004`, `SIM105`, `S110`, `S112`, and `S603`. `E501` is allowed only in `tests/**/*.py` and `plugins/**/tests/**/*.py`.
`S101` is scoped to those same test paths because it is the standard pytest
`assert` idiom. The enabled maintainability rules use local suppressions where the
existing public API or COM behavior intentionally requires the current shape.

### Development in sandboxed environments

In some restricted or sandboxed Windows environments, `uv` cannot use its default user cache.
For that case, the repository includes PowerShell wrappers in `scripts/`.

- `uv-run.ps1` wraps `uv` and preserves `UV_CACHE_DIR` if it is already set.
- If `UV_CACHE_DIR` is not set and the default user cache is not writable, it
  falls back to `%TEMP%\msword-cli-uv-cache`.
- `run-pytest.ps1` is the stable entry point for the test suite and delegates
  to `uv-run.ps1`.

This keeps automated and agent-driven runs stable without creating `.uv-cache`
or similar temp directories in the repository root.
In a normal local setup, plain `uv` commands are still fine.

Integration tests:
- Live Word COM tests live in `tests/integration/`.
- They are skipped unless `MSWORD_RUN_INTEGRATION=1` is set.
- The scenarios run serially against one hidden Word instance, which is closed after the integration session.
  Do not parallelize this test group.
- Run them explicitly on a Windows machine with Microsoft Word installed:

```bash
$env:MSWORD_RUN_INTEGRATION="1"
./scripts/run-pytest.ps1 -Integration
```
