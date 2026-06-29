# Tests

What is covered:
- CLI command behavior through the Click command group in `msword_cli.cli`
- High-level API behavior for `WordClient` and `Document`
- Newer compare/merge flows and helper behavior such as lazy client init and
  error conversion

Design notes:
- Tests use mocked COM / `pywin32` objects and do not require Microsoft Word to
  be installed.
- The default suite is unit-level verification of command wiring and API
  behavior. Live Word checks are opt-in.

Run the suite from the project root:

```bash
uv sync --dev
uv run pytest -q
```


Integration tests:
- Live Word COM tests live in `tests/integration/`.
- They are skipped unless `MSWORD_RUN_INTEGRATION=1` is set.
- The scenarios run serially against one hidden Word instance, which is closed
  after the integration session. Do not parallelize this test group.
- Run them explicitly on a Windows machine with Microsoft Word installed:

```bash
$env:MSWORD_RUN_INTEGRATION="1"
uv run pytest -q -m integration
```
