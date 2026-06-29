# MS Word CLI & API

*Note: This project is a modern evolution of the original [MSWord-CLI](https://github.com/waylan/msword-cli) package authored by [Waylan Limberg](https://github.com/waylan).* 

A command-line interface (CLI) and object-oriented Python API for automating Microsoft Word using COM technology (`win32com`).

MS Word CLI & API allows you to control Microsoft Word from the command line and/or automate it from scripts. Among other things, you may create, open, print, export to PDF/XPS, save, and close Word documents. Note that this tool does not actually edit the content of any documents.

## Features

- **Dual mode:** powerful CLI based on `click` and a clean Core API for use as an importable module.
- **Command pipelines:** subcommands can be chained together in a single line.
- **Batch processing:** convenient classes and context managers for safely automating many documents.
- **Smart resource management:** lazy client initialization prevents unnecessary Word startup for simple CLI help output.
- **Plugin support:** third-party packages can add new CLI subcommands through `msw.plugin` entry points.

## Differences from the original

This project intentionally stays close to the spirit of the original `msword-cli`, but it is no longer just a thin CLI wrapper around Word.

Key differences:
- It drops Python 2 compatibility and targets modern Python (`>=3.8`).
- It introduces a high-level library API (`WordClient`, `Document`) in addition to the CLI.
- It uses a cleaner internal architecture with explicit API exceptions and a CLI wrapper layer.
- It is documented with modern packaging and `uv run` usage in mind.
- It is expected to evolve independently, so some original tests and behaviors may require adaptation.

## Requirements and limitations

### Requirements
- **OS:** Windows.
- **Software:** Microsoft Word installed.
- **Python:** `>=3.8`.
- **Dependencies:** `pywin32`, `click`.

### Limitations
- This tool does not work on Linux or macOS.
- In case of a hard crash outside a context manager, `WINWORD.EXE` may remain running in the background.
- The current implementation focuses on document lifecycle operations and export/print workflows, not rich content editing.

## CLI usage

You can run the script directly with `uv run`:

```bash
uv run msword_cli.py --help
uv run msword_cli.py open my.docx export --pdf . close
```

Unless otherwise specified, subcommands operate on the active document.

### Basic commands
- `open <path>` — open an existing document and activate it.
- `new` — create a new document, optionally from a template.
- `export <path>` — export the active document to PDF or XPS.
- `print` — print the active document.
- `save` — save changes.
- `close` — close the current document.
- `docs` — list all open documents.
- `activate <index>` — activate an open document by index.

### Chaining example

```bash
uv run msword_cli.py open somedoc.docx print --copies 2 --pages "2-4, 6" close
```

If options are specified for a subcommand, they must appear after that subcommand and before the next subcommand in the chain.

### Exporting to multiple formats

```bash
uv run msword_cli.py export --pdf . export --xps .
```

## Library usage

The project also exposes a high-level Python API built around `WordClient` and `Document`.

```python
from pathlib import Path
from msword_cli import WordClient


def convert_folder_to_pdf(folder_path: str) -> None:
    folder = Path(folder_path)
    with WordClient(visible=False, quit_on_exit=True) as word:
        for docx_file in folder.glob("*.docx"):
            doc = None
            try:
                doc = word.open(str(docx_file), visible=False)
                doc.export_fixed_format(str(folder))
            finally:
                if doc is not None:
                    doc.close(force=True)
```

## Plugins

Plugins are registered through the `msw.plugin` entry-point group and can add new CLI commands.

Example `pyproject.toml` snippet for a plugin:

```toml
[project.entry-points."msw.plugin"]
import = "msw_import:imprt"
```

Original project links:
- [Original repository](https://github.com/waylan/msword-cli)
- [Original README](https://github.com/waylan/msword-cli/raw/refs/heads/master/README.rst)
