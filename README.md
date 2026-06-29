# MS Word CLI & API

*This project is a modern evolution of the original
[MSWord-CLI](https://github.com/waylan/msword-cli) package authored by
[Waylan Limberg](https://github.com/waylan).*

A command-line interface (CLI) and object-oriented Python API for automating
Microsoft Word using COM technology (`win32com`).

MS Word CLI & API allows you to control Microsoft Word from the command line
and/or automate it from your own scripts. Among other things, you may create,
open, print, export to PDF/XPS, save, and close Word documents. Note that this
tool does not actually edit the content of any documents.

## Features

- **Dual mode:** powerful CLI based on [Click](https://click.palletsprojects.com/)
  and a high-level Python API for use as an importable module.
- **Command pipelines:** subcommands can be chained together in a single call,
  loading Word only once regardless of how many operations are performed.
- **Batch processing:** convenient classes and context managers for automating
  many documents with explicit Word shutdown via `quit_on_exit=True`.
- **Smart resource management:** lazy client initialization — Word is not started
  unless an actual command requires it (e.g. `--help` is instant).
- **Plugin support:** third-party packages can register new CLI subcommands
  through the `msw.plugin` entry-point group.

## Differences from the original

This project intentionally stays close to the spirit of the original
`msword-cli`, but it is no longer just a thin CLI wrapper around Word.

Key differences:
- Drops Python 2 compatibility; targets Python `>=3.8`.
- Introduces a high-level library API (`WordClient`, `Document`) alongside the CLI.
- Cleaner internal architecture: explicit `WordAPIError` exceptions, a dedicated
  CLI wrapper layer, and a `@handle_api_error` decorator eliminating repetitive
  `try/except` blocks.
- Uses `pathlib.Path` throughout instead of `os.path`.
- Modern packaging via `pyproject.toml` and `uv run` usage.
- Adds `compare` and `merge` commands (not present in the original).
- Original tests were adapted to the new API/CLI architecture and extended for
  the newer compare/merge and library API behavior.

## Requirements and limitations

### Requirements

- **OS:** Windows (COM interface is Windows-only).
- **Software:** A working copy of Microsoft Word installed.
- **Python:** `>=3.8`.
- **Dependencies:** [pywin32](https://github.com/mhammond/pywin32),
  [Click](https://click.palletsprojects.com/).

### Limitations

- Does not work on Linux or macOS.
- In case of a hard crash (e.g. `Ctrl+C` outside a context manager),
  `WINWORD.EXE` may remain running in the background.
- Focuses on document lifecycle and export/print workflows; does not edit
  document content.

## Installation

The project is not currently published on PyPI. Clone the repository and use
[uv](https://github.com/astral-sh/uv) to install the `msw` command:

```bash
git clone https://github.com/johnd0e/msword-cli
cd msword-cli
uv tool install .
msw --help
```

For development, `uv sync` creates the project environment and installs the
package in editable mode:

```bash
git clone https://github.com/johnd0e/msword-cli
cd msword-cli
uv sync --dev
uv run msw --help
```

To use the Python API from another uv project, add the Git dependency:

```bash
uv add git+https://github.com/johnd0e/msword-cli.git
```

## CLI usage

If installed as a package, use the `msw` entry point:

```bash
msw --help
msw open my.docx export --pdf . close
```

Or run the script directly with `uv run` (no install needed):

```bash
uv run msword_cli.py --help
```

Unless otherwise specified, all subcommands operate on the currently active
document.

### Basic commands

| Command | Description |
|---|---|
| `open <path>` | Open an existing document and activate it |
| `new` | Create a new document, optionally from a template |
| `export <path>` | Export the active document to PDF or XPS |
| `print` | Print the active document |
| `save` | Save changes |
| `close` | Close the current document |
| `docs` | List all open documents |
| `activate <index>` | Activate an open document by its index |
| `compare <original> <revised>` | Compare two documents, showing differences as tracked changes |
| `merge <original> <revised>` | Merge two documents, combining their tracked changes |

For a complete list of options for any subcommand, run:

```bash
msw <command> --help
```

### Listing open documents

```bash
$ msw docs

Open Documents:

 * [1] doc1.docx
   [2] doc2.docx*
```

The `*` prefix marks the currently active document. The trailing `*` on a
filename indicates unsaved changes.

### Chaining commands

Subcommands can be chained in a single call. Word is loaded only once, which
is significantly faster than invoking the script multiple times:

```bash
msw open somedoc.docx print --copies 2 --pages "2-4, 6" close
```

Options for a subcommand must appear immediately after that subcommand and
before the next one in the chain.

Chaining also lets you run the same command twice with different options.
For example, to export to both PDF and XPS in one go:

```bash
msw export --pdf . export --xps .
```

Note: `.` (a single dot) refers to the current working directory. The export
command will resolve the output filename from the active document's name.

### Comparing documents

The `compare` command wraps Word's
[`Application.CompareDocuments`](https://learn.microsoft.com/en-us/office/vba/api/word.application.comparedocuments)
and produces a new document with all differences shown as tracked changes:

```bash
msw compare original.docx revised.docx
```

By default the result opens as a new document. Use `--to-original` or
`--to-revised` to put the diff inline:

```bash
msw compare original.docx revised.docx --to-revised
```

Comparison granularity is word-level by default; use `--char-level` for
character-level diff. Individual difference types can be excluded:

```bash
msw compare original.docx revised.docx --char-level --no-formatting --no-whitespace
```

Available `--no-*` flags: `--no-formatting`, `--no-case-changes`,
`--no-whitespace`, `--no-tables`, `--no-headers`, `--no-footnotes`,
`--no-textboxes`, `--no-fields`, `--no-comments`, `--no-moves`.

Use `--author <name>` to override the author attributed to tracked changes
(defaults to the Word username). Use `--ignore-warnings` to suppress any
Word comparison warning dialogs.

### Merging documents

The `merge` command wraps
[`Application.MergeDocuments`](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._application.mergedocuments)
and combines the tracked changes from both documents:

```bash
msw merge original.docx revised.docx
```

It accepts the same `--no-*`, `--char-level`, `--author`, and
`--ignore-warnings` options as `compare`. The destination flags are
`--to-original` and `--to-revised` (default: new document).

## Library usage

The project exposes a high-level Python API built around `WordClient` and
`Document`. It writes nothing to `stdout` and raises `WordAPIError` on
failures, making errors straightforward to handle in larger scripts.

### Example: batch folder conversion to PDF

The `with` context manager ensures Word is closed correctly even if an error
occurs:

```python
from pathlib import Path
from msword_cli import WordClient, WordAPIError


def convert_folder_to_pdf(folder_path: str) -> None:
    folder = Path(folder_path)
    with WordClient(visible=False, quit_on_exit=True) as word:
        for docx_file in folder.glob("*.docx"):
            doc = None
            try:
                doc = word.open(str(docx_file), visible=False)
                pdf_path = doc.export_fixed_format(str(folder))
                print(f"Saved: {pdf_path}")
            except WordAPIError as e:
                print(f"Error processing {docx_file.name}: {e}")
            finally:
                if doc is not None:
                    doc.close(force=True)


if __name__ == "__main__":
    convert_folder_to_pdf(r"C:\Users\User\Documents")
```

### Example: compare two documents via API

```python
from msword_cli import WordClient

with WordClient(visible=True, quit_on_exit=True) as word:
    diff = word.compare("original.docx", "revised.docx")
    print(f"Diff document: {diff.name}")
    diff.save("diff.docx", force=True)
```

## Plugins

MSWord-CLI supports third-party plugins. A plugin adds one or more Click
commands that can be included in any chain alongside the built-in ones.

Create a file `msw_import.py`:

```python
import click

@click.command('import')
def imprt():
    '''Import data into the active document.'''
    click.echo('Importing data...')
```

Register it in your `pyproject.toml`:

```toml
[project.entry-points."msw.plugin"]
import = "msw_import:imprt"
```

Sync the editable development environment so the plugin entry point is picked
up immediately:

```bash
uv sync --dev
```

The new command will appear in the project environment and can be used in
chains:

```bash
uv run msw open data.docx import close
```

## Development notes

For test setup and execution, see [tests/README.md](tests/README.md).

---

*Original project: [waylan/msword-cli](https://github.com/waylan/msword-cli)
— [README](https://github.com/waylan/msword-cli/raw/refs/heads/master/README.rst)*
