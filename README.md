# MS Word CLI & API

This project is a modern evolution of the original [MSWord-CLI][https://github.com/waylan/msword-cli] package.

A command-line interface (CLI) and object-oriented Python API for automating
Microsoft Word using Windows [COM] technology and the [Word Object Model],
using Python's `pywin32`/`win32com` integration layer.

MS Word CLI & API allows you to control Microsoft Word from the command line and/or automate it from your own scripts.
Among other things, you may create, open, print, export, save, compare, merge, search, replace, review, and inspect Word documents.

[COM]: https://learn.microsoft.com/en-us/windows/win32/com/the-component-object-model
[Word Object Model]: https://learn.microsoft.com/en-us/office/vba/api/overview/word/object-model


## Features

- **Dual mode:** powerful CLI based on [Click]
  and a high-level Python API for use as an importable module.
- **Command pipelines:** subcommands can be chained together in a single call,
  loading Word only once regardless of how many operations are performed.
- **Smart resource management:** lazy client initialization — Word is not started
  unless an actual command requires it (e.g. `--help` is instant).
- **Plugin support:** third-party packages can register new CLI subcommands
  through the `msw.plugin` entry-point group.
- **Batch processing:** API workflows can reuse one Word client across many
  documents with explicit lifecycle control.


## Requirements and limitations

### Requirements

- **OS:** Windows - the project depends on Microsoft Word automation via Windows [COM].
- **Software:** A working copy of Microsoft Word installed.
- **Python:** `>=3.8`.
- **Dependencies:** [pywin32], [Click].

[click]: https://click.palletsprojects.com/
[pywin32]: https://github.com/mhammond/pywin32


### Limitations

- In case of a hard crash (e.g. `Ctrl+C` outside a context manager),
  `WINWORD.EXE` may remain running in the background.


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

First-party plugins are installed by default with `msword-cli`. The `save-as`,
`compare`, and `merge` commands are provided by bundled plugins, while
`export` remains the PDF/XPS-specific `ExportAsFixedFormat` route in the core.


### Command groups

| Group | Commands |
|---|---|
| Documents | `open`, `new`, `save`, `save-as`, `save-copy`, `close`, `list-documents`, `activate`, `compare`, `merge` |
| Content | `find`, `replace`, `print`, `export`, `update-fields` |
| Review | `track-changes`, `accept-revisions`, `reject-revisions`, `list-comments`, `export-comments`, `delete-comments` |
| Document Data | `summary`, `statistics`, `list-properties`, `get-property`, `set-property`, `delete-property` |

Bundled plugin commands appear in the CLI automatically and are listed in the
`Plugins` section of `msw --help`.

For a complete list of options for any subcommand, run:

```bash
msw <command> --help
```


### Common examples

```bash
msw open my.docx save-as renamed.docx close
msw open my.docx save-as renamed.pdf --format pdf close
msw save-as --list-formats
msw open --readonly my.docx summary close
msw open --hide my.docx summary
msw open my.docx find --format json invoice close
msw open draft.docx track-changes --on replace old new save close
msw open review.docx list-comments export-comments --json comments.json close
msw open report.docx update-fields export --pdf-a --with-properties . close
```

`summary` is the curated overview for document state and common metadata. The
`*-property*` commands expose raw built-in/custom Word properties directly.
### Listing open documents

```bash
$ msw list-documents

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

Options for a subcommand must appear immediately after that subcommand and before its positional arguments or the next command in the chain.

Hidden opens are session-scoped. If you use `open --hide` or `new --hide`,
the CLI will automatically close any document that it opened hidden and that
was not already open before the command started. This avoids leaving behind a
hidden Word document or background Word instance after the chain finishes.

Chaining also lets you run the same command twice with different options.
For example, to export to both PDF and XPS in one go:

```bash
msw export --pdf . export --xps .
```

Note: `.` (a single dot) refers to the current working directory. The export
command will resolve the output filename from the active document's name.


## Plugins

MSWord-CLI supports third-party plugins. Plugins add Click commands that can be
used in the same chain as built-in commands, and first-party plugins live under
[`plugins/`](plugins/).

Plugin packaging, discovery, install modes, precedence rules, and development
guidance are documented in [plugins/README.md](plugins/README.md).

Bundled plugin command docs:

- [`save-as`](plugins/save-as/README.md)
- [`compare-merge`](plugins/compare-merge/README.md)


## Library usage

The project exposes a high-level Python API built around `WordClient` and
`Document`. It writes nothing to `stdout` and raises `WordAPIError` on
failures, making errors straightforward to handle in larger scripts.

Plugins can also extend a specific `WordClient` instance. Load them explicitly
before using any plugin-provided methods:

```python
from msword_cli import WordClient

with WordClient(visible=False, quit_on_exit=True) as word:
    word.load_plugins(include="save-as")
    output = word.save_as("report.pdf", save_format="pdf")
    print(output)
```

`load_plugins()` accepts:

- `plugin_dir`: one optional local plugin root path as a string;
- `include`: `None`, one plugin name as a string, or a sequence of names.

Library plugins inject methods onto that `WordClient` instance only. They do
not modify the `WordClient` class globally.


### Example: batch processing folder conversion to PDF

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


### Advanced API: direct COM access via `native`

For advanced automation, `WordClient.native` exposes the underlying
`Word.Application` COM object and `Document.native` exposes the wrapped
`Word.Document`. This is intended as an escape hatch for plugins or scripts that
need Word features not yet covered by the high-level API.

```python
from msword_cli import WordClient

with WordClient(visible=False, quit_on_exit=True) as word:
    doc = word.open("report.docx", visible=False)
    selection = word.native.Selection
    selection.EndKey(Unit=6)  # wdStory
    selection.TypeText("\nAppended through COM")
    doc.native.Save()
```

`WordClient.native` raises `WordAPIError` after `quit()`. `Document.native`
returns the raw COM proxy as-is; any later COM errors from using that proxy are
not wrapped by `msword-cli`.


## Development notes

For test setup and execution, see [tests/README.md](tests/README.md).
