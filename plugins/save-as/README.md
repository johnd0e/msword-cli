# Save-As Plugin

The bundled `save-as` plugin provides the `save-as` command plus the
dynamically injected `WordClient.save_as(...)` library method.


## CLI command

Use `save-as` to save the active document to another path. You can keep the
same Word format or convert the document to another save format supported by
the current Word installation.

```bash
msw open draft.docx save-as renamed.docx close
msw open draft.docx save-as report.pdf --format pdf close
```

To inspect the formats available on the current machine:

```bash
msw save-as --list-formats
```

The command supports:

- built-in aliases such as `doc`, `docx`, `pdf`, `rtf`, and `xps`
- compositional `xml`, `html`, and `text` modes
- direct `wdFormat...` constants when Word exposes them through the
  [`WdSaveFormat` API][word-wdsaveformat]

Examples:

```bash
msw save-as report.pdf --format pdf
msw save-as archive.xml --format xml --document
msw save-as strict.xml --format xml --strict
msw save-as notes.txt --format text --unicode
msw save-as page.html --format html --filtered
```


## Library usage

Load the plugin explicitly before calling its library method:

```python
from msword_cli import WordClient

with WordClient(visible=False, quit_on_exit=True) as word:
    word.load_plugins(include="save-as")
    output = word.save_as("report.pdf", save_format="pdf")
    print(output)
```

[word-wdsaveformat]: https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
