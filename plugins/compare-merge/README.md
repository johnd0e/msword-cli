# Compare/Merge Plugin

The bundled `compare-merge` plugin provides the `compare` and `merge`
commands plus the dynamically injected `WordClient.compare(...)` and
`WordClient.merge(...)` library methods.


## CLI commands

The `compare` command wraps Word's
[`Application.CompareDocuments`][word-compare-docs] and produces a new
document with all differences shown as tracked changes:

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
(defaults to the Word username). Use `--ignore-warnings` to suppress any Word
comparison warning dialogs.


The `merge` command wraps [`Application.MergeDocuments`][word-merge-docs] and
combines the tracked changes from both documents:

```bash
msw merge original.docx revised.docx
```

It accepts the same `--no-*`, `--char-level`, `--author`, and
`--ignore-warnings` options as `compare`. The destination flags are
`--to-original` and `--to-revised` (default: new document).


## Library usage

Load the plugin explicitly before calling its library methods:

```python
from msword_cli import WordClient

with WordClient(visible=True, quit_on_exit=True) as word:
    word.load_plugins(include="compare-merge")
    diff = word.compare("original.docx", "revised.docx")
    print(f"Diff document: {diff.name}")
    diff.save("diff.docx", force=True)
```

[word-compare-docs]: https://learn.microsoft.com/en-us/office/vba/api/word.application.comparedocuments
[word-merge-docs]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._application.mergedocuments
