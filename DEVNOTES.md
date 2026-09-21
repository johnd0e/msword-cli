# Developer Notes

This file captures repository constraints that materially affect design
decisions. Keep it focused on implementation boundaries, runtime quirks, and
"why the CLI looks like this" notes that are easy to forget later.


## Click chaining constraints

The root `msw` CLI is intentionally defined as a chained Click group:

- `msword_cli.py` declares the root command with `@click.group(chain=True)`

This enables the project's command style:

```text
msw open input.docx save-as output.pdf --format pdf close
```

That choice imposes an important Click limitation:

- Click does not allow nested groups below a `chain=True` group.

Practical consequence:

- commands that must participate in the existing chained workflow cannot be
  modeled as nested groups such as `msw export pdf ...`
- they need to be exposed as flat commands or as option-driven behavior on
  existing commands

Document lifecycle invariant:

- `WordClient.open()` and `WordClient.new()` activate the returned document,
  including when `visible=False`; `--hide` controls visibility, not which
  document subsequent chained commands target.
- Hidden-document cleanup runs after the command chain, including after an
  error. Cleanup failures are warnings and must not replace the original
  command error.

This is a framework constraint, not a stylistic preference. If a future design
wants nested command namespaces under `msw`, that would require reworking the
root CLI structure away from the current chained model.

Reference:

- Click documentation, "Advanced Groups and Context -> Command Chaining":
  https://click.palletsprojects.com/en/stable/commands/#command-chaining


## Prefer flags over command growth

When a feature is still one conceptual operation, prefer extending an existing
command with options instead of adding more top-level commands.

Use additional commands only when they represent a meaningfully different
workflow rather than a mode switch.

Examples:

- prefer extending a command with a mode switch when the behavior is still the
  same conceptual operation
- prefer format or validation switches such as `--data-format`,
  `--strict/--no-strict`, and `--overwrite` over command proliferation

Examples from core-style CLI design:

- current style:
  `msw open draft.docx track-changes --off save close`
- alternate command-oriented design that would also be plausible:
  `msw open draft.docx disable-track-changes save close`
- current style:
  `msw open review.docx export-comments --json comments.json close`
- alternate command-oriented design that would also be plausible:
  `msw open review.docx export-comments-json comments.json close`
- current style:
  `msw open report.docx list-properties --kind custom close`
- alternate command-oriented design that would also be plausible:
  `msw open report.docx list-custom-properties close`

Use a new command only when it is actually a different workflow, not just a
different mode of an existing one.


## Why The Core Does Not Follow This Rigidly

`Prefer flags over command growth` is a default design bias, not an absolute
rule. The current core already shows why exceptions are sometimes correct.

Use a separate command when one or more of these are true:

- the operation has a meaningfully different effect on the document
- the operation changes the lifecycle model or result type
- combining the behaviors would mix inspection and mutation in one command
- the command maps to a distinct Word operation with its own semantics
- a subgroup would be the cleanest model, but Click chaining rules rule that
  out, so flat commands are the next-best option

Examples from the current core:

- `find` vs `replace`
  - these are closely related, but `find` is inspection and `replace` mutates
    document content
- `save`, `save-copy`, `save-as`, and `export`
  - these all concern document output, but they differ in lifecycle,
    destination semantics, and underlying Word APIs
- `summary`, `statistics`, `list-properties`, `get-property`, `set-property`,
  and `delete-property`
  - these form a family, but they are not just output modes of one operation;
    they represent different read/write intents, and a nested namespace is not
    available under the chained root CLI

In short:

- prefer flags when the user is still performing one conceptual operation
- prefer a separate command when the user intent, side effects, or result model
  are materially different


## Plugin notes live under plugins/

Plugin-specific architectural rules now live in `plugins/DEVNOTES.md`.

Keep root `DEVNOTES.md` focused on core CLI/API constraints. Use the plugin
notes file for:

- plugin manifest rules;
- `WordClient.load_plugins()` behavior;
- CLI/library sharing patterns inside plugins;
- plugin-specific laziness and validation requirements.
