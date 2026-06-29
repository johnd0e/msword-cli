# Future Expansion Analysis

## Current Functional Surface

The repository is no longer a thin historical wrapper around Microsoft Word.
It already has three distinct layers:

- a chained CLI built on `click`
- a Python automation API built around `WordClient` and `Document`
- a plugin model via the `msw.plugin` entry-point group

The current feature set includes:

- opening, creating, saving, closing, and activating documents
- listing open documents
- printing documents
- exporting to PDF and XPS
- comparing documents with tracked changes
- merging documents with tracked changes
- lazy startup and reuse of a single Word client during command chains
- plugin-based CLI extension
- unit tests with mocked COM objects and opt-in live integration tests against real Word

## Architectural Reading

The strongest parts of the current design are:

- clear separation between CLI error handling and API behavior
- a compact public surface concentrated in `msword_cli.py`
- good alignment between tests and runtime behavior
- explicit lifecycle handling for Word COM sessions

The main functional constraint is also clear:

- the project manages document lifecycle and Word-native operations well, but it does not yet model repeatable higher-level workflows across many files

That gap matters more than low-level editing support. Deep content editing would pull the project into the full complexity of the Word object model. The current architecture is better suited to orchestration than to rich document authoring.

## Recommended Expansion

The most natural next feature is batch job execution through a manifest-driven pipeline.

Proposed user-facing shape:

- `msw run job.yaml`

The manifest would describe a repeatable workflow such as:

- opening one or more files
- comparing document pairs
- merging reviewed variants
- exporting outputs to PDF or XPS
- saving result documents
- closing intermediates
- producing a final machine-readable report

## Why This Fits This Codebase

This extension fits the current repository better than most alternatives:

- the CLI already has a pipeline mindset through chained subcommands
- the API already encapsulates the operations needed for a batch runner
- the error model (`WordAPIError`) is already suitable for per-step reporting
- the live integration suite can validate real COM behavior for a small number of end-to-end scenarios
- plugins remain useful, but a manifest runner adds value even with no third-party ecosystem

## Suggested MVP

Keep the first version narrow and operationally useful:

- support YAML or JSON manifest input
- support `open`, `new`, `compare`, `merge`, `export`, `save`, and `close`
- support variable interpolation such as `{stem}`, `{suffix}`, `{outdir}`
- support `--dry-run`
- emit a JSON report with step status, output files, and failures

Example direction:

```yaml
jobs:
  - compare:
      original: input/original.docx
      revised: input/revised.docx
      destination: new
      output: out/diff.docx
  - export:
      source: out/diff.docx
      format: pdf
      output: out/
```

## Non-Goals For The First Iteration

Avoid expanding scope into:

- rich document content editing
- template rendering with business data
- parallel Word automation
- headless server execution outside Windows desktop Word

Each of those is either much riskier or a different product direction.

## Follow-Up Design Topics

When returning to this, the main design questions to settle are:

- whether manifests should describe document identities explicitly or implicitly through the active document
- whether batch execution should continue after a failed step or stop immediately
- whether reports should be JSON only or also human-readable Markdown
- whether plugin commands should be invocable from batch manifests

## Implementation Priority

If this is promoted into actual development work, the order should be:

1. Define the manifest schema and execution semantics.
2. Add a thin batch executor in the API layer.
3. Add a `run` CLI command that loads and validates the manifest.
4. Add mocked unit tests for planning/execution behavior.
5. Add one or two live Word integration scenarios for end-to-end validation.
