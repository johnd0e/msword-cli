# Plugins

`msword-cli` supports plugins through the `msw.plugin` entry-point group.

First-party plugins live under `plugins/*` in this repository.
A plugin is a separate Python package that can expose:

- a Click command for the chained `msw` CLI;
- one or more library methods that can be injected onto a `WordClient`
  instance.

Plugin-specific architectural constraints and invariants live in
[`DEVNOTES.md`](DEVNOTES.md).

Bundled first-party plugins can ship by default with `msword-cli` while still
remaining outside `msword_cli.py`. Current bundled examples include:

- `save-as` for richer save-format handling;
- `compare-merge` for two-document compare/merge review workflows.

Plugin-specific command usage lives with each plugin package:

- [`plugins/save-as/README.md`](save-as/README.md)
- [`plugins/compare-merge/README.md`](compare-merge/README.md)


## Plugin package structure

Each plugin is its own package with a `pyproject.toml`. The package declares an
entry point in the `msw.plugin` group:

```toml
[project.entry-points."msw.plugin"]
save-as = "save_as:plugin_manifest"
```

The value uses `module:attribute` syntax and should point at a plain dictionary
named however you like. The project currently uses `plugin_manifest`.

Minimal shape:

```python
plugin_manifest = {
    "name": "save-as",
    "command": save_as_cmd,
    "client_methods": {
        "save_as": save_as,
    },
}
```

Field semantics:

- `name`: required unique plugin identifier.
- `command`: optional `click.Command` for CLI registration.
- `client_methods`: optional mapping of method names to callables.

Each library callable should accept the active `WordClient` instance as its
first argument:

```python
def save_as(client, path, save_format=None, ...):
    ...
```


## Discovery model

The runtime has two discovery paths:

- installed Python package metadata through `msw.plugin`
- explicit local plugin roots passed through `--plugin-dir PATH`

Installed plugins are discovered first through `importlib.metadata.entry_points()`.
Every discovered plugin is loaded independently, and a failing plugin does not block later plugins from loading.

When `--plugin-dir PATH` is provided, `msword-cli` scans that directory for
plugin subdirectories containing a `pyproject.toml` with `msw.plugin` entry
points.

Without `--plugin-dir`, a plugin is discovered automatically only when it is
installed into the active Python environment.

CLI plugin commands are registered automatically when discovered.

Library plugin methods are loaded explicitly per client:

```python
from msword_cli import WordClient

with WordClient() as word:
    word.load_plugins(include=("save-as", "compare-merge"))
    word.save_as("out.pdf", save_format="pdf")
```


## Installed vs editable installs

How the plugin is installed matters.


### Regular install

Examples:

```bash
pip install ./plugins/save-as
uv pip install ./plugins/save-as
```

This is a normal package installation. The installed package behaves like a
regular built distribution. In this mode, the installed plugin can diverge from
the source tree in this repository if you edit files after installation.

Use this mode when you want to test the packaged artifact as users would
consume it.


### Editable install

Examples:

```bash
pip install -e ./plugins/save-as
uv pip install -e ./plugins/save-as
```

Editable install keeps imports pointed at the source tree instead of a copied
package payload. In this mode, code changes in the repository are reflected by
the installed plugin immediately, while metadata changes may still require a
reinstall.

Use this mode for local plugin development.


### Workspace install with `uv`

In this repository, `uv sync` installs workspace packages as editable local
packages. That means first-party plugins such as `plugins/save-as` are typically
linked to the source tree during normal repository development.


## Development recommendations

If you are developing a plugin locally, prefer one of these workflows:

1. Do not install a separate copy of the plugin package at all.
2. Install the plugin as editable.

Avoid regular non-editable installation during active plugin development unless
you explicitly want to test the packaged install behavior, because the
installed code can drift away from the files you are editing.


## Command design under the chained CLI

Plugins participate in the same chained root CLI as built-in commands.
When adding a plugin command, prefer a command shape that works both:

- as a standalone invocation;
- as a step in an existing command chain after `open` or `new`.

Because the root `msw` command uses Click command chaining, nested command
namespaces are a poor fit for plugin features that need to participate in the
existing chain.
Prefer flat commands or option-driven behavior when that keeps
the plugin aligned with the main CLI model.


## Precedence policy

The active rule is:

- installed plugins are loaded by default
- local source files under `plugins/*` are not implicitly preferred over an
  installed plugin
- `--plugin-dir PATH` is an explicit override and takes priority over installed
  plugins with the same command name

This keeps runtime behavior aligned with the active Python environment and
avoids silently mixing installed metadata with unrelated source files, while
still providing a deliberate local development override.


## Example

`plugins/save-as` is a first-party plugin package that provides the `save-as`
command and the `WordClient.save_as(...)` library method. When the package is
installed in the active environment, `msword-cli` discovers it through
`msw.plugin` and registers the command automatically.

`plugins/compare-merge` follows the same model for the `compare` and `merge`
commands plus the dynamically injected `WordClient.compare(...)` and
`WordClient.merge(...)` methods.

For local development without reinstalling, you can point the CLI at the
repository plugin root directly:

```bash
msw --plugin-dir ./plugins save-as --help
```
