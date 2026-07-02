# Plugins

`msword-cli` supports command plugins through the `msw.plugin` entry-point
group.

First-party plugins live under `plugins/*` in this repository. A plugin is a
separate Python package that exposes one or more Click commands.

## Plugin package structure

Each plugin is its own package with a `pyproject.toml`. The package declares an
entry point in the `msw.plugin` group:

```toml
[project.entry-points."msw.plugin"]
save-as = "msword_save_as_plugin:save_as_cmd"
```

The value uses `module:function` syntax. When the plugin is discovered,
`msword-cli` loads that object and registers it as a CLI command.

## Discovery model

The runtime has two discovery paths:

- installed Python package metadata through `msw.plugin`
- explicit local plugin roots passed through `--plugin-dir PATH`

Installed plugins are discovered first through
`importlib.metadata.entry_points()`. Every discovered plugin is loaded
independently, and a failing plugin does not block later plugins from loading.

When `--plugin-dir PATH` is provided, `msword-cli` scans that directory for
plugin subdirectories containing a `pyproject.toml` with `msw.plugin` entry
points.

Without `--plugin-dir`, a plugin is discovered automatically only when it is
installed into the active Python environment.

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
command. When the package is installed in the active environment, `msword-cli`
discovers it through `msw.plugin` and registers the command automatically.

For local development without reinstalling, you can point the CLI at the
repository plugin root directly:

```bash
msw --plugin-dir ./plugins save-as --help
```
