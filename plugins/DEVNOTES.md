# Plugin Developer Notes

This file captures plugin-specific architectural constraints that are easy to
forget when extending the plugin system or adding first-party plugins under
`plugins/*`.


## Unified plugin manifest

Plugins share one discovery path through the `msw.plugin` entry-point group.
The entry point resolves to a manifest dictionary rather than directly to a
Click command.

Current manifest shape:

```python
plugin_manifest = {
    "name": "save-as",
    "command": save_as_cmd,
    "client_methods": {
        "save_as": save_as,
    },
}
```

Keep these rules:

- `name` is required and must be unique across loaded plugins.
- `command` is optional and is consumed by the CLI loader.
- `client_methods` is optional and is consumed by `WordClient.load_plugins()`.
- A plugin may be CLI-only, library-only, or both.


## Library method injection model

Library plugin methods are injected onto one `WordClient` instance, not onto
the `WordClient` class globally.

Keep these rules:

- `WordClient.load_plugins(...)` is explicit and instance-scoped.
- Client method callables take the active `WordClient` instance as their first
  argument.
- Method name conflicts are hard errors; plugins must not override core methods
  or previously injected plugin methods.
- `include` filtering is by plugin manifest `name`, not by Python module name.


## CLI and library logic should share behavior

When a plugin exposes both a CLI command and a library method, keep the Click
wrapper thin and keep the real operation in the library callable where
practical.

The preferred shape is:

- shared validation and option resolution in normal Python helpers;
- one library callable used by direct Python users;
- one thin Click command that adapts CLI arguments and output.

This keeps CLI and library behavior aligned and reduces drift between the two
surfaces.


## Validate before touching Word

Plugin commands should validate arguments before calling `get_client()` or
touching COM-backed objects.

Why this matters:

- usage errors should fail fast without starting or querying Word;
- `--help`/validation flows stay lazy and predictable;
- tests can assert bad input is rejected before any Word interaction.

The `save-as` plugin now follows this rule by resolving and validating format
options before it requests the active client.


## Core vs bundled first-party plugins

Bundled first-party plugins are still part of the product, but they do not
belong in core automatically just because the underlying Word feature is
"native."

Prefer a bundled first-party plugin when one or more of these are true:

- the workflow is specialized rather than part of the default single-document
  lifecycle;
- the plugin needs a noticeable semantic shell around a direct COM call
  (validation, policy, option mapping, special result handling);
- the operation has an unusual input/output model such as multiple input files
  or a result document with custom cleanup rules.

Keep these rules:

- core helpers are part of the product contract, so the bar for adding them is
  high;
- do not promote plugin-local helpers into core preemptively;
- repeated use or a strong, stable architectural signal is required before a
  plugin helper becomes a core helper;
- trivially short local code is not a reason to create a shared abstraction by
  itself.
