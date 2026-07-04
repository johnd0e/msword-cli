# Helper Candidates

This file is intentionally non-normative. It tracks plugin-local patterns that
may eventually justify a shared helper, but are not part of the core contract
today.

Current candidates observed while building first-party plugins:

- Hidden multi-document open flows that take more than one input path.
- Cleanup logic that closes temporary input documents while preserving the
  input document that became the result.
- Reusable option-to-COM policy translation when a plugin grows a large
  semantic shell around a direct Word call.

Do not move any of these into `msword_cli.py` just because one plugin needs
them. A candidate should stay local until repeated use or a strong, stable
architectural signal proves the abstraction.
