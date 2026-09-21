from __future__ import annotations

import warnings
from typing import Any

import click
from pywintypes import com_error

from msword_cli import (
    Document,
    WordAPIError,
    _com_error_message,
    _normalize_output_path,
    _resolve_constant,
    get_client,
    handle_api_error,
)

_COMPARE_OPTIONS = [
    click.option(
        "--char-level",
        "granularity",
        flag_value="wdGranularityCharLevel",
        help="Compare at character level (default: word level).",
    ),
    click.option(
        "--no-formatting",
        "formatting",
        is_flag=True,
        default=False,
        help="Ignore formatting differences.",
    ),
    click.option(
        "--no-case-changes",
        "case_changes",
        is_flag=True,
        default=False,
        help="Ignore case change differences.",
    ),
    click.option(
        "--no-whitespace",
        "whitespace",
        is_flag=True,
        default=False,
        help="Ignore whitespace differences.",
    ),
    click.option(
        "--no-tables",
        "tables",
        is_flag=True,
        default=False,
        help="Ignore table differences.",
    ),
    click.option(
        "--no-headers",
        "headers",
        is_flag=True,
        default=False,
        help="Ignore header/footer differences.",
    ),
    click.option(
        "--no-footnotes",
        "footnotes",
        is_flag=True,
        default=False,
        help="Ignore footnote differences.",
    ),
    click.option(
        "--no-textboxes",
        "textboxes",
        is_flag=True,
        default=False,
        help="Ignore text box differences.",
    ),
    click.option(
        "--no-fields",
        "fields",
        is_flag=True,
        default=False,
        help="Ignore field differences.",
    ),
    click.option(
        "--no-comments",
        "comments",
        is_flag=True,
        default=False,
        help="Ignore comment differences.",
    ),
    click.option(
        "--no-moves",
        "moves",
        is_flag=True,
        default=False,
        help="Ignore move differences.",
    ),
    click.option(
        "--author",
        type=str,
        default=None,
        help="Author name for tracked changes (defaults to Word username).",
    ),
    click.option(
        "--ignore-warnings",
        is_flag=True,
        default=False,
        help="Suppress all comparison warning dialogs.",
    ),
]


def add_compare_options(func):
    for option in reversed(_COMPARE_OPTIONS):
        func = option(func)
    return func


def _close_temporary_documents(
    *, original_doc: Any, revised_doc: Any, result_doc: Any
) -> list:
    errors = []
    for temp_doc in (original_doc, revised_doc):
        if temp_doc is not None and temp_doc is not result_doc:
            try:
                temp_doc.Close(_resolve_constant("wdDoNotSaveChanges"))
            except Exception as error:  # noqa: BLE001
                errors.append(error)
    return errors


def _run_document_comparison(  # noqa: PLR0913
    *,
    client: Any,
    operation_name: str,
    com_method_name: str,
    original: str,
    revised: str,
    destination: Any,
    new_destination: str,
    original_destination: str,
    revised_destination: str,
    granularity: Any = "wdGranularityWordLevel",
    formatting: bool = True,
    case_changes: bool = True,
    whitespace: bool = True,
    tables: bool = True,
    headers: bool = True,
    footnotes: bool = True,
    textboxes: bool = True,
    fields: bool = True,
    comments: bool = True,
    moves: bool = True,
    author: str | None = None,
    ignore_warnings: bool = False,
) -> Document:
    original_doc = None
    revised_doc = None
    result_doc = None
    try:
        word = client.native
        original_doc = word.Documents.Open(
            FileName=_normalize_output_path(original), Visible=False
        )
        revised_doc = word.Documents.Open(
            FileName=_normalize_output_path(revised), Visible=False
        )
        resolved_destination = _resolve_constant(destination or new_destination)
        result = getattr(word, com_method_name)(
            OriginalDocument=original_doc,
            RevisedDocument=revised_doc,
            Destination=resolved_destination,
            Granularity=_resolve_constant(granularity or "wdGranularityWordLevel"),
            CompareFormatting=formatting,
            CompareCaseChanges=case_changes,
            CompareWhitespace=whitespace,
            CompareTables=tables,
            CompareHeaders=headers,
            CompareFootnotes=footnotes,
            CompareTextboxes=textboxes,
            CompareFields=fields,
            CompareComments=comments,
            CompareMoves=moves,
            RevisedAuthor=author or word.UserName,
            IgnoreAllComparisonWarnings=ignore_warnings,
        )
        if resolved_destination == _resolve_constant(original_destination):
            result_doc = original_doc
        elif resolved_destination == _resolve_constant(revised_destination):
            result_doc = revised_doc
        return Document(result)
    except com_error as error:
        raise WordAPIError(
            f"{operation_name} failed: {_com_error_message(error)}"
        ) from error
    finally:
        cleanup_errors = _close_temporary_documents(
            original_doc=original_doc,
            revised_doc=revised_doc,
            result_doc=result_doc,
        )
        for error in cleanup_errors:
            warnings.warn(
                f"Failed to close temporary comparison document: {error}",
                RuntimeWarning,
                stacklevel=2,
            )


def compare(  # noqa: PLR0913, PLR0917
    client: Any,
    original: str,
    revised: str,
    destination: Any = "wdCompareDestinationNew",
    granularity: Any = "wdGranularityWordLevel",
    formatting: bool = True,
    case_changes: bool = True,
    whitespace: bool = True,
    tables: bool = True,
    headers: bool = True,
    footnotes: bool = True,
    textboxes: bool = True,
    fields: bool = True,
    comments: bool = True,
    moves: bool = True,
    author: str | None = None,
    ignore_warnings: bool = False,
) -> Document:
    return _run_document_comparison(
        client=client,
        operation_name="Compare",
        com_method_name="CompareDocuments",
        original=original,
        revised=revised,
        destination=destination,
        new_destination="wdCompareDestinationNew",
        original_destination="wdCompareDestinationOriginal",
        revised_destination="wdCompareDestinationRevised",
        granularity=granularity,
        formatting=formatting,
        case_changes=case_changes,
        whitespace=whitespace,
        tables=tables,
        headers=headers,
        footnotes=footnotes,
        textboxes=textboxes,
        fields=fields,
        comments=comments,
        moves=moves,
        author=author,
        ignore_warnings=ignore_warnings,
    )


def merge(  # noqa: PLR0913, PLR0917
    client: Any,
    original: str,
    revised: str,
    destination: Any = "wdMergeDestinationNewDocument",
    granularity: Any = "wdGranularityWordLevel",
    formatting: bool = True,
    case_changes: bool = True,
    whitespace: bool = True,
    tables: bool = True,
    headers: bool = True,
    footnotes: bool = True,
    textboxes: bool = True,
    fields: bool = True,
    comments: bool = True,
    moves: bool = True,
    author: str | None = None,
    ignore_warnings: bool = False,
) -> Document:
    return _run_document_comparison(
        client=client,
        operation_name="Merge",
        com_method_name="MergeDocuments",
        original=original,
        revised=revised,
        destination=destination,
        new_destination="wdMergeDestinationNewDocument",
        original_destination="wdMergeDestinationOriginalDocument",
        revised_destination="wdMergeDestinationRevisedDocument",
        granularity=granularity,
        formatting=formatting,
        case_changes=case_changes,
        whitespace=whitespace,
        tables=tables,
        headers=headers,
        footnotes=footnotes,
        textboxes=textboxes,
        fields=fields,
        comments=comments,
        moves=moves,
        author=author,
        ignore_warnings=ignore_warnings,
    )


@click.command(
    "compare",
    short_help="Compare two documents.",
    help=(
        "Compare ORIGINAL and REVISED documents, showing differences as "
        "tracked changes."
    ),
)
@click.argument("original", type=click.Path(exists=True, resolve_path=True))
@click.argument("revised", type=click.Path(exists=True, resolve_path=True))
@click.option(
    "--to-original",
    "destination",
    flag_value="wdCompareDestinationOriginal",
    help="Put diff into the original document.",
)
@click.option(
    "--to-revised",
    "destination",
    flag_value="wdCompareDestinationRevised",
    help="Put diff into the revised document.",
)
@add_compare_options
@handle_api_error
def compare_cmd(  # noqa: PLR0913, PLR0917
    original: str,
    revised: str,
    destination: Any,
    granularity: Any,
    formatting: bool,
    case_changes: bool,
    whitespace: bool,
    tables: bool,
    headers: bool,
    footnotes: bool,
    textboxes: bool,
    fields: bool,
    comments: bool,
    moves: bool,
    author: str | None,
    ignore_warnings: bool,
) -> None:
    click.echo(f'Comparing "{original}" with "{revised}"')
    result = compare(
        get_client(),
        original=original,
        revised=revised,
        destination=_resolve_constant(destination or "wdCompareDestinationNew"),
        granularity=_resolve_constant(granularity or "wdGranularityWordLevel"),
        formatting=not formatting,
        case_changes=not case_changes,
        whitespace=not whitespace,
        tables=not tables,
        headers=not headers,
        footnotes=not footnotes,
        textboxes=not textboxes,
        fields=not fields,
        comments=not comments,
        moves=not moves,
        author=author,
        ignore_warnings=ignore_warnings,
    )
    click.echo(f'Result document: "{result.name}"')


@click.command(
    "merge",
    short_help="Merge two documents.",
    help="Merge ORIGINAL and REVISED documents, combining their tracked changes.",
)
@click.argument("original", type=click.Path(exists=True, resolve_path=True))
@click.argument("revised", type=click.Path(exists=True, resolve_path=True))
@click.option(
    "--to-original",
    "destination",
    flag_value="wdMergeDestinationOriginalDocument",
    help="Merge result into the original document.",
)
@click.option(
    "--to-revised",
    "destination",
    flag_value="wdMergeDestinationRevisedDocument",
    help="Merge result into the revised document.",
)
@add_compare_options
@handle_api_error
def merge_cmd(  # noqa: PLR0913, PLR0917
    original: str,
    revised: str,
    destination: Any,
    granularity: Any,
    formatting: bool,
    case_changes: bool,
    whitespace: bool,
    tables: bool,
    headers: bool,
    footnotes: bool,
    textboxes: bool,
    fields: bool,
    comments: bool,
    moves: bool,
    author: str | None,
    ignore_warnings: bool,
) -> None:
    click.echo(f'Merging "{original}" with "{revised}"')
    result = merge(
        get_client(),
        original=original,
        revised=revised,
        destination=_resolve_constant(destination or "wdMergeDestinationNewDocument"),
        granularity=_resolve_constant(granularity or "wdGranularityWordLevel"),
        formatting=not formatting,
        case_changes=not case_changes,
        whitespace=not whitespace,
        tables=not tables,
        headers=not headers,
        footnotes=not footnotes,
        textboxes=not textboxes,
        fields=not fields,
        comments=not comments,
        moves=not moves,
        author=author,
        ignore_warnings=ignore_warnings,
    )
    click.echo(f'Result document: "{result.name}"')


plugin_manifest = {
    "name": "compare-merge",
    "commands": {
        "compare": compare_cmd,
        "merge": merge_cmd,
    },
    "client_methods": {
        "compare": compare,
        "merge": merge,
    },
}
