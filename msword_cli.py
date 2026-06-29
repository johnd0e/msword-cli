# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pywin32>=306",
#     "click>=8.0",
# ]
# ///

import sys
from functools import wraps
from importlib.metadata import entry_points
from pathlib import Path
from typing import Any, List, Optional, Tuple

import click
from pywintypes import com_error
from win32com import client as com
from win32com.client import constants as C

VERSION = '0.3.0'
__version__ = VERSION


class WordAPIError(Exception):
    """Base exception for Word API errors."""


def _resolve_constant(value: Any) -> Any:
    if isinstance(value, str):
        try:
            return getattr(C, value)
        except AttributeError as error:
            raise WordAPIError(f'Unknown Word constant: {value}') from error
    return value


def _com_error_message(error: Exception, fallback: str = 'COM error') -> str:
    excepinfo = getattr(error, 'excepinfo', None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return excepinfo[2]
    return str(error) or fallback


class Document:
    """Represents a Microsoft Word document."""

    def __init__(self, com_doc: Any):
        self._doc = com_doc

    @property
    def name(self) -> str:
        return self._doc.Name

    @property
    def saved(self) -> bool:
        return self._doc.Saved

    def activate(self) -> None:
        try:
            self._doc.Activate()
        except com_error as e:
            raise WordAPIError(f"Activation failed: {_com_error_message(e)}") from e

    def save(self, path: Optional[str] = None, force: bool = False) -> None:
        try:
            if path:
                self._doc.SaveAs(str(Path(path).resolve()))
            else:
                self._doc.Save(NoPrompt=force)
        except com_error as e:
            raise WordAPIError(f"Save failed: {_com_error_message(e)}") from e

    def close(self, force: bool = False) -> None:
        try:
            save_changes = _resolve_constant('wdDoNotSaveChanges') if force else _resolve_constant('wdPromptToSaveChanges')
            self._doc.Close(save_changes)
        except com_error as e:
            raise WordAPIError(f"Close failed: {_com_error_message(e)}") from e

    def export_fixed_format(
        self,
        path: str,
        format_val: Any = 'wdExportFormatPDF',
        show: bool = False,
        optimize: Any = 'wdExportOptimizeForPrint',
        pages: Optional[Tuple[int, int]] = None,
        rng: Any = None,
        markup: bool = False,
        properties: bool = False,
        irm: bool = False,
        bookmarks: Any = 'wdExportCreateNoBookmarks',
        struct: bool = True,
        bitmap: bool = True,
        useiso19005_1: bool = False,
    ) -> str:
        try:
            path_obj = Path(path)
            if path_obj.is_dir():
                path_obj = path_obj / Path(self.name).stem
            if path_obj.suffix.lower() not in ['.pdf', '.xps']:
                ext = '.pdf' if _resolve_constant(format_val) == _resolve_constant('wdExportFormatPDF') else '.xps'
                path_obj = path_obj.with_suffix(ext)

            final_path = str(path_obj.resolve())
            options = {
                'OutputFileName': final_path,
                'ExportFormat': _resolve_constant(format_val),
                'OpenAfterExport': show,
                'OptimizeFor': _resolve_constant(optimize),
                'Range': _resolve_constant(rng) if rng else _resolve_constant('wdExportFromTo') if pages else _resolve_constant('wdExportAllDocument'),
                'Item': _resolve_constant('wdExportDocumentWithMarkup') if markup else _resolve_constant('wdExportDocumentContent'),
                'IncludeDocProps': properties,
                'KeepIRM': not irm,
                'CreateBookmarks': _resolve_constant(bookmarks),
                'DocStructureTags': not struct,
                'BitmapMissingFonts': not bitmap,
                'UseISO19005_1': useiso19005_1,
            }
            if pages:
                options['From'] = pages[0]
                options['To'] = pages[1]

            self._doc.ExportAsFixedFormat(**options)
            return final_path
        except com_error as e:
            raise WordAPIError(f"Export failed: {_com_error_message(e)}") from e

    def print_out(
        self,
        copies: int = 1,
        pages: Optional[str] = None,
        pagetype: Any = 'wdPrintAllPages',
        rng: Any = 'wdPrintAllDocument',
        item: Any = 'wdPrintDocumentContent',
        no_collate: bool = False,
        to_file: Optional[str] = None,
        append: bool = False,
        columns: int = 1,
        rows: int = 1,
    ) -> None:
        try:
            options = {
                'Background': True,
                'Copies': copies,
                'Collate': not no_collate,
                'Item': _resolve_constant(item),
                'PrintToFile': False,
                'PageType': _resolve_constant(pagetype),
                'Range': _resolve_constant(rng),
                'PrintZoomColumn': columns,
                'PrintZoomRow': rows,
            }
            if pages:
                options['Pages'] = pages
                options['Range'] = _resolve_constant('wdPrintRangeOfPages')
            if to_file:
                options['PrintToFile'] = True
                options['OutputFileName'] = str(Path(to_file).resolve())
                if append:
                    options['Append'] = True

            self._doc.PrintOut(**options)
        except com_error as e:
            raise WordAPIError(f"Print failed: {_com_error_message(e)}") from e


class WordClient:
    """High-level client for Microsoft Word automation."""

    def __init__(self, visible: bool = False, quit_on_exit: bool = False):
        self._quit_on_exit = quit_on_exit
        self._word = None
        try:
            self._word = com.gencache.EnsureDispatch('Word.Application')
            self._word.Visible = visible
        except com_error as e:
            raise WordAPIError(f"Failed to initialize Word: {_com_error_message(e)}") from e
        except Exception:
            raise WordAPIError("Unable to load 'Word.Application'.")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._quit_on_exit:
            self.quit()

    @property
    def closed(self) -> bool:
        return self._word is None

    @property
    def visible(self) -> bool:
        return self._word.Visible

    @visible.setter
    def visible(self, value: bool) -> None:
        self._word.Visible = value

    @property
    def template_dir(self) -> Path:
        return Path(self._word.Options.DefaultFilePath(_resolve_constant('wdUserTemplatesPath')))

    @property
    def active_document(self) -> Document:
        try:
            return Document(self._word.ActiveDocument)
        except Exception:
            raise WordAPIError("No active document available.")

    @property
    def documents(self) -> List[Document]:
        return [Document(self._word.Documents.Item(i)) for i in range(1, self._word.Documents.Count + 1)]

    @property
    def document_count(self) -> int:
        return self._word.Documents.Count

    def open(self, path: str, visible: bool = True) -> Document:
        try:
            doc = self._word.Documents.Open(FileName=str(Path(path).resolve()), Visible=visible)
            if visible and not self._word.Visible:
                self._word.Visible = True
            return Document(doc)
        except com_error as e:
            raise WordAPIError(f"Failed to open document: {_com_error_message(e)}") from e

    def new(self, template: Optional[str] = None, visible: bool = True) -> Document:
        try:
            if template:
                doc = self._word.Documents.Add(Template=str(Path(template).resolve()), Visible=visible)
            else:
                doc = self._word.Documents.Add(Visible=visible)
            if visible and not self._word.Visible:
                self._word.Visible = True
            return Document(doc)
        except com_error as e:
            raise WordAPIError(f"Failed to create new document: {_com_error_message(e)}") from e

    def compare(
        self,
        original: str,
        revised: str,
        destination: Any = 'wdCompareDestinationNew',
        granularity: Any = 'wdGranularityWordLevel',
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
        author: Optional[str] = None,
        ignore_warnings: bool = False,
    ) -> Document:
        """Compare two documents; returns a Document with tracked-change diff."""
        orig_doc = None
        rev_doc = None
        result_doc = None
        try:
            orig_doc = self._word.Documents.Open(FileName=str(Path(original).resolve()), Visible=False)
            rev_doc = self._word.Documents.Open(FileName=str(Path(revised).resolve()), Visible=False)
            resolved_destination = _resolve_constant(destination)
            result = self._word.CompareDocuments(
                OriginalDocument=orig_doc,
                RevisedDocument=rev_doc,
                Destination=resolved_destination,
                Granularity=_resolve_constant(granularity),
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
                RevisedAuthor=author or self._word.UserName,
                IgnoreAllComparisonWarnings=ignore_warnings,
            )
            if resolved_destination == _resolve_constant('wdCompareDestinationOriginal'):
                result_doc = orig_doc
            elif resolved_destination == _resolve_constant('wdCompareDestinationRevised'):
                result_doc = rev_doc
            return Document(result)
        except com_error as e:
            raise WordAPIError(f"Compare failed: {_com_error_message(e)}") from e
        finally:
            for temp_doc in (orig_doc, rev_doc):
                if temp_doc is not None and temp_doc is not result_doc:
                    try:
                        temp_doc.Close(_resolve_constant('wdDoNotSaveChanges'))
                    except Exception:
                        pass

    def merge(
        self,
        original: str,
        revised: str,
        destination: Any = 'wdMergeDestinationNewDocument',
        granularity: Any = 'wdGranularityWordLevel',
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
        author: Optional[str] = None,
        ignore_warnings: bool = False,
    ) -> Document:
        """Merge two documents; returns a Document with combined tracked changes."""
        orig_doc = None
        rev_doc = None
        result_doc = None
        try:
            orig_doc = self._word.Documents.Open(FileName=str(Path(original).resolve()), Visible=False)
            rev_doc = self._word.Documents.Open(FileName=str(Path(revised).resolve()), Visible=False)
            resolved_destination = _resolve_constant(destination)
            result = self._word.MergeDocuments(
                OriginalDocument=orig_doc,
                RevisedDocument=rev_doc,
                Destination=resolved_destination,
                Granularity=_resolve_constant(granularity),
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
                RevisedAuthor=author or self._word.UserName,
                IgnoreAllComparisonWarnings=ignore_warnings,
            )
            if resolved_destination == _resolve_constant('wdMergeDestinationOriginalDocument'):
                result_doc = orig_doc
            elif resolved_destination == _resolve_constant('wdMergeDestinationRevisedDocument'):
                result_doc = rev_doc
            return Document(result)
        except com_error as e:
            raise WordAPIError(f"Merge failed: {_com_error_message(e)}") from e
        finally:
            for temp_doc in (orig_doc, rev_doc):
                if temp_doc is not None and temp_doc is not result_doc:
                    try:
                        temp_doc.Close(_resolve_constant('wdDoNotSaveChanges'))
                    except Exception:
                        pass

    def quit(self) -> None:
        word = self._word
        self._word = None
        if word is None:
            return
        try:
            word.Quit()
        except Exception:
            pass


_CLI_CLIENT = None


def get_client() -> WordClient:
    global _CLI_CLIENT
    if _CLI_CLIENT is None or _CLI_CLIENT.closed:
        _CLI_CLIENT = WordClient(visible=False)
    return _CLI_CLIENT


def handle_api_error(func):
    """Convert API errors to Click exceptions."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except WordAPIError as e:
            raise click.ClickException(str(e))
    return wrapper


class CliTemplate(click.Path):
    def convert(self, value: str, param: Optional[click.Parameter], ctx: Optional[click.Context]) -> str:
        p = Path(value)
        if not p.is_absolute():
            if p.resolve().exists():
                value = str(p.resolve())
            else:
                value = str(get_client().template_dir / p)
        return super().convert(value, param, ctx)


def validate_range(ctx: click.Context, param: click.Parameter, value: Optional[str]) -> Optional[Tuple[int, int]]:
    try:
        if value is not None:
            frm, to = map(int, value.split('-', 2))
            if frm < 1 or frm > to:
                raise ValueError
            return (frm, to)
    except ValueError:
        raise click.BadParameter('Range must be in the format "x-y" where "x" and "y" are positive integers.')
    return None


def print_version(ctx: click.Context, param: click.Parameter, value: bool) -> None:
    if not value or ctx.resilient_parsing:
        return
    click.echo(f'Version {VERSION}')
    ctx.exit()


# Shared options for compare/merge commands
_COMPARE_OPTIONS = [
    click.option('--char-level', 'granularity', flag_value='wdGranularityCharLevel',
                 help='Compare at character level (default: word level).'),
    click.option('--no-formatting', 'formatting', is_flag=True, default=False,
                 help='Ignore formatting differences.'),
    click.option('--no-case-changes', 'case_changes', is_flag=True, default=False,
                 help='Ignore case change differences.'),
    click.option('--no-whitespace', 'whitespace', is_flag=True, default=False,
                 help='Ignore whitespace differences.'),
    click.option('--no-tables', 'tables', is_flag=True, default=False,
                 help='Ignore table differences.'),
    click.option('--no-headers', 'headers', is_flag=True, default=False,
                 help='Ignore header/footer differences.'),
    click.option('--no-footnotes', 'footnotes', is_flag=True, default=False,
                 help='Ignore footnote differences.'),
    click.option('--no-textboxes', 'textboxes', is_flag=True, default=False,
                 help='Ignore text box differences.'),
    click.option('--no-fields', 'fields', is_flag=True, default=False,
                 help='Ignore field differences.'),
    click.option('--no-comments', 'comments', is_flag=True, default=False,
                 help='Ignore comment differences.'),
    click.option('--no-moves', 'moves', is_flag=True, default=False,
                 help='Ignore move differences.'),
    click.option('--author', type=str, default=None,
                 help='Author name for tracked changes (defaults to Word username).'),
    click.option('--ignore-warnings', is_flag=True, default=False,
                 help='Suppress all comparison warning dialogs.'),
]


def add_compare_options(func):
    """Attach all shared compare/merge options to a command."""
    for option in reversed(_COMPARE_OPTIONS):
        func = option(func)
    return func


@click.group(chain=True)
@click.option('--version', is_flag=True, callback=print_version, expose_value=False, is_eager=True)
def cli() -> None:
    """Command line interface for Microsoft Word."""
    pass


@cli.command('open', help='Open an existing document and make it active.')
@click.argument('path', type=click.Path(exists=True, resolve_path=True))
@click.option('--show/--hide', default=True, help='Display or hide the document.')
@handle_api_error
def open_cmd(path: str, show: bool) -> None:
    click.echo(f'Opening document at "{path}"')
    get_client().open(path, visible=show)


@cli.command('new', help='Create a new document, optionally from a template.')
@click.option('-t', '--template', type=CliTemplate(exists=True, dir_okay=False, resolve_path=True),
              help='Path to a Word template file.')
@click.option('--show/--hide', default=True, help='Display or hide the document.')
@handle_api_error
def new_cmd(template: Optional[str], show: bool) -> None:
    if template:
        click.echo(f'Opening new document using template: "{template}"')
    else:
        click.echo('Opening new blank document.')
    get_client().new(template=template, visible=show)


PRINT_OUT_ITEMS = {
    'document_content': 'wdPrintDocumentContent',
    'doc_with_markup': 'wdPrintDocumentWithMarkup',
    'comments': 'wdPrintComments',
    'properties': 'wdPrintProperties',
    'markup': 'wdPrintMarkup',
    'styles': 'wdPrintStyles',
    'auto_text_entries': 'wdPrintAutoTextEntries',
    'key_assignments': 'wdPrintKeyAssignments',
    'envelope': 'wdPrintEnvelope',
}


@cli.command('print', help='Print the active document.')
@click.option('-c', '--copies', type=click.IntRange(min=1), default=1, help='Number of copies to print.')
@click.option('-p', '--pages', type=str, help='Specific pages or page ranges, for example 2-4, 6.')
@click.option('--even', 'pagetype', flag_value='wdPrintEvenPagesOnly', help='Print even-numbered pages only.')
@click.option('--odd', 'pagetype', flag_value='wdPrintOddPagesOnly', help='Print odd-numbered pages only.')
@click.option('--current-page', 'rng', flag_value='wdPrintCurrentPage', help='Print only the current page.')
@click.option('--selection', 'rng', flag_value='wdPrintSelection', help='Print only the current selection.')
@click.option('--no-collate', is_flag=True, help='Disable collation when printing multiple copies.')
@click.option('--to-file', type=click.Path(dir_okay=False, resolve_path=True),
              help='Print to a file instead of a printer.')
@click.option('--append', is_flag=True, help='Append to the output file when supported.')
@click.option('--columns', type=click.Choice(['1', '2', '3', '4']), default='1',
              help='Number of pages to print across each row.')
@click.option('--rows', type=click.Choice(['1', '2', '4']), default='1',
              help='Number of pages to print down each column.')
@click.option('--item', type=click.Choice(list(PRINT_OUT_ITEMS.keys())), default='document_content',
              help='Select which document content Word should print.')
@handle_api_error
def print_cmd(copies: int, pages: Optional[str], pagetype: Any, rng: Any, item: str, no_collate: bool, to_file: Optional[str], append: bool, columns: str, rows: str) -> None:
    click.echo(f"Printing {copies} copies of pages: {pages or 'all'}")
    doc = get_client().active_document
    doc.print_out(
        copies=copies,
        pages=pages,
        pagetype=_resolve_constant(pagetype or 'wdPrintAllPages'),
        rng=_resolve_constant(rng or 'wdPrintAllDocument'),
        item=_resolve_constant(PRINT_OUT_ITEMS[item]),
        no_collate=no_collate,
        to_file=to_file,
        append=append,
        columns=int(columns),
        rows=int(rows),
    )


@cli.command('export', help='Export the active document to PDF or XPS.')
@click.option('--pdf', 'format_val', flag_value='wdExportFormatPDF', default=True, help='Export as PDF.')
@click.option('--xps', 'format_val', flag_value='wdExportFormatXPS', help='Export as XPS.')
@click.option('--show', is_flag=True, help='Open the exported file after creating it.')
@click.option('--for-print', 'optimize', flag_value='wdExportOptimizeForPrint', default=True,
              help='Optimize the export for printing.')
@click.option('--for-screen', 'optimize', flag_value='wdExportOptimizeForOnScreen',
              help='Optimize the export for on-screen viewing.')
@click.option('--pages', type=str, callback=validate_range, help='Export only a page range in the form x-y.')
@click.option('--current-page', 'rng', flag_value='wdExportCurrentPage', help='Export only the current page.')
@click.option('--selection', 'rng', flag_value='wdExportSelection', help='Export only the current selection.')
@click.option('--with-markup', 'markup', is_flag=True, help='Include tracked changes and markup.')
@click.option('--with-props', 'properties', is_flag=True, help='Include document properties.')
@click.option('--without-irm', 'irm', is_flag=True, help='Do not preserve IRM permissions in the export.')
@click.option('--with-heading-bookmarks', 'bookmarks', flag_value='wdExportCreateHeadingBookmarks',
              help='Create bookmarks from document headings.')
@click.option('--with-word-bookmarks', 'bookmarks', flag_value='wdExportCreateWordBookmarks',
              help='Create bookmarks from Word bookmarks.')
@click.option('--without-structure-tags', 'struct', is_flag=True,
              help='Disable document structure tags in the export.')
@click.option('--without-bitmaped-fonts', 'bitmap', is_flag=True,
              help='Disable bitmap fallback for fonts that cannot be embedded.')
@click.option('--useiso19005-1', 'useiso19005_1', is_flag=True, help='Create a PDF/A-compatible export.')
@click.argument('path', type=click.Path(dir_okay=True, resolve_path=True))
@handle_api_error
def export_cmd(path: str, format_val: Any, show: bool, optimize: Any, pages: Optional[Tuple[int, int]], rng: Any, markup: bool, properties: bool, irm: bool, bookmarks: Any, struct: bool, bitmap: bool, useiso19005_1: bool) -> None:
    doc = get_client().active_document
    final_path = doc.export_fixed_format(
        path=path,
        format_val=_resolve_constant(format_val),
        show=show,
        optimize=_resolve_constant(optimize),
        pages=pages,
        rng=_resolve_constant(rng) if rng else None,
        markup=markup,
        properties=properties,
        irm=irm,
        bookmarks=_resolve_constant(bookmarks or 'wdExportCreateNoBookmarks'),
        struct=struct,
        bitmap=bitmap,
        useiso19005_1=useiso19005_1,
    )
    click.echo(f'Exported to "{final_path}"')


@cli.command('save', help='Save the active document or all open documents.')
@click.option('-a', '--all', 'save_all', is_flag=True, help='Save all open documents.')
@click.option('-f', '--force', is_flag=True, help='Save without prompting when Word supports it.')
@click.option('-p', '--path', type=click.Path(resolve_path=True), help='Save the active document to a new path.')
@handle_api_error
def save_cmd(save_all: bool, force: bool, path: Optional[str]) -> None:
    client = get_client()
    if path:
        click.echo(f'Saving document to: "{path}"')
        client.active_document.save(path=path)
    else:
        click.echo('Saving changes to existing document(s).')
        docs_to_save = client.documents if save_all else [client.active_document]
        for d in docs_to_save:
            d.save(force=force)


@cli.command('close', help='Close the active document or all open documents.')
@click.option('-a', '--all', 'close_all', is_flag=True, help='Close all open documents.')
@click.option('-f', '--force', is_flag=True, help='Discard unsaved changes without prompting.')
@handle_api_error
def close_cmd(close_all: bool, force: bool) -> None:
    client = get_client()
    docs_to_close = client.documents if close_all else [client.active_document]
    action = 'Force closing' if force else 'Closing'
    click.echo(f'{action} document(s)...')
    for d in docs_to_close:
        d.close(force=force)
    if client.document_count == 0:
        client.quit()


@cli.command('activate', help='Activate an open document by its index from the docs list.')
@click.argument('index', type=int)
@handle_api_error
def activate_cmd(index: int) -> None:
    click.echo(f'Activate document at index "{index}"')
    client = get_client()
    if 1 <= index <= client.document_count:
        client.documents[index - 1].activate()
    else:
        raise click.ClickException(f'Index {index} out of range.')


@cli.command('docs', help='List open documents and show which one is active.')
@handle_api_error
def docs_cmd() -> None:
    client = get_client()
    if client.document_count > 0:
        click.echo('\nOpen Documents:\n')
        pad_len = len(str(client.document_count))
        active_name = client.active_document.name
        for i, doc in enumerate(client.documents, start=1):
            active = '*' if doc.name == active_name else ' '
            saved = '*' if not doc.saved else ''
            click.echo(f' {active} [{i: ={pad_len}}] {doc.name}{saved}')
    else:
        click.echo('\nNo open documents found.')


@cli.command('compare')
@click.argument('original', type=click.Path(exists=True, resolve_path=True))
@click.argument('revised', type=click.Path(exists=True, resolve_path=True))
@click.option('--to-original', 'destination', flag_value='wdCompareDestinationOriginal',
              help='Put diff into the original document.')
@click.option('--to-revised', 'destination', flag_value='wdCompareDestinationRevised',
              help='Put diff into the revised document.')
@add_compare_options
@handle_api_error
def compare_cmd(
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
    author: Optional[str],
    ignore_warnings: bool,
) -> None:
    """Compare ORIGINAL and REVISED documents, showing differences as tracked changes."""
    click.echo(f'Comparing "{original}" with "{revised}"')
    result = get_client().compare(
        original=original,
        revised=revised,
        destination=_resolve_constant(destination or 'wdCompareDestinationNew'),
        granularity=_resolve_constant(granularity or 'wdGranularityWordLevel'),
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


@cli.command('merge')
@click.argument('original', type=click.Path(exists=True, resolve_path=True))
@click.argument('revised', type=click.Path(exists=True, resolve_path=True))
@click.option('--to-original', 'destination', flag_value='wdMergeDestinationOriginalDocument',
              help='Merge result into the original document.')
@click.option('--to-revised', 'destination', flag_value='wdMergeDestinationRevisedDocument',
              help='Merge result into the revised document.')
@add_compare_options
@handle_api_error
def merge_cmd(
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
    author: Optional[str],
    ignore_warnings: bool,
) -> None:
    """Merge ORIGINAL and REVISED documents, combining their tracked changes."""
    click.echo(f'Merging "{original}" with "{revised}"')
    result = get_client().merge(
        original=original,
        revised=revised,
        destination=_resolve_constant(destination or 'wdMergeDestinationNewDocument'),
        granularity=_resolve_constant(granularity or 'wdGranularityWordLevel'),
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


try:
    eps = entry_points()
    plugin_eps = eps.get('msw.plugin', []) if sys.version_info < (3, 10) else eps.select(group='msw.plugin')
except Exception as e:
    click.echo(f'Warning: Failed to discover plugins: {e}', err=True)
else:
    for plugin in plugin_eps:
        try:
            cli.add_command(plugin.load())
        except Exception as e:
            plugin_name = getattr(plugin, 'name', '<unknown>')
            click.echo(f'Warning: Failed to load plugin {plugin_name}: {e}', err=True)


if __name__ == '__main__':
    cli()

