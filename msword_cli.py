# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pywin32",
#     "click",
# ]
# ///

import sys
from pathlib import Path
from typing import Optional, Tuple, Any, List
from functools import wraps

from win32com import client as com
from win32com.client import constants as C
from pywintypes import com_error
import click

if sys.version_info < (3, 10):
    from importlib.metadata import entry_points
else:
    from importlib.metadata import entry_points

VERSION = '0.2.0'


class WordAPIError(Exception):
    """Base exception for Word API errors."""
    pass


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
            raise WordAPIError(f"Activation failed: {e.excepinfo[2]}")

    def save(self, path: Optional[str] = None, force: bool = False) -> None:
        try:
            if path:
                self._doc.SaveAs(str(Path(path).resolve()))
            else:
                self._doc.Save(NoPrompt=force)
        except com_error as e:
            raise WordAPIError(f"Save failed: {e.excepinfo[2]}")

    def close(self, force: bool = False) -> None:
        try:
            save_changes = C.wdDoNotSaveChanges if force else C.wdPromptToSaveChanges
            self._doc.Close(save_changes)
        except com_error as e:
            raise WordAPIError(f"Close failed: {e.excepinfo[2]}")

    def export_fixed_format(
        self,
        path: str,
        format_val: Any = C.wdExportFormatPDF,
        show: bool = False,
        optimize: Any = C.wdExportOptimizeForPrint,
        pages: Optional[Tuple[int, int]] = None,
        rng: Any = None,
        markup: bool = False,
        properties: bool = False,
        irm: bool = False,
        bookmarks: Any = C.wdExportCreateNoBookmarks,
        struct: bool = True,
        bitmap: bool = True,
        useiso19005_1: bool = False,
    ) -> str:
        try:
            path_obj = Path(path)
            if path_obj.is_dir():
                path_obj = path_obj / Path(self.name).stem
            if path_obj.suffix.lower() not in ['.pdf', '.xps']:
                ext = '.pdf' if format_val == C.wdExportFormatPDF else '.xps'
                path_obj = path_obj.with_suffix(ext)

            final_path = str(path_obj.resolve())
            options = {
                'OutputFileName': final_path,
                'ExportFormat': format_val,
                'OpenAfterExport': show,
                'OptimizeFor': optimize,
                'Range': rng if rng else C.wdExportFromTo if pages else C.wdExportAllDocument,
                'Item': C.wdExportDocumentWithMarkup if markup else C.wdExportDocumentContent,
                'IncludeDocProps': properties,
                'KeepIRM': not irm,
                'CreateBookmarks': bookmarks,
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
            raise WordAPIError(f"Export failed: {e.excepinfo[2]}")

    def print_out(
        self,
        copies: int = 1,
        pages: Optional[str] = None,
        pagetype: Any = C.wdPrintAllPages,
        rng: Any = C.wdPrintAllDocument,
        item: Any = C.wdPrintDocumentContent,
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
                'Item': item,
                'PrintToFile': False,
                'PageType': pagetype,
                'Range': rng,
                'PrintZoomColumn': columns,
                'PrintZoomRow': rows,
            }
            if pages:
                options['Pages'] = pages
                options['Range'] = C.wdPrintRangeOfPages
            if to_file:
                options['PrintToFile'] = True
                options['OutputFileName'] = str(Path(to_file).resolve())
                if append:
                    options['Append'] = True

            self._doc.PrintOut(**options)
        except com_error as e:
            raise WordAPIError(f"Print failed: {e.excepinfo[2]}")


class WordClient:
    """High-level client for Microsoft Word automation."""

    def __init__(self, visible: bool = False, quit_on_exit: bool = False):
        self._quit_on_exit = quit_on_exit
        try:
            self._word = com.gencache.EnsureDispatch('Word.Application')
            self._word.Visible = visible
        except com_error as e:
            raise WordAPIError(f"Failed to initialize Word: {e.excepinfo[2]}")
        except Exception:
            raise WordAPIError("Unable to load 'Word.Application'.")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._quit_on_exit:
            self.quit()

    @property
    def visible(self) -> bool:
        return self._word.Visible

    @visible.setter
    def visible(self, value: bool) -> None:
        self._word.Visible = value

    @property
    def template_dir(self) -> Path:
        return Path(self._word.Options.DefaultFilePath(C.wdUserTemplatesPath))

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
            raise WordAPIError(f"Failed to open document: {e.excepinfo[2]}")

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
            raise WordAPIError(f"Failed to create new document: {e.excepinfo[2]}")

    def compare(
        self,
        original: str,
        revised: str,
        destination: Any = C.wdCompareDestinationNew,
        granularity: Any = C.wdGranularityWordLevel,
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
        try:
            orig_doc = self._word.Documents.Open(FileName=str(Path(original).resolve()), Visible=False)
            rev_doc = self._word.Documents.Open(FileName=str(Path(revised).resolve()), Visible=False)
            result = self._word.CompareDocuments(
                OriginalDocument=orig_doc,
                RevisedDocument=rev_doc,
                Destination=destination,
                Granularity=granularity,
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
            return Document(result)
        except com_error as e:
            raise WordAPIError(f"Compare failed: {e.excepinfo[2]}")

    def merge(
        self,
        original: str,
        revised: str,
        destination: Any = C.wdMergeDestinationNewDocument,
        granularity: Any = C.wdGranularityWordLevel,
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
        try:
            orig_doc = self._word.Documents.Open(FileName=str(Path(original).resolve()), Visible=False)
            rev_doc = self._word.Documents.Open(FileName=str(Path(revised).resolve()), Visible=False)
            result = self._word.MergeDocuments(
                OriginalDocument=orig_doc,
                RevisedDocument=rev_doc,
                Destination=destination,
                Granularity=granularity,
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
            return Document(result)
        except com_error as e:
            raise WordAPIError(f"Merge failed: {e.excepinfo[2]}")

    def quit(self) -> None:
        try:
            self._word.Quit()
        except Exception:
            pass


_CLI_CLIENT = None


def get_client() -> WordClient:
    global _CLI_CLIENT
    if _CLI_CLIENT is None:
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
            if frm > to:
                raise ValueError
            return (frm, to)
    except ValueError:
        raise click.BadParameter('Range must be in the format "x-y" where "x" and "y" are integers.')
    return None


def print_version(ctx: click.Context, param: click.Parameter, value: bool) -> None:
    if not value or ctx.resilient_parsing:
        return
    click.echo(f'Version {VERSION}')
    ctx.exit()


# Shared options for compare/merge commands
_COMPARE_OPTIONS = [
    click.option('--char-level', 'granularity', flag_value=C.wdGranularityCharLevel,
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


@cli.command('open')
@click.argument('path', type=click.Path(exists=True, resolve_path=True))
@click.option('--show/--hide', default=True, help='Display or hide the document.')
@handle_api_error
def open_cmd(path: str, show: bool) -> None:
    click.echo(f'Opening document at "{path}"')
    get_client().open(path, visible=show)


@cli.command('new')
@click.option('-t', '--template', type=CliTemplate(exists=True, dir_okay=False, resolve_path=True))
@click.option('--show/--hide', default=True, help='Display or hide the document.')
@handle_api_error
def new_cmd(template: Optional[str], show: bool) -> None:
    if template:
        click.echo(f'Opening new document using template: "{template}"')
    else:
        click.echo('Opening new blank document.')
    get_client().new(template=template, visible=show)


PRINT_OUT_ITEMS = {
    'document_content': C.wdPrintDocumentContent,
    'doc_with_markup': C.wdPrintDocumentWithMarkup,
    'comments': C.wdPrintComments,
    'properties': C.wdPrintProperties,
    'markup': C.wdPrintMarkup,
    'styles': C.wdPrintStyles,
    'auto_text_entries': C.wdPrintAutoTextEntries,
    'key_assignments': C.wdPrintKeyAssignments,
    'envelope': C.wdPrintEnvelope,
}


@cli.command('print')
@click.option('-c', '--copies', default=1)
@click.option('-p', '--pages', type=str)
@click.option('--even', 'pagetype', flag_value=C.wdPrintEvenPagesOnly)
@click.option('--odd', 'pagetype', flag_value=C.wdPrintOddPagesOnly)
@click.option('--current-page', 'rng', flag_value=C.wdPrintCurrentPage)
@click.option('--selection', 'rng', flag_value=C.wdPrintSelection)
@click.option('--no-collate', is_flag=True)
@click.option('--to-file', type=click.Path(dir_okay=False, resolve_path=True))
@click.option('--append', is_flag=True)
@click.option('--columns', type=click.Choice(['1', '2', '3', '4']), default='1')
@click.option('--rows', type=click.Choice(['1', '2', '4']), default='1')
@click.option('--item', type=click.Choice(list(PRINT_OUT_ITEMS.keys())), default='document_content')
@handle_api_error
def print_cmd(copies: int, pages: Optional[str], pagetype: Any, rng: Any, item: str, no_collate: bool, to_file: Optional[str], append: bool, columns: str, rows: str) -> None:
    click.echo(f"Printing {copies} copies of pages: {pages or 'all'}")
    doc = get_client().active_document
    doc.print_out(
        copies=copies,
        pages=pages,
        pagetype=pagetype or C.wdPrintAllPages,
        rng=rng or C.wdPrintAllDocument,
        item=PRINT_OUT_ITEMS[item],
        no_collate=no_collate,
        to_file=to_file,
        append=append,
        columns=int(columns),
        rows=int(rows),
    )


@cli.command('export')
@click.option('--pdf', 'format_val', flag_value=C.wdExportFormatPDF, default=True)
@click.option('--xps', 'format_val', flag_value=C.wdExportFormatXPS)
@click.option('--show', is_flag=True)
@click.option('--for-print', 'optimize', flag_value=C.wdExportOptimizeForPrint, default=True)
@click.option('--for-screen', 'optimize', flag_value=C.wdExportOptimizeForOnScreen)
@click.option('--pages', type=str, callback=validate_range)
@click.option('--current-page', 'rng', flag_value=C.wdExportCurrentPage)
@click.option('--selection', 'rng', flag_value=C.wdExportSelection)
@click.option('--with-markup', 'markup', is_flag=True)
@click.option('--with-props', 'properties', is_flag=True)
@click.option('--without-irm', 'irm', is_flag=True)
@click.option('--with-heading-bookmarks', 'bookmarks', flag_value=C.wdExportCreateHeadingBookmarks)
@click.option('--with-word-bookmarks', 'bookmarks', flag_value=C.wdExportCreateWordBookmarks)
@click.option('--without-structure-tags', 'struct', is_flag=True)
@click.option('--without-bitmaped-fonts', 'bitmap', is_flag=True)
@click.option('--useiso19005-1', 'useiso19005_1', is_flag=True)
@click.argument('path', type=click.Path(dir_okay=True, resolve_path=True))
@handle_api_error
def export_cmd(path: str, format_val: Any, show: bool, optimize: Any, pages: Optional[Tuple[int, int]], rng: Any, markup: bool, properties: bool, irm: bool, bookmarks: Any, struct: bool, bitmap: bool, useiso19005_1: bool) -> None:
    doc = get_client().active_document
    final_path = doc.export_fixed_format(
        path=path,
        format_val=format_val,
        show=show,
        optimize=optimize,
        pages=pages,
        rng=rng,
        markup=markup,
        properties=properties,
        irm=irm,
        bookmarks=bookmarks or C.wdExportCreateNoBookmarks,
        struct=struct,
        bitmap=bitmap,
        useiso19005_1=useiso19005_1,
    )
    click.echo(f'Exported to "{final_path}"')


@cli.command('save')
@click.option('-a', '--all', 'save_all', is_flag=True)
@click.option('-f', '--force', is_flag=True)
@click.option('-p', '--path', type=click.Path(resolve_path=True))
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


@cli.command('close')
@click.option('-a', '--all', 'close_all', is_flag=True)
@click.option('-f', '--force', is_flag=True)
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


@cli.command('activate')
@click.argument('index', type=int)
@handle_api_error
def activate_cmd(index: int) -> None:
    click.echo(f'Activate document at index "{index}"')
    client = get_client()
    if 1 <= index <= client.document_count:
        client.documents[index - 1].activate()
    else:
        raise click.ClickException(f'Index {index} out of range.')


@cli.command('docs')
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
@click.option('--to-original', 'destination', flag_value=C.wdCompareDestinationOriginal,
              help='Put diff into the original document.')
@click.option('--to-revised', 'destination', flag_value=C.wdCompareDestinationRevised,
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
        destination=destination or C.wdCompareDestinationNew,
        granularity=granularity or C.wdGranularityWordLevel,
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
@click.option('--to-original', 'destination', flag_value=C.wdMergeDestinationOriginalDocument,
              help='Merge result into the original document.')
@click.option('--to-revised', 'destination', flag_value=C.wdMergeDestinationRevisedDocument,
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
        destination=destination or C.wdMergeDestinationNewDocument,
        granularity=granularity or C.wdGranularityWordLevel,
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
    for plugin in plugin_eps:
        cli.add_command(plugin.load())
except Exception as e:
    click.echo(f'Warning: Failed to load plugins: {e}', err=True)


if __name__ == '__main__':
    cli()
