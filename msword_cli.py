# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pywin32>=306",
#     "click>=8.0",
# ]
# ///

import json
import sys
from datetime import date, datetime
from functools import wraps
from importlib.metadata import entry_points
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import click
from pywintypes import com_error
from win32com import client as com
from win32com.client import constants as C

VERSION = "0.4.0"
__version__ = VERSION


class WordAPIError(Exception):
    """Base exception for Word API errors."""


def _resolve_constant(value: Any) -> Any:
    if isinstance(value, str):
        try:
            return getattr(C, value)
        except AttributeError as error:
            raise WordAPIError(f"Unknown Word constant: {value}") from error
    return value


def _com_error_message(error: Exception, fallback: str = "COM error") -> str:
    excepinfo = getattr(error, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return excepinfo[2]
    return str(error) or fallback


def _safe_getattr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _iter_collection(collection: Any) -> Iterable[Any]:
    if collection is None:
        return []
    try:
        return list(collection)
    except Exception:
        pass
    count = _safe_getattr(collection, "Count", 0) or 0
    items = []
    for index in range(1, count + 1):
        try:
            if hasattr(collection, "Item"):
                items.append(collection.Item(index))
            else:
                items.append(collection(index))
        except Exception:
            continue
    return items


def _duplicate_range(range_obj: Any) -> Any:
    duplicate = _safe_getattr(range_obj, "Duplicate", None)
    return duplicate if duplicate is not None else range_obj


def _guess_property_type(value: Any) -> Any:
    if isinstance(value, bool):
        return _resolve_constant("msoPropertyTypeBoolean")
    if isinstance(value, int) and not isinstance(value, bool):
        return _resolve_constant("msoPropertyTypeNumber")
    if isinstance(value, float):
        return _resolve_constant("msoPropertyTypeFloat")
    return _resolve_constant("msoPropertyTypeString")


def _parse_property_value(value: str) -> Any:
    lowered = value.lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value


def _json_dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True)


def _normalize_output_path(path: str) -> str:
    return str(Path(path).resolve())


def _normalize_summary_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


class Document:
    """Represents a Microsoft Word document."""

    def __init__(self, com_doc: Any):
        self._doc = com_doc

    @property
    def native(self) -> Any:
        """Expose the underlying Word COM document for advanced automation."""
        return self._doc

    @property
    def name(self) -> str:
        return self._doc.Name

    @property
    def saved(self) -> bool:
        return self._doc.Saved

    def activate(self) -> None:
        try:
            self._doc.Activate()
        except com_error as error:
            raise WordAPIError(f"Activation failed: {_com_error_message(error)}") from error

    def save(self, path: Optional[str] = None, force: bool = False) -> None:
        try:
            if path:
                self._doc.SaveAs(_normalize_output_path(path))
            else:
                self._doc.Save(NoPrompt=force)
        except com_error as error:
            raise WordAPIError(f"Save failed: {_com_error_message(error)}") from error

    def save_copy(self, path: str) -> str:
        final_path = _normalize_output_path(path)
        try:
            self._doc.SaveCopyAs(final_path)
            return final_path
        except com_error as error:
            raise WordAPIError(f"Save copy failed: {_com_error_message(error)}") from error

    def close(self, force: bool = False) -> None:
        try:
            save_changes = (
                _resolve_constant("wdDoNotSaveChanges")
                if force
                else _resolve_constant("wdPromptToSaveChanges")
            )
            self._doc.Close(save_changes)
        except com_error as error:
            raise WordAPIError(f"Close failed: {_com_error_message(error)}") from error

    def export_fixed_format(
        self,
        path: str,
        format_val: Any = "wdExportFormatPDF",
        show: bool = False,
        optimize: Any = "wdExportOptimizeForPrint",
        pages: Optional[Tuple[int, int]] = None,
        rng: Any = None,
        markup: bool = False,
        properties: bool = False,
        irm: bool = False,
        bookmarks: Any = "wdExportCreateNoBookmarks",
        struct: bool = True,
        bitmap: bool = True,
        useiso19005_1: bool = False,
    ) -> str:
        try:
            path_obj = Path(path)
            if path_obj.is_dir():
                path_obj = path_obj / Path(self.name).stem
            if path_obj.suffix.lower() not in [".pdf", ".xps"]:
                ext = ".pdf" if _resolve_constant(format_val) == _resolve_constant("wdExportFormatPDF") else ".xps"
                path_obj = path_obj.with_suffix(ext)

            final_path = str(path_obj.resolve())
            options = {
                "OutputFileName": final_path,
                "ExportFormat": _resolve_constant(format_val),
                "OpenAfterExport": show,
                "OptimizeFor": _resolve_constant(optimize),
                "Range": _resolve_constant(rng) if rng else _resolve_constant("wdExportFromTo") if pages else _resolve_constant("wdExportAllDocument"),
                "Item": _resolve_constant("wdExportDocumentWithMarkup") if markup else _resolve_constant("wdExportDocumentContent"),
                "IncludeDocProps": properties,
                "KeepIRM": not irm,
                "CreateBookmarks": _resolve_constant(bookmarks),
                "DocStructureTags": not struct,
                "BitmapMissingFonts": not bitmap,
                "UseISO19005_1": useiso19005_1,
            }
            if pages:
                options["From"] = pages[0]
                options["To"] = pages[1]

            self._doc.ExportAsFixedFormat(**options)
            return final_path
        except com_error as error:
            raise WordAPIError(f"Export failed: {_com_error_message(error)}") from error

    def print_out(
        self,
        copies: int = 1,
        pages: Optional[str] = None,
        pagetype: Any = "wdPrintAllPages",
        rng: Any = "wdPrintAllDocument",
        item: Any = "wdPrintDocumentContent",
        no_collate: bool = False,
        to_file: Optional[str] = None,
        append: bool = False,
        columns: int = 1,
        rows: int = 1,
    ) -> None:
        try:
            options = {
                "Background": True,
                "Copies": copies,
                "Collate": not no_collate,
                "Item": _resolve_constant(item),
                "PrintToFile": False,
                "PageType": _resolve_constant(pagetype),
                "Range": _resolve_constant(rng),
                "PrintZoomColumn": columns,
                "PrintZoomRow": rows,
            }
            if pages:
                options["Pages"] = pages
                options["Range"] = _resolve_constant("wdPrintRangeOfPages")
            if to_file:
                options["PrintToFile"] = True
                options["OutputFileName"] = _normalize_output_path(to_file)
                if append:
                    options["Append"] = True

            self._doc.PrintOut(**options)
        except com_error as error:
            raise WordAPIError(f"Print failed: {_com_error_message(error)}") from error


class WordClient:
    """High-level client for Microsoft Word automation."""

    def __init__(self, visible: bool = False, quit_on_exit: bool = False):
        self._quit_on_exit = quit_on_exit
        self._word = None
        try:
            self._word = com.gencache.EnsureDispatch("Word.Application")
            if visible:
                self._word.Visible = True
        except com_error as error:
            raise WordAPIError(f"Failed to initialize Word: {_com_error_message(error)}") from error
        except Exception:
            raise WordAPIError("Unable to load 'Word.Application'.")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._quit_on_exit:
            self.quit()

    def _require_word(self) -> Any:
        if self._word is None:
            raise WordAPIError("Word client is closed.")
        return self._word

    def _active_range(self, scope: str) -> Any:
        range_obj = self.native.Selection.Range if scope == "selection" else self.active_document.native.Content
        return _duplicate_range(range_obj)

    def _comments_for_scope(self, scope: str) -> Any:
        return self.native.Selection.Range.Comments if scope == "selection" else self.active_document.native.Comments

    def _revisions_for_scope(self, scope: str) -> Any:
        return self.native.Selection.Range.Revisions if scope == "selection" else self.active_document.native.Revisions

    def _get_property_from_collection(self, collection: Any, name: str) -> Any:
        try:
            if hasattr(collection, "Item"):
                return collection.Item(name)
            return collection(name)
        except Exception:
            for item in _iter_collection(collection):
                item_name = _safe_getattr(item, "Name", "")
                if item_name and item_name.lower() == name.lower():
                    return item
        raise WordAPIError(f'Property "{name}" not found.')

    def _serialize_comment(self, comment: Any, index: int) -> Dict[str, Any]:
        return {
            "index": index,
            "author": _safe_getattr(comment, "Author", None),
            "initials": _safe_getattr(comment, "Initial", None),
            "date": str(_safe_getattr(comment, "Date", "")) or None,
            "text": _safe_getattr(_safe_getattr(comment, "Range", None), "Text", None),
            "scope_text": _safe_getattr(_safe_getattr(comment, "Scope", None), "Text", None),
        }

    def _built_in_property_value(self, document: Any, name: str) -> Any:
        try:
            item = self._get_property_from_collection(document.BuiltInDocumentProperties, name)
        except WordAPIError:
            return None
        return _normalize_summary_value(_safe_getattr(item, "Value", None))

    @property
    def closed(self) -> bool:
        return self._word is None

    @property
    def native(self) -> Any:
        """Expose the underlying Word COM application for advanced automation."""
        return self._require_word()

    @property
    def visible(self) -> bool:
        return self.native.Visible

    @visible.setter
    def visible(self, value: bool) -> None:
        self.native.Visible = value

    @property
    def template_dir(self) -> Path:
        return Path(self.native.Options.DefaultFilePath(_resolve_constant("wdUserTemplatesPath")))

    @property
    def active_document(self) -> Document:
        try:
            return Document(self.native.ActiveDocument)
        except Exception as error:
            raise WordAPIError("No active document available.") from error

    @property
    def documents(self) -> List[Document]:
        word = self.native
        return [Document(word.Documents.Item(i)) for i in range(1, word.Documents.Count + 1)]

    @property
    def document_count(self) -> int:
        return self.native.Documents.Count

    def open(self, path: str, visible: bool = True, read_only: bool = False, repair: bool = False) -> Document:
        try:
            word = self.native
            doc = word.Documents.Open(
                FileName=_normalize_output_path(path),
                Visible=visible,
                ReadOnly=read_only,
                OpenAndRepair=repair,
            )
            if visible and not word.Visible:
                word.Visible = True
            return Document(doc)
        except com_error as error:
            raise WordAPIError(f"Failed to open document: {_com_error_message(error)}") from error

    def new(self, template: Optional[str] = None, visible: bool = True) -> Document:
        try:
            word = self.native
            if template:
                doc = word.Documents.Add(Template=_normalize_output_path(template), Visible=visible)
            else:
                doc = word.Documents.Add(Visible=visible)
            if visible and not word.Visible:
                word.Visible = True
            return Document(doc)
        except com_error as error:
            raise WordAPIError(f"Failed to create new document: {_com_error_message(error)}") from error

    def compare(
        self,
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
        author: Optional[str] = None,
        ignore_warnings: bool = False,
    ) -> Document:
        orig_doc = None
        rev_doc = None
        result_doc = None
        try:
            word = self.native
            orig_doc = word.Documents.Open(FileName=_normalize_output_path(original), Visible=False)
            rev_doc = word.Documents.Open(FileName=_normalize_output_path(revised), Visible=False)
            resolved_destination = _resolve_constant(destination)
            result = word.CompareDocuments(
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
                RevisedAuthor=author or word.UserName,
                IgnoreAllComparisonWarnings=ignore_warnings,
            )
            if resolved_destination == _resolve_constant("wdCompareDestinationOriginal"):
                result_doc = orig_doc
            elif resolved_destination == _resolve_constant("wdCompareDestinationRevised"):
                result_doc = rev_doc
            return Document(result)
        except com_error as error:
            raise WordAPIError(f"Compare failed: {_com_error_message(error)}") from error
        finally:
            for temp_doc in (orig_doc, rev_doc):
                if temp_doc is not None and temp_doc is not result_doc:
                    try:
                        temp_doc.Close(_resolve_constant("wdDoNotSaveChanges"))
                    except Exception:
                        pass

    def merge(
        self,
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
        author: Optional[str] = None,
        ignore_warnings: bool = False,
    ) -> Document:
        orig_doc = None
        rev_doc = None
        result_doc = None
        try:
            word = self.native
            orig_doc = word.Documents.Open(FileName=_normalize_output_path(original), Visible=False)
            rev_doc = word.Documents.Open(FileName=_normalize_output_path(revised), Visible=False)
            resolved_destination = _resolve_constant(destination)
            result = word.MergeDocuments(
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
                RevisedAuthor=author or word.UserName,
                IgnoreAllComparisonWarnings=ignore_warnings,
            )
            if resolved_destination == _resolve_constant("wdMergeDestinationOriginalDocument"):
                result_doc = orig_doc
            elif resolved_destination == _resolve_constant("wdMergeDestinationRevisedDocument"):
                result_doc = rev_doc
            return Document(result)
        except com_error as error:
            raise WordAPIError(f"Merge failed: {_com_error_message(error)}") from error
        finally:
            for temp_doc in (orig_doc, rev_doc):
                if temp_doc is not None and temp_doc is not result_doc:
                    try:
                        temp_doc.Close(_resolve_constant("wdDoNotSaveChanges"))
                    except Exception:
                        pass

    def set_track_changes(self, enabled: bool) -> bool:
        try:
            document = self.active_document.native
            document.TrackRevisions = enabled
            return bool(document.TrackRevisions)
        except com_error as error:
            raise WordAPIError(f"Track changes update failed: {_com_error_message(error)}") from error

    def accept_revisions(self, scope: str = "document") -> int:
        try:
            revisions = self._revisions_for_scope(scope)
            count = _safe_getattr(revisions, "Count", 0) or 0
            revisions.AcceptAll()
            return count
        except com_error as error:
            raise WordAPIError(f"Accept revisions failed: {_com_error_message(error)}") from error

    def reject_revisions(self, scope: str = "document") -> int:
        try:
            revisions = self._revisions_for_scope(scope)
            count = _safe_getattr(revisions, "Count", 0) or 0
            revisions.RejectAll()
            return count
        except com_error as error:
            raise WordAPIError(f"Reject revisions failed: {_com_error_message(error)}") from error

    def list_comments(self, scope: str = "document") -> List[Dict[str, Any]]:
        try:
            comments = self._comments_for_scope(scope)
            return [self._serialize_comment(comment, index) for index, comment in enumerate(_iter_collection(comments), start=1)]
        except com_error as error:
            raise WordAPIError(f"List comments failed: {_com_error_message(error)}") from error

    def delete_comments(self, scope: str = "document") -> int:
        try:
            comments = list(_iter_collection(self._comments_for_scope(scope)))
            for comment in reversed(comments):
                comment.Delete()
            return len(comments)
        except com_error as error:
            raise WordAPIError(f"Delete comments failed: {_com_error_message(error)}") from error

    def export_comments(self, path: str, scope: str = "document", output_format: str = "text") -> str:
        comments = self.list_comments(scope=scope)
        final_path = _normalize_output_path(path)
        output = Path(final_path)
        if output_format == "json":
            output.write_text(_json_dump(comments) + "\n", encoding="utf-8")
        else:
            lines = [f'[{item["index"]}] {item.get("author") or "Unknown"}: {item.get("text") or ""}' for item in comments]
            output.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
        return final_path

    def update_fields(self, scope: str = "document") -> Dict[str, int]:
        try:
            if scope == "selection":
                fields = self.native.Selection.Range.Fields
                count = _safe_getattr(fields, "Count", 0) or 0
                fields.Update()
                return {"fields": count, "tables_of_contents": 0}

            document = self.active_document.native
            fields = document.Fields
            field_count = _safe_getattr(fields, "Count", 0) or 0
            fields.Update()
            toc_count = 0
            for toc in _iter_collection(_safe_getattr(document, "TablesOfContents", None)):
                toc_count += 1
                toc.Update()
            return {"fields": field_count, "tables_of_contents": toc_count}
        except com_error as error:
            raise WordAPIError(f"Update fields failed: {_com_error_message(error)}") from error

    def find(self, text: str, scope: str = "document", match_case: bool = False, whole_word: bool = False) -> List[Dict[str, Any]]:
        try:
            search_range = self._active_range(scope)
            end_limit = _safe_getattr(search_range, "End", 0) or 0
            results = []
            while search_range.Find.Execute(FindText=text, MatchCase=match_case, MatchWholeWord=whole_word, Forward=True, Wrap=_resolve_constant("wdFindStop"), Format=False):
                start = _safe_getattr(search_range, "Start", 0) or 0
                end = _safe_getattr(search_range, "End", 0) or 0
                results.append({"start": start, "end": end, "text": _safe_getattr(search_range, "Text", None)})
                if end <= start or end >= end_limit:
                    break
                search_range.SetRange(end, end_limit)
            return results
        except com_error as error:
            raise WordAPIError(f"Find failed: {_com_error_message(error)}") from error

    def replace(self, find_text: str, replace_with: str, scope: str = "document", match_case: bool = False, whole_word: bool = False) -> int:
        try:
            matches = self.find(find_text, scope=scope, match_case=match_case, whole_word=whole_word)
            replace_range = self._active_range(scope)
            replace_range.Find.Execute(FindText=find_text, ReplaceWith=replace_with, MatchCase=match_case, MatchWholeWord=whole_word, Forward=True, Wrap=_resolve_constant("wdFindStop"), Format=False, Replace=_resolve_constant("wdReplaceAll"))
            return len(matches)
        except com_error as error:
            raise WordAPIError(f"Replace failed: {_com_error_message(error)}") from error

    def list_properties(self, kind: str = "all") -> List[Dict[str, Any]]:
        try:
            document = self.active_document.native
            results = []
            if kind in {"all", "built-in"}:
                for item in _iter_collection(document.BuiltInDocumentProperties):
                    results.append({"kind": "built-in", "name": _safe_getattr(item, "Name", None), "value": _safe_getattr(item, "Value", None)})
            if kind in {"all", "custom"}:
                for item in _iter_collection(document.CustomDocumentProperties):
                    results.append({"kind": "custom", "name": _safe_getattr(item, "Name", None), "value": _safe_getattr(item, "Value", None)})
            return results
        except com_error as error:
            raise WordAPIError(f"List properties failed: {_com_error_message(error)}") from error

    def get_property(self, name: str) -> Dict[str, Any]:
        document = self.active_document.native
        for kind, collection in (("built-in", document.BuiltInDocumentProperties), ("custom", document.CustomDocumentProperties)):
            try:
                item = self._get_property_from_collection(collection, name)
                return {"kind": kind, "name": _safe_getattr(item, "Name", name), "value": _safe_getattr(item, "Value", None)}
            except WordAPIError:
                continue
        raise WordAPIError(f'Property "{name}" not found.')

    def set_property(self, name: str, value: Any) -> Dict[str, Any]:
        document = self.active_document.native
        for kind, collection in (("built-in", document.BuiltInDocumentProperties), ("custom", document.CustomDocumentProperties)):
            try:
                item = self._get_property_from_collection(collection, name)
                item.Value = value
                return {"kind": kind, "name": _safe_getattr(item, "Name", name), "value": _safe_getattr(item, "Value", value)}
            except WordAPIError:
                continue
        try:
            document.CustomDocumentProperties.Add(Name=name, LinkToContent=False, Type=_guess_property_type(value), Value=value)
        except com_error as error:
            raise WordAPIError(f"Set property failed: {_com_error_message(error)}") from error
        return self.get_property(name)

    def delete_property(self, name: str) -> bool:
        document = self.active_document.native
        try:
            item = self._get_property_from_collection(document.CustomDocumentProperties, name)
            item.Delete()
            return True
        except WordAPIError as error:
            raise WordAPIError(f'Custom property "{name}" not found.') from error
        except com_error as error:
            raise WordAPIError(f"Delete property failed: {_com_error_message(error)}") from error

    def statistics(self) -> Dict[str, Optional[int]]:
        document = self.active_document.native
        stats = {}
        for key, constant_name in (("pages", "wdStatisticPages"), ("words", "wdStatisticWords"), ("characters", "wdStatisticCharacters"), ("paragraphs", "wdStatisticParagraphs")):
            try:
                stats[key] = document.ComputeStatistics(_resolve_constant(constant_name))
            except Exception:
                stats[key] = None
        return stats

    def summary(self) -> Dict[str, Any]:
        doc = self.active_document
        native = doc.native
        return {
            "name": doc.name,
            "path": _safe_getattr(native, "FullName", None),
            "saved": doc.saved,
            "read_only": bool(_safe_getattr(native, "ReadOnly", False)),
            "track_changes": bool(_safe_getattr(native, "TrackRevisions", False)),
            "revisions": _safe_getattr(_safe_getattr(native, "Revisions", None), "Count", 0) or 0,
            "comments": _safe_getattr(_safe_getattr(native, "Comments", None), "Count", 0) or 0,
            "template": _safe_getattr(_safe_getattr(native, "AttachedTemplate", None), "FullName", None),
            "author": self._built_in_property_value(native, "Author"),
            "last_author": self._built_in_property_value(native, "Last Author"),
            "created_at": self._built_in_property_value(native, "Creation Date"),
            "modified_at": self._built_in_property_value(native, "Last Save Time"),
        }

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
        except WordAPIError as error:
            raise click.ClickException(str(error))

    return wrapper


class CliTemplate(click.Path):
    def convert(self, value: str, param: Optional[click.Parameter], ctx: Optional[click.Context]) -> str:
        path_obj = Path(value)
        if not path_obj.is_absolute():
            if path_obj.resolve().exists():
                value = str(path_obj.resolve())
            else:
                value = str(get_client().template_dir / path_obj)
        return super().convert(value, param, ctx)


def validate_range(ctx: click.Context, param: click.Parameter, value: Optional[str]) -> Optional[Tuple[int, int]]:
    try:
        if value is not None:
            frm, to = map(int, value.split("-", 2))
            if frm < 1 or frm > to:
                raise ValueError
            return (frm, to)
    except ValueError as error:
        raise click.BadParameter('Range must be in the format "x-y" where "x" and "y" are positive integers.') from error
    return None


def print_version(ctx: click.Context, param: click.Parameter, value: bool) -> None:
    if not value or ctx.resilient_parsing:
        return
    click.echo(f"Version {VERSION}")
    ctx.exit()


def _scope_option(func):
    return click.option("--scope", type=click.Choice(["document", "selection"]), default="document", show_default=True, help="Apply the operation to the whole document or the current selection.")(func)


def _format_option(func):
    return click.option("--format", "output_format", type=click.Choice(["text", "json"]), default="text", show_default=True, help="Choose human-readable or machine-readable output.")(func)


def _emit_data(value: Any, output_format: str) -> None:
    if output_format == "json":
        click.echo(_json_dump(value))
    else:
        click.echo(value)


def _render_documents(client: WordClient) -> str:
    if client.document_count == 0:
        return "\nNo open documents found."
    lines = ["", "Open Documents:", ""]
    pad_len = len(str(client.document_count))
    active_name = client.active_document.name
    for index, doc in enumerate(client.documents, start=1):
        active = "*" if doc.name == active_name else " "
        saved = "*" if not doc.saved else ""
        lines.append(f" {active} [{index: ={pad_len}}] {doc.name}{saved}")
    return "\n".join(lines)


class SectionedHelpGroup(click.Group):
    COMMAND_SECTIONS = {
        "Documents": ["open", "new", "save", "save-as", "save-copy", "close", "activate", "list-documents", "compare", "merge"],
        "Content": ["find", "replace", "update-fields", "print", "export"],
        "Review": ["track-changes", "accept-revisions", "reject-revisions", "list-comments", "export-comments", "delete-comments"],
        "Document Data": ["summary", "statistics", "list-properties", "get-property", "set-property", "delete-property"],
    }

    def format_commands(self, ctx: click.Context, formatter: click.HelpFormatter) -> None:
        commands = []
        for subcommand_name in self.list_commands(ctx):
            command = self.get_command(ctx, subcommand_name)
            if command is None or command.hidden:
                continue
            commands.append((subcommand_name, command))
        remaining = {name: command for name, command in commands}
        for section_name, names in self.COMMAND_SECTIONS.items():
            rows = []
            for name in names:
                command = remaining.pop(name, None)
                if command is not None:
                    rows.append((name, command.get_short_help_str()))
            if rows:
                with formatter.section(section_name):
                    formatter.write_dl(rows)
        if remaining:
            plugin_rows = [(name, command.get_short_help_str()) for name, command in sorted(remaining.items(), key=lambda item: item[0])]
            with formatter.section("Plugins"):
                formatter.write_dl(plugin_rows)


_COMPARE_OPTIONS = [
    click.option("--char-level", "granularity", flag_value="wdGranularityCharLevel", help="Compare at character level (default: word level)."),
    click.option("--no-formatting", "formatting", is_flag=True, default=False, help="Ignore formatting differences."),
    click.option("--no-case-changes", "case_changes", is_flag=True, default=False, help="Ignore case change differences."),
    click.option("--no-whitespace", "whitespace", is_flag=True, default=False, help="Ignore whitespace differences."),
    click.option("--no-tables", "tables", is_flag=True, default=False, help="Ignore table differences."),
    click.option("--no-headers", "headers", is_flag=True, default=False, help="Ignore header/footer differences."),
    click.option("--no-footnotes", "footnotes", is_flag=True, default=False, help="Ignore footnote differences."),
    click.option("--no-textboxes", "textboxes", is_flag=True, default=False, help="Ignore text box differences."),
    click.option("--no-fields", "fields", is_flag=True, default=False, help="Ignore field differences."),
    click.option("--no-comments", "comments", is_flag=True, default=False, help="Ignore comment differences."),
    click.option("--no-moves", "moves", is_flag=True, default=False, help="Ignore move differences."),
    click.option("--author", type=str, default=None, help="Author name for tracked changes (defaults to Word username)."),
    click.option("--ignore-warnings", is_flag=True, default=False, help="Suppress all comparison warning dialogs."),
]


def add_compare_options(func):
    for option in reversed(_COMPARE_OPTIONS):
        func = option(func)
    return func


@click.group(chain=True, cls=SectionedHelpGroup)
@click.option("--version", is_flag=True, callback=print_version, expose_value=False, is_eager=True)
def cli() -> None:
    """Command line interface for Microsoft Word."""


@cli.command("open", short_help="Open a document.", help="Open an existing document and make it active.")
@click.argument("path", type=click.Path(exists=True, resolve_path=True))
@click.option("--show/--hide", default=True, help="Display or hide the document.")
@click.option("--readonly", "read_only", is_flag=True, help="Open without write access when Word supports it.")
@click.option("--repair", is_flag=True, help="Ask Word to repair the document while opening it.")
@handle_api_error
def open_cmd(path: str, show: bool, read_only: bool, repair: bool) -> None:
    click.echo(f'Opening document at "{path}"')
    get_client().open(path, visible=show, read_only=read_only, repair=repair)


@cli.command("new", short_help="Create a new document.", help="Create a new document, optionally from a template.")
@click.option("-t", "--template", type=CliTemplate(exists=True, dir_okay=False, resolve_path=True), help="Path to a Word template file.")
@click.option("--show/--hide", default=True, help="Display or hide the document.")
@handle_api_error
def new_cmd(template: Optional[str], show: bool) -> None:
    click.echo(f'Opening new document using template: "{template}"' if template else "Opening new blank document.")
    get_client().new(template=template, visible=show)


PRINT_OUT_ITEMS = {
    "document_content": "wdPrintDocumentContent",
    "doc_with_markup": "wdPrintDocumentWithMarkup",
    "comments": "wdPrintComments",
    "properties": "wdPrintProperties",
    "markup": "wdPrintMarkup",
    "styles": "wdPrintStyles",
    "auto_text_entries": "wdPrintAutoTextEntries",
    "key_assignments": "wdPrintKeyAssignments",
    "envelope": "wdPrintEnvelope",
}


@cli.command("print", short_help="Print the active document.", help="Print the active document.")
@click.option("-c", "--copies", type=click.IntRange(min=1), default=1, help="Number of copies to print.")
@click.option("-p", "--pages", type=str, help="Specific pages or page ranges, for example 2-4, 6.")
@click.option("--even", "pagetype", flag_value="wdPrintEvenPagesOnly", help="Print even-numbered pages only.")
@click.option("--odd", "pagetype", flag_value="wdPrintOddPagesOnly", help="Print odd-numbered pages only.")
@click.option("--current-page", "rng", flag_value="wdPrintCurrentPage", help="Print only the current page.")
@click.option("--selection", "rng", flag_value="wdPrintSelection", help="Print only the current selection.")
@click.option("--no-collate", is_flag=True, help="Disable collation when printing multiple copies.")
@click.option("--to-file", type=click.Path(dir_okay=False, resolve_path=True), help="Print to a file instead of a printer.")
@click.option("--append", is_flag=True, help="Append to the output file when supported.")
@click.option("--columns", type=click.Choice(["1", "2", "3", "4"]), default="1", help="Pages across each row.")
@click.option("--rows", type=click.Choice(["1", "2", "4"]), default="1", help="Pages down each column.")
@click.option("--item", type=click.Choice(list(PRINT_OUT_ITEMS.keys())), default="document_content", help="Select which document content Word should print.")
@handle_api_error
def print_cmd(copies: int, pages: Optional[str], pagetype: Any, rng: Any, item: str, no_collate: bool, to_file: Optional[str], append: bool, columns: str, rows: str) -> None:
    click.echo(f"Printing {copies} copies of pages: {pages or 'all'}")
    doc = get_client().active_document
    doc.print_out(copies=copies, pages=pages, pagetype=_resolve_constant(pagetype or "wdPrintAllPages"), rng=_resolve_constant(rng or "wdPrintAllDocument"), item=_resolve_constant(PRINT_OUT_ITEMS[item]), no_collate=no_collate, to_file=to_file, append=append, columns=int(columns), rows=int(rows))


@cli.command("export", short_help="Export to PDF or XPS.", help="Export the active document to PDF or XPS.")
@click.option("--pdf", "format_val", flag_value="wdExportFormatPDF", default=True, help="Export as PDF.")
@click.option("--xps", "format_val", flag_value="wdExportFormatXPS", help="Export as XPS.")
@click.option("--show", is_flag=True, help="Open the exported file after creating it.")
@click.option("--for-print", "optimize", flag_value="wdExportOptimizeForPrint", default=True, help="Optimize the export for printing.")
@click.option("--for-screen", "optimize", flag_value="wdExportOptimizeForOnScreen", help="Optimize for on-screen viewing.")
@click.option("--pages", type=str, callback=validate_range, help="Export only a page range in the form x-y.")
@click.option("--current-page", "rng", flag_value="wdExportCurrentPage", help="Export only the current page.")
@click.option("--selection", "rng", flag_value="wdExportSelection", help="Export only the current selection.")
@click.option("--with-markup", "markup", is_flag=True, help="Include tracked changes and markup.")
@click.option("--with-properties", "properties", is_flag=True, help="Include document properties.")
@click.option("--without-irm", "irm", is_flag=True, help="Do not preserve IRM permissions in the export.")
@click.option("--with-heading-bookmarks", "bookmarks", flag_value="wdExportCreateHeadingBookmarks", help="Create bookmarks from headings.")
@click.option("--with-word-bookmarks", "bookmarks", flag_value="wdExportCreateWordBookmarks", help="Create bookmarks from Word bookmarks.")
@click.option("--without-structure-tags", "struct", is_flag=True, help="Disable document structure tags in the export.")
@click.option("--without-bitmap-fonts", "bitmap", is_flag=True, help="Disable bitmap fallback for missing fonts.")
@click.option("--pdf-a", "useiso19005_1", is_flag=True, help="Create a PDF/A-compatible export.")
@click.argument("path", type=click.Path(dir_okay=True, resolve_path=True))
@handle_api_error
def export_cmd(path: str, format_val: Any, show: bool, optimize: Any, pages: Optional[Tuple[int, int]], rng: Any, markup: bool, properties: bool, irm: bool, bookmarks: Any, struct: bool, bitmap: bool, useiso19005_1: bool) -> None:
    doc = get_client().active_document
    final_path = doc.export_fixed_format(path=path, format_val=_resolve_constant(format_val), show=show, optimize=_resolve_constant(optimize), pages=pages, rng=_resolve_constant(rng) if rng else None, markup=markup, properties=properties, irm=irm, bookmarks=_resolve_constant(bookmarks or "wdExportCreateNoBookmarks"), struct=struct, bitmap=bitmap, useiso19005_1=useiso19005_1)
    click.echo(f'Exported to "{final_path}"')


@cli.command("save", short_help="Save current changes.", help="Save the active document or all open documents.")
@click.option("-a", "--all", "save_all", is_flag=True, help="Save all open documents.")
@click.option("-f", "--force", is_flag=True, help="Save without prompting when Word supports it.")
@handle_api_error
def save_cmd(save_all: bool, force: bool) -> None:
    client = get_client()
    click.echo("Saving changes to existing document(s).")
    docs_to_save = client.documents if save_all else [client.active_document]
    for document in docs_to_save:
        document.save(force=force)


@cli.command("save-as", short_help="Save the active document to a new path.", help="Save the active document to a new path.")
@click.argument("path", type=click.Path(resolve_path=True))
@handle_api_error
def save_as_cmd(path: str) -> None:
    click.echo(f'Saving active document as "{path}"')
    get_client().active_document.save(path=path)


@cli.command("save-copy", short_help="Save a copy without switching the active document.", help="Save a copy of the active document without changing the original.")
@click.argument("path", type=click.Path(resolve_path=True))
@handle_api_error
def save_copy_cmd(path: str) -> None:
    final_path = get_client().active_document.save_copy(path)
    click.echo(f'Saved copy to "{final_path}"')


@cli.command("close", short_help="Close documents.", help="Close the active document or all open documents.")
@click.option("-a", "--all", "close_all", is_flag=True, help="Close all open documents.")
@click.option("-f", "--force", is_flag=True, help="Discard unsaved changes without prompting.")
@handle_api_error
def close_cmd(close_all: bool, force: bool) -> None:
    client = get_client()
    docs_to_close = client.documents if close_all else [client.active_document]
    action = "Force closing" if force else "Closing"
    click.echo(f"{action} document(s)...")
    for document in docs_to_close:
        document.close(force=force)
    if client.document_count == 0:
        client.quit()


@cli.command("activate", short_help="Activate an open document by index.", help="Activate an open document by its index from the list-documents output.")
@click.argument("index", type=int)
@handle_api_error
def activate_cmd(index: int) -> None:
    click.echo(f'Activate document at index "{index}"')
    client = get_client()
    if 1 <= index <= client.document_count:
        client.documents[index - 1].activate()
        return
    raise click.ClickException(f"Index {index} out of range.")


@cli.command("list-documents", short_help="List open documents.", help="List open documents and show which one is active.")
@handle_api_error
def list_documents_cmd() -> None:
    click.echo(_render_documents(get_client()))



@cli.command("compare", short_help="Compare two documents.", help="Compare ORIGINAL and REVISED documents, showing differences as tracked changes.")
@click.argument("original", type=click.Path(exists=True, resolve_path=True))
@click.argument("revised", type=click.Path(exists=True, resolve_path=True))
@click.option("--to-original", "destination", flag_value="wdCompareDestinationOriginal", help="Put diff into the original document.")
@click.option("--to-revised", "destination", flag_value="wdCompareDestinationRevised", help="Put diff into the revised document.")
@add_compare_options
@handle_api_error
def compare_cmd(original: str, revised: str, destination: Any, granularity: Any, formatting: bool, case_changes: bool, whitespace: bool, tables: bool, headers: bool, footnotes: bool, textboxes: bool, fields: bool, comments: bool, moves: bool, author: Optional[str], ignore_warnings: bool) -> None:
    click.echo(f'Comparing "{original}" with "{revised}"')
    result = get_client().compare(original=original, revised=revised, destination=_resolve_constant(destination or "wdCompareDestinationNew"), granularity=_resolve_constant(granularity or "wdGranularityWordLevel"), formatting=not formatting, case_changes=not case_changes, whitespace=not whitespace, tables=not tables, headers=not headers, footnotes=not footnotes, textboxes=not textboxes, fields=not fields, comments=not comments, moves=not moves, author=author, ignore_warnings=ignore_warnings)
    click.echo(f'Result document: "{result.name}"')


@cli.command("merge", short_help="Merge two documents.", help="Merge ORIGINAL and REVISED documents, combining their tracked changes.")
@click.argument("original", type=click.Path(exists=True, resolve_path=True))
@click.argument("revised", type=click.Path(exists=True, resolve_path=True))
@click.option("--to-original", "destination", flag_value="wdMergeDestinationOriginalDocument", help="Merge result into the original document.")
@click.option("--to-revised", "destination", flag_value="wdMergeDestinationRevisedDocument", help="Merge result into the revised document.")
@add_compare_options
@handle_api_error
def merge_cmd(original: str, revised: str, destination: Any, granularity: Any, formatting: bool, case_changes: bool, whitespace: bool, tables: bool, headers: bool, footnotes: bool, textboxes: bool, fields: bool, comments: bool, moves: bool, author: Optional[str], ignore_warnings: bool) -> None:
    click.echo(f'Merging "{original}" with "{revised}"')
    result = get_client().merge(original=original, revised=revised, destination=_resolve_constant(destination or "wdMergeDestinationNewDocument"), granularity=_resolve_constant(granularity or "wdGranularityWordLevel"), formatting=not formatting, case_changes=not case_changes, whitespace=not whitespace, tables=not tables, headers=not headers, footnotes=not footnotes, textboxes=not textboxes, fields=not fields, comments=not comments, moves=not moves, author=author, ignore_warnings=ignore_warnings)
    click.echo(f'Result document: "{result.name}"')


@cli.command("find", short_help="Find text.", help="Find text in the active document.")
@click.argument("text")
@_scope_option
@click.option("--match-case", is_flag=True, help="Match letter case.")
@click.option("--whole-word", is_flag=True, help="Match whole words only.")
@_format_option
@handle_api_error
def find_cmd(text: str, scope: str, match_case: bool, whole_word: bool, output_format: str) -> None:
    matches = get_client().find(text, scope=scope, match_case=match_case, whole_word=whole_word)
    if output_format == "json":
        _emit_data(matches, output_format)
        return
    if not matches:
        click.echo("No matches found.")
        return
    for match in matches:
        click.echo(f'[{match["start"]}:{match["end"]}] {match["text"]}')


@cli.command("replace", short_help="Replace text.", help="Replace text in the active document.")
@click.argument("find_text")
@click.argument("replace_with")
@_scope_option
@click.option("--match-case", is_flag=True, help="Match letter case.")
@click.option("--whole-word", is_flag=True, help="Match whole words only.")
@_format_option
@handle_api_error
def replace_cmd(find_text: str, replace_with: str, scope: str, match_case: bool, whole_word: bool, output_format: str) -> None:
    replaced = get_client().replace(find_text, replace_with, scope=scope, match_case=match_case, whole_word=whole_word)
    payload = {"replaced": replaced}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f"Replaced {replaced} occurrence(s).")


@cli.command("track-changes", short_help="Enable or disable tracked changes.", help="Enable or disable tracked changes on the active document.")
@click.option("--on/--off", "enabled", default=True, show_default=True, help="Turn tracked changes on or off.")
@_format_option
@handle_api_error
def track_changes_cmd(enabled: bool, output_format: str) -> None:
    current = get_client().set_track_changes(enabled)
    payload = {"track_changes": current}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f'Track changes {"enabled" if current else "disabled"}.')


@cli.command("accept-revisions", short_help="Accept revisions.", help="Accept revisions in the active document or selection.")
@_scope_option
@_format_option
@handle_api_error
def accept_revisions_cmd(scope: str, output_format: str) -> None:
    count = get_client().accept_revisions(scope=scope)
    payload = {"accepted": count, "scope": scope}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f"Accepted {count} revision(s).")


@cli.command("reject-revisions", short_help="Reject revisions.", help="Reject revisions in the active document or selection.")
@_scope_option
@_format_option
@handle_api_error
def reject_revisions_cmd(scope: str, output_format: str) -> None:
    count = get_client().reject_revisions(scope=scope)
    payload = {"rejected": count, "scope": scope}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f"Rejected {count} revision(s).")


@cli.command("list-comments", short_help="List comments.", help="List comments in the active document or selection.")
@_scope_option
@_format_option
@handle_api_error
def list_comments_cmd(scope: str, output_format: str) -> None:
    comments = get_client().list_comments(scope=scope)
    if output_format == "json":
        _emit_data(comments, output_format)
        return
    if not comments:
        click.echo("No comments found.")
        return
    for item in comments:
        click.echo(f'[{item["index"]}] {item.get("author") or "Unknown"}: {item.get("text") or ""}')


@cli.command("export-comments", short_help="Export comments.", help="Export comments from the active document or selection.")
@click.argument("path", type=click.Path(resolve_path=True))
@_scope_option
@click.option("--json", "output_format", flag_value="json", help="Export comments as JSON.")
@click.option("--text", "output_format", flag_value="text", default=True, help="Export comments as plain text.")
@handle_api_error
def export_comments_cmd(path: str, scope: str, output_format: str) -> None:
    final_path = get_client().export_comments(path, scope=scope, output_format=output_format)
    click.echo(f'Exported comments to "{final_path}"')


@cli.command("delete-comments", short_help="Delete comments.", help="Delete comments in the active document or selection.")
@_scope_option
@_format_option
@handle_api_error
def delete_comments_cmd(scope: str, output_format: str) -> None:
    count = get_client().delete_comments(scope=scope)
    payload = {"deleted": count, "scope": scope}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f"Deleted {count} comment(s).")


@cli.command("update-fields", short_help="Update fields.", help="Update fields in the active document or selection.")
@_scope_option
@_format_option
@handle_api_error
def update_fields_cmd(scope: str, output_format: str) -> None:
    result = get_client().update_fields(scope=scope)
    if output_format == "json":
        _emit_data(result, output_format)
    else:
        click.echo(f'Updated {result["fields"]} field(s)' + (f' and {result["tables_of_contents"]} table(s) of contents.' if result["tables_of_contents"] else "."))


@cli.command("summary", short_help="Show active document summary.", help="Show active document state plus a small set of common metadata.")
@_format_option
@handle_api_error
def summary_cmd(output_format: str) -> None:
    summary = get_client().summary()
    if output_format == "json":
        _emit_data(summary, output_format)
        return
    for key, value in summary.items():
        click.echo(f"{key}: {value}")


@cli.command("statistics", short_help="Show document statistics.", help="Show page, word, character, and paragraph counts.")
@_format_option
@handle_api_error
def statistics_cmd(output_format: str) -> None:
    stats = get_client().statistics()
    if output_format == "json":
        _emit_data(stats, output_format)
        return
    for key, value in stats.items():
        click.echo(f"{key}: {value}")


@cli.command("list-properties", short_help="List raw document properties.", help="List built-in and custom Word document properties without curating the output.")
@click.option("--kind", type=click.Choice(["all", "built-in", "custom"]), default="all", show_default=True, help="Choose which raw property set to inspect.")
@_format_option
@handle_api_error
def list_properties_cmd(kind: str, output_format: str) -> None:
    properties = get_client().list_properties(kind=kind)
    if output_format == "json":
        _emit_data(properties, output_format)
        return
    if not properties:
        click.echo("No properties found.")
        return
    for item in properties:
        click.echo(f'{item["kind"]}: {item["name"]} = {item["value"]}')


@cli.command("get-property", short_help="Read one raw property.", help="Read one built-in or custom Word document property by name.")
@click.argument("name")
@_format_option
@handle_api_error
def get_property_cmd(name: str, output_format: str) -> None:
    prop = get_client().get_property(name)
    _emit_data(prop, output_format) if output_format == "json" else click.echo(f'{prop["kind"]}: {prop["name"]} = {prop["value"]}')


@cli.command("set-property", short_help="Write one raw property.", help="Write a built-in or custom Word document property by name.")
@click.argument("name")
@click.argument("value")
@_format_option
@handle_api_error
def set_property_cmd(name: str, value: str, output_format: str) -> None:
    prop = get_client().set_property(name, _parse_property_value(value))
    _emit_data(prop, output_format) if output_format == "json" else click.echo(f'Set {prop["kind"]} property {prop["name"]} = {prop["value"]}')


@cli.command("delete-property", short_help="Delete one custom raw property.", help="Delete a custom Word document property by name.")
@click.argument("name")
@_format_option
@handle_api_error
def delete_property_cmd(name: str, output_format: str) -> None:
    get_client().delete_property(name)
    payload = {"deleted": True, "name": name}
    _emit_data(payload, output_format) if output_format == "json" else click.echo(f'Deleted custom property "{name}".')


try:
    eps = entry_points()
    plugin_eps = eps.get("msw.plugin", []) if sys.version_info < (3, 10) else eps.select(group="msw.plugin")
except Exception as error:
    click.echo(f"Warning: Failed to discover plugins: {error}", err=True)
else:
    for plugin in plugin_eps:
        try:
            cli.add_command(plugin.load())
        except Exception as error:
            plugin_name = getattr(plugin, "name", "<unknown>")
            click.echo(f"Warning: Failed to load plugin {plugin_name}: {error}", err=True)


if __name__ == "__main__":
    cli()

