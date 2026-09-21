import datetime as dt
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import click
import pytest
from click.testing import CliRunner

from tests.conftest import FakeCollection, FakeComment, FakeCommentsCollection, FakeComDocument, FakeFieldsCollection, FakePropertyCollection, FakeRevisionsCollection, FakeWordApp


def test_print_version_does_not_initialize_client(msword_cli, monkeypatch):
    runner = CliRunner()
    get_client = Mock(side_effect=AssertionError("client should not be initialized"))
    monkeypatch.setattr(msword_cli, "get_client", get_client)

    result = runner.invoke(msword_cli.cli, ["--version"])

    assert result.exit_code == 0
    assert f"Version {msword_cli.VERSION}" in result.output
    get_client.assert_not_called()


def test_version_matches_project_metadata(msword_cli):
    pyproject = Path(msword_cli.__file__).with_name("pyproject.toml").read_text(encoding="utf-8")

    assert f'version = "{msword_cli.__version__}"' in pyproject


def test_validate_range_parses_valid_input(msword_cli):
    assert msword_cli.validate_range(None, None, "2-5") == (2, 5)


@pytest.mark.parametrize("value", ["5-2", "0-2", "1-0"])
def test_validate_range_rejects_invalid_input(msword_cli, value):
    with pytest.raises(click.BadParameter):
        msword_cli.validate_range(None, None, value)


def test_handle_api_error_converts_word_errors(msword_cli):
    @msword_cli.handle_api_error
    def boom():
        raise msword_cli.WordAPIError("broken")

    with pytest.raises(click.ClickException, match="broken"):
        boom()


def test_get_client_is_singleton(msword_cli, monkeypatch):
    client_instance = Mock()
    client_instance.closed = False
    word_client_ctor = Mock(return_value=client_instance)
    monkeypatch.setattr(msword_cli, "_CLI_CLIENT", None)
    monkeypatch.setattr(msword_cli, "WordClient", word_client_ctor)

    first = msword_cli.get_client()
    second = msword_cli.get_client()

    assert first is client_instance
    assert second is client_instance
    word_client_ctor.assert_called_once_with(visible=False)


def test_document_native_exposes_wrapped_com_document(msword_cli):
    com_doc = FakeComDocument("report.docx")

    assert msword_cli.Document(com_doc).native is com_doc


def test_document_export_fixed_format_normalizes_path_and_options(msword_cli):
    com_doc = FakeComDocument("report.docx")
    document = msword_cli.Document(com_doc)

    with CliRunner().isolated_filesystem():
        out_dir = Path("exports")
        out_dir.mkdir()
        exported = document.export_fixed_format(
            path=str(out_dir),
            format_val=msword_cli.C.wdExportFormatXPS,
            pages=(2, 4),
            markup=True,
            properties=True,
            irm=True,
            bookmarks=msword_cli.C.wdExportCreateHeadingBookmarks,
            struct=False,
            bitmap=False,
            useiso19005_1=True,
        )

        expected = str((out_dir / "report.xps").resolve())

    assert exported == expected
    com_doc.ExportAsFixedFormat.assert_called_once_with(
        OutputFileName=expected,
        ExportFormat=msword_cli.C.wdExportFormatXPS,
        OpenAfterExport=False,
        OptimizeFor=msword_cli.C.wdExportOptimizeForPrint,
        Range=msword_cli.C.wdExportFromTo,
        From=2,
        To=4,
        Item=msword_cli.C.wdExportDocumentWithMarkup,
        IncludeDocProps=True,
        KeepIRM=False,
        CreateBookmarks=msword_cli.C.wdExportCreateHeadingBookmarks,
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=True,
    )


def test_document_close_force_uses_no_save_changes(msword_cli):
    com_doc = FakeComDocument("report.docx")
    document = msword_cli.Document(com_doc)

    document.close(force=True)

    com_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)


def test_word_client_open_passes_readonly_and_repair(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    app = FakeWordApp([com_doc])
    app.Documents.Open.return_value = com_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    result = client.open("report.docx", visible=False, read_only=True, repair=True)

    assert isinstance(result, msword_cli.Document)
    com_doc.Activate.assert_called_once_with()
    app.Documents.Open.assert_called_once_with(
        FileName=str(Path("report.docx").resolve()),
        Visible=False,
        ReadOnly=True,
        OpenAndRepair=True,
    )


def test_word_client_new_activates_hidden_document(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    app = FakeWordApp([com_doc])
    app.Documents.Add.return_value = com_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    result = client.new(visible=False)

    assert isinstance(result, msword_cli.Document)
    com_doc.Activate.assert_called_once_with()
    app.Documents.Add.assert_called_once_with(Visible=False)


def test_word_client_visible_false_does_not_hide_existing_word(msword_cli, monkeypatch):
    app = FakeWordApp(visible=True)
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    msword_cli.WordClient(visible=False)

    assert app.Visible is True


def test_word_client_context_manager_quits_on_exit(msword_cli, monkeypatch):
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))
    client = msword_cli.WordClient(quit_on_exit=True)

    with client:
        pass
    client.quit()

    assert client.closed is True
    app.Quit.assert_called_once_with()


def test_word_client_quit_keeps_client_after_failure_and_retries(msword_cli, monkeypatch):
    app = FakeWordApp()
    app.Quit.side_effect = [RuntimeError("busy"), None]
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))
    client = msword_cli.WordClient(visible=False)

    with pytest.raises(msword_cli.WordAPIError, match="Failed to quit Word"):
        client.quit()
    assert client.closed is False

    client.quit()
    assert client.closed is True
    assert app.Quit.call_count == 2


def test_word_client_context_preserves_body_error_when_quit_fails(msword_cli, monkeypatch):
    app = FakeWordApp()
    app.Quit.side_effect = RuntimeError("quit boom")
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    with pytest.raises(ValueError, match="body boom"):
        with msword_cli.WordClient(quit_on_exit=True):
            raise ValueError("body boom")


def test_word_client_closes_word_when_visible_setup_fails(msword_cli, monkeypatch):
    class BrokenVisibleApp(FakeWordApp):
        def __init__(self):
            self.Quit = Mock()

        @property
        def Visible(self):
            return False

        @Visible.setter
        def Visible(self, value):
            raise RuntimeError("visibility boom")

    app = BrokenVisibleApp()
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    with pytest.raises(msword_cli.WordAPIError, match="Unable to load"):
        msword_cli.WordClient(visible=True)
    app.Quit.assert_called_once_with()


def test_word_client_open_rolls_back_when_activation_fails(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.Activate.side_effect = msword_cli.com_error("activate boom")
    app = FakeWordApp([com_doc])
    app.Documents.Open.return_value = com_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))
    client = msword_cli.WordClient(visible=False)

    with pytest.raises(msword_cli.WordAPIError, match="Activation failed"):
        client.open("report.docx", visible=False)
    com_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)


def test_word_client_new_rolls_back_when_activation_fails(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.Activate.side_effect = msword_cli.com_error("activate boom")
    app = FakeWordApp([com_doc])
    app.Documents.Add.return_value = com_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))
    client = msword_cli.WordClient(visible=False)

    with pytest.raises(msword_cli.WordAPIError, match="Activation failed"):
        client.new(visible=False)
    com_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)


def test_word_client_native_exposes_word_application(msword_cli, monkeypatch):
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)

    assert client.native is app


def test_word_client_native_raises_after_quit(msword_cli, monkeypatch):
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    client.quit()

    with pytest.raises(msword_cli.WordAPIError, match="Word client is closed"):
        _ = client.native


def test_get_client_replaces_closed_client(msword_cli, monkeypatch):
    closed_client = Mock(closed=True)
    replacement = Mock(closed=False)
    word_client_ctor = Mock(return_value=replacement)
    monkeypatch.setattr(msword_cli, "_CLI_CLIENT", closed_client)
    monkeypatch.setattr(msword_cli, "WordClient", word_client_ctor)

    assert msword_cli.get_client() is replacement
    word_client_ctor.assert_called_once_with(visible=False)


def test_word_client_does_not_expose_compare_or_merge_in_core(msword_cli):
    assert not hasattr(msword_cli.WordClient, "compare")
    assert not hasattr(msword_cli.WordClient, "merge")


def test_word_client_load_plugins_injects_compare_merge_methods(msword_cli, monkeypatch):
    plugin_dir = Path(msword_cli.__file__).with_name("plugins") / "compare-merge"
    client = msword_cli.WordClient.__new__(msword_cli.WordClient)
    client._loaded_library_plugins = set()
    monkeypatch.setattr(msword_cli, "_installed_plugin_entry_points", lambda: [])

    msword_cli.WordClient.load_plugins(client, plugin_dir=str(plugin_dir), include="compare-merge")

    assert callable(getattr(client, "compare"))
    assert callable(getattr(client, "merge"))


def test_unknown_constant_does_not_start_word(msword_cli):
    msword_cli.com.gencache.EnsureDispatch.reset_mock()

    with pytest.raises(msword_cli.WordAPIError, match="Unknown Word constant"):
        msword_cli._resolve_constant("wdDoesNotExist")

    msword_cli.com.gencache.EnsureDispatch.assert_not_called()


def test_document_save_copy_normalizes_path(msword_cli):
    com_doc = FakeComDocument("report.docx")
    document = msword_cli.Document(com_doc)

    with CliRunner().isolated_filesystem():
        expected = str(Path("copy.docx").resolve())
        actual = document.save_copy("copy.docx")

    assert actual == expected
    com_doc.SaveCopyAs.assert_called_once_with(expected)


def test_word_client_track_changes_and_revisions(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.Revisions = msword_cli.Document(com_doc).native.Revisions if False else com_doc.Revisions
    app = FakeWordApp([com_doc])
    app.Selection.Range.Revisions = FakeRevisionsCollection(2)
    com_doc.Revisions = FakeRevisionsCollection(3)
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)

    assert client.set_track_changes(True) is True
    assert client.accept_revisions() == 3
    assert client.reject_revisions(scope="selection") == 2
    com_doc.Revisions.AcceptAll.assert_called_once_with()
    app.Selection.Range.Revisions.RejectAll.assert_called_once_with()


def test_word_client_comment_operations(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    first = FakeComment(author="Alice", text="One")
    second = FakeComment(author="Bob", text="Two")
    com_doc.Comments = FakeCommentsCollection([first, second])
    app = FakeWordApp([com_doc])
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    comments = client.list_comments()
    deleted = client.delete_comments()

    assert comments[0]["author"] == "Alice"
    assert comments[1]["text"] == "Two"
    assert deleted == 2
    first.Delete.assert_called_once_with()
    second.Delete.assert_called_once_with()


def test_word_client_export_comments_writes_json(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.Comments = FakeCommentsCollection([FakeComment(author="Alice", text="One")])
    app = FakeWordApp([com_doc])
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    with CliRunner().isolated_filesystem():
        output = Path("comments.json")
        exported = client.export_comments(str(output), output_format="json")
        content = output.read_text(encoding="utf-8")

    assert Path(exported).name == "comments.json"
    assert '"author": "Alice"' in content


def test_word_client_summary_normalizes_metadata(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.ReadOnly = True
    com_doc.TrackRevisions = True
    com_doc.Revisions = FakeRevisionsCollection(3)
    com_doc.Comments = FakeCommentsCollection([FakeComment(author="Alice", text="One")])
    app = FakeWordApp([com_doc])
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    summary = client.summary()

    assert summary == {
        "name": "report.docx",
        "path": str(Path("report.docx").resolve()),
        "saved": True,
        "read_only": True,
        "track_changes": True,
        "revisions": 3,
        "comments": 1,
        "template": "C:/Templates/normal.dotm",
        "author": "Alice",
        "last_author": "Bob",
        "created_at": "2026-06-01T09:30:00",
        "modified_at": "2026-06-29T18:45:00",
    }


def test_word_client_summary_returns_none_for_missing_metadata(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.BuiltInDocumentProperties = FakePropertyCollection([])
    com_doc.AttachedTemplate = None
    app = FakeWordApp([com_doc])
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    summary = client.summary()

    assert summary["author"] is None
    assert summary["last_author"] is None
    assert summary["created_at"] is None
    assert summary["modified_at"] is None
    assert summary["template"] is None


def test_word_client_properties_support_iterable_only_com_collections(msword_cli, monkeypatch):
    class IterableOnlyPropertyCollection:
        def __init__(self, items):
            self._items = list(items)

        def __iter__(self):
            return iter(self._items)

    def prop(name, value):
        return SimpleNamespace(Name=name, Value=value)

    com_doc = FakeComDocument("report.docx")
    com_doc.BuiltInDocumentProperties = IterableOnlyPropertyCollection(
        [
            prop("Author", "Alice"),
            prop("Last Author", "Bob"),
            prop("Creation Date", dt.datetime(2026, 6, 1, 9, 30, 0)),
            prop("Last Save Time", dt.datetime(2026, 6, 29, 18, 45, 0)),
        ]
    )
    com_doc.CustomDocumentProperties = IterableOnlyPropertyCollection([prop("Project", "CLI")])
    app = FakeWordApp([com_doc])
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)

    assert client.list_properties() == [
        {"kind": "built-in", "name": "Author", "value": "Alice"},
        {"kind": "built-in", "name": "Last Author", "value": "Bob"},
        {"kind": "built-in", "name": "Creation Date", "value": dt.datetime(2026, 6, 1, 9, 30, 0)},
        {"kind": "built-in", "name": "Last Save Time", "value": dt.datetime(2026, 6, 29, 18, 45, 0)},
        {"kind": "custom", "name": "Project", "value": "CLI"},
    ]

    summary = client.summary()

    assert summary["author"] == "Alice"
    assert summary["last_author"] == "Bob"
    assert summary["created_at"] == "2026-06-01T09:30:00"
    assert summary["modified_at"] == "2026-06-29T18:45:00"


def test_json_dump_serializes_datetimes(msword_cli):
    payload = {"created_at": msword_cli._normalize_summary_value("2026-06-01T09:30:00")}

    dumped = msword_cli._json_dump(payload)

    assert '"created_at": "2026-06-01T09:30:00"' in dumped


def test_word_client_update_fields_find_replace_properties_and_summary(msword_cli, monkeypatch):
    com_doc = FakeComDocument("report.docx")
    com_doc.Fields = FakeFieldsCollection(2)
    toc = Mock()
    toc.Update = Mock()
    com_doc.TablesOfContents = FakeCollection([toc])
    app = FakeWordApp([com_doc])
    app.Selection.Range.Fields = FakeFieldsCollection(1)
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    client.find = Mock(return_value=[{"start": 0, "end": 3, "text": "foo"}])
    app.Selection.Range.Find.Execute = Mock(return_value=True)

    document_result = client.update_fields()
    selection_result = client.update_fields(scope="selection")
    replaced = client.replace("foo", "bar", scope="selection")
    prop = client.set_property("Flag", True)
    summary = client.summary()
    stats = client.statistics()

    assert document_result == {"fields": 2, "tables_of_contents": 1}
    assert selection_result == {"fields": 1, "tables_of_contents": 0}
    assert replaced == 1
    assert prop["value"] is True
    assert summary["name"] == "report.docx"
    assert summary["author"] == "Alice"
    assert stats["pages"] == 2
