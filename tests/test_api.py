from pathlib import Path
from unittest.mock import Mock

import click
import pytest
from click.testing import CliRunner

from tests.conftest import FakeComDocument, FakeWordApp


def test_merge_command_passes_new_options(msword_cli, monkeypatch):
    runner = CliRunner()
    client = Mock()
    result_doc = Mock()
    result_doc.name = "merged.docx"
    client.merge.return_value = result_doc
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("original.docx").touch()
        Path("revised.docx").touch()
        original = str(Path("original.docx").resolve())
        revised = str(Path("revised.docx").resolve())
        result = runner.invoke(
            msword_cli.cli,
            [
                "merge",
                "--to-original",
                "--char-level",
                "--no-comments",
                "--ignore-warnings",
                "original.docx",
                "revised.docx",
            ],
        )

    assert result.exit_code == 0
    client.merge.assert_called_once_with(
        original=original,
        revised=revised,
        destination=msword_cli.C.wdMergeDestinationOriginalDocument,
        granularity=msword_cli.C.wdGranularityCharLevel,
        formatting=True,
        case_changes=True,
        whitespace=True,
        tables=True,
        headers=True,
        footnotes=True,
        textboxes=True,
        fields=True,
        comments=False,
        moves=True,
        author=None,
        ignore_warnings=True,
    )


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


def test_word_client_compare_uses_word_username_by_default(msword_cli, monkeypatch):
    original_doc = Mock()
    revised_doc = Mock()
    result_doc = Mock()
    app = FakeWordApp()
    app.Documents.Open = Mock(side_effect=[original_doc, revised_doc])
    app.CompareDocuments.return_value = result_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    with CliRunner().isolated_filesystem():
        Path("a.docx").touch()
        Path("b.docx").touch()
        original = str(Path("a.docx").resolve())
        revised = str(Path("b.docx").resolve())
        client = msword_cli.WordClient(visible=False)
        result = client.compare(original, revised)

    assert isinstance(result, msword_cli.Document)
    app.Documents.Open.assert_any_call(FileName=original, Visible=False)
    app.Documents.Open.assert_any_call(FileName=revised, Visible=False)
    app.CompareDocuments.assert_called_once_with(
        OriginalDocument=original_doc,
        RevisedDocument=revised_doc,
        Destination=msword_cli.C.wdCompareDestinationNew,
        Granularity=msword_cli.C.wdGranularityWordLevel,
        CompareFormatting=True,
        CompareCaseChanges=True,
        CompareWhitespace=True,
        CompareTables=True,
        CompareHeaders=True,
        CompareFootnotes=True,
        CompareTextboxes=True,
        CompareFields=True,
        CompareComments=True,
        CompareMoves=True,
        RevisedAuthor=app.UserName,
        IgnoreAllComparisonWarnings=False,
    )
    original_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)
    revised_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)


def test_word_client_compare_accepts_explicit_author_and_flags(msword_cli, monkeypatch):
    original_doc = Mock()
    revised_doc = Mock()
    result_doc = Mock()
    app = FakeWordApp()
    app.Documents.Open = Mock(side_effect=[original_doc, revised_doc])
    app.CompareDocuments.return_value = result_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    with CliRunner().isolated_filesystem():
        Path("a.docx").touch()
        Path("b.docx").touch()
        original = str(Path("a.docx").resolve())
        revised = str(Path("b.docx").resolve())
        client = msword_cli.WordClient(visible=False)
        client.compare(
            original,
            revised,
            destination=msword_cli.C.wdCompareDestinationRevised,
            granularity=msword_cli.C.wdGranularityCharLevel,
            formatting=False,
            comments=False,
            moves=False,
            author="Tester",
            ignore_warnings=True,
        )

    app.CompareDocuments.assert_called_once_with(
        OriginalDocument=original_doc,
        RevisedDocument=revised_doc,
        Destination=msword_cli.C.wdCompareDestinationRevised,
        Granularity=msword_cli.C.wdGranularityCharLevel,
        CompareFormatting=False,
        CompareCaseChanges=True,
        CompareWhitespace=True,
        CompareTables=True,
        CompareHeaders=True,
        CompareFootnotes=True,
        CompareTextboxes=True,
        CompareFields=True,
        CompareComments=False,
        CompareMoves=False,
        RevisedAuthor="Tester",
        IgnoreAllComparisonWarnings=True,
    )
    original_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)
    revised_doc.Close.assert_not_called()


def test_word_client_context_manager_quits_on_exit(msword_cli, monkeypatch):
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))
    client = msword_cli.WordClient(quit_on_exit=True)

    with client:
        pass
    client.quit()

    assert client.closed is True
    app.Quit.assert_called_once_with()


def test_get_client_replaces_closed_client(msword_cli, monkeypatch):
    closed_client = Mock(closed=True)
    replacement = Mock(closed=False)
    word_client_ctor = Mock(return_value=replacement)
    monkeypatch.setattr(msword_cli, "_CLI_CLIENT", closed_client)
    monkeypatch.setattr(msword_cli, "WordClient", word_client_ctor)

    assert msword_cli.get_client() is replacement
    word_client_ctor.assert_called_once_with(visible=False)


def test_unknown_constant_does_not_start_word(msword_cli):
    msword_cli.com.gencache.EnsureDispatch.reset_mock()

    with pytest.raises(msword_cli.WordAPIError, match="Unknown Word constant"):
        msword_cli._resolve_constant("wdDoesNotExist")

    msword_cli.com.gencache.EnsureDispatch.assert_not_called()


def test_word_client_merge_keeps_original_destination_open(msword_cli, monkeypatch):
    original_doc = Mock()
    revised_doc = Mock()
    app = FakeWordApp()
    app.Documents.Open = Mock(side_effect=[original_doc, revised_doc])
    app.MergeDocuments.return_value = original_doc
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    result = client.merge(
        "original.docx",
        "revised.docx",
        destination=msword_cli.C.wdMergeDestinationOriginalDocument,
    )

    assert result._doc is original_doc
    original_doc.Close.assert_not_called()
    revised_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)
