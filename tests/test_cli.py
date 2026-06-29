from pathlib import Path
from unittest.mock import Mock

from click.testing import CliRunner

from tests.conftest import FakeClient, make_document


def invoke(runner, msword_cli, args):
    return runner.invoke(msword_cli.cli, args)


def test_open_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        expected_path = str(Path("foo.docx").resolve())
        result = invoke(runner, msword_cli, ["open", "foo.docx"])

    assert result.exit_code == 0
    client.open.assert_called_once_with(expected_path, visible=True)


def test_open_hide(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        expected_path = str(Path("foo.docx").resolve())
        result = invoke(runner, msword_cli, ["open", "--hide", "foo.docx"])

    assert result.exit_code == 0
    client.open.assert_called_once_with(expected_path, visible=False)


def test_new_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["new"])

    assert result.exit_code == 0
    client.new.assert_called_once_with(template=None, visible=True)


def test_new_hide(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["new", "--hide"])

    assert result.exit_code == 0
    client.new.assert_called_once_with(template=None, visible=False)


def test_new_template_from_cwd(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient(template_dir="C:/Templates")
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.dot").touch()
        expected_path = str(Path("foo.dot").resolve())
        result = invoke(runner, msword_cli, ["new", "--template", "foo.dot"])

    assert result.exit_code == 0
    client.new.assert_called_once_with(template=expected_path, visible=True)


def test_new_template_from_default_template_dir(msword_cli, monkeypatch):
    runner = CliRunner()
    with runner.isolated_filesystem():
        template_dir = Path.cwd() / "templates"
        template_dir.mkdir()
        template_path = template_dir / "normal.dotm"
        template_path.touch()
        client = FakeClient(template_dir=str(template_dir))
        monkeypatch.setattr(msword_cli, "get_client", lambda: client)

        result = invoke(runner, msword_cli, ["new", "--template", "normal.dotm"])

    assert result.exit_code == 0
    client.new.assert_called_once_with(template=str(template_path), visible=True)


def test_print_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["print"])

    assert result.exit_code == 0
    doc.print_out.assert_called_once_with(
        copies=1,
        pages=None,
        pagetype=msword_cli.C.wdPrintAllPages,
        rng=msword_cli.C.wdPrintAllDocument,
        item=msword_cli.C.wdPrintDocumentContent,
        no_collate=False,
        to_file=None,
        append=False,
        columns=1,
        rows=1,
    )


def test_print_with_options(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(
        runner,
        msword_cli,
        ["print", "--copies", "3", "--pages", "2-3, 6", "--odd", "--no-collate", "--columns", "2", "--rows", "4"],
    )

    assert result.exit_code == 0
    doc.print_out.assert_called_once_with(
        copies=3,
        pages="2-3, 6",
        pagetype=msword_cli.C.wdPrintOddPagesOnly,
        rng=msword_cli.C.wdPrintAllDocument,
        item=msword_cli.C.wdPrintDocumentContent,
        no_collate=True,
        to_file=None,
        append=False,
        columns=2,
        rows=4,
    )


def test_print_to_file(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        expected_path = str(Path("out.prn").resolve())
        result = invoke(runner, msword_cli, ["print", "--to-file", "out.prn", "--append"])

    assert result.exit_code == 0
    doc.print_out.assert_called_once_with(
        copies=1,
        pages=None,
        pagetype=msword_cli.C.wdPrintAllPages,
        rng=msword_cli.C.wdPrintAllDocument,
        item=msword_cli.C.wdPrintDocumentContent,
        no_collate=False,
        to_file=expected_path,
        append=True,
        columns=1,
        rows=1,
    )


def test_export_requires_path(msword_cli):
    runner = CliRunner()

    result = invoke(runner, msword_cli, ["export"])

    assert result.exit_code == 2


def test_export_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    doc.export_fixed_format.return_value = str(Path("foo.pdf").resolve())
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        expected_path = str(Path("foo.pdf").resolve())
        result = invoke(runner, msword_cli, ["export", "foo.pdf"])

    assert result.exit_code == 0
    doc.export_fixed_format.assert_called_once_with(
        path=expected_path,
        format_val=msword_cli.C.wdExportFormatPDF,
        show=False,
        optimize=msword_cli.C.wdExportOptimizeForPrint,
        pages=None,
        rng=None,
        markup=False,
        properties=False,
        irm=False,
        bookmarks=msword_cli.C.wdExportCreateNoBookmarks,
        struct=False,
        bitmap=False,
        useiso19005_1=False,
    )


def test_export_with_options(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    doc.export_fixed_format.return_value = str(Path("foo.xps").resolve())
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        expected_path = str(Path("foo").resolve())
        result = invoke(
            runner,
            msword_cli,
            [
                "export",
                "--xps",
                "--show",
                "--for-screen",
                "--pages",
                "2-3",
                "--with-markup",
                "--with-props",
                "--without-irm",
                "--with-word-bookmarks",
                "--without-structure-tags",
                "--without-bitmaped-fonts",
                "--useiso19005-1",
                "foo",
            ],
        )

    assert result.exit_code == 0
    doc.export_fixed_format.assert_called_once_with(
        path=expected_path,
        format_val=msword_cli.C.wdExportFormatXPS,
        show=True,
        optimize=msword_cli.C.wdExportOptimizeForOnScreen,
        pages=(2, 3),
        rng=None,
        markup=True,
        properties=True,
        irm=True,
        bookmarks=msword_cli.C.wdExportCreateWordBookmarks,
        struct=True,
        bitmap=True,
        useiso19005_1=True,
    )


def test_export_bad_range(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["export", "--pages", "4-3", "foo.pdf"])

    assert result.exit_code == 2
    doc.export_fixed_format.assert_not_called()


def test_save_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["save"])

    assert result.exit_code == 0
    doc.save.assert_called_once_with(force=False)


def test_save_with_path(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        expected_path = str(Path("bar.docx").resolve())
        result = invoke(runner, msword_cli, ["save", "--path", "bar.docx"])

    assert result.exit_code == 0
    doc.save.assert_called_once_with(path=expected_path)


def test_force_save(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["save", "--force"])

    assert result.exit_code == 0
    doc.save.assert_called_once_with(force=True)


def test_save_all(msword_cli, monkeypatch):
    runner = CliRunner()
    first = make_document("foo.docx")
    second = make_document("bar.docx")
    client = FakeClient([first, second])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["save", "--all", "--force"])

    assert result.exit_code == 0
    first.save.assert_called_once_with(force=True)
    second.save.assert_called_once_with(force=True)


def test_close_defaults(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc], document_count=1)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close"])

    assert result.exit_code == 0
    doc.close.assert_called_once_with(force=False)
    client.quit.assert_not_called()


def test_force_close(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc], document_count=1)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close", "--force"])

    assert result.exit_code == 0
    doc.close.assert_called_once_with(force=True)


def test_close_all(msword_cli, monkeypatch):
    runner = CliRunner()
    first = make_document("foo.docx")
    second = make_document("bar.docx")
    client = FakeClient([first, second], document_count=2)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close", "--all", "--force"])

    assert result.exit_code == 0
    first.close.assert_called_once_with(force=True)
    second.close.assert_called_once_with(force=True)


def test_close_quits_when_last_document_is_closed(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc], document_count=0)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close"])

    assert result.exit_code == 0
    client.quit.assert_called_once_with()


def test_docs_with_no_open_documents(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient([], document_count=0)
    client.active_document = None
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["docs"])

    assert result.exit_code == 0
    assert result.output == "\nNo open documents found.\n"


def test_docs_lists_and_marks_active_and_unsaved(msword_cli, monkeypatch):
    runner = CliRunner()
    first = make_document("foo.docx", saved=True)
    second = make_document("bar.docx", saved=False)
    client = FakeClient([first, second], document_count=2)
    client.active_document = second
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["docs"])

    assert result.exit_code == 0
    assert "  [1] foo.docx" in result.output
    assert " * [2] bar.docx*" in result.output


def test_activate_document_by_index(msword_cli, monkeypatch):
    runner = CliRunner()
    first = make_document("foo.docx")
    second = make_document("bar.docx")
    client = FakeClient([first, second], document_count=2)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["activate", "2"])

    assert result.exit_code == 0
    second.activate.assert_called_once_with()


def test_activate_bad_index(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient([make_document("foo.docx")], document_count=1)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["activate", "6"])

    assert result.exit_code != 0
    assert "Index 6 out of range." in result.output


def test_compare_command_passes_new_options(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    result_doc = Mock()
    result_doc.name = "diff.docx"
    client.compare.return_value = result_doc
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("original.docx").touch()
        Path("revised.docx").touch()
        original = str(Path("original.docx").resolve())
        revised = str(Path("revised.docx").resolve())
        result = invoke(
            runner,
            msword_cli,
            ["compare", "--to-revised", "--char-level", "--no-formatting", "--author", "Tester", "original.docx", "revised.docx"],
        )

    assert result.exit_code == 0
    client.compare.assert_called_once_with(
        original=original,
        revised=revised,
        destination=msword_cli.C.wdCompareDestinationRevised,
        granularity=msword_cli.C.wdGranularityCharLevel,
        formatting=False,
        case_changes=True,
        whitespace=True,
        tables=True,
        headers=True,
        footnotes=True,
        textboxes=True,
        fields=True,
        comments=True,
        moves=True,
        author="Tester",
        ignore_warnings=False,
    )


def test_print_rejects_non_positive_copies(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["print", "--copies", "0"])

    assert result.exit_code == 2
    doc.print_out.assert_not_called()


