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
    client.open.assert_called_once_with(expected_path, visible=True, read_only=False, repair=False)


def test_open_hide(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        expected_path = str(Path("foo.docx").resolve())
        result = invoke(runner, msword_cli, ["open", "--hide", "foo.docx"])

    assert result.exit_code == 0
    client.open.assert_called_once_with(expected_path, visible=False, read_only=False, repair=False)


def test_open_hide_auto_closes_new_document(msword_cli, monkeypatch):
    runner = CliRunner()
    hidden_doc = make_document("foo.docx")
    client = FakeClient(document_count=0)
    client.open.return_value = hidden_doc
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        result = invoke(runner, msword_cli, ["open", "--hide", "foo.docx"])

    assert result.exit_code == 0
    hidden_doc.close.assert_called_once_with(force=True)
    client.quit.assert_called_once_with()
    assert 'Auto closing hidden document "foo.docx"' in result.output


def test_open_hide_does_not_auto_close_preexisting_document(msword_cli, monkeypatch):
    runner = CliRunner()
    existing_doc = make_document("foo.docx")
    reopened_doc = make_document("foo.docx")
    client = FakeClient([existing_doc], document_count=1)
    client.open.return_value = reopened_doc
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        result = invoke(runner, msword_cli, ["open", "--hide", "foo.docx"])

    assert result.exit_code == 0
    reopened_doc.close.assert_not_called()
    client.quit.assert_not_called()


def test_open_readonly_and_repair(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("foo.docx").touch()
        expected_path = str(Path("foo.docx").resolve())
        result = invoke(runner, msword_cli, ["open", "--hide", "--readonly", "--repair", "foo.docx"])

    assert result.exit_code == 0
    client.open.assert_called_once_with(expected_path, visible=False, read_only=True, repair=True)


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
                "--with-properties",
                "--without-irm",
                "--with-word-bookmarks",
                "--without-structure-tags",
                "--without-bitmap-fonts",
                "--pdf-a",
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
    assert 'Closing document "foo.docx"' in result.output


def test_force_close(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc], document_count=1)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close", "--force"])

    assert result.exit_code == 0
    doc.close.assert_called_once_with(force=True)
    assert 'Force closing document "foo.docx"' in result.output


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
    assert 'Force closing document "foo.docx"' in result.output
    assert 'Force closing document "bar.docx"' in result.output


def test_close_quits_when_last_document_is_closed(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    client = FakeClient([doc], document_count=0)
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["close"])

    assert result.exit_code == 0
    client.quit.assert_called_once_with()




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




def test_help_groups_commands_and_does_not_initialize_client(msword_cli, monkeypatch):
    runner = CliRunner()
    get_client = Mock(side_effect=AssertionError("client should not be initialized"))
    monkeypatch.setattr(msword_cli, "get_client", get_client)

    result = invoke(runner, msword_cli, ["--help"])

    assert result.exit_code == 0
    assert "Documents:" in result.output
    assert "Review:" in result.output
    assert "list-documents" in result.output
    assert "track-changes" in result.output
    get_client.assert_not_called()

def test_save_copy_command(msword_cli, monkeypatch):
    runner = CliRunner()
    doc = make_document("foo.docx")
    doc.save_copy.return_value = str(Path("copy.docx").resolve())
    client = FakeClient([doc])
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        expected_path = str(Path("copy.docx").resolve())
        result = invoke(runner, msword_cli, ["save-copy", "copy.docx"])

    assert result.exit_code == 0
    doc.save_copy.assert_called_once_with(expected_path)


def test_list_documents_command(msword_cli, monkeypatch):
    runner = CliRunner()
    first = make_document("foo.docx", saved=True)
    second = make_document("bar.docx", saved=False)
    client = FakeClient([first, second], document_count=2)
    client.active_document = second
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["list-documents"])

    assert result.exit_code == 0
    assert "foo.docx" in result.output
    assert "bar.docx*" in result.output


def test_find_command_json(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.find.return_value = [{"start": 1, "end": 4, "text": "foo"}]
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["find", "--scope", "selection", "--match-case", "--format", "json", "foo"])

    assert result.exit_code == 0
    client.find.assert_called_once_with("foo", scope="selection", match_case=True, whole_word=False)
    assert '"text": "foo"' in result.output


def test_replace_command_text(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.replace.return_value = 3
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["replace", "--whole-word", "foo", "bar"])

    assert result.exit_code == 0
    client.replace.assert_called_once_with("foo", "bar", scope="document", match_case=False, whole_word=True)
    assert "Replaced 3 occurrence(s)." in result.output


def test_track_changes_command(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.set_track_changes.return_value = False
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["track-changes", "--off", "--format", "json"])

    assert result.exit_code == 0
    client.set_track_changes.assert_called_once_with(False)
    assert '"track_changes": false' in result.output


def test_comments_commands(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.list_comments.return_value = [{"index": 1, "author": "Tester", "text": "Note"}]
    client.export_comments.return_value = str(Path("comments.json").resolve())
    client.delete_comments.return_value = 1
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    with runner.isolated_filesystem():
        exported_path = str(Path("comments.json").resolve())
        listed = invoke(runner, msword_cli, ["list-comments"])
        exported = invoke(runner, msword_cli, ["export-comments", "--json", "comments.json"])
        deleted = invoke(runner, msword_cli, ["delete-comments", "--format", "json"])

    assert listed.exit_code == 0
    assert "Tester: Note" in listed.output
    assert exported.exit_code == 0
    client.export_comments.assert_called_once_with(exported_path, scope="document", output_format="json")
    assert deleted.exit_code == 0
    assert '"deleted": 1' in deleted.output


def test_summary_command_text(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.summary.return_value = {
        "name": "foo.docx",
        "path": "C:/docs/foo.docx",
        "saved": True,
        "read_only": False,
        "track_changes": True,
        "revisions": 2,
        "comments": 1,
        "template": "C:/Templates/normal.dotm",
        "author": "Alice",
        "last_author": "Bob",
        "created_at": "2026-06-01T09:30:00",
        "modified_at": "2026-06-29T18:45:00",
    }
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["summary"])

    assert result.exit_code == 0
    client.summary.assert_called_once_with()
    assert "name: foo.docx" in result.output
    assert "author: Alice" in result.output
    assert "created_at: 2026-06-01T09:30:00" in result.output


def test_summary_command_json(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.summary.return_value = {"name": "foo.docx", "created_at": None}
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    result = invoke(runner, msword_cli, ["summary", "--format", "json"])

    assert result.exit_code == 0
    client.summary.assert_called_once_with()
    assert '"name": "foo.docx"' in result.output
    assert '"created_at": null' in result.output


def test_help_shows_summary_and_hides_info(msword_cli, monkeypatch):
    runner = CliRunner()
    get_client = Mock(side_effect=AssertionError("client should not be initialized"))
    monkeypatch.setattr(msword_cli, "get_client", get_client)

    result = invoke(runner, msword_cli, ["--help"])

    assert result.exit_code == 0
    assert "summary" in result.output
    assert "info" not in result.output
    get_client.assert_not_called()


def test_update_fields_and_property_commands(msword_cli, monkeypatch):
    runner = CliRunner()
    client = FakeClient()
    client.update_fields.return_value = {"fields": 2, "tables_of_contents": 1}
    client.list_properties.return_value = [{"kind": "custom", "name": "Project", "value": "CLI"}]
    client.get_property.return_value = {"kind": "custom", "name": "Project", "value": "CLI"}
    client.set_property.return_value = {"kind": "custom", "name": "Project", "value": True}
    monkeypatch.setattr(msword_cli, "get_client", lambda: client)

    updated = invoke(runner, msword_cli, ["update-fields"])
    listed = invoke(runner, msword_cli, ["list-properties"])
    got = invoke(runner, msword_cli, ["get-property", "Project"])
    set_result = invoke(runner, msword_cli, ["set-property", "--format", "json", "Project", "true"])

    assert updated.exit_code == 0
    assert "2 field(s)" in updated.output
    assert listed.exit_code == 0
    assert "custom: Project = CLI" in listed.output
    assert got.exit_code == 0
    assert "Project = CLI" in got.output
    assert set_result.exit_code == 0
    client.set_property.assert_called_once_with("Project", True)
    assert '"value": true' in set_result.output


def test_summary_help_mentions_document_state(msword_cli):
    runner = CliRunner()

    result = invoke(runner, msword_cli, ["summary", "--help"])

    assert result.exit_code == 0
    assert "state" in result.output.lower()
    assert "properties" not in result.output.lower()

