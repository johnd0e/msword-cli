import importlib.util
import sys
import uuid
from pathlib import Path
from unittest.mock import Mock

from click.testing import CliRunner
from tests.conftest import FakeWordApp


def _load_compare_merge_plugin(msword_cli):
    plugin_path = Path(msword_cli.__file__).with_name("plugins") / "compare-merge" / "compare_merge.py"
    module_name = f"test_compare_merge_plugin_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(module_name, plugin_path)
    module = importlib.util.module_from_spec(spec)
    assert spec is not None
    assert spec.loader is not None
    sys.modules.pop(module_name, None)
    spec.loader.exec_module(module)
    return module


def test_compare_plugin_cli_passes_options(msword_cli, monkeypatch):
    plugin = _load_compare_merge_plugin(msword_cli)
    runner = CliRunner()
    client = object()
    result_doc = Mock()
    result_doc.name = "diff.docx"
    compare_call = Mock(return_value=result_doc)
    monkeypatch.setattr(plugin, "compare", compare_call)
    monkeypatch.setattr(plugin, "get_client", lambda: client)

    with runner.isolated_filesystem():
        Path("original.docx").touch()
        Path("revised.docx").touch()
        original = str(Path("original.docx").resolve())
        revised = str(Path("revised.docx").resolve())
        result = runner.invoke(
            plugin.compare_cmd,
            [
                "--to-revised",
                "--char-level",
                "--no-formatting",
                "--author",
                "Tester",
                "original.docx",
                "revised.docx",
            ],
        )

    assert result.exit_code == 0
    compare_call.assert_called_once_with(
        client,
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


def test_compare_plugin_method_uses_word_username_by_default(msword_cli, monkeypatch):
    original_doc = Mock()
    revised_doc = Mock()
    result_doc = Mock()
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli, "_installed_plugin_entry_points", lambda: [])
    app.Documents.Open = Mock(side_effect=[original_doc, revised_doc])
    app.CompareDocuments = Mock(return_value=result_doc)
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    with CliRunner().isolated_filesystem():
        Path("a.docx").touch()
        Path("b.docx").touch()
        original = str(Path("a.docx").resolve())
        revised = str(Path("b.docx").resolve())
        client = msword_cli.WordClient(visible=False)
        client.load_plugins(plugin_dir=str(Path(msword_cli.__file__).with_name("plugins") / "compare-merge"), include="compare-merge")
        result = client.compare(original, revised)

    assert isinstance(result, msword_cli.Document)
    app.CompareDocuments.assert_called_once()
    original_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)
    revised_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)


def test_merge_plugin_keeps_original_destination_open(msword_cli, monkeypatch):
    original_doc = Mock()
    revised_doc = Mock()
    app = FakeWordApp()
    monkeypatch.setattr(msword_cli, "_installed_plugin_entry_points", lambda: [])
    app.Documents.Open = Mock(side_effect=[original_doc, revised_doc])
    app.MergeDocuments = Mock(return_value=original_doc)
    monkeypatch.setattr(msword_cli.com.gencache, "EnsureDispatch", Mock(return_value=app))

    client = msword_cli.WordClient(visible=False)
    client.load_plugins(plugin_dir=str(Path(msword_cli.__file__).with_name("plugins") / "compare-merge"), include="compare-merge")
    result = client.merge(
        "original.docx",
        "revised.docx",
        destination=msword_cli.C.wdMergeDestinationOriginalDocument,
    )

    assert result.native is original_doc
    original_doc.Close.assert_not_called()
    revised_doc.Close.assert_called_once_with(msword_cli.C.wdDoNotSaveChanges)
