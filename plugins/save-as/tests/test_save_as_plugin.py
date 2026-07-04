import importlib
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import Mock

import click
import pytest
from click.testing import CliRunner


@pytest.fixture
def runtime(monkeypatch):
    pywintypes = ModuleType("pywintypes")

    class FakeComError(Exception):
        def __init__(self, *args, excepinfo=None):
            super().__init__(*args)
            self.excepinfo = excepinfo

    pywintypes.com_error = FakeComError
    constants = SimpleNamespace(
        wdFormatRTF=51,
        wdFormatHTML=56,
        wdFormatFilteredHTML=57,
        wdFormatPDF=60,
        wdFormatText=61,
        wdFormatTextLineBreaks=62,
        wdFormatDOSText=63,
        wdFormatDOSTextLineBreaks=64,
        wdFormatEncodedText=65,
        wdFormatUnicodeText=66,
        wdFormatXPS=67,
        wdFormatXML=68,
        wdFormatXMLDocument=69,
        wdFormatXMLDocumentMacroEnabled=70,
        wdFormatXMLTemplate=71,
        wdFormatXMLTemplateMacroEnabled=72,
        wdFormatFlatXML=73,
        wdFormatFlatXMLMacroEnabled=74,
        wdFormatFlatXMLTemplate=75,
        wdFormatFlatXMLTemplateMacroEnabled=76,
        wdFormatStrictOpenXMLDocument=77,
    )
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.constants = constants
    client.gencache = SimpleNamespace(EnsureDispatch=Mock())
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pywintypes", pywintypes)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.delitem(sys.modules, "msword_cli", raising=False)
    core = importlib.import_module("msword_cli")
    return SimpleNamespace(core=core, constants=constants)


@pytest.fixture
def plugin(runtime, monkeypatch):
    plugin_root = Path(__file__).resolve().parents[1]
    monkeypatch.syspath_prepend(str(plugin_root))
    monkeypatch.delitem(sys.modules, "save_as", raising=False)
    return importlib.import_module("save_as")


def test_format_internals_are_private(plugin):
    private_names = {
        "_SAVE_AS_FORMAT_ALIASES",
        "_XML_FORMAT_CONSTANTS",
        "_TEXT_FORMAT_CONSTANTS",
        "_FIXED_FORMAT_CONSTANTS",
        "_COMPOSITIONAL_FORMATS",
        "_resolve_xml_format",
        "_resolve_text_format",
        "_resolve_html_format",
        "_is_detected_format_constant",
        "_iter_collection",
        "_normalize_extensions",
        "_list_save_as_formats",
        "_render_save_as_formats",
        "_com_error_message",
        "_save_as2",
    }

    assert private_names <= set(vars(plugin))
    assert not {name[1:] for name in private_names} & set(vars(plugin))


@pytest.mark.parametrize(
    ("options", "expected"),
    [
        ({}, "wdFormatXML"),
        ({"document": True}, "wdFormatXMLDocument"),
        ({"document": True, "macro": True}, "wdFormatXMLDocumentMacroEnabled"),
        ({"template": True}, "wdFormatXMLTemplate"),
        ({"template": True, "macro": True}, "wdFormatXMLTemplateMacroEnabled"),
        ({"flat": True}, "wdFormatFlatXML"),
        ({"flat": True, "macro": True}, "wdFormatFlatXMLMacroEnabled"),
        ({"flat": True, "template": True}, "wdFormatFlatXMLTemplate"),
        (
            {"flat": True, "template": True, "macro": True},
            "wdFormatFlatXMLTemplateMacroEnabled",
        ),
        ({"strict": True}, "wdFormatStrictOpenXMLDocument"),
    ],
)
def test_resolve_xml_format(plugin, options, expected):
    arguments = {"document": False, "template": False, "flat": False, "macro": False, "strict": False}
    arguments.update(options)

    assert plugin._resolve_xml_format(**arguments) == expected


@pytest.mark.parametrize(
    "options",
    [
        {"document": True, "template": True},
        {"flat": True, "document": True},
        {"strict": True, "document": True},
        {"strict": True, "template": True},
        {"strict": True, "flat": True},
        {"strict": True, "macro": True},
        {"macro": True},
    ],
)
def test_resolve_xml_format_rejects_conflicts(plugin, options):
    arguments = {"document": False, "template": False, "flat": False, "macro": False, "strict": False}
    arguments.update(options)

    with pytest.raises(click.UsageError):
        plugin._resolve_xml_format(**arguments)


@pytest.mark.parametrize(
    ("options", "expected"),
    [
        ({}, "wdFormatText"),
        ({"line_breaks": True}, "wdFormatTextLineBreaks"),
        ({"dos": True}, "wdFormatDOSText"),
        ({"dos": True, "line_breaks": True}, "wdFormatDOSTextLineBreaks"),
        ({"encoded": True}, "wdFormatEncodedText"),
        ({"unicode": True}, "wdFormatUnicodeText"),
    ],
)
def test_resolve_text_format(plugin, options, expected):
    arguments = {"line_breaks": False, "dos": False, "encoded": False, "unicode": False}
    arguments.update(options)

    assert plugin._resolve_text_format(**arguments) == expected


@pytest.mark.parametrize(
    "options",
    [
        {"encoded": True, "dos": True},
        {"encoded": True, "line_breaks": True},
        {"encoded": True, "unicode": True},
        {"unicode": True, "dos": True},
        {"unicode": True, "line_breaks": True},
    ],
)
def test_resolve_text_format_rejects_conflicts(plugin, options):
    arguments = {"line_breaks": False, "dos": False, "encoded": False, "unicode": False}
    arguments.update(options)

    with pytest.raises(click.UsageError):
        plugin._resolve_text_format(**arguments)


@pytest.mark.parametrize(
    ("filtered", "expected"),
    [
        (False, "wdFormatHTML"),
        (True, "wdFormatFilteredHTML"),
    ],
)
def test_resolve_html_format(plugin, filtered, expected):
    assert plugin._resolve_html_format(filtered) == expected


def test_detected_format_accepts_only_additional_system_constants(plugin):
    plugin.C.wdFormatCustomSystemFormat = 901

    assert plugin._is_detected_format_constant("wdFormatCustomSystemFormat") is True
    assert plugin._is_detected_format_constant("wdFormatRTF") is False
    assert plugin._is_detected_format_constant("wdFormatXML") is False
    assert plugin._is_detected_format_constant("wdFormatText") is False
    assert plugin._is_detected_format_constant("wdFormatPDF") is False
    assert plugin._is_detected_format_constant("wdFormatXPS") is False
    assert plugin._is_detected_format_constant("wdFormatMissing") is False
    assert plugin._is_detected_format_constant("customSystemFormat") is False
    assert plugin._is_detected_format_constant(None) is False


def test_save_as2_normalizes_path_calls_com_and_returns_path(plugin, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    native = SimpleNamespace(SaveAs2=Mock())
    document = SimpleNamespace(native=native)

    result = plugin._save_as2(document, "output.rtf", 51)

    expected = str((tmp_path / "output.rtf").resolve())
    native.SaveAs2.assert_called_once_with(FileName=expected, FileFormat=51)
    assert result == expected


def test_save_as2_wraps_com_error_and_preserves_cause(plugin, runtime, tmp_path):
    error = plugin.com_error("failed", excepinfo=(None, None, "Word rejected the format"))
    native = SimpleNamespace(SaveAs2=Mock(side_effect=error))
    document = SimpleNamespace(native=native)

    with pytest.raises(runtime.core.WordAPIError, match="Save as failed: Word rejected the format") as caught:
        plugin._save_as2(document, tmp_path / "output.rtf", 51)

    assert caught.value.__cause__ is error


@pytest.fixture
def command_runtime(plugin, monkeypatch):
    document = SimpleNamespace(save=Mock(), native=SimpleNamespace(SaveAs2=Mock()))
    client = SimpleNamespace(active_document=document)
    get_client = Mock(return_value=client)
    monkeypatch.setattr(plugin, "get_client", get_client)
    return SimpleNamespace(document=document, get_client=get_client, runner=CliRunner())


def test_save_as_command_is_exported(plugin):
    assert isinstance(plugin.save_as_cmd, click.Command)
    assert plugin.save_as_cmd.name == "save-as"


def test_plugin_manifest_exposes_cli_and_library_entries(plugin):
    assert plugin.plugin_manifest["name"] == "save-as"
    assert plugin.plugin_manifest["command"] is plugin.save_as_cmd
    assert plugin.plugin_manifest["client_methods"]["save_as"] is plugin.save_as


def test_save_as_requires_path_for_save_operation(plugin, command_runtime):
    result = command_runtime.runner.invoke(plugin.save_as_cmd, [])

    assert result.exit_code == 2
    assert "Missing argument PATH" in result.output
    command_runtime.get_client.assert_not_called()


def test_save_as_list_formats_requires_no_path_and_no_active_document(plugin, command_runtime):
    result = command_runtime.runner.invoke(plugin.save_as_cmd, ["--list-formats"])

    assert result.exit_code == 0, result.output
    assert "Built-in formats:" in result.output
    command_runtime.get_client.assert_called_once_with()
    assert not command_runtime.document.save.called
    assert not command_runtime.document.native.SaveAs2.called


def test_save_as_list_formats_rejects_path(plugin, command_runtime, tmp_path):
    result = command_runtime.runner.invoke(plugin.save_as_cmd, ["--list-formats", str(tmp_path / "out.docx")])

    assert result.exit_code == 2
    assert "Path is not allowed with --list-formats." in result.output
    command_runtime.get_client.assert_not_called()


@pytest.mark.parametrize(
    "arguments",
    [
        ["--list-formats", "--format", "docx"],
        ["--list-formats", "--document"],
        ["--list-formats", "--template"],
        ["--list-formats", "--flat"],
        ["--list-formats", "--strict"],
        ["--list-formats", "--macro"],
        ["--list-formats", "--line-breaks"],
        ["--list-formats", "--dos"],
        ["--list-formats", "--encoded"],
        ["--list-formats", "--unicode"],
    ],
)
def test_save_as_list_formats_rejects_format_selection_options(plugin, command_runtime, arguments):
    result = command_runtime.runner.invoke(plugin.save_as_cmd, arguments)

    assert result.exit_code == 2
    assert "--list-formats cannot be combined with format selection options." in result.output
    command_runtime.get_client.assert_not_called()


def test_plain_save_as_uses_active_document_save(plugin, command_runtime, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    result = command_runtime.runner.invoke(plugin.save_as_cmd, ["output.docx"])

    expected = str((tmp_path / "output.docx").resolve())
    assert result.exit_code == 0, result.output
    command_runtime.document.save.assert_called_once_with(path=expected)
    assert result.output == f'Saving active document as "{expected}"\n'


def test_library_save_as_uses_active_document_save(plugin, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    document = SimpleNamespace(save=Mock(), native=SimpleNamespace(SaveAs2=Mock()))
    client = SimpleNamespace(active_document=document)

    result = plugin.save_as(client, "output.docx")

    expected = str((tmp_path / "output.docx").resolve())
    document.save.assert_called_once_with(path=expected)
    assert result == expected


@pytest.mark.parametrize(
    ("arguments", "constant_name"),
    [
        (["--format", "xml"], "wdFormatXML"),
        (["--format", "xml", "--document"], "wdFormatXMLDocument"),
        (["--format", "xml", "--document", "--macro"], "wdFormatXMLDocumentMacroEnabled"),
        (["--format", "xml", "--template"], "wdFormatXMLTemplate"),
        (["--format", "xml", "--template", "--macro"], "wdFormatXMLTemplateMacroEnabled"),
        (["--format", "xml", "--flat"], "wdFormatFlatXML"),
        (["--format", "xml", "--flat", "--macro"], "wdFormatFlatXMLMacroEnabled"),
        (["--format", "xml", "--flat", "--template"], "wdFormatFlatXMLTemplate"),
        (["--format", "xml", "--flat", "--template", "--macro"], "wdFormatFlatXMLTemplateMacroEnabled"),
        (["--format", "xml", "--strict"], "wdFormatStrictOpenXMLDocument"),
        (["--format", "text"], "wdFormatText"),
        (["--format", "text", "--line-breaks"], "wdFormatTextLineBreaks"),
        (["--format", "text", "--dos"], "wdFormatDOSText"),
        (["--format", "text", "--dos", "--line-breaks"], "wdFormatDOSTextLineBreaks"),
        (["--format", "text", "--encoded"], "wdFormatEncodedText"),
        (["--format", "text", "--unicode"], "wdFormatUnicodeText"),
        (["--format", "html"], "wdFormatHTML"),
        (["--format", "html", "--filtered"], "wdFormatFilteredHTML"),
        (["--format", "wdFormatRTF"], "wdFormatRTF"),
        (["--format", "pdf"], "wdFormatPDF"),
        (["--format", "xps"], "wdFormatXPS"),
        (["--format", "wdFormatXMLDocument"], "wdFormatXMLDocument"),
        (["--format", "wdFormatText"], "wdFormatText"),
        (["--format", "wdFormatPDF"], "wdFormatPDF"),
        (["--format", "wdFormatXPS"], "wdFormatXPS"),
    ],
)
def test_formatted_save_as_resolves_format(plugin, command_runtime, tmp_path, arguments, constant_name):
    output = tmp_path / "output.any"

    result = command_runtime.runner.invoke(plugin.save_as_cmd, [str(output), *arguments])

    expected = str(output.resolve())
    assert result.exit_code == 0, result.output
    command_runtime.document.native.SaveAs2.assert_called_once_with(
        FileName=expected,
        FileFormat=getattr(plugin.C, constant_name),
    )
    assert result.output == f'Saved as "{expected}"\n'


@pytest.mark.parametrize(
    "arguments",
    [
        ["--document"],
        ["--template"],
        ["--flat"],
        ["--strict"],
        ["--macro"],
        ["--filtered"],
        ["--format", "text", "--document"],
        ["--line-breaks"],
        ["--dos"],
        ["--encoded"],
        ["--unicode"],
        ["--format", "html", "--document"],
        ["--format", "html", "--filtered", "--unicode"],
        ["--format", "html-filtered"],
        ["--format", "xml", "--dos"],
        ["--format", "missing"],
        ["--format", "wdFormatMissing"],
        ["--format", "xml", "--document", "--template"],
        ["--format", "xml", "--flat", "--document"],
        ["--format", "xml", "--strict", "--macro"],
        ["--format", "xml", "--macro"],
        ["--format", "text", "--encoded", "--unicode"],
        ["--format", "text", "--unicode", "--dos"],
    ],
)
def test_save_as_rejects_invalid_format_options(plugin, command_runtime, tmp_path, arguments):
    result = command_runtime.runner.invoke(plugin.save_as_cmd, [str(tmp_path / "output.any"), *arguments])

    assert result.exit_code == 2
    assert "Error:" in result.output
    command_runtime.get_client.assert_not_called()


def test_save_as_converts_word_api_error_to_click_exception(plugin, command_runtime, runtime, tmp_path):
    command_runtime.document.save.side_effect = runtime.core.WordAPIError("save rejected")

    result = command_runtime.runner.invoke(plugin.save_as_cmd, [str(tmp_path / "output.docx")])

    assert result.exit_code == 1
    assert "Error: save rejected" in result.output


def test_library_save_as_converts_word_api_error(plugin, runtime, tmp_path):
    document = SimpleNamespace(
        save=Mock(side_effect=runtime.core.WordAPIError("save rejected")),
        native=SimpleNamespace(SaveAs2=Mock()),
    )
    client = SimpleNamespace(active_document=document)

    with pytest.raises(runtime.core.WordAPIError, match="save rejected"):
        plugin.save_as(client, str(tmp_path / "output.docx"))


def test_list_save_as_formats_groups_runtime_formats_and_converters(plugin):
    constants = SimpleNamespace(**vars(plugin.C))
    constants.wdFormatDocument97 = 52
    constants.wdFormatDocumentDefault = 53
    constants.wdFormatTemplate97 = 54
    constants.wdFormatTemplate = 55
    constants.wdFormatHTML = 56
    constants.wdFormatFilteredHTML = 57
    constants.wdFormatWebArchive = 58
    constants.wdFormatOpenDocumentText = 59
    constants.wdFormatCustomSystemFormat = 901
    plugin.C = constants
    client = SimpleNamespace(
        native=SimpleNamespace(
            FileConverters=[
                SimpleNamespace(
                    CanSave=True,
                    FormatName="OpenDocument Text",
                    ClassName="OdfTextConverter",
                    SaveFormat=23,
                    Extensions="*.odt;*.fodt",
                ),
                SimpleNamespace(
                    CanSave=False,
                    FormatName="Import Only",
                    ClassName="ImportOnlyConverter",
                    SaveFormat=88,
                    Extensions="imp",
                ),
            ]
        )
    )

    formats = plugin._list_save_as_formats(client)

    assert formats["built_in"][0] == {"alias": "doc", "constant": "wdFormatDocument97"}
    assert {"alias": "pdf", "constant": "wdFormatPDF"} in formats["built_in"]
    assert {"label": "html", "constant": "wdFormatHTML"} in formats["html"]
    assert {"label": "html --filtered", "constant": "wdFormatFilteredHTML"} in formats["html"]
    assert {"label": "xml", "constant": "wdFormatXML"} in formats["xml"]
    assert {"label": "xml --document", "constant": "wdFormatXMLDocument"} in formats["xml"]
    assert {"label": "xml --flat --template --macro", "constant": "wdFormatFlatXMLTemplateMacroEnabled"} in formats["xml"]
    assert {"label": "text", "constant": "wdFormatText"} in formats["text"]
    assert {"label": "text --dos --line-breaks", "constant": "wdFormatDOSTextLineBreaks"} in formats["text"]
    assert formats["detected_constants"] == [
        {"alias": "wdFormatCustomSystemFormat", "constant": "wdFormatCustomSystemFormat"}
    ]
    assert formats["converters"] == [
        {
            "name": "OpenDocument Text",
            "class_name": "OdfTextConverter",
            "save_format": 23,
            "extensions": ["odt", "fodt"],
        }
    ]


def test_render_save_as_formats_omits_empty_sections_and_formats_labels(plugin):
    rendered = plugin._render_save_as_formats(
        {
            "built_in": [
                {"alias": "docx", "constant": "wdFormatDocumentDefault"},
                {"alias": "pdf", "constant": "wdFormatPDF"},
            ],
            "xml": [
                {"label": "xml", "constant": "wdFormatXML"},
                {"label": "xml --strict", "constant": "wdFormatStrictOpenXMLDocument"},
            ],
            "html": [
                {"label": "html", "constant": "wdFormatHTML"},
                {"label": "html --filtered", "constant": "wdFormatFilteredHTML"},
            ],
            "text": [
                {"label": "text", "constant": "wdFormatText"},
                {"label": "text --unicode", "constant": "wdFormatUnicodeText"},
            ],
            "detected_constants": [],
            "converters": [],
        }
    )

    assert "Built-in formats:" in rendered
    assert "  docx -> wdFormatDocumentDefault" in rendered
    assert "  pdf -> wdFormatPDF" in rendered
    assert "HTML formats:" in rendered
    assert "  html -> wdFormatHTML" in rendered
    assert "  html --filtered -> wdFormatFilteredHTML" in rendered
    assert "XML formats:" in rendered
    assert "  xml -> wdFormatXML" in rendered
    assert "  xml --strict -> wdFormatStrictOpenXMLDocument" in rendered
    assert "Text formats:" in rendered
    assert "  text -> wdFormatText" in rendered
    assert "  text --unicode -> wdFormatUnicodeText" in rendered
    assert "Detected constants:" not in rendered
    assert "Converter formats:" not in rendered
