from pathlib import Path
from typing import Any, Dict, List, Optional

import click
from msword_cli import WordAPIError, get_client, handle_api_error
from pywintypes import com_error
from win32com.client import constants as C


_SAVE_AS_FORMAT_ALIASES = [
    ("doc", "wdFormatDocument97"),
    ("docx", "wdFormatDocumentDefault"),
    ("dot", "wdFormatTemplate97"),
    ("dotx", "wdFormatTemplate"),
    ("mht", "wdFormatWebArchive"),
    ("odt", "wdFormatOpenDocumentText"),
    ("pdf", "wdFormatPDF"),
    ("rtf", "wdFormatRTF"),
    ("xps", "wdFormatXPS"),
]

_XML_FORMAT_CONSTANTS = {
    "wdFormatXML",
    "wdFormatXMLDocument",
    "wdFormatXMLDocumentMacroEnabled",
    "wdFormatStrictOpenXMLDocument",
    "wdFormatXMLTemplate",
    "wdFormatXMLTemplateMacroEnabled",
    "wdFormatFlatXML",
    "wdFormatFlatXMLMacroEnabled",
    "wdFormatFlatXMLTemplate",
    "wdFormatFlatXMLTemplateMacroEnabled",
}

_TEXT_FORMAT_CONSTANTS = {
    "wdFormatText",
    "wdFormatTextLineBreaks",
    "wdFormatDOSText",
    "wdFormatDOSTextLineBreaks",
    "wdFormatEncodedText",
    "wdFormatUnicodeText",
}

_FIXED_FORMAT_CONSTANTS = {"wdFormatPDF", "wdFormatXPS"}
_COMPOSITIONAL_FORMATS = {
    "html": [
        ("html", "wdFormatHTML"),
        ("html --filtered", "wdFormatFilteredHTML"),
    ],
    "xml": [
        ("xml", "wdFormatXML"),
        ("xml --document", "wdFormatXMLDocument"),
        ("xml --document --macro", "wdFormatXMLDocumentMacroEnabled"),
        ("xml --strict", "wdFormatStrictOpenXMLDocument"),
        ("xml --flat", "wdFormatFlatXML"),
        ("xml --flat --macro", "wdFormatFlatXMLMacroEnabled"),
        ("xml --flat --template", "wdFormatFlatXMLTemplate"),
        ("xml --flat --template --macro", "wdFormatFlatXMLTemplateMacroEnabled"),
        ("xml --template", "wdFormatXMLTemplate"),
        ("xml --template --macro", "wdFormatXMLTemplateMacroEnabled"),
    ],
    "text": [
        ("text", "wdFormatText"),
        ("text --line-breaks", "wdFormatTextLineBreaks"),
        ("text --dos", "wdFormatDOSText"),
        ("text --dos --line-breaks", "wdFormatDOSTextLineBreaks"),
        ("text --encoded", "wdFormatEncodedText"),
        ("text --unicode", "wdFormatUnicodeText"),
    ],
}


def _resolve_xml_format(document: bool, template: bool, flat: bool, macro: bool, strict: bool) -> str:
    if document and template:
        raise click.UsageError("--document and --template cannot be used together.")
    if flat and document:
        raise click.UsageError("--flat cannot be combined with --document.")
    if strict and (document or template or flat or macro):
        raise click.UsageError("--strict cannot be combined with --document, --template, --flat, or --macro.")
    if macro and not (document or template or flat):
        raise click.UsageError("--macro requires --document, --template, or --flat with --format xml.")
    if strict:
        return "wdFormatStrictOpenXMLDocument"
    if flat and template and macro:
        return "wdFormatFlatXMLTemplateMacroEnabled"
    if flat and template:
        return "wdFormatFlatXMLTemplate"
    if flat and macro:
        return "wdFormatFlatXMLMacroEnabled"
    if flat:
        return "wdFormatFlatXML"
    if template and macro:
        return "wdFormatXMLTemplateMacroEnabled"
    if template:
        return "wdFormatXMLTemplate"
    if document and macro:
        return "wdFormatXMLDocumentMacroEnabled"
    if document:
        return "wdFormatXMLDocument"
    return "wdFormatXML"


def _resolve_text_format(line_breaks: bool, dos: bool, encoded: bool, unicode: bool) -> str:
    if encoded and (dos or line_breaks or unicode):
        raise click.UsageError("--encoded cannot be combined with --dos, --line-breaks, or --unicode.")
    if unicode and (dos or line_breaks or encoded):
        raise click.UsageError("--unicode cannot be combined with --dos, --line-breaks, or --encoded.")
    if dos and line_breaks:
        return "wdFormatDOSTextLineBreaks"
    if dos:
        return "wdFormatDOSText"
    if encoded:
        return "wdFormatEncodedText"
    if unicode:
        return "wdFormatUnicodeText"
    if line_breaks:
        return "wdFormatTextLineBreaks"
    return "wdFormatText"


def _resolve_html_format(filtered: bool) -> str:
    if filtered:
        return "wdFormatFilteredHTML"
    return "wdFormatHTML"


def _is_detected_format_constant(value: Optional[str]) -> bool:
    compositional_constants = {
        constant
        for family in _COMPOSITIONAL_FORMATS.values()
        for _, constant in family
    }
    excluded = (
        compositional_constants
        | {constant for _, constant in _SAVE_AS_FORMAT_ALIASES}
    )
    return isinstance(value, str) and value.startswith("wdFormat") and hasattr(C, value) and value not in excluded


def _iter_collection(collection: Any) -> List[Any]:
    if collection is None:
        return []
    try:
        return list(collection)
    except TypeError:
        pass
    count = getattr(collection, "Count", None)
    item = getattr(collection, "Item", None)
    if count is None or item is None:
        return []
    return [item(index) for index in range(1, int(count) + 1)]


def _normalize_extensions(raw: Any) -> List[str]:
    extensions = []
    for part in str(raw or "").replace(";", ",").split(","):
        cleaned = part.strip().lstrip("*.").lstrip(".")
        if cleaned:
            extensions.append(cleaned)
    return extensions


def _list_save_as_formats(client: Any) -> Dict[str, List[Dict[str, Any]]]:
    built_in = [
        {"alias": alias, "constant": constant}
        for alias, constant in _SAVE_AS_FORMAT_ALIASES
        if hasattr(C, constant)
    ]
    xml_formats = [
        {"label": label, "constant": constant}
        for label, constant in _COMPOSITIONAL_FORMATS["xml"]
        if hasattr(C, constant)
    ]
    html_formats = [
        {"label": label, "constant": constant}
        for label, constant in _COMPOSITIONAL_FORMATS["html"]
        if hasattr(C, constant)
    ]
    text_formats = [
        {"label": label, "constant": constant}
        for label, constant in _COMPOSITIONAL_FORMATS["text"]
        if hasattr(C, constant)
    ]
    detected_constants = [
        {"alias": name, "constant": name}
        for name in sorted(dir(C))
        if _is_detected_format_constant(name)
    ]
    converters = []
    for converter in _iter_collection(getattr(getattr(client, "native", None), "FileConverters", None)):
        if not getattr(converter, "CanSave", False):
            continue
        converters.append(
            {
                "name": getattr(converter, "FormatName", None) or "Unknown",
                "class_name": getattr(converter, "ClassName", None) or "Unknown",
                "save_format": getattr(converter, "SaveFormat", None),
                "extensions": _normalize_extensions(getattr(converter, "Extensions", None)),
            }
        )
    converters.sort(key=lambda item: (item["name"].lower(), item["class_name"].lower(), item["save_format"] or 0))
    return {
        "built_in": built_in,
        "html": html_formats,
        "xml": xml_formats,
        "text": text_formats,
        "detected_constants": detected_constants,
        "converters": converters,
    }


def _render_save_as_formats(formats: Dict[str, List[Dict[str, Any]]]) -> str:
    lines = ["Built-in formats:"]
    for item in formats.get("built_in", []):
        lines.append(f'  {item["alias"]} -> {item["constant"]}')
    html_formats = formats.get("html", [])
    if html_formats:
        lines.extend(["", "HTML formats:"])
        for item in html_formats:
            lines.append(f'  {item["label"]} -> {item["constant"]}')
    xml_formats = formats.get("xml", [])
    if xml_formats:
        lines.extend(["", "XML formats:"])
        for item in xml_formats:
            lines.append(f'  {item["label"]} -> {item["constant"]}')
    text_formats = formats.get("text", [])
    if text_formats:
        lines.extend(["", "Text formats:"])
        for item in text_formats:
            lines.append(f'  {item["label"]} -> {item["constant"]}')
    detected_constants = formats.get("detected_constants", [])
    if detected_constants:
        lines.extend(["", "Detected constants:"])
        for item in detected_constants:
            lines.append(f'  {item["alias"]} -> {item["constant"]}')
    converters = formats.get("converters", [])
    if converters:
        lines.extend(["", "Converter formats:"])
        for item in converters:
            suffix = ""
            if item["extensions"]:
                suffix = " (" + ", ".join(f'*.{ext}' for ext in item["extensions"]) + ")"
            lines.append(f'  {item["name"]} -> {item["save_format"]} [{item["class_name"]}]{suffix}')
    return "\n".join(lines)


def _com_error_message(error: Exception, fallback: str = "COM error") -> str:
    excepinfo = getattr(error, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return str(excepinfo[2])
    return str(error) or fallback


def _save_as2(document: Any, path: Any, file_format: Any) -> str:
    final_path = str(Path(path).resolve())
    try:
        document.native.SaveAs2(FileName=final_path, FileFormat=file_format)
    except com_error as error:
        raise WordAPIError(f"Save as failed: {_com_error_message(error)}") from error
    return final_path


def _resolve_save_as_constant(
    save_format: str,
    document: bool,
    template: bool,
    flat: bool,
    strict: bool,
    macro: bool,
    filtered: bool,
    line_breaks: bool,
    dos: bool,
    encoded: bool,
    unicode: bool,
) -> str:
    aliased_constant = dict(_SAVE_AS_FORMAT_ALIASES).get(save_format)
    if save_format == "html":
        return _resolve_html_format(filtered)
    if save_format == "xml":
        return _resolve_xml_format(document, template, flat, macro, strict)
    if save_format == "text":
        return _resolve_text_format(line_breaks, dos, encoded, unicode)
    if aliased_constant is not None and hasattr(C, aliased_constant):
        return aliased_constant
    if save_format.startswith("wdFormat") and hasattr(C, save_format):
        return save_format
    raise click.UsageError(f'Unknown or unsupported format "{save_format}".')


def _prepare_save_as(
    path: str,
    save_format: Optional[str] = None,
    document: bool = False,
    template: bool = False,
    flat: bool = False,
    strict: bool = False,
    macro: bool = False,
    filtered: bool = False,
    line_breaks: bool = False,
    dos: bool = False,
    encoded: bool = False,
    unicode: bool = False,
) -> Any:
    final_path = str(Path(path).resolve())
    html_options = filtered
    xml_options = document or template or flat or strict or macro
    text_options = line_breaks or dos or encoded or unicode
    if html_options and save_format != "html":
        raise click.UsageError("HTML options require --format html.")
    if xml_options and save_format != "xml":
        raise click.UsageError("XML options require --format xml.")
    if text_options and save_format != "text":
        raise click.UsageError("Text options require --format text.")
    constant_name = None
    if save_format is not None:
        constant_name = _resolve_save_as_constant(
            save_format=save_format,
            document=document,
            template=template,
            flat=flat,
            strict=strict,
            macro=macro,
            filtered=filtered,
            line_breaks=line_breaks,
            dos=dos,
            encoded=encoded,
            unicode=unicode,
        )
    return final_path, constant_name


def save_as(
    client: Any,
    path: str,
    save_format: Optional[str] = None,
    document: bool = False,
    template: bool = False,
    flat: bool = False,
    strict: bool = False,
    macro: bool = False,
    filtered: bool = False,
    line_breaks: bool = False,
    dos: bool = False,
    encoded: bool = False,
    unicode: bool = False,
) -> str:
    final_path, constant_name = _prepare_save_as(
        path=path,
        save_format=save_format,
        document=document,
        template=template,
        flat=flat,
        strict=strict,
        macro=macro,
        filtered=filtered,
        line_breaks=line_breaks,
        dos=dos,
        encoded=encoded,
        unicode=unicode,
    )
    if constant_name is None:
        client.active_document.save(path=final_path)
        return final_path
    return _save_as2(client.active_document, final_path, getattr(C, constant_name))


@click.command(
    "save-as",
    short_help="Save the active document to another path or format.",
    help="Save the active document to another path or convert it to another supported Word format. Use --list-formats to inspect save formats available through the current Word installation.",
)
@click.argument("path", required=False, type=click.Path(resolve_path=True))
@click.option("--list-formats", is_flag=True, help="List save formats available through Word on this system.")
@click.option("save_format", "--format", metavar="FORMAT")
@click.option("document", "--document", is_flag=True)
@click.option("template", "--template", is_flag=True)
@click.option("flat", "--flat", is_flag=True)
@click.option("strict", "--strict", is_flag=True)
@click.option("macro", "--macro", is_flag=True)
@click.option("filtered", "--filtered", is_flag=True)
@click.option("line_breaks", "--line-breaks", is_flag=True)
@click.option("dos", "--dos", is_flag=True)
@click.option("encoded", "--encoded", is_flag=True)
@click.option("unicode", "--unicode", is_flag=True)
@handle_api_error
def save_as_cmd(
    path: Optional[str],
    list_formats: bool,
    save_format: Optional[str],
    document: bool,
    template: bool,
    flat: bool,
    strict: bool,
    macro: bool,
    filtered: bool,
    line_breaks: bool,
    dos: bool,
    encoded: bool,
    unicode: bool,
) -> None:
    """Save the active document under a new name."""
    if list_formats:
        if path is not None:
            raise click.UsageError("Path is not allowed with --list-formats.")
        if save_format or document or template or flat or strict or macro or filtered or line_breaks or dos or encoded or unicode:
            raise click.UsageError("--list-formats cannot be combined with format selection options.")
        click.echo(_render_save_as_formats(_list_save_as_formats(get_client())))
        return

    if path is None:
        raise click.MissingParameter(param_hint="PATH", param_type="argument")

    final_path, constant_name = _prepare_save_as(
        path=path,
        save_format=save_format,
        document=document,
        template=template,
        flat=flat,
        strict=strict,
        macro=macro,
        filtered=filtered,
        line_breaks=line_breaks,
        dos=dos,
        encoded=encoded,
        unicode=unicode,
    )

    if constant_name is None:
        save_as(get_client(), final_path)
        click.echo(f'Saving active document as "{final_path}"')
        return

    final_path = save_as(
        get_client(),
        final_path,
        save_format=save_format,
        document=document,
        template=template,
        flat=flat,
        strict=strict,
        macro=macro,
        filtered=filtered,
        line_breaks=line_breaks,
        dos=dos,
        encoded=encoded,
        unicode=unicode,
    )
    click.echo(f'Saved as "{final_path}"')


plugin_manifest = {
    "name": "save-as",
    "command": save_as_cmd,
    "client_methods": {
        "save_as": save_as,
    },
}
