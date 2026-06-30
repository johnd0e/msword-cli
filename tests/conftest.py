import importlib
import datetime as dt
import os
import shutil
import sys
import uuid
from contextlib import contextmanager
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import Mock

import pytest
from click.testing import CliRunner


@pytest.fixture(scope="session", autouse=True)
def test_tempdir():
    root = Path(__file__).resolve().parents[1] / ".pytest-tmp"
    root.mkdir(exist_ok=True)
    os.environ["TMPDIR"] = str(root)
    yield


@pytest.fixture(autouse=True)
def patch_click_isolated_filesystem(monkeypatch):
    @contextmanager
    def isolated_filesystem(self, temp_dir=None):
        base = Path(temp_dir) if temp_dir is not None else Path(__file__).resolve().parents[1] / ".pytest-tmp"
        base.mkdir(parents=True, exist_ok=True)
        cwd = Path.cwd()
        target = base / f"tmp-{uuid.uuid4().hex}"
        target.mkdir()
        try:
            os.chdir(target)
            yield str(target)
        finally:
            os.chdir(cwd)
            shutil.rmtree(target, ignore_errors=True)

    monkeypatch.setattr(CliRunner, "isolated_filesystem", isolated_filesystem)
    yield


def _make_pywin32_stubs():
    pywintypes = ModuleType("pywintypes")

    class FakeComError(Exception):
        def __init__(self, *args, excepinfo=None, **kwargs):
            super().__init__(*args)
            self.excepinfo = excepinfo or (None, None, "boom")

    pywintypes.com_error = FakeComError

    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.gencache = SimpleNamespace(EnsureDispatch=Mock())
    client.constants = SimpleNamespace(
        wdUserTemplatesPath=1,
        wdDoNotSaveChanges=2,
        wdPromptToSaveChanges=3,
        wdExportFormatPDF=4,
        wdExportFormatXPS=5,
        wdExportOptimizeForPrint=6,
        wdExportOptimizeForOnScreen=7,
        wdExportFromTo=8,
        wdExportAllDocument=9,
        wdExportCurrentPage=10,
        wdExportSelection=11,
        wdExportDocumentWithMarkup=12,
        wdExportDocumentContent=13,
        wdExportCreateNoBookmarks=14,
        wdExportCreateHeadingBookmarks=15,
        wdExportCreateWordBookmarks=16,
        wdPrintAllPages=17,
        wdPrintAllDocument=18,
        wdPrintDocumentContent=19,
        wdPrintDocumentWithMarkup=20,
        wdPrintComments=21,
        wdPrintProperties=22,
        wdPrintMarkup=23,
        wdPrintStyles=24,
        wdPrintAutoTextEntries=25,
        wdPrintKeyAssignments=26,
        wdPrintEnvelope=27,
        wdPrintOddPagesOnly=28,
        wdPrintEvenPagesOnly=29,
        wdPrintCurrentPage=30,
        wdPrintSelection=31,
        wdPrintRangeOfPages=32,
        wdCompareDestinationNew=33,
        wdCompareDestinationOriginal=34,
        wdCompareDestinationRevised=35,
        wdMergeDestinationNewDocument=36,
        wdMergeDestinationOriginalDocument=37,
        wdMergeDestinationRevisedDocument=38,
        wdGranularityWordLevel=39,
        wdGranularityCharLevel=40,
        wdReplaceAll=41,
        wdFindStop=42,
        wdStatisticPages=43,
        wdStatisticWords=44,
        wdStatisticCharacters=45,
        wdStatisticParagraphs=46,
        msoPropertyTypeNumber=47,
        msoPropertyTypeBoolean=48,
        msoPropertyTypeString=49,
        msoPropertyTypeFloat=50,
    )
    win32com.client = client
    return pywintypes, win32com, client


@pytest.fixture
def msword_cli(monkeypatch):
    pywintypes, win32com, client = _make_pywin32_stubs()
    monkeypatch.setitem(sys.modules, "pywintypes", pywintypes)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    sys.modules.pop("msword_cli", None)
    module = importlib.import_module("msword_cli")
    return importlib.reload(module)


@pytest.fixture(scope="session")
def real_msword_cli():
    if os.getenv("MSWORD_RUN_INTEGRATION") != "1":
        pytest.skip("Set MSWORD_RUN_INTEGRATION=1 to run live Word integration tests.")

    sys.modules.pop("msword_cli", None)
    sys.modules.pop("pywintypes", None)
    for module_name in tuple(sys.modules):
        if module_name == "win32com" or module_name.startswith("win32com."):
            sys.modules.pop(module_name, None)

    try:
        import pythoncom
        import pywintypes  # noqa: F401
        import win32com.client  # noqa: F401
        import win32com.client.CLSIDToClass  # noqa: F401
        import win32com.client.util  # noqa: F401
    except ImportError as exc:
        pytest.skip(f"pywin32 is not available: {exc}")

    pythoncom.CoInitialize()
    try:
        yield importlib.import_module("msword_cli")
    finally:
        pythoncom.CoUninitialize()


class FakeCollection:
    def __init__(self, items=None):
        self._items = list(items or [])

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        if isinstance(index, str):
            for item in self._items:
                if getattr(item, "Name", None) == index:
                    return item
            raise KeyError(index)
        return self._items[index - 1]


class FakeProperty:
    def __init__(self, name, value):
        self.Name = name
        self.Value = value
        self.Delete = Mock()


class FakePropertyCollection(FakeCollection):
    def Add(self, Name, LinkToContent, Type, Value):
        prop = FakeProperty(Name, Value)
        self._items.append(prop)
        return prop


class FakeCommentsCollection(FakeCollection):
    pass


class FakeRevisionsCollection(FakeCollection):
    def __init__(self, count=0):
        super().__init__([object()] * count)
        self.AcceptAll = Mock()
        self.RejectAll = Mock()


class FakeFieldsCollection(FakeCollection):
    def __init__(self, count=0):
        super().__init__([object()] * count)
        self.Update = Mock()


class FakeRange:
    def __init__(self, text="", start=0, end=None):
        self.Text = text
        self.Start = start
        self.End = len(text) if end is None else end
        self.Find = SimpleNamespace(Execute=Mock(return_value=False))
        self.SetRange = Mock(side_effect=self._set_range)
        self.Comments = FakeCommentsCollection()
        self.Revisions = FakeRevisionsCollection(0)
        self.Fields = FakeFieldsCollection(0)

    @property
    def Duplicate(self):
        duplicate = FakeRange(self.Text, self.Start, self.End)
        duplicate.Find = self.Find
        duplicate.SetRange = self.SetRange
        duplicate.Comments = self.Comments
        duplicate.Revisions = self.Revisions
        duplicate.Fields = self.Fields
        return duplicate

    def _set_range(self, start, end):
        self.Start = start
        self.End = end


class FakeComment:
    def __init__(self, author="Tester", text="Comment", scope_text="Body", initials="TT"):
        self.Author = author
        self.Initial = initials
        self.Date = "2026-06-30"
        self.Range = SimpleNamespace(Text=text)
        self.Scope = SimpleNamespace(Text=scope_text)
        self.Delete = Mock()


class FakeComDocument:
    def __init__(self, name: str, saved: bool = True):
        self.Name = name
        self.Saved = saved
        self.FullName = str(Path(name).resolve())
        self.ReadOnly = False
        self.TrackRevisions = False
        self.Activate = Mock()
        self.Save = Mock()
        self.SaveAs = Mock()
        self.SaveCopyAs = Mock()
        self.Close = Mock()
        self.ExportAsFixedFormat = Mock()
        self.PrintOut = Mock()
        self.Comments = FakeCommentsCollection()
        self.Revisions = FakeRevisionsCollection(0)
        self.Fields = FakeFieldsCollection(0)
        self.TablesOfContents = FakeCollection()
        self.Content = FakeRange()
        self.AttachedTemplate = SimpleNamespace(FullName="C:/Templates/normal.dotm")
        self.BuiltInDocumentProperties = FakePropertyCollection(
            [
                FakeProperty("Title", "Report"),
                FakeProperty("Author", "Alice"),
                FakeProperty("Last Author", "Bob"),
                FakeProperty("Creation Date", dt.datetime(2026, 6, 1, 9, 30, 0)),
                FakeProperty("Last Save Time", dt.datetime(2026, 6, 29, 18, 45, 0)),
            ]
        )
        self.CustomDocumentProperties = FakePropertyCollection([FakeProperty("Project", "CLI")])
        self.ComputeStatistics = Mock(side_effect=lambda constant: {43: 2, 44: 10, 45: 50, 46: 3}[constant])


class FakeDocumentCollection:
    def __init__(self, docs):
        self._docs = list(docs)
        self.Open = Mock()
        self.Add = Mock()

    @property
    def Count(self):
        return len(self._docs)

    def Item(self, index: int):
        return self._docs[index - 1]


class FakeWordApp:
    def __init__(self, docs=None, template_dir="C:/Templates", visible=False):
        self.Visible = visible
        self.UserName = "Word User"
        self._docs = list(docs or [])
        self.Documents = FakeDocumentCollection(self._docs)
        self.Options = SimpleNamespace(DefaultFilePath=Mock(return_value=template_dir))
        self.Quit = Mock()
        self.CompareDocuments = Mock()
        self.MergeDocuments = Mock()
        self.Selection = SimpleNamespace(Range=FakeRange())

    @property
    def ActiveDocument(self):
        if not self._docs:
            raise RuntimeError("No active document")
        return self._docs[-1]


class FakeClient:
    def __init__(self, docs=None, template_dir="C:/Templates", visible=False, document_count=None):
        self.template_dir = Path(template_dir)
        self.active_document = Mock()
        self.documents = list(docs or [])
        if self.documents:
            self.active_document = self.documents[-1]
        self.document_count = len(self.documents) if document_count is None else document_count
        self.visible = visible
        self.open = Mock()
        self.new = Mock()
        self.quit = Mock()
        self.compare = Mock()
        self.merge = Mock()
        self.find = Mock(return_value=[])
        self.replace = Mock(return_value=0)
        self.set_track_changes = Mock(return_value=True)
        self.accept_revisions = Mock(return_value=0)
        self.reject_revisions = Mock(return_value=0)
        self.list_comments = Mock(return_value=[])
        self.export_comments = Mock()
        self.delete_comments = Mock(return_value=0)
        self.update_fields = Mock(return_value={"fields": 0, "tables_of_contents": 0})
        self.summary = Mock(return_value={"name": "foo.docx"})
        self.statistics = Mock(return_value={"pages": 1})
        self.list_properties = Mock(return_value=[])
        self.get_property = Mock(return_value={"kind": "custom", "name": "Project", "value": "CLI"})
        self.set_property = Mock(return_value={"kind": "custom", "name": "Project", "value": "CLI"})
        self.delete_property = Mock(return_value=True)


def make_document(name: str, saved: bool = True):
    document = Mock()
    document.name = name
    document.saved = saved
    document.activate = Mock()
    document.save = Mock()
    document.save_copy = Mock(return_value=str(Path(name).resolve()))
    document.close = Mock()
    document.export_fixed_format = Mock()
    document.print_out = Mock()
    return document
