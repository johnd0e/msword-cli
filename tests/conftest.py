import importlib
import os
import sys
from click.testing import CliRunner
import uuid
from contextlib import contextmanager
import shutil
import tempfile
from pathlib import Path
from types import ModuleType, SimpleNamespace
from unittest.mock import Mock

import pytest

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


class FakeComDocument:
    def __init__(self, name: str, saved: bool = True):
        self.Name = name
        self.Saved = saved
        self.Activate = Mock()
        self.Save = Mock()
        self.SaveAs = Mock()
        self.Close = Mock()
        self.ExportAsFixedFormat = Mock()
        self.PrintOut = Mock()


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


def make_document(name: str, saved: bool = True):
    document = Mock()
    document.name = name
    document.saved = saved
    document.activate = Mock()
    document.save = Mock()
    document.close = Mock()
    document.export_fixed_format = Mock()
    document.print_out = Mock()
    return document

