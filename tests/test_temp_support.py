import shutil
import uuid
from pathlib import Path

import tests._temp_support as _temp_support


def _make_workspace() -> Path:
    root = Path(__file__).resolve().parents[1] / ".tmp-test-temp-support"
    workspace = root / uuid.uuid4().hex
    workspace.mkdir(parents=True)
    return workspace


def test_choose_temp_root_prefers_system_temp_when_writable(monkeypatch):
    workspace = _make_workspace()
    system_temp = workspace / "system-temp"
    repo_root = workspace / "repo"
    repo_root.mkdir()

    try:
        monkeypatch.setattr(_temp_support.tempfile, "gettempdir", lambda: str(system_temp))

        chosen = _temp_support.choose_temp_root(repo_root)

        assert chosen == system_temp / "msword-cli-tests"
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def test_choose_temp_root_falls_back_to_repo_temp_when_system_temp_is_unwritable(monkeypatch):
    workspace = _make_workspace()
    system_temp = workspace / "system-temp"
    repo_root = workspace / "repo"
    repo_root.mkdir()

    try:
        monkeypatch.setattr(_temp_support.tempfile, "gettempdir", lambda: str(system_temp))

        def fake_is_writable(path: Path) -> bool:
            return path != system_temp / "msword-cli-tests"

        monkeypatch.setattr(_temp_support, "_is_writable_directory", fake_is_writable)

        chosen = _temp_support.choose_temp_root(repo_root)

        assert chosen == repo_root / ".pytest-tmp"
    finally:
        shutil.rmtree(workspace, ignore_errors=True)
