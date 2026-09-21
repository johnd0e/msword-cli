import shutil
import uuid
from pathlib import Path

from tests import _temp_support


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


def test_choose_run_temp_root_reuses_existing_env(monkeypatch):
    workspace = _make_workspace()
    repo_root = workspace / "repo"
    repo_root.mkdir()
    existing = workspace / "run-root"

    try:
        monkeypatch.setenv("MSWORD_TEST_RUN_TEMP_ROOT", str(existing))

        chosen = _temp_support.choose_run_temp_root(repo_root)

        assert chosen == existing
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def test_choose_run_temp_root_creates_unique_directory(monkeypatch):
    workspace = _make_workspace()
    repo_root = workspace / "repo"
    repo_root.mkdir()
    base_root = workspace / "base-temp"

    try:
        monkeypatch.setattr(_temp_support, "choose_temp_root", lambda path: base_root)
        monkeypatch.delenv("MSWORD_TEST_RUN_TEMP_ROOT", raising=False)

        first = _temp_support.choose_run_temp_root(repo_root)
        second = _temp_support.choose_run_temp_root(repo_root)

        assert first == second
        assert first.parent == base_root
        assert first.name.startswith("msword-cli-")
        assert first.exists()
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def test_remove_temp_tree_retries_after_transient_error(monkeypatch):
    workspace = _make_workspace()
    target = workspace / "to-delete"
    target.mkdir()
    (target / "file.txt").write_text("x", encoding="ascii")
    calls = {"count": 0}
    real_rmtree = _temp_support.shutil.rmtree

    def flaky_rmtree(path):
        calls["count"] += 1
        if calls["count"] == 1:
            raise PermissionError("locked")
        return real_rmtree(path)

    try:
        monkeypatch.setattr(_temp_support.shutil, "rmtree", flaky_rmtree)
        monkeypatch.setattr(_temp_support.time, "sleep", lambda _: None)

        removed = _temp_support.remove_temp_tree(target, attempts=2, delay_seconds=0)

        assert removed is True
        assert calls["count"] == 2  # noqa: PLR2004
        assert not target.exists()
    finally:
        real_rmtree(workspace, ignore_errors=True)
