import os
import tempfile
import uuid
from pathlib import Path


def repo_temp_root(repo_root: Path) -> Path:
    return repo_root / ".pytest-tmp"


def system_temp_root() -> Path:
    return Path(tempfile.gettempdir()) / "msword-cli-tests"


def choose_temp_root(repo_root: Path) -> Path:
    system_temp = system_temp_root()
    if _is_writable_directory(system_temp):
        return system_temp

    fallback = repo_temp_root(repo_root)
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def _is_writable_directory(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / f".msword-cli-temp-probe-{uuid.uuid4().hex}"
        probe.write_text("ok", encoding="ascii")
        probe.unlink()
        return True
    except OSError:
        return False


def apply_temp_environment(root: Path) -> None:
    value = os.fspath(root)
    os.environ["TMPDIR"] = value
    os.environ["TEMP"] = value
    os.environ["TMP"] = value
