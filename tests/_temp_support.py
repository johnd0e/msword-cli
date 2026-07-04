import os
import shutil
import tempfile
import time
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


def choose_run_temp_root(repo_root: Path) -> Path:
    existing = os.environ.get("MSWORD_TEST_RUN_TEMP_ROOT")
    if existing:
        return Path(existing)

    base_dir = choose_temp_root(repo_root)
    base_dir.mkdir(parents=True, exist_ok=True)
    root = Path(tempfile.mkdtemp(prefix="msword-cli-", dir=os.fspath(base_dir)))
    os.environ["MSWORD_TEST_RUN_TEMP_ROOT"] = os.fspath(root)
    return root


def remove_temp_tree(path: Path, attempts: int = 6, delay_seconds: float = 0.5) -> bool:
    for attempt in range(attempts):
        try:
            shutil.rmtree(path)
            return True
        except FileNotFoundError:
            return True
        except OSError:
            if attempt == attempts - 1:
                return False
            time.sleep(delay_seconds)
    return False


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
