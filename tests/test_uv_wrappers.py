import os
import subprocess
from pathlib import Path


def _write_fake_uv(bin_dir: Path) -> None:
    script = "\n".join(
        [
            "@echo off",
            "echo UV_CACHE_DIR=%UV_CACHE_DIR%",
            "echo ARGS=%*",
            "exit /b 0",
        ]
    )
    (bin_dir / "uv.cmd").write_text(script, encoding="ascii")


def _run_powershell(script: Path, *args: str, env: dict[str, str]) -> subprocess.CompletedProcess[str]:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        *args,
    ]
    return subprocess.run(command, capture_output=True, text=True, check=False, env=env)


def test_uv_run_preserves_explicit_uv_cache_dir(tmp_path):
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    _write_fake_uv(bin_dir)

    env = os.environ.copy()
    env["PATH"] = f"{bin_dir}{os.pathsep}{env['PATH']}"
    env["UV_CACHE_DIR"] = str(tmp_path / "preconfigured-cache")

    script = Path(__file__).resolve().parents[1] / "scripts" / "uv-run.ps1"
    result = _run_powershell(script, "--version", env=env)

    assert result.returncode == 0
    assert f"UV_CACHE_DIR={env['UV_CACHE_DIR']}" in result.stdout
    assert "ARGS=--version" in result.stdout


def test_uv_run_falls_back_to_override_cache_dir_when_default_cache_is_unusable(tmp_path):
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    _write_fake_uv(bin_dir)

    blocked_default = tmp_path / "blocked-cache"
    blocked_default.write_text("not a directory", encoding="ascii")
    fallback_cache = tmp_path / "fallback-cache"

    env = os.environ.copy()
    env["PATH"] = f"{bin_dir}{os.pathsep}{env['PATH']}"
    env.pop("UV_CACHE_DIR", None)

    script = Path(__file__).resolve().parents[1] / "scripts" / "uv-run.ps1"
    result = _run_powershell(
        script,
        "-DefaultCacheDir",
        str(blocked_default),
        "-CacheDir",
        str(fallback_cache),
        "--version",
        env=env,
    )

    assert result.returncode == 0
    assert f"UV_CACHE_DIR={fallback_cache}" in result.stdout
    assert "ARGS=--version" in result.stdout
    assert fallback_cache.is_dir()


def test_run_pytest_wraps_uv_run_with_default_pytest_arguments(tmp_path):
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    _write_fake_uv(bin_dir)

    blocked_default = tmp_path / "blocked-cache"
    blocked_default.write_text("not a directory", encoding="ascii")
    fallback_cache = tmp_path / "fallback-cache"

    env = os.environ.copy()
    env["PATH"] = f"{bin_dir}{os.pathsep}{env['PATH']}"
    env.pop("UV_CACHE_DIR", None)

    script = Path(__file__).resolve().parents[1] / "scripts" / "run-pytest.ps1"
    result = _run_powershell(
        script,
        "-DefaultCacheDir",
        str(blocked_default),
        "-CacheDir",
        str(fallback_cache),
        "-k",
        "smoke",
        env=env,
    )

    assert result.returncode == 0
    assert f"UV_CACHE_DIR={fallback_cache}" in result.stdout
    assert "ARGS=run --no-sync pytest -q -k smoke" in result.stdout


def test_run_pytest_uses_default_uv_arguments_without_named_overrides(tmp_path):
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    _write_fake_uv(bin_dir)

    env = os.environ.copy()
    env["PATH"] = f"{bin_dir}{os.pathsep}{env['PATH']}"
    env["UV_CACHE_DIR"] = str(tmp_path / "preconfigured-cache")

    script = Path(__file__).resolve().parents[1] / "scripts" / "run-pytest.ps1"
    result = _run_powershell(script, env=env)

    assert result.returncode == 0
    assert f"UV_CACHE_DIR={env['UV_CACHE_DIR']}" in result.stdout
    assert "ARGS=run --no-sync pytest -q" in result.stdout
