from importlib import import_module
from pathlib import Path
from types import SimpleNamespace
import sys
from unittest.mock import Mock

import click

try:
    import tomllib
except ImportError:  # pragma: no cover - Python < 3.11
    import tomli as tomllib

from tests.conftest import _make_pywin32_stubs


ROOT = Path(__file__).resolve().parents[1]


def import_with_plugins(monkeypatch, entry_points_result):
    pywintypes, win32com, client = _make_pywin32_stubs()
    monkeypatch.setitem(sys.modules, "pywintypes", pywintypes)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    metadata = import_module('importlib.metadata')
    monkeypatch.setattr(metadata, 'entry_points', Mock(return_value=entry_points_result))
    sys.modules.pop('msword_cli', None)
    return import_module('msword_cli')


def write_local_plugin(plugin_root: Path, package_name: str, command_name: str, body: str = "click.echo('local plugin')") -> Path:
    plugin_dir = plugin_root / package_name
    plugin_dir.mkdir(parents=True)
    (plugin_dir / 'pyproject.toml').write_text(
        '\n'.join(
            [
                '[project]',
                f'name = "{package_name}"',
                'version = "0.1.0"',
                '',
                '[project.entry-points."msw.plugin"]',
                f'{command_name} = "{package_name}:{command_name.replace("-", "_")}_cmd"',
                '',
            ]
        ),
        encoding='utf-8',
    )
    (plugin_dir / f'{package_name}.py').write_text(
        '\n'.join(
            [
                'import click',
                '',
                f'@click.command("{command_name}")',
                f'def {command_name.replace("-", "_")}_cmd():',
                f'    {body}',
                '',
            ]
        ),
        encoding='utf-8',
    )
    return plugin_dir


def test_save_as_plugin_is_a_default_workspace_dependency():
    with (ROOT / 'pyproject.toml').open('rb') as config_file:
        config = tomllib.load(config_file)

    assert 'msword-cli-save-as~=0.4.0' in config['project']['dependencies']
    assert config['tool']['uv']['sources']['msword-cli-save-as'] == {'workspace': True}
    assert config['tool']['uv']['workspace']['members'] == ['plugins/*']


def test_plugin_command_is_registered(monkeypatch):
    @click.command('hello-plugin')
    def hello_plugin():
        click.echo('hello')

    plugin = SimpleNamespace(load=Mock(return_value=hello_plugin))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    ctx = click.Context(msword_cli.cli)

    assert msword_cli.cli.get_command(ctx, 'hello-plugin') is hello_plugin
    plugin.load.assert_called_once_with()
    entry_points_result.select.assert_called_once_with(group='msw.plugin')


def test_save_as_command_is_not_defined_by_core_without_plugins(monkeypatch):
    entry_points_result = SimpleNamespace(select=Mock(return_value=[]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)

    assert 'save-as' not in msword_cli.cli.commands


def test_plugin_can_register_save_as_command(monkeypatch, tmp_path):
    @click.command('save-as')
    def plugin_save_as():
        click.echo('plugin save-as')

    plugin = SimpleNamespace(load=Mock(return_value=plugin_save_as))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    monkeypatch.setattr(msword_cli, '__file__', str(tmp_path / 'msword_cli.py'))
    ctx = click.Context(msword_cli.cli)

    assert msword_cli.cli.get_command(ctx, 'save-as') is plugin_save_as


def test_plugin_load_failure_emits_warning(monkeypatch):
    plugin = SimpleNamespace(name='broken', load=Mock(side_effect=RuntimeError('bad plugin')))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))
    echo = Mock()
    monkeypatch.setattr(click, 'echo', echo)

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    msword_cli.cli.list_commands(click.Context(msword_cli.cli))

    echo.assert_any_call('Warning: Failed to load plugin broken: bad plugin', err=True)


def test_plugin_load_failure_does_not_block_later_plugins(monkeypatch):
    @click.command('working-plugin')
    def working_plugin():
        pass

    broken = SimpleNamespace(name='broken', load=Mock(side_effect=RuntimeError('bad plugin')))
    working = SimpleNamespace(name='working', load=Mock(return_value=working_plugin))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[broken, working]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    ctx = click.Context(msword_cli.cli)

    assert msword_cli.cli.get_command(ctx, 'working-plugin') is working_plugin
    working.load.assert_called_once_with()


def test_plugin_dir_loads_local_plugins_without_install(monkeypatch, tmp_path):
    entry_points_result = SimpleNamespace(select=Mock(return_value=[]))
    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    write_local_plugin(tmp_path, 'local_plugin', 'hello-plugin')

    result = click.testing.CliRunner().invoke(msword_cli.cli, ['--plugin-dir', str(tmp_path), 'hello-plugin'])

    assert result.exit_code == 0, result.output
    assert result.output == 'local plugin\n'


def test_plugin_dir_overrides_installed_plugin(monkeypatch, tmp_path):
    @click.command('save-as')
    def installed_plugin():
        click.echo('installed plugin')

    plugin = SimpleNamespace(load=Mock(return_value=installed_plugin))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))
    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    write_local_plugin(tmp_path, 'local_save_as', 'save-as')

    result = click.testing.CliRunner().invoke(msword_cli.cli, ['--plugin-dir', str(tmp_path), 'save-as'])

    assert result.exit_code == 0, result.output
    assert result.output == 'local plugin\n'


def test_plugin_dir_failure_does_not_block_later_local_plugins(monkeypatch, tmp_path):
    entry_points_result = SimpleNamespace(select=Mock(return_value=[]))
    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    broken_dir = tmp_path / 'broken'
    broken_dir.mkdir()
    (broken_dir / 'pyproject.toml').write_text(
        '\n'.join(
            [
                '[project]',
                'name = "broken"',
                'version = "0.1.0"',
                '',
                '[project.entry-points."msw.plugin"]',
                'broken = "broken:not_a_command"',
                '',
            ]
        ),
        encoding='utf-8',
    )
    write_local_plugin(tmp_path, 'working_plugin', 'working-plugin')

    result = click.testing.CliRunner().invoke(msword_cli.cli, ['--plugin-dir', str(tmp_path), 'working-plugin'])

    assert result.exit_code == 0, result.output
    assert result.output.endswith('local plugin\n')
    assert 'Warning: Failed to load local plugin' in result.output


def test_script_mode_loads_repo_local_plugins_by_default(monkeypatch, tmp_path):
    entry_points_result = SimpleNamespace(select=Mock(return_value=[]))
    msword_cli = import_with_plugins(monkeypatch, entry_points_result)
    plugin_root = tmp_path / 'plugins'
    write_local_plugin(plugin_root, 'local_plugin', 'hello-plugin')
    monkeypatch.setattr(msword_cli, '__file__', str(tmp_path / 'msword_cli.py'))

    result = click.testing.CliRunner().invoke(msword_cli.cli, ['hello-plugin'])

    assert result.exit_code == 0, result.output
    assert result.output == 'local plugin\n'
