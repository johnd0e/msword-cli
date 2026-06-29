from importlib import import_module
from types import SimpleNamespace
import sys
from unittest.mock import Mock

import click

from tests.conftest import _make_pywin32_stubs


def import_with_plugins(monkeypatch, entry_points_result):
    pywintypes, win32com, client = _make_pywin32_stubs()
    monkeypatch.setitem(sys.modules, "pywintypes", pywintypes)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)

    metadata = import_module('importlib.metadata')
    monkeypatch.setattr(metadata, 'entry_points', Mock(return_value=entry_points_result))
    sys.modules.pop('msword_cli', None)
    return import_module('msword_cli')


def test_plugin_command_is_registered(monkeypatch):
    @click.command('hello-plugin')
    def hello_plugin():
        click.echo('hello')

    plugin = SimpleNamespace(load=Mock(return_value=hello_plugin))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)

    assert 'hello-plugin' in msword_cli.cli.commands
    plugin.load.assert_called_once_with()
    entry_points_result.select.assert_called_once_with(group='msw.plugin')


def test_plugin_load_failure_emits_warning(monkeypatch):
    plugin = SimpleNamespace(name='broken', load=Mock(side_effect=RuntimeError('bad plugin')))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[plugin]))
    echo = Mock()
    monkeypatch.setattr(click, 'echo', echo)

    import_with_plugins(monkeypatch, entry_points_result)

    echo.assert_any_call('Warning: Failed to load plugin broken: bad plugin', err=True)


def test_plugin_load_failure_does_not_block_later_plugins(monkeypatch):
    @click.command('working-plugin')
    def working_plugin():
        pass

    broken = SimpleNamespace(name='broken', load=Mock(side_effect=RuntimeError('bad plugin')))
    working = SimpleNamespace(name='working', load=Mock(return_value=working_plugin))
    entry_points_result = SimpleNamespace(select=Mock(return_value=[broken, working]))

    msword_cli = import_with_plugins(monkeypatch, entry_points_result)

    assert 'working-plugin' in msword_cli.cli.commands
    working.load.assert_called_once_with()
