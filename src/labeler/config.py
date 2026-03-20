from __future__ import annotations

import json
import re
import shutil
import sys
from datetime import datetime
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict

from .models import ALLOWED_FIELD_TYPES, ALLOWED_RULE_ACTIONS, ALLOWED_RULE_OPERATORS

APP_NAME = 'Offline Dialogue Labeler Pro'

DEFAULT_APP_SETTINGS = {
    'autosave_seconds': 30,
    'history_recent_limit': 20,
    'history_global_limit': 200,
    'window_geometry': '1500x980',
    'last_mode_id': '',
    'last_input_file': '',
    'last_annotator_name': '',
}


class ModesConfigError(Exception):
    def __init__(self, message: str, path=None, line=None, column=None):
        super().__init__(message)
        self.path = Path(path) if path else None
        self.line = line
        self.column = column


class ModesConfigFileMissingError(ModesConfigError):
    pass


class ModesConfigJsonError(ModesConfigError):
    pass


class ModesConfigValidationError(ModesConfigError):
    pass


def _project_root():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def ensure_runtime_dirs():
    base_dir = _project_root()
    input_dir = base_dir / 'input'
    output_dir = base_dir / 'output'
    drafts_dir = output_dir / '_drafts'
    meta_dir = output_dir / '_meta'
    config_dir = base_dir / 'config'
    backups_dir = config_dir / '_history'
    for path in [input_dir, output_dir, drafts_dir, meta_dir, config_dir, backups_dir]:
        path.mkdir(parents=True, exist_ok=True)
    modes_path = config_dir / 'modes.json'
    app_settings_path = config_dir / 'app_settings.json'
    if not app_settings_path.exists():
        save_app_settings(dict(DEFAULT_APP_SETTINGS), app_settings_path)
    return {
        'base_dir': base_dir,
        'input_dir': input_dir,
        'output_dir': output_dir,
        'drafts_dir': drafts_dir,
        'meta_dir': meta_dir,
        'config_dir': config_dir,
        'backups_dir': backups_dir,
        'modes_path': modes_path,
        'app_settings_path': app_settings_path,
    }


def _load_json(path, default):
    path = Path(path)
    if not path.exists():
        return default
    try:
        with path.open('r', encoding='utf-8') as fh:
            return json.load(fh)
    except Exception:
        return default


def _save_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def load_app_settings(path):
    payload = _load_json(path, {})
    settings = dict(DEFAULT_APP_SETTINGS)
    if isinstance(payload, dict):
        settings.update(payload)
    return settings


def save_app_settings(settings, path):
    payload = dict(DEFAULT_APP_SETTINGS)
    payload.update(settings or {})
    _save_json(path, payload)


def bump_mode_version(version):
    match = re.match(r'^(\d+)(?:\.(\d+))?$', str(version or '1.0').strip())
    if not match:
        return '1.0'
    major = int(match.group(1))
    minor = int(match.group(2) or 0)
    minor += 1
    return '{0}.{1}'.format(major, minor)


def _validate_field(field):
    if not isinstance(field, dict):
        raise ValueError('Описание поля должно быть объектом.')
    for key in ['key', 'label', 'type']:
        if not str(field.get(key, '')).strip():
            raise ValueError(u"У поля должны быть заполнены key, label и type.")
    if field.get('type') not in ALLOWED_FIELD_TYPES:
        raise ValueError(u"Неподдерживаемый тип поля: {0}".format(field.get('type')))
    if field.get('type') in ('select', 'multiselect') and not (field.get('options') or []):
        raise ValueError(u"Для поля {0} требуется непустой список options.".format(field.get('key')))


def _validate_rule(rule, field_keys):
    if not isinstance(rule, dict):
        raise ValueError('Правило должно быть объектом.')
    if not str(rule.get('id', '')).strip():
        raise ValueError('У правила должен быть id.')
    if rule.get('severity', 'error') not in ('error', 'warning'):
        raise ValueError(u"У правила {0} severity должен быть error или warning.".format(rule.get('id')))
    for condition in rule.get('if', []) or []:
        if condition.get('field') not in field_keys:
            raise ValueError(u"Правило {0}: неизвестное поле в if: {1}".format(rule.get('id'), condition.get('field')))
        if condition.get('op') not in ALLOWED_RULE_OPERATORS:
            raise ValueError(u"Правило {0}: неизвестный оператор {1}".format(rule.get('id'), condition.get('op')))
    for action in rule.get('then', []) or []:
        if action.get('field') not in field_keys:
            raise ValueError(u"Правило {0}: неизвестное поле в then: {1}".format(rule.get('id'), action.get('field')))
        if action.get('type') not in ALLOWED_RULE_ACTIONS:
            raise ValueError(u"Правило {0}: неизвестное действие {1}".format(rule.get('id'), action.get('type')))


def validate_mode(mode):
    if not isinstance(mode, dict):
        raise ValueError('Режим должен быть объектом.')
    for key in ['id', 'name', 'fields']:
        if key not in mode:
            raise ValueError(u"В режиме отсутствует обязательный ключ: {0}".format(key))
    if not isinstance(mode.get('fields'), list) or not mode.get('fields'):
        raise ValueError(u"В режиме {0} должен быть непустой список fields.".format(mode.get('id')))
    field_keys = []
    for field in mode.get('fields', []):
        _validate_field(field)
        key = field.get('key')
        if key in field_keys:
            raise ValueError(u"В режиме {0} дублируется field.key: {1}".format(mode.get('id'), key))
        field_keys.append(key)
    for rule in mode.get('rules', []) or []:
        _validate_rule(rule, field_keys)


def validate_modes_payload(payload):
    if not isinstance(payload, dict):
        raise ValueError('Файл modes.json должен содержать JSON-объект.')
    modes = payload.get('modes')
    if not isinstance(modes, list) or not modes:
        raise ValueError('В modes.json должен быть непустой список modes.')
    seen = set()
    for mode in modes:
        validate_mode(mode)
        if mode['id'] in seen:
            raise ValueError(u"Дублируется mode.id: {0}".format(mode['id']))
        seen.add(mode['id'])
    return True


def load_modes(path):
    path = Path(path)
    if not path.exists():
        raise ModesConfigFileMissingError(
            u"Не найден файл конфигурации режимов: {0}".format(path),
            path=path,
        )
    try:
        with path.open('r', encoding='utf-8') as fh:
            payload = json.load(fh)
    except JSONDecodeError as exc:
        raise ModesConfigJsonError(
            u"Ошибка в формате modes.json: строка {0}, столбец {1}. {2}".format(exc.lineno, exc.colno, exc.msg),
            path=path,
            line=exc.lineno,
            column=exc.colno,
        )
    except Exception as exc:
        raise ModesConfigError(u"Не удалось прочитать modes.json: {0}".format(exc), path=path)
    try:
        validate_modes_payload(payload)
    except Exception as exc:
        raise ModesConfigValidationError(str(exc), path=path)
    return payload


def save_modes(payload, modes_path, backups_dir):
    validate_modes_payload(payload)
    modes_path = Path(modes_path)
    backups_dir = Path(backups_dir)
    backups_dir.mkdir(parents=True, exist_ok=True)
    if modes_path.exists():
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = backups_dir / ('modes_{0}.json'.format(ts))
        shutil.copy2(str(modes_path), str(backup_path))
    _save_json(modes_path, payload)


def latest_modes_backup(backups_dir):
    backups_dir = Path(backups_dir)
    if not backups_dir.exists():
        return None
    candidates = [path for path in backups_dir.glob('modes_*.json') if path.is_file()]
    if not candidates:
        return None
    return sorted(candidates, key=lambda p: p.name, reverse=True)[0]


def restore_latest_modes_backup(modes_path, backups_dir):
    modes_path = Path(modes_path)
    backup_path = latest_modes_backup(backups_dir)
    if backup_path is None:
        raise FileNotFoundError('Не найдено ни одной резервной копии modes.json.')
    modes_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(backup_path), str(modes_path))
    return backup_path
