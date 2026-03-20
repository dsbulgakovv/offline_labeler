from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Dict

ALLOWED_FIELD_TYPES = {
    'text',
    'textarea',
    'select',
    'multiselect',
    'checkbox',
    'number',
}

ALLOWED_RULE_OPERATORS = {
    'equals',
    'not_equals',
    'in',
    'not_empty',
    'empty',
    'is_true',
    'is_false',
}

ALLOWED_RULE_ACTIONS = {
    'require_filled',
    'require_empty',
    'require_value_in',
}

NUMBER_MISSING_MARKERS = {'none', 'null', 'nan'}


@dataclass
class ValidationIssue:
    level: str
    message: str
    field_key: str = ''
    rule_id: str = ''


def clone_mode(mode):
    return deepcopy(mode)


def is_number_missing_marker(value):
    return isinstance(value, str) and value.strip().lower() in NUMBER_MISSING_MARKERS


def _field_default(field):
    if 'default' in field:
        default = field.get('default')
        if field.get('type') == 'number' and (default is None or is_number_missing_marker(default)):
            return None
        if isinstance(default, list):
            return list(default)
        return default
    field_type = field.get('type', 'text')
    if field_type == 'checkbox':
        return False
    if field_type == 'multiselect':
        return []
    return ''


def empty_annotation(mode):
    values = {}
    for field in mode.get('fields', []):
        values[field['key']] = _field_default(field)
    values['confidence'] = ''
    values['needs_review'] = False
    values['annotator_comment'] = ''
    values['completed'] = False
    values['updated_at'] = ''
    return values
