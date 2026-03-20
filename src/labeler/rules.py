from __future__ import annotations

from typing import Any, Dict, List

from .models import ALLOWED_RULE_ACTIONS, ALLOWED_RULE_OPERATORS, ValidationIssue


def _is_filled(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return True
    if isinstance(value, (list, tuple, set, dict)):
        return bool(value)
    return str(value).strip() != ''


def _normalize_scalar(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return ''
    return str(value).strip()


def _condition_matches(condition, values):
    field = condition.get('field', '')
    op = condition.get('op', '')
    expected = condition.get('value')
    actual = values.get(field)

    if op not in ALLOWED_RULE_OPERATORS:
        return False
    if op == 'equals':
        return _normalize_scalar(actual) == _normalize_scalar(expected)
    if op == 'not_equals':
        return _normalize_scalar(actual) != _normalize_scalar(expected)
    if op == 'in':
        if isinstance(expected, list):
            options = [_normalize_scalar(item) for item in expected]
        else:
            options = [item.strip() for item in str(expected).split(';') if item.strip()]
        return _normalize_scalar(actual) in options
    if op == 'not_empty':
        return _is_filled(actual)
    if op == 'empty':
        return not _is_filled(actual)
    if op == 'is_true':
        return bool(actual) is True
    if op == 'is_false':
        return bool(actual) is False
    return False


def _evaluate_rule(rule, values):
    conditions = rule.get('if', []) or []
    return all(_condition_matches(condition, values) for condition in conditions)


def matched_rule_actions(mode, values, action_type=None):
    actions = []
    for rule in mode.get('rules', []) or []:
        if not _evaluate_rule(rule, values):
            continue
        for action in rule.get('then', []) or []:
            if action_type is not None and action.get('type') != action_type:
                continue
            actions.append(action)
    return actions


def disabled_field_keys(mode, values):
    result = []
    seen = set()
    for action in matched_rule_actions(mode, values, action_type='require_empty'):
        field_key = str(action.get('field', '')).strip()
        if not field_key or field_key in seen:
            continue
        seen.add(field_key)
        result.append(field_key)
    return result


def _apply_rule_action(rule, action, values):
    issues = []
    action_type = action.get('type', '')
    field = action.get('field', '')
    level = rule.get('severity', 'error') or 'error'
    message = rule.get('message') or rule.get('description') or 'Нарушено правило.'
    value = values.get(field)
    if action_type not in ALLOWED_RULE_ACTIONS:
        return issues
    if action_type == 'require_filled' and not _is_filled(value):
        issues.append(ValidationIssue(level=level, message=message, field_key=field, rule_id=rule.get('id', '')))
    elif action_type == 'require_empty' and _is_filled(value):
        issues.append(ValidationIssue(level=level, message=message, field_key=field, rule_id=rule.get('id', '')))
    elif action_type == 'require_value_in':
        allowed = action.get('values', []) or []
        normalized_allowed = [_normalize_scalar(item) for item in allowed]
        if _normalize_scalar(value) not in normalized_allowed:
            issues.append(ValidationIssue(level=level, message=message, field_key=field, rule_id=rule.get('id', '')))
    return issues


def validate_annotation(mode, values):
    issues = []
    for field in mode.get('fields', []):
        key = field.get('key', '')
        value = values.get(key)
        field_type = field.get('type', 'text')
        if field.get('required') and not _is_filled(value):
            issues.append(ValidationIssue(level='error', message=u"Поле '{0}' обязательно для заполнения.".format(field.get('label', key)), field_key=key))
        if field_type == 'number' and _is_filled(value):
            try:
                float(str(value).replace(',', '.'))
            except Exception:
                issues.append(ValidationIssue(level='error', message=u"Поле '{0}' должно содержать число.".format(field.get('label', key)), field_key=key))
        if field_type == 'select' and _is_filled(value):
            options = field.get('options', []) or []
            if options and str(value) not in options:
                issues.append(ValidationIssue(level='error', message=u"Поле '{0}' должно содержать одно из допустимых значений.".format(field.get('label', key)), field_key=key))
        if field_type == 'multiselect' and value:
            options = set(field.get('options', []) or [])
            wrong = [item for item in value if item not in options]
            if wrong:
                issues.append(ValidationIssue(level='error', message=u"Поле '{0}' содержит недопустимые значения.".format(field.get('label', key)), field_key=key))
    for rule in mode.get('rules', []) or []:
        if _evaluate_rule(rule, values):
            for action in rule.get('then', []) or []:
                issues.extend(_apply_rule_action(rule, action, values))
    return issues


def is_annotation_complete(mode, values):
    issues = validate_annotation(mode, values)
    return not any(issue.level == 'error' for issue in issues)
