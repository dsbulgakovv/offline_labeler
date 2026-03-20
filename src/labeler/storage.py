from __future__ import annotations

import json
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List

HISTORY_FILENAME = 'field_history.json'


def _history_path(meta_dir):
    return Path(meta_dir) / HISTORY_FILENAME


def _default_history_payload():
    return {
        'schema_version': '1.0',
        'updated_at': '',
        'by_mode': {},
    }


def load_history(meta_dir):
    path = _history_path(meta_dir)
    if not path.exists():
        return _default_history_payload()
    try:
        with path.open('r', encoding='utf-8') as fh:
            payload = json.load(fh)
    except Exception:
        return _default_history_payload()
    if not isinstance(payload, dict):
        return _default_history_payload()
    payload.setdefault('by_mode', {})
    return payload


def save_history(meta_dir, payload):
    path = _history_path(meta_dir)
    Path(meta_dir).mkdir(parents=True, exist_ok=True)
    payload = dict(payload)
    payload['updated_at'] = datetime.now().isoformat(timespec='seconds')
    with path.open('w', encoding='utf-8') as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def _iter_values_for_history(field, value):
    field_type = field.get('type', 'text')
    if value is None:
        return []
    if field_type in ('text', 'textarea', 'select', 'number'):
        text = str(value).strip()
        return [text] if text else []
    if field_type == 'multiselect':
        return [str(item).strip() for item in (value or []) if str(item).strip()]
    if field_type == 'checkbox':
        return ['Да'] if bool(value) else []
    return []


def build_current_file_history(fields, annotations):
    result = {}
    field_map = dict((field.get('key', ''), field) for field in fields)
    for values in annotations.values():
        for key, field in field_map.items():
            result.setdefault(key, {})
            for item in _iter_values_for_history(field, values.get(key)):
                result[key][item] = result[key].get(item, 0) + 1
    return result


def _ensure_mode_field_payload(payload, mode_id, field_key):
    by_mode = payload.setdefault('by_mode', {})
    mode_payload = by_mode.setdefault(mode_id, {'fields': {}})
    fields_payload = mode_payload.setdefault('fields', {})
    field_payload = fields_payload.setdefault(field_key, {'global_counts': {}, 'recent': []})
    field_payload.setdefault('global_counts', {})
    field_payload.setdefault('recent', [])
    return field_payload


def update_history_from_annotations(payload, mode_id, fields, annotations, recent_limit=20, global_limit=200):
    payload = payload or _default_history_payload()
    payload.setdefault('by_mode', {})
    field_map = dict((field.get('key', ''), field) for field in fields)
    current_seen = OrderedDict()
    for values in annotations.values():
        for field_key, field in field_map.items():
            for item in _iter_values_for_history(field, values.get(field_key)):
                current_seen[item] = True
                fp = _ensure_mode_field_payload(payload, mode_id, field_key)
                fp['global_counts'][item] = int(fp['global_counts'].get(item, 0)) + 1
                recent = [entry for entry in fp.get('recent', []) if entry != item]
                recent.insert(0, item)
                fp['recent'] = recent[:max(1, int(recent_limit))]
    # truncate global counts per field
    for field_key in field_map:
        fp = _ensure_mode_field_payload(payload, mode_id, field_key)
        items = sorted(fp['global_counts'].items(), key=lambda x: (-int(x[1]), x[0].lower()))[:max(1, int(global_limit))]
        fp['global_counts'] = dict(items)
    payload['updated_at'] = datetime.now().isoformat(timespec='seconds')
    return payload


def _prefix_filter(values, prefix):
    prefix = (prefix or '').strip().lower()
    if not prefix:
        return list(values)
    return [value for value in values if value.lower().startswith(prefix)]


def get_grouped_history(payload, current_file_history, mode_id, field_key):
    payload = payload or _default_history_payload()
    field_payload = payload.get('by_mode', {}).get(mode_id, {}).get('fields', {}).get(field_key, {})
    current_counts = current_file_history.get(field_key, {}) or {}
    current_file = [value for value, _ in sorted(current_counts.items(), key=lambda x: (-int(x[1]), x[0].lower()))]
    global_counts = field_payload.get('global_counts', {}) or {}
    global_top = [value for value, _ in sorted(global_counts.items(), key=lambda x: (-int(x[1]), x[0].lower()))][:50]
    recent = list(field_payload.get('recent', []) or [])[:50]
    return {
        'current_file': current_file,
        'global_top': global_top,
        'recent': recent,
    }


def get_field_suggestions(payload, current_file_history, mode_id, field_key, prefix='', max_items=40):
    grouped = get_grouped_history(payload, current_file_history, mode_id, field_key)
    merged = []
    seen = set()
    for bucket in ('current_file', 'global_top', 'recent'):
        for value in grouped.get(bucket, []):
            if value and value not in seen:
                seen.add(value)
                merged.append(value)
    merged = _prefix_filter(merged, prefix)
    return merged[:max_items]
