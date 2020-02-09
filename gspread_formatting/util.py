# -*- coding: utf-8 -*-

from gspread.utils import a1_to_rowcol as _a1_to_rowcol

def _build_repeat_cell_request(worksheet, range, cell_format, celldata_field='userEnteredFormat'):
    return {
        'repeatCell': {
            'range': _range_to_gridrange_object(range, worksheet.id),
            'cell': { celldata_field: cell_format.to_props() },
            'fields': ",".join(cell_format.affected_fields(celldata_field))
        }
    }

def _fetch_with_updated_properties(spreadsheet, key, params=None):
    try:
        return spreadsheet._properties[key]
    except KeyError:
        metadata = spreadsheet.fetch_sheet_metadata(params)
        spreadsheet._properties.update(metadata['properties'])
        return spreadsheet._properties[key]

def _range_to_gridrange_object(range, worksheet_id):
    parts = range.split(':')
    start = parts[0]
    end = parts[1] if len(parts) > 1 else ''
    (row_offset, column_offset) = _a1_to_rowcol(start)
    (last_row, last_column) = _a1_to_rowcol(end) if end else (row_offset, column_offset)
    return {
        'sheetId': worksheet_id,
        'startRowIndex': row_offset-1,
        'endRowIndex': last_row,
        'startColumnIndex': column_offset-1,
        'endColumnIndex': last_column
    }

def _props_to_component(class_registry, class_alias, value, none_if_empty=False):
    if class_alias not in class_registry:
        raise ValueError("No format component named '%s'" % class_alias)
    cls = class_registry[class_alias]
    kwargs = {}
    for k, v in value.items():
        if isinstance(v, dict):
            if isinstance(cls._FIELDS, dict) and cls._FIELDS.get(k) is not None:
                item_alias = cls._FIELDS[k]
            else:
                item_alias = k
            v = _props_to_component(class_registry, item_alias, v, True)
        if v is not None:
            kwargs[k] = v
    return cls(**kwargs) if (kwargs or not none_if_empty) else None

def _ul_repl(m):
    return '_' + m.group(1).lower()

def _underlower(name):
    return name[0].lower() + name[1:]

def _parse_string_enum(name, value, set_of_values, required=False):
    if value is None and required:
        raise ValueError("%s value is required" % name)
    if value is not None and value.upper() not in set_of_values:
        raise ValueError("%s value must be one of: %s" % (name, set_of_values))
    return value.upper() if value is not None else None

def _enforce_type(name, cls, value, required=False):
    if value is None and required:
        raise ValueError("%s value is required" % name)
    if value is not None and not isinstance(value, cls):
        raise ValueError("%s value must be instance of: %s" % (name, cls))
    return value

def _extract_props(value):
    if hasattr(value, 'to_props'):
        return value.to_props()
    return value

def _extract_fieldrefs(name, value, prefix):
    if hasattr(value, 'affected_fields'):
        return value.affected_fields(".".join([prefix, name]))
    elif value is not None:
        return [".".join([prefix, name])]
    else:
        return []

