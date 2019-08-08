# -*- coding: utf-8 -*-

from .util import _props_to_component, _extract_props, _extract_fieldrefs, \
    _parse_string_enum, _underlower, _range_to_gridrange_object
                  
class FormattingComponent(object):
    _FIELDS = ()

    @classmethod
    def from_props(cls, props):
        return _props_to_component(_CLASSES, _underlower(cls.__name__), props)

    def __repr__(self):
        return '<' + self.__class__.__name__ + ' ' + str(self) + '>'

    def __str__(self):
        p = []
        for a in self._FIELDS:
            v = getattr(self, a)
            if v is not None:
                if isinstance(v, FormattingComponent):
                    p.append( (a, "(" + str(v) + ")") )
                else:
                    p.append( (a, str(v)) )
        return ";".join(["%s=%s" % (k, v) for k, v in p])

    def to_props(self):
        p = {}
        for a in self._FIELDS:
            if getattr(self, a) is not None:
                p[a] = _extract_props(getattr(self, a))
        return p

    def affected_fields(self, prefix):
        fields = []
        for a in self._FIELDS:
            if getattr(self, a) is not None:
                fields.extend( _extract_fieldrefs(a, getattr(self, a), prefix) )
        return fields

    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        for a in self._FIELDS:
            self_v = getattr(self, a, None)
            other_v = getattr(other, a, None)
            if self_v != other_v:
                return False
        return True

    def __ne__(self, other):
        return not self.__eq__(other)

    def add(self, other):
        new_props = {}
        for a in self._FIELDS:
            self_v = getattr(self, a, None)
            other_v = getattr(other, a, None)
            if isinstance(self_v, CellFormatComponent):
                this_v = self_v.add(other_v)
            elif other_v is not None:
                this_v = other_v
            else:
                this_v = self_v
            if this_v is not None:
                new_props[a] = _extract_props(this_v)
        return self.__class__.from_props(new_props)

    __add__ = add

    def intersection(self, other):
        new_props = {}
        for a in self._FIELDS:
            self_v = getattr(self, a, None)
            other_v = getattr(other, a, None)
            this_v = None
            if isinstance(self_v, CellFormatComponent):
                this_v = self_v.intersection(other_v)
            elif self_v == other_v:
                this_v = self_v
            if this_v is not None:
                new_props[a] = _extract_props(this_v)
        return self.__class__.from_props(new_props) if new_props else None

    __and__ = intersection

    def difference(self, other):
        new_props = {}
        for a in self._FIELDS:
            self_v = getattr(self, a, None)
            other_v = getattr(other, a, None)
            this_v = None
            if isinstance(self_v, CellFormatComponent):
                this_v = self_v.difference(other_v)
            elif other_v != self_v:
                this_v = self_v
            if this_v is not None:
                new_props[a] = _extract_props(this_v)
        return self.__class__.from_props(new_props) if new_props else None

    __sub__ = difference

class GridRange(FormattingComponent):
    _FIELDS = ('sheetId', 'startRowIndex', 'endRowIndex', 'startColumnIndex', 'endColumnIndex')

    @classmethod
    def from_a1_range(cls, range):
        return GridRange.from_props(_range_to_gridrange_object(range))

    def __init__(self, sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex):
        self.sheetId = sheetId
        self.startRowIndex = startRowIndex
        self.endRowIndex = endRowIndex
        self.startColumnIndex = startColumnIndex
        self.endColumnIndex = endColumnIndex

class CellFormatComponent(FormattingComponent):
    pass

class CellFormat(CellFormatComponent):
    _FIELDS = (
        'numberFormat', 'backgroundColor', 'borders', 'padding', 
        'horizontalAlignment', 'verticalAlignment', 'wrapStrategy', 
        'textDirection', 'textFormat', 'hyperlinkDisplayType', 'textRotation'
    )

    def __init__(self, 
        numberFormat=None,
        backgroundColor=None,
        borders=None,
        padding=None,
        horizontalAlignment=None,
        verticalAlignment=None,
        wrapStrategy=None,
        textDirection=None,
        textFormat=None,
        hyperlinkDisplayType=None,
        textRotation=None
        ):
        self.numberFormat = numberFormat
        self.backgroundColor = backgroundColor
        self.borders = borders
        self.padding = padding
        self.horizontalAlignment = _parse_string_enum('horizontalAlignment', horizontalAlignment, set(['LEFT', 'CENTER', 'RIGHT']))
        self.verticalAlignment = _parse_string_enum('verticalAlignment', verticalAlignment, set(['TOP', 'MIDDLE', 'BOTTOM']))
        self.wrapStrategy = _parse_string_enum('wrapStrategy', wrapStrategy, set(['OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP']))
        self.textDirection = _parse_string_enum('textDirection', textDirection, set(['LEFT_TO_RIGHT', 'RIGHT_TO_LEFT']))
        self.textFormat = textFormat
        self.hyperlinkDisplayType = _parse_string_enum('hyperlinkDisplayType', hyperlinkDisplayType, set(['LINKED', 'PLAIN_TEXT']))
        self.textRotation = textRotation

class NumberFormat(CellFormatComponent):
    _FIELDS = ('type', 'pattern')

    TYPES = set(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'])

    def __init__(self, type, pattern=None):
        self.type = _parse_string_enum("type", type, NumberFormat.TYPES, required=True)
        self.pattern = pattern

class Color(CellFormatComponent):
    _FIELDS = ('red', 'green', 'blue', 'alpha')

    def __init__(self, red=None, green=None, blue=None, alpha=None):
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha

class Border(CellFormatComponent):
    _FIELDS = ('style', 'color')

    STYLES = set(['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE'])

    def __init__(self, style, color=None, width=None):
        self.style = _parse_string_enum("style", style, Border.STYLES, required=True)
        self.width = width
        self.color = color

class Borders(CellFormatComponent):
    _FIELDS = {
        'top': 'border', 
        'bottom': 'border', 
        'left': 'border', 
        'right': 'border'
    }

    def __init__(self, top=None, bottom=None, left=None, right=None):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right

class Padding(CellFormatComponent):
    _FIELDS = ('top', 'right', 'bottom', 'left')

    def __init__(self, top=None, right=None, bottom=None, left=None):
        self.top = top
        self.right = right
        self.bottom = bottom
        self.left = left

class TextFormat(CellFormatComponent):
    _FIELDS = ('foregroundColor', 'fontFamily', 'fontSize', 'bold', 'italic', 'strikethrough', 'underline')

    def __init__(self, 
        foregroundColor=None, 
        fontFamily=None, 
        fontSize=None, 
        bold=None, 
        italic=None, 
        strikethrough=None, 
        underline=None
        ):
        self.foregroundColor = foregroundColor
        self.fontFamily = fontFamily
        self.fontSize = fontSize
        self.bold = bold
        self.italic = italic
        self.strikethrough = strikethrough
        self.underline = underline

class TextRotation(CellFormatComponent):
    _FIELDS = ('angle', 'vertical')

    def __init__(self, angle=None, vertical=None):
        if len([expr for expr in (angle is not None, vertical is not None) if expr]) != 1:
            raise ValueError("Either angle or vertical must be specified, not both or neither")
        self.angle = angle
        self.vertical = vertical

# Conditional formatting objects

class BooleanRule(CellFormatComponent):
    _FIELDS = ('condition', 'format')

    def __init__(self, condition, format):
        self.condition = condition
        self.format = format

class BooleanCondition(CellFormatComponent):
    _FIELDS = ('type', 'values')

    TYPES = {
        'NUMBER_GREATER': 1,
        'NUMBER_GREATER_THAN_EQ': 1,
        'NUMBER_LESS': 1,
        'NUMBER_LESS_THAN_EQ': 1,
        'NUMBER_EQ': 1,
        'NUMBER_NOT_EQ': 1,
        'NUMBER_BETWEEN': 2,
        'NUMBER_NOT_BETWEEN': 2,
        'TEXT_CONTAINS': 1,
        'TEXT_NOT_CONTAINS': 1,
        'TEXT_STARTS_WITH': 1,
        'TEXT_ENDS_WITH': 1,
        'TEXT_EQ': 1,
        'TEXT_IS_EMAIL': 0,
        'TEXT_IS_URL': 0,
        'DATE_EQ': 1,
        'DATE_BEFORE': 1,
        'DATE_AFTER': 1,
        'DATE_ON_OR_BEFORE': 1,
        'DATE_ON_OR_AFTER': 1,
        'DATE_BETWEEN': 2,
        'DATE_NOT_BETWEEN': 2,
        'DATE_IS_VALID': 0,
        'ONE_OF_RANGE': 1,
        'ONE_OF_LIST': (lambda x: isinstance(x, (list, tuple)) and len(x) > 0),
        'BLANK': 0,
        'NOT_BLANK': 0,
        'CUSTOM_FORMULA': 1,
        'BOOLEAN': (lambda x: isinstance(x, (list, tuple)) and len(x) >= 0 and len(x) <= 2)
    }

    def __init__(self, type, values):
        self.type = _parse_string_enum("type", type, BooleanCondition.TYPES.keys())
        validator = BooleanCondition.TYPES[self.type]
        if (callable(validator) and not validator(values)) or len(values) != validator:
            raise ValueError("BooleanCondition.values has inappropriate length/content for condition type %s" % self.type)
        # values are either RelativeDate enum values or user-entered values
        self.values = [ 
            v if isinstance(v, ConditionValue) else ConditionValue(v) 
            for v in values 
        ]

class RelativeDate(object):
    VALUES = set(['PAST_YEAR', 'PAST_MONTH', 'PAST_WEEK', 'YESTERDAY', 'TODAY', 'TOMORROW'])

    def __init__(self, value):
        self.value = _parse_string_enum("value", value, RelativeDate.VALUES, required=True)

    def to_props(self):
        return self.value

class ConditionValue(CellFormatComponent):
    _FIELDS = ('relativeDate', 'userEnteredValue')

    def __init__(self, value):
        self.relativeDate = None
        self.userEnteredValue = None
        if isinstance(value, RelativeDate):
            self.relativeDate = value.value
        else:
            self.userEnteredValue = value

class InterpolationPoint(CellFormatComponent):
    _FIELDS = ('color', 'type', 'value')

    TYPES = set(['MIN', 'MAX', 'NUMBER', 'PERCENT', 'PERCENTILE'])

    def __init__(self, color, type, value=None):
        self.color = color
        self.type = _parse_string_enum("type", type, InterpolationPoint.TYPES, required=True)
        if value is None and self.type not in set(['MIN', 'MAX']):
            raise ValueError("InterpolationPoint.type %s requires a value" % self.type)
        self.value = value

class GradientRule(CellFormatComponent):
    _FIELDS = ('minpoint', 'midpoint', 'maxpoint')

    def __init__(self, minpoint, midpoint, maxpoint):
        self.minpoint = minpoint
        self.midpoint = midpoint
        self.maxpoint = maxpoint

# provide camelCase aliases for all component classes.

_CLASSES = {}
for _c in [ 
        obj for name, obj in locals().items() 
        if isinstance(obj, type) and issubclass(obj, FormattingComponent) 
    ]:
    _k = _underlower(_c.__name__)
    _CLASSES[_k] = _c
    locals()[_k] = _c
_CLASSES['foregroundColor'] = Color
_CLASSES['backgroundColor'] = Color

