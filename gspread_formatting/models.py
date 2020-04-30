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
    def from_a1_range(cls, range, worksheet):
        return GridRange.from_props(_range_to_gridrange_object(range, worksheet.id))

    def __init__(self, sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex):
        self.sheetId = sheetId
        self.startRowIndex = startRowIndex
        self.endRowIndex = endRowIndex
        self.startColumnIndex = startColumnIndex
        self.endColumnIndex = endColumnIndex

class CellFormatComponent(FormattingComponent):
    pass

class CellFormat(CellFormatComponent):
    _FIELDS = {
        'numberFormat': None,
        'backgroundColor': 'color',
        'borders': None,
        'padding': None,
        'horizontalAlignment': None,
        'verticalAlignment': None,
        'wrapStrategy': None,
        'textDirection': None,
        'textFormat': None,
        'hyperlinkDisplayType': None,
        'textRotation': None,
        'foregroundColorStyle': 'colorStyle',
        'backgroundColorStyle': 'colorStyle'
    }

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
        textRotation=None,
        foregroundColorStyle=None,
        backgroundColorStyle=None
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
        self.foregroundColorStyle = foregroundColorStyle
        self.backgroundColorStyle = backgroundColorStyle

class NumberFormat(CellFormatComponent):
    _FIELDS = ('type', 'pattern')

    TYPES = set(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'])

    def __init__(self, type, pattern=None):
        if type.upper() not in NumberFormat.TYPES:
            raise ValueError("NumberFormat.type must be one of: %s" % NumberFormat.TYPES)
        self.type = type.upper()
        self.pattern = pattern

class ColorStyle(CellFormatComponent):
    _FIELDS = {
        'themeColor': None,
        'rgbColor': 'color'
    }

    def __init__(self, themeColor=None, rgbColor=None):
        self.themeColor = themeColor
        self.rgbColor = rgbColor

class Color(CellFormatComponent):
    _FIELDS = ('red', 'green', 'blue', 'alpha')

    def __init__(self, red=None, green=None, blue=None, alpha=None):
        self.red = red
        self.green = green
        self.blue = blue
        self.alpha = alpha

    @classmethod
    def fromHex(cls,hexcolor):
        # Check Hex String
        if not hexcolor.startswith('#'):
            raise ValueError(f'Color string given: {hexcolor}: Hex color strings must start with #')
        hexlen = len(hexcolor)
        if hexlen != 7 and hexlen != 9:
            raise ValueError(f'Color string given: {hexcolor}: Hex string must be of the form: "#RRGGBB" or "#RRGGBBAA')
        # Remainder of string should be parsable as hex
        try:
            if int(hexcolor[1:],16):
                pass
        except Exception as e:
            raise ValueError(f'Color string given: {hexcolor}: Bad color string entered')
        # Convert Hex range 0-255 to 0-1.0
        RR = int(hexcolor[1:3],16) / 255
        GG = int(hexcolor[3:5],16) / 255
        BB = int(hexcolor[5:7],16) / 255
        # Slices wont causes IndexErrors
        A = hexcolor[7:9]
        if not A:
            AA = None
        else:
            AA = int(A,16) / 255
        return cls(RR,GG,BB,AA)

    def toHex(self):
        RR = format(int((self.red if self.red else 0) * 255), '02x')
        GG = format(int((self.green if self.green else 0) * 255), '02x')
        BB = format(int((self.blue if self.blue else 0) * 255), '02x')
        AA = format(int((self.alpha if self.alpha else 0) * 255), '02x')

        if self.alpha != None:
            hexformat = f'#{RR}{GG}{BB}{AA}'
        else:
            hexformat = f'#{RR}{GG}{BB}'
        return hexformat

class Border(CellFormatComponent):
    _FIELDS = ('style', 'color', 'colorStyle')

    STYLES = set(['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE'])

    def __init__(self, style, color=None, width=None, colorStyle=None):
        if style.upper() not in Border.STYLES:
            raise ValueError("Border.style must be one of: %s" % Border.STYLES)
        self.style = style.upper()
        self.width = width
        self.color = color
        self.colorStyle = colorStyle

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
    _FIELDS = {
        'foregroundColor': 'color',
        'fontFamily': None,
        'fontSize': None,
        'bold': None,
        'italic': None,
        'strikethrough': None,
        'underline': None,
        'foregroundColorStyle': 'colorStyle'
    }

    def __init__(self,
        foregroundColor=None,
        fontFamily=None,
        fontSize=None,
        bold=None,
        italic=None,
        strikethrough=None,
        underline=None,
        foregroundColorStyle=None
        ):
        self.foregroundColor = foregroundColor
        self.fontFamily = fontFamily
        self.fontSize = fontSize
        self.bold = bold
        self.italic = italic
        self.strikethrough = strikethrough
        self.underline = underline
        self.foregroundColorStyle = foregroundColorStyle

class TextRotation(CellFormatComponent):
    _FIELDS = ('angle', 'vertical')

    def __init__(self, angle=None, vertical=None):
        if len([expr for expr in (angle is not None, vertical is not None) if expr]) != 1:
            raise ValueError("Either angle or vertical must be specified, not both or neither")
        self.angle = angle
        self.vertical = vertical

# provide camelCase aliases for all component classes.

_CLASSES = {}
for _c in [ obj for name, obj in locals().items() if isinstance(obj, type) and issubclass(obj, CellFormatComponent) ]:
    _k = _underlower(_c.__name__)
    _CLASSES[_k] = _c
    locals()[_k] = _c
