# -*- coding: utf-8 -*-

from .util import _props_to_component, _extract_props, _extract_fieldrefs, \
    _parse_string_enum, _underlower, _range_to_gridrange_object
                  
class _field(object):
    def __init__(self, type, values=None, required=False):
        self.type = type
        self.values = values
        self.required = required

    def validate(self, value):
        if value == None:
            if self.required:
                raise ValueError("is a required value")
            else:
                return value
        if not isinstance(value, self.type):
            raise ValueError("must be an instance of %s" % self.type)
        if self.values != None and value not in self.values:
            raise ValueError("must be one of the following: %s" % self.values)
        return value

def _required(type, values=None):
    return _field(type, values, required=True)

def _optional(type, values=None):
    return _field(type, values, required=False)

def _enum(values, required=False):
    return _field(type(next(iter(values))), values, required)


class FormattingComponent(object):
    _FIELDS = {}

    @classmethod
    def from_props(cls, props):
        return _props_to_component(_CLASSES, _underlower(cls.__name__), props)

    def __init__(self, *args, **kwargs):
        combined_kwargs = {}
        field_names = list(self._FIELDS.keys())
        # TODO positional args match either by underlower name or by position in _FIELDS or by underlower name!
        for idx, arg in enumerate(args):
            # TODO below
            # if type is correct (basic type or FormattingComponent type) for positional field in _FIELDS, 
            # it's a match.
            
            # elif a FormattingComponent, and its _underlower name matches a previously unmatched 
            # field in _FIELDS, it's a match.

            # else this argument is invalid.
            if not isinstance(arg, FormattingComponent):
                raise ValueError("Positional argument must be an instance of FormattingComponent, not %r" % arg)
            # its camelCase name must match that of something in _FIELDS
            arg_ul_name = _underlower(arg.__class__.__name__)
            if arg_ul_name not in self._FIELDS:
                raise ValueError("Positional argument (a %s) does not match any of the expected arguments: %s" % (arg_ul_name, self._FIELDS))
            combined_kwargs[arg_ul_name] = arg
        for k, v in kwargs:
            if k not in self._FIELDS:
                raise ValueError(
                    "Keyword argument '%s' does not match any of "
                    "the expected arguments: %s" % (k, self._FIELDS)
                )
            combined_kwargs[k] = v
        for attrname in self._FIELDS:
            # TODO check for requiredness
            # TODO type/value check
            setattr(self, attrname, combined_kwargs.get(attrname))

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
    _FIELDS = {
        'sheetId': _required(int), 
        'startRowIndex': _optional(int), 
        'endRowIndex': _optional(int), 
        'startColumnIndex': _optional(int), 
        'endColumnIndex': _optional(int)
    }

    @classmethod
    def from_a1_range(cls, range, worksheet):
        return GridRange.from_props(_range_to_gridrange_object(range, worksheet.id))

class CellFormatComponent(FormattingComponent):
    pass


class Color(CellFormatComponent):
    _FIELDS = {
        'red': _optional((int, float)), 
        'green': _optional((int, float)), 
        'blue': _optional((int, float)), 
        'alpha': _optional((int, float))
    }

    @classmethod
    def fromHex(cls,hexcolor):
        # Check Hex String
        if not hexcolor.startswith('#'):
            raise ValueError('Color string given: %s: Hex color strings must start with #' % hexcolor)
        hexlen = len(hexcolor)
        if hexlen != 7 and hexlen != 9:
            raise ValueError('Color string given: %s: Hex string must be of the form: "#RRGGBB" or "#RRGGBBAA' % hexcolor)
        # Remainder of string should be parsable as hex
        try:
            if int(hexcolor[1:],16):
                pass
        except Exception as e:
            raise ValueError('Color string given: %s: Bad color string entered' % hexcolor)
        # Convert Hex range 0-255 to 0-1.0
        RR = int(hexcolor[1:3],16) / 255.0
        GG = int(hexcolor[3:5],16) / 255.0
        BB = int(hexcolor[5:7],16) / 255.0
        # Slices wont causes IndexErrors
        A = hexcolor[7:9]
        if not A:
            AA = None
        else:
            AA = int(A,16) / 255.0
        return cls(red=RR,green=GG,blue=BB,alpha=AA)

    def toHex(self):
        RR = format(int((self.red if self.red else 0) * 255), '02x')
        GG = format(int((self.green if self.green else 0) * 255), '02x')
        BB = format(int((self.blue if self.blue else 0) * 255), '02x')
        AA = format(int((self.alpha if self.alpha else 0) * 255), '02x')

        if self.alpha != None:
            hexformat = '#{RR}{GG}{BB}{AA}'.format(RR=RR,GG=GG,BB=BB,AA=AA)
        else:
            hexformat = '#{RR}{GG}{BB}'.format(RR=RR,GG=GG,BB=BB)
        return hexformat

class NumberFormat(CellFormatComponent):
    _FIELDS = {
        'type': _enum(set(['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'])),
        'pattern': _optional(str)
    }


class ColorStyle(CellFormatComponent):
    _FIELDS = {
        'themeColor': _optional(str),
        'rgbColor': _optional(Color)
    }


class Border(CellFormatComponent):
    _FIELDS = {
        'style': _enum(
            set(['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE']), 
            required=True
            ),
        'color': _optional(Color),
        'width': _optional(int),
        'colorStyle': _optional(ColorStyle)
    }

    STYLES = set(['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'NONE', 'DOUBLE'])


class Borders(CellFormatComponent):
    _FIELDS = {
        'top': _optional(Border),
        'bottom': _optional(Border),
        'left': _optional(Border),
        'right': _optional(Border)
    }


class Padding(CellFormatComponent):
    _FIELDS = {
        'top': _optional(int), 
        'right': _optional(int), 
        'bottom': _optional(int), 
        'left': _optional(int)
    }


class TextFormat(CellFormatComponent):
    _FIELDS = {
        'foregroundColor': _optional(Color),
        'fontFamily': _optional(str),
        'fontSize': _optional(int),
        'bold': _optional(bool),
        'italic': _optional(bool),
        'strikethrough': _optional(bool),
        'underline': _optional(bool),
        'foregroundColorStyle': _optional(ColorStyle)
    }


class TextRotation(CellFormatComponent):
    _FIELDS = {
        'angle': _optional(int), 
        'vertical': _optional(bool)
    }


class CellFormat(CellFormatComponent):
    _FIELDS = {
        'numberFormat': _optional(NumberFormat),
        'backgroundColor': _optional(Color),
        'borders': _optional(Borders),
        'padding': _optional(Padding),
        'horizontalAlignment': _enum(set(['LEFT', 'CENTER', 'RIGHT'])),
        'verticalAlignment': _enum(set(['TOP', 'MIDDLE', 'BOTTOM'])),
        'wrapStrategy': _enum(set(['OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP'])),
        'textDirection': _enum(set(['LEFT_TO_RIGHT', 'RIGHT_TO_LEFT'])),
        'textFormat': _optional(str),
        'hyperlinkDisplayType': _enum(set(['LINKED', 'PLAIN_TEXT'])),
        'textRotation': None,
        'foregroundColorStyle': _optional(ColorStyle),
        'backgroundColorStyle': _optional(ColorStyle)
    }


_CLASSES = {}
for _c in [ obj for name, obj in locals().items() if isinstance(obj, type) and issubclass(obj, CellFormatComponent) ]:
    _k = _underlower(_c.__name__)
    _CLASSES[_k] = _c
    locals()[_k] = _c
