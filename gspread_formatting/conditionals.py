# -*- coding: utf-8 -*-

from .util import _parse_string_enum, _underlower, _enforce_type
from .models import FormattingComponent, GridRange, _CLASSES

class ConditionalFormattingComponent(FormattingComponent):
    pass

class BooleanRule(ConditionalFormattingComponent):
    _FIELDS = ('condition', 'format')

    def __init__(self, condition, format):
        self.condition = condition
        self.format = format

class BooleanCondition(ConditionalFormattingComponent):
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
        self.type = _parse_string_enum("type", type, BooleanCondition.TYPES)
        validator = BooleanCondition.TYPES[self.type]
        valid = validator(values) if callable(validator) else len(values) == validator
        if not valid:
            raise ValueError(
                "BooleanCondition.values has inappropriate "
                "length/content for condition type %s" % self.type
            )
        # values are either RelativeDate enum values or user-entered values
        self.values = [ 
            v if isinstance(v, ConditionValue) else ConditionValue.from_props(v) 
            for v in values 
        ]

    def to_props(self):
        return {
            'type': self.type,
            'values': [ v.to_props() for v in self.values ]
        }

class RelativeDate(FormattingComponent):
    VALUES = set(['PAST_YEAR', 'PAST_MONTH', 'PAST_WEEK', 'YESTERDAY', 'TODAY', 'TOMORROW'])

    def __init__(self, value):
        if value.upper() not in RelativeDate.VALUES:
            raise ValueError("RelativeDate must be one of: %s" % RelativeDate.VALUES)
        self.value = _parse_string_enum("value", value, RelativeDate.VALUES)

    def to_props(self):
        return self.value

class ConditionValue(ConditionalFormattingComponent):
    _FIELDS = ('relativeDate', 'userEnteredValue')

    def __init__(self, value):
        self.relativeDate = None
        self.userEnteredValue = None
        if isinstance(value, dict):
            if 'relativeDate' in value:
                value = RelativeDate(value['relativeDate'])
            elif 'userEnteredValue' in value:
                value = value['userEnteredValue']
            else:
                raise ValueError(
                    "ConditionValue must be either "
                    "relativeDate or userEnteredValue, not: %s" % value
                )
        if isinstance(value, RelativeDate):
            self.relativeDate = value.value
        else:
            self.userEnteredValue = value

class InterpolationPoint(ConditionalFormattingComponent):
    _FIELDS = ('color', 'type', 'value')

    TYPES = set(['MIN', 'MAX', 'NUMBER', 'PERCENT', 'PERCENTILE'])

    def __init__(self, color, type, value=None):
        self.color = color
        self.type = _parse_string_enum("type", type, InterpolationPoint.TYPES, required=True)
        if value is None and self.type not in set(['MIN', 'MAX']):
            raise ValueError("InterpolationPoint.type %s requires a value" % self.type)
        self.value = value

class GradientRule(ConditionalFormattingComponent):
    _FIELDS = ('minpoint', 'midpoint', 'maxpoint')

    def __init__(self, minpoint, midpoint, maxpoint):
        self.minpoint = minpoint
        self.midpoint = midpoint
        self.maxpoint = maxpoint

class ConditionalFormatRule(ConditionalFormattingComponent):
    _FIELDS = ('ranges', 'booleanRule', 'gradientRule')

    def __init__(self, ranges, booleanRule=None, gradientRule=None):
        self.booleanRule = _enforce_type("booleanRule", BooleanRule, booleanRule, required=False)
        self.gradientRule = _enforce_type("gradientRule", GradientRule, gradientRule, required=False)
        if len([x for x in (self.booleanRule, self.gradientRule) if x is not None]) != 1:
            raise ValueError("Must specify exactly one of: booleanRule, gradientRule")
        # values are either GridRange objects or bare properties
        self.ranges = [ 
            v if isinstance(v, GridRange) else GridRange.from_props(v) 
            for v in ranges 
        ]

    def to_props(self):
        p = {
            'ranges': [ v.to_props() for v in self.ranges ]
        }
        if self.booleanRule:
            p['booleanRule'] = self.booleanRule.to_props()
        if self.gradientRule:
            p['gradientRule'] = self.gradientRule.to_props()
        return p

class DataValidationRule(FormattingComponent):
    _FIELDS = {
        'condition': 'booleanCondition', 
        'inputMessage': str, 
        'strict': bool, 
        'showCustomUi': bool
    }

    def __init__(self, condition, inputMessage=None, strict=None, showCustomUi=None):
        self.condition = _enforce_type("condition", BooleanCondition, condition, True)
        self.inputMessage = _enforce_type("inputMessage", str, inputMessage, False)
        self.strict = _enforce_type("strict", bool, strict, False)
        self.showCustomUi = _enforce_type("showCustomUi", bool, showCustomUi, False)

# provide camelCase aliases for all component classes.

for _c in [ 
        obj for name, obj in locals().items() 
        if isinstance(obj, type) 
        and issubclass(obj, FormattingComponent) 
    ]:
    _k = _underlower(_c.__name__)
    _CLASSES[_k] = _c
    locals()[_k] = _c
