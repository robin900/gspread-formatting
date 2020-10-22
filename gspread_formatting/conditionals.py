# -*- coding: utf-8 -*-

from .util import _parse_string_enum, _underlower, _enforce_type
from .models import FormattingComponent, GridRange, _CLASSES

try:
    from collections.abc import MutableSequence, Iterable
except ImportError:
    from collections import MutableSequence, Iterable


def get_conditional_format_rules(worksheet):
    resp = worksheet.spreadsheet.fetch_sheet_metadata()
    rules = []
    for sheet in resp['sheets']:
        if sheet['properties']['sheetId'] == worksheet.id:
            rules = [ ConditionalFormatRule.from_props(p) for p in sheet.get('conditionalFormats', []) ]
            break
    return ConditionalFormatRules(worksheet, rules)

def _make_delete_rule_request(worksheet, rule, ruleIndex):
   return {
       'deleteConditionalFormatRule': {
           'sheetId': worksheet.id,
           'index': ruleIndex
       }
   }

def _make_add_rule_request(worksheet, rule, ruleIndex):
   return {
       'addConditionalFormatRule': {
           'rule': rule.to_props(),
           'index': ruleIndex
       }
   }

class ConditionalFormatRules(MutableSequence):
    def __init__(self, worksheet, *rules):
        self.worksheet = worksheet
        if len(rules) == 1 and isinstance(rules[0], Iterable):
            rules = rules[0]
        self.rules = list(rules)
        self._original_rules = list(rules)

    def __getitem__(self, idx):
        return self.rules[idx]

    def __setitem__(self, idx, value):
        self.rules[idx] = value

    def __delitem__(self, idx):
        del self.rules[idx]

    def __len__(self):
        return len(self.rules)

    # py2.7 MutableSequence does not offer clear()
    def clear(self):
        del self.rules[:]

    def insert(self, idx, value):
        return self.rules.insert(idx, value)

    def save(self):
        # ideally, we would determine the longest "increasing" subsequence
        # between the original and new rule lists, then determine the add/upd/del
        # operations to position the remaining items.
        # But until I implement that correctly, we are just going to delete all rules
        # and re-add them.
        delete_requests = [ 
            _make_delete_rule_request(self.worksheet, r, idx) for idx, r in enumerate(self._original_rules) 
        ]
        # want to delete from highest index to lowest...
        delete_requests.reverse()
        add_requests = [ 
            _make_add_rule_request(self.worksheet, r, idx) for idx, r in enumerate(self.rules) 
        ]
        if not delete_requests and not add_requests:
            return
        body = {
            'requests': delete_requests + add_requests
        }
        resp = self.worksheet.spreadsheet.batch_update(body)
        self._original_rules = list(self.rules)
        return resp


        
class ConditionalFormattingComponent(FormattingComponent):
    pass

class BooleanRule(ConditionalFormattingComponent):
    _FIELDS = {
        'condition': 'booleanCondition', 
        'format': 'cellFormat'
    }

    def __init__(self, condition, format):
        self.condition = condition
        self.format = format

class BooleanCondition(ConditionalFormattingComponent):

    illegal_types_for_data_validation = { 
        'TEXT_STARTS_WITH', 
        'TEXT_ENDS_WITH', 
        'BLANK', 
        'NOT_BLANK' 
    }

    illegal_types_for_conditional_formatting = { 
        'TEXT_IS_EMAIL',
        'TEXT_IS_URL',
        'DATE_ON_OR_BEFORE',
        'DATE_ON_OR_AFTER',
        'DATE_BETWEEN',
        'DATE_NOT_BETWEEN',
        'DATE_IS_VALID',
        'ONE_OF_RANGE',
        'ONE_OF_LIST'
        'BOOLEAN' 
    }

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

    def __init__(self, type, values=()):
        self.type = _parse_string_enum("type", type, BooleanCondition.TYPES)
        validator = BooleanCondition.TYPES[self.type]
        if not isinstance(values, (list, tuple)):
            raise ValueError("values parameter must always be list/tuple of values, even for a single element")
        valid = validator(values) if callable(validator) else len(values) == validator
        if not valid:
            raise ValueError(
                "BooleanCondition.values has inappropriate "
                "length/content for condition type %s" % self.type
            )
        # values are either RelativeDate enum values or user-entered values
        self.values = [ 
            v if isinstance(v, ConditionValue) else (
                ConditionValue.from_props(v) 
                if isinstance(v, dict) 
                else ConditionValue(userEnteredValue=v)
            )
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

    def __init__(self, relativeDate=None, userEnteredValue=None):
        self.relativeDate = relativeDate
        self.userEnteredValue = userEnteredValue

class InterpolationPoint(ConditionalFormattingComponent):
    _FIELDS = ('color', 'colorStyle', 'type', 'value')

    TYPES = set(['MIN', 'MAX', 'NUMBER', 'PERCENT', 'PERCENTILE'])

    def __init__(self, color=None, colorStyle=None, type=None, value=None):
        self.color = color
        self.colorStyle = colorStyle
        self.type = _parse_string_enum("type", type, InterpolationPoint.TYPES, required=True)
        if value is None and self.type not in set(['MIN', 'MAX']):
            raise ValueError("InterpolationPoint.type %s requires a value" % self.type)
        self.value = value

class GradientRule(ConditionalFormattingComponent):
    _FIELDS = {
        'minpoint': 'interpolationPoint', 
        'midpoint': 'interpolationPoint', 
        'maxpoint': 'interpolationPoint'
    }

    def __init__(self, minpoint, maxpoint, midpoint=None):
        self.minpoint = _enforce_type("minpoint", InterpolationPoint, minpoint, required=True)
        self.midpoint = _enforce_type("midpoint", InterpolationPoint, midpoint, required=False)
        self.maxpoint = _enforce_type("maxpoint", InterpolationPoint, maxpoint, required=True)

class ConditionalFormatRule(ConditionalFormattingComponent):
    _FIELDS = ('ranges', 'booleanRule', 'gradientRule')

    def __init__(self, ranges, booleanRule=None, gradientRule=None):
        self.booleanRule = _enforce_type("booleanRule", BooleanRule, booleanRule, required=False)
        if self.booleanRule:
            if self.booleanRule.condition.type in BooleanCondition.illegal_types_for_conditional_formatting:
                raise ValueError(
                    "BooleanCondition.type for conditional formatting must not be one of: %s" % 
                    BooleanCondition.illegal_types_for_conditional_formatting
                )
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
        if self.condition.type in BooleanCondition.illegal_types_for_data_validation:
            raise ValueError(
                "BooleanCondition.type for data validation must not be one of: %s" % 
                BooleanCondition.illegal_types_for_data_validation
            )
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
