import requests
import gspread_formatting.models
from gspread_formatting.util import _underlower

import inspect
import pprint

r = requests.get("https://sheets.googleapis.com/$discovery/rest?version=v4")
j = r.json()

schemas = j['schemas']

classes = gspread_formatting.models._CLASSES

def resolve_schema_property(sch_prop):
    if '$ref' in sch_prop:
        return resolve_schema_property(schemas[sch_prop['$ref']])
    else:
        return sch_prop

def resolve_class_field(fields, field_name):
    if isinstance(fields, dict):
        field_ref = (fields[field_name] or field_name) if (field_name in fields) else None
    else:
        field_ref = field_name if (field_name in fields) else None
    if field_ref is None:
        return None
    else:
        ref_class = classes.get(_underlower(field_ref))
        if ref_class is None:
            return { 'type': 'unknown' }
        else:
            return ref_class

def compare_property(name, sch_prop, cls_prop):
    errors = []
    sch_type = sch_prop['type']
    cls_type = None
    if inspect.isclass(cls_prop):
        cls_type = 'object' 
    elif isinstance(cls_prop, dict):
        cls_type = cls_prop['type']
    if sch_type != cls_type: 
        errors.append( (name, 'schema and class property type differs', '%r != %r' % (sch_type, cls_type)) )
    elif sch_type == 'object':
        errors.extend( compare_object(sch_prop, cls_prop) )
    return errors

def compare_object(schema, cls):
    errors = []
    # 1. names must match
    schema_name = schema['id']
    cls_name = cls.__name__
    if schema_name != cls_name:
        errors.append( (schema_name, 'class name differs', '%r != %r' % (sch_name, cls_name)) )
        
    # 2. report extraneous properties in cls
    for cls_propname in cls._FIELDS:
        if cls_propname not in schema['properties']:
            errors.append( (schema_name, 'extraneous property in class', cls_propname) )

    # 3. report missing properties in cls
    for sch_propname, sch_prop in schema['properties'].items():
        cls_field = resolve_class_field(cls._FIELDS, sch_propname)
        if cls_field is None:
            errors.append( (schema_name, 'property missing in class', sch_propname) )
        # 4. each property must be of correct type
        errors.extend( 
            compare_property("%s.%s" % (schema_name, sch_propname), resolve_schema_property(sch_prop), cls_field) 
        )

    return errors
    

diffs = compare_object(schemas['CellFormat'], classes['cellFormat'])
pprint.pprint(diffs)
