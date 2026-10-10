"""Validate input records only. All scoring is performed by Excel formulas."""
from decimal import Decimal, InvalidOperation

class InputValidationError(ValueError):pass

def number(value):
    if isinstance(value,bool):raise InputValidationError('布尔值不能作为数值测量')
    try:result=Decimal(str(value))
    except (InvalidOperation,ValueError) as exc:raise InputValidationError('输入必须是有效数字') from exc
    if not result.is_finite():raise InputValidationError('输入必须是有限数字')
    return result

def validate_facts(spec,facts):
    if not isinstance(facts,dict):raise InputValidationError('输入记录必须是 JSON 对象')
    definitions={item['id']:item for item in spec['scoring']['inputs']}
    extra=set(facts)-set(definitions)
    if extra:raise InputValidationError('包含未声明字段：'+', '.join(sorted(extra)))
    clean={}
    for key,definition in definitions.items():
        value=facts.get(key)
        if value is None:
            if definition.get('required',True):raise InputValidationError('缺少必填字段：'+key)
            continue
        kind=definition.get('type','number')
        if kind=='number':
            value=number(value)
            if definition.get('minimum') is not None and value<number(definition['minimum']):raise InputValidationError(key+' 小于允许范围')
            if definition.get('maximum') is not None and value>number(definition['maximum']):raise InputValidationError(key+' 大于允许范围')
        elif kind=='boolean':
            if not isinstance(value,bool):raise InputValidationError(key+' 必须是 true/false')
        elif kind=='string':
            if not isinstance(value,str) or (definition.get('enum') and value not in definition['enum']):raise InputValidationError(key+' 文本不在允许选项中')
        else:raise InputValidationError(key+' 的输入类型不受支持')
        clean[key]=value
    return clean
