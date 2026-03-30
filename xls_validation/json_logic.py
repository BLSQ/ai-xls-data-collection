"""Minimal JsonLogic evaluator for cell-level and row-level validation rules.

Supports the standard operators (``==``, ``!=``, ``>``, ``>=``, ``<``, ``<=``,
``+``, ``-``, ``*``, ``/``, ``!``, ``if``, ``and``, ``or``) plus ``var`` for
resolving column values by canonical name from a data context dict.

Pure functions — no side effects, no OpenHEXA dependency.
"""

from __future__ import annotations


def evaluate_json_logic(rule, data):
    """Evaluate a JsonLogic rule dictionary against a data context.

    Args:
        rule: A JsonLogic rule (dict) or a literal value (passthrough).
        data: A dict mapping variable names (e.g. canonical column names) to
            their values.

    Returns:
        The evaluated result — may be a bool, number, string, or None.

    Raises:
        ValueError: If an unknown operator is encountered.
    """
    if not isinstance(rule, dict):
        return rule

    operator = next(iter(rule))
    raw_arguments = rule[operator]
    if not isinstance(raw_arguments, list):
        raw_arguments = [raw_arguments]

    # -- var: resolve a dotted path in the data context --
    if operator == "var":
        path = raw_arguments[0] if raw_arguments else ""
        default = raw_arguments[1] if len(raw_arguments) > 1 else None
        result = _resolve_variable(data, path)
        return result if result is not None else default

    # -- Short-circuit logical operators --
    if operator in ("and", "or"):
        return _evaluate_short_circuit(operator, raw_arguments, data)

    # -- Eagerly evaluate all arguments for remaining operators --
    arguments = [evaluate_json_logic(arg, data) for arg in raw_arguments]

    if operator == "==":
        return arguments[0] == arguments[1] if len(arguments) >= 2 else False
    if operator == "!=":
        return arguments[0] != arguments[1] if len(arguments) >= 2 else True
    if operator == ">":
        return float(arguments[0]) > float(arguments[1])
    if operator == ">=":
        return float(arguments[0]) >= float(arguments[1])
    if operator == "<":
        return float(arguments[0]) < float(arguments[1])
    if operator == "<=":
        return float(arguments[0]) <= float(arguments[1])
    if operator == "+":
        return sum(float(a) for a in arguments if a is not None)
    if operator == "-":
        if len(arguments) == 1:
            return -float(arguments[0])
        return float(arguments[0]) - float(arguments[1])
    if operator == "*":
        result = 1.0
        for argument in arguments:
            result *= float(argument)
        return result
    if operator == "/":
        return float(arguments[0]) / float(arguments[1])
    if operator == "!":
        return not arguments[0]
    if operator == "if":
        index = 0
        while index < len(arguments) - 1:
            if arguments[index]:
                return arguments[index + 1]
            index += 2
        return arguments[index] if index < len(arguments) else None

    raise ValueError(f"Unknown JsonLogic operator: {operator}")


def _resolve_variable(data, path):
    """Resolve a dotted variable path in a nested dict.

    Args:
        data: The data context dict.
        path: A dot-separated path string (e.g. ``"module.field"``).

    Returns:
        The resolved value, or None if any segment is missing.
    """
    if not path:
        return data
    result = data
    for key in str(path).split("."):
        if isinstance(result, dict):
            result = result.get(key)
        else:
            return None
    return result


def _evaluate_short_circuit(operator: str, raw_arguments: list, data):
    """Evaluate ``and`` / ``or`` with short-circuit semantics.

    Args:
        operator: ``"and"`` or ``"or"``.
        raw_arguments: The unevaluated argument list.
        data: The data context dict.

    Returns:
        The last truthy/falsy value per short-circuit rules.
    """
    if operator == "and":
        result = True
        for argument in raw_arguments:
            result = evaluate_json_logic(argument, data)
            if not result:
                return result
        return result
    else:
        result = False
        for argument in raw_arguments:
            result = evaluate_json_logic(argument, data)
            if result:
                return result
        return result
