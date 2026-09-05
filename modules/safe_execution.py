# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Deny-by-default interpreter for model-generated spreadsheet operations."""

import ast
import itertools
import operator
import reprlib
from collections.abc import Iterator
from typing import Any, Callable, Dict, Iterable, Optional


SAFE_BUILTINS = {
    "abs": abs,
    "all": all,
    "any": any,
    "bool": bool,
    "dict": dict,
    "enumerate": enumerate,
    "float": float,
    "int": int,
    "len": len,
    "list": list,
    "max": max,
    "min": min,
    "range": range,
    "round": round,
    "set": set,
    "sorted": sorted,
    "sum": sum,
    "tuple": tuple,
    "zip": zip,
}


class UnsafeCodeError(ValueError):
    """Raised when generated code uses syntax outside the approved language."""


class SafeInterpreter:
    """Interpret a small Python-like language without invoking Python execution APIs."""

    _ALLOWED_NODE_TYPES = (
        ast.Module, ast.Expr, ast.Assign, ast.AugAssign, ast.If, ast.For, ast.Pass,
        ast.Constant, ast.Name, ast.Load, ast.Store, ast.List, ast.Tuple, ast.Set,
        ast.Dict, ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare, ast.IfExp,
        ast.Subscript, ast.Slice, ast.Call, ast.keyword,
        ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod,
        ast.UAdd, ast.USub, ast.Not, ast.And, ast.Or,
        ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
        ast.In, ast.NotIn, ast.Is, ast.IsNot,
    )

    _BINARY_OPERATORS = {
        ast.Add: operator.add,
        ast.Sub: operator.sub,
        ast.Mult: operator.mul,
        ast.Div: operator.truediv,
        ast.FloorDiv: operator.floordiv,
        ast.Mod: operator.mod,
    }
    _UNARY_OPERATORS = {
        ast.UAdd: operator.pos,
        ast.USub: operator.neg,
        ast.Not: operator.not_,
    }
    _COMPARISON_OPERATORS = {
        ast.Eq: operator.eq,
        ast.NotEq: operator.ne,
        ast.Lt: operator.lt,
        ast.LtE: operator.le,
        ast.Gt: operator.gt,
        ast.GtE: operator.ge,
        ast.In: lambda item, container: operator.contains(container, item),
        ast.NotIn: lambda item, container: not operator.contains(container, item),
        ast.Is: operator.is_,
        ast.IsNot: operator.is_not,
    }

    def __init__(self, functions: Dict[str, Callable[..., Any]], max_source_length: int = 20000,
                 max_nodes: int = 1000, max_loop_iterations: int = 10000,
                 max_value_size: int = 10000, max_integer_bits: int = 4096,
                 function_call_limits: Optional[Dict[str, int]] = None):
        self.functions = {**SAFE_BUILTINS, **functions}
        self.functions["print"] = self._print
        self.max_source_length = max_source_length
        self.max_nodes = max_nodes
        self.max_loop_iterations = max_loop_iterations
        self.max_value_size = max_value_size
        self.max_integer_bits = max_integer_bits
        self.function_call_limits = dict(function_call_limits or {})
        self.variables: Dict[str, Any] = {}
        self.output = []
        self._loop_iterations = 0
        self._function_calls: Dict[str, int] = {}
        self._repr = reprlib.Repr()
        self._repr.maxlevel = 2
        self._repr.maxstring = 200
        self._repr.maxother = 200
        self._repr.maxlist = 20
        self._repr.maxtuple = 20
        self._repr.maxset = 20
        self._repr.maxfrozenset = 20
        self._repr.maxdict = 20
        self.functions["str"] = self._display

    def reset_session(self) -> None:
        """Reset variables and cumulative helper quotas for a new analysis."""
        self.variables = {}
        self._function_calls = {}
        self._loop_iterations = 0

    def execute(self, source: str) -> str:
        """Validate and interpret source, returning captured textual output."""
        if len(source) > self.max_source_length:
            raise UnsafeCodeError("Generated code exceeds the allowed size")

        try:
            tree = ast.parse(source, mode="exec")
        except SyntaxError as error:
            raise UnsafeCodeError(f"Invalid generated code: {error.msg}") from error

        if sum(1 for _ in ast.walk(tree)) > self.max_nodes:
            raise UnsafeCodeError("Generated code is too complex")
        self._validate_policy(tree)

        self.output = []
        last_value = None
        for statement in tree.body:
            last_value = self._execute_statement(statement)

        if self.output:
            return "Output:\n" + "\n".join(self.output)
        if last_value is not None:
            self._validate_value(last_value)
            return f"Expression result: {self._display(last_value)}"
        return "Code executed successfully (no output)"

    def _validate_policy(self, tree: ast.Module) -> None:
        direct_callees = {
            id(node.func) for node in ast.walk(tree)
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
        }
        for node in ast.walk(tree):
            if not isinstance(node, self._ALLOWED_NODE_TYPES):
                raise UnsafeCodeError(f"Syntax '{type(node).__name__}' is not allowed")
            if isinstance(node, ast.Assign):
                if len(node.targets) != 1:
                    raise UnsafeCodeError("Chained assignment is not allowed")
                self._validate_assignment_target(node.targets[0])
            elif isinstance(node, ast.AugAssign):
                self._validate_assignment_target(node.target)
            elif isinstance(node, ast.For):
                self._validate_assignment_target(node.target)
            elif isinstance(node, ast.Dict) and any(key is None for key in node.keys):
                raise UnsafeCodeError("Dictionary unpacking is not allowed")
            elif isinstance(node, ast.Call):
                if not isinstance(node.func, ast.Name) or node.func.id not in self.functions:
                    raise UnsafeCodeError("Only approved named functions may be called")
                if any(keyword.arg is None for keyword in node.keywords):
                    raise UnsafeCodeError("Expanded keyword arguments are not allowed")
            elif (isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load)
                  and node.id in self.functions and id(node) not in direct_callees):
                raise UnsafeCodeError(f"Name '{node.id}' is not available")

    def _validate_assignment_target(self, target: ast.expr) -> None:
        if isinstance(target, ast.Name):
            if target.id.startswith("_") or target.id in self.functions:
                raise UnsafeCodeError(f"Assignment to '{target.id}' is not allowed")
            return
        if isinstance(target, (ast.Tuple, ast.List)):
            for item in target.elts:
                self._validate_assignment_target(item)
            return
        raise UnsafeCodeError("Only variable and unpacking assignments are allowed")

    def _execute_statement(self, node: ast.stmt) -> Any:
        if isinstance(node, ast.Expr):
            return self._evaluate(node.value)
        if isinstance(node, ast.Assign):
            if len(node.targets) != 1:
                raise UnsafeCodeError("Chained assignment is not allowed")
            value = self._evaluate(node.value)
            self._assign(node.targets[0], value)
            return None
        if isinstance(node, ast.AugAssign):
            current = self._evaluate(node.target)
            operation = self._get_operator(node.op, self._BINARY_OPERATORS)
            self._assign(node.target, operation(current, self._evaluate(node.value)))
            return None
        if isinstance(node, ast.If):
            branch = node.body if self._evaluate(node.test) else node.orelse
            for statement in branch:
                self._execute_statement(statement)
            return None
        if isinstance(node, ast.For):
            iterable = self._evaluate(node.iter)
            for item in self._bounded_iterable(iterable):
                self._assign(node.target, item)
                for statement in node.body:
                    self._execute_statement(statement)
            for statement in node.orelse:
                self._execute_statement(statement)
            return None
        if isinstance(node, ast.Pass):
            return None
        raise UnsafeCodeError(f"Statement '{type(node).__name__}' is not allowed")

    def _evaluate(self, node: ast.expr) -> Any:
        if isinstance(node, ast.Constant):
            return self._validate_value(node.value)
        if isinstance(node, ast.Name):
            if node.id in self.variables:
                return self.variables[node.id]
            raise UnsafeCodeError(f"Name '{node.id}' is not available")
        if isinstance(node, ast.List):
            return self._validate_value([self._evaluate(item) for item in node.elts])
        if isinstance(node, ast.Tuple):
            return self._validate_value(tuple(self._evaluate(item) for item in node.elts))
        if isinstance(node, ast.Set):
            return self._validate_value({self._evaluate(item) for item in node.elts})
        if isinstance(node, ast.Dict):
            result = {}
            for key, value in zip(node.keys, node.values):
                if key is None:
                    raise UnsafeCodeError("Dictionary unpacking is not allowed")
                result[self._evaluate(key)] = self._evaluate(value)
            return self._validate_value(result)
        if isinstance(node, ast.BinOp):
            left = self._evaluate(node.left)
            right = self._evaluate(node.right)
            self._validate_binary_operation(node.op, left, right)
            operation = self._get_operator(node.op, self._BINARY_OPERATORS)
            return self._validate_value(operation(left, right))
        if isinstance(node, ast.UnaryOp):
            operation = self._get_operator(node.op, self._UNARY_OPERATORS)
            return self._validate_value(operation(self._evaluate(node.operand)))
        if isinstance(node, ast.BoolOp):
            return self._evaluate_boolean(node)
        if isinstance(node, ast.Compare):
            return self._evaluate_comparison(node)
        if isinstance(node, ast.IfExp):
            return self._evaluate(node.body if self._evaluate(node.test) else node.orelse)
        if isinstance(node, ast.Subscript):
            return self._validate_value(
                self._evaluate(node.value)[self._evaluate_slice(node.slice)]
            )
        if isinstance(node, ast.Call):
            return self._call(node)
        raise UnsafeCodeError(f"Expression '{type(node).__name__}' is not allowed")

    def _print(self, *values: Any, sep: str = " ", end: str = "\n") -> None:
        if not isinstance(sep, str) or not isinstance(end, str):
            raise UnsafeCodeError("print separators must be strings")
        text = sep.join(self._display(value) for value in values) + end
        if sum(len(line) for line in self.output) + len(text) > 10000:
            raise UnsafeCodeError("Generated code exceeded the output limit")
        self.output.append(text.rstrip("\n"))

    def _display(self, value: Any) -> str:
        if isinstance(value, str):
            rendered = value
        else:
            rendered = self._repr.repr(value)
        display_limit = min(self.max_value_size, 8000)
        if len(rendered) > display_limit:
            return rendered[:display_limit - 3] + "..."
        return rendered

    def _call(self, node: ast.Call) -> Any:
        if not isinstance(node.func, ast.Name) or node.func.id not in self.functions:
            raise UnsafeCodeError("Only approved named functions may be called")
        function_name = node.func.id
        call_count = self._function_calls.get(function_name, 0) + 1
        call_limit = self.function_call_limits.get(function_name)
        if call_limit is not None and call_count > call_limit:
            raise UnsafeCodeError(f"Function '{function_name}' exceeded its call limit")
        self._function_calls[function_name] = call_count
        arguments = [self._evaluate(argument) for argument in node.args]
        keywords: Dict[str, Any] = {}
        for keyword in node.keywords:
            if keyword.arg is None:
                raise UnsafeCodeError("Expanded keyword arguments are not allowed")
            keywords[keyword.arg] = self._evaluate(keyword.value)
        for value in arguments:
            self._validate_value(value)
        for value in keywords.values():
            self._validate_value(value)
        return self._validate_value(self.functions[function_name](*arguments, **keywords))

    def _validate_binary_operation(self, operation: ast.operator, left: Any, right: Any) -> None:
        if isinstance(operation, ast.Mod) and isinstance(left, (str, bytes, bytearray)):
            raise UnsafeCodeError("String percent formatting is not allowed")
        if not isinstance(operation, ast.Mult):
            return
        sequence_types = (str, bytes, bytearray, list, tuple)
        if isinstance(left, sequence_types) and isinstance(right, int):
            requested_size = len(left) * max(right, 0)
        elif isinstance(right, sequence_types) and isinstance(left, int):
            requested_size = len(right) * max(left, 0)
        else:
            return
        if requested_size > self.max_value_size:
            raise UnsafeCodeError("Generated value exceeds the allowed size")

    def _validate_value(self, value: Any) -> Any:
        if isinstance(value, Iterator):
            materialized = list(itertools.islice(value, self.max_value_size + 1))
            if len(materialized) > self.max_value_size:
                raise UnsafeCodeError("Generated iterator exceeds the allowed size")
            value = materialized
        if isinstance(value, int) and value.bit_length() > self.max_integer_bits:
            raise UnsafeCodeError("Generated integer exceeds the allowed size")
        sized_types = (str, bytes, bytearray, range, list, tuple, set, frozenset, dict)
        if isinstance(value, sized_types) and len(value) > self.max_value_size:
            raise UnsafeCodeError("Generated value exceeds the allowed size")
        return value

    def _assign(self, target: ast.expr, value: Any) -> None:
        self._validate_assignment_target(target)
        if isinstance(target, ast.Name):
            self.variables[target.id] = value
            return
        if isinstance(target, (ast.Tuple, ast.List)):
            values = list(value)
            if len(target.elts) != len(values):
                raise UnsafeCodeError("Assignment target and value lengths differ")
            for item_target, item_value in zip(target.elts, values):
                self._assign(item_target, item_value)
            return

    def _bounded_iterable(self, value: Any) -> Iterable[Any]:
        try:
            iterator = iter(value)
        except TypeError as error:
            raise UnsafeCodeError("For-loop value is not iterable") from error
        for item in iterator:
            self._loop_iterations += 1
            if self._loop_iterations > self.max_loop_iterations:
                raise UnsafeCodeError("Generated code exceeded the loop iteration limit")
            yield item

    def _evaluate_boolean(self, node: ast.BoolOp) -> Any:
        if isinstance(node.op, ast.And):
            result = True
            for value in node.values:
                result = self._evaluate(value)
                if not result:
                    return result
            return result
        if isinstance(node.op, ast.Or):
            result = False
            for value in node.values:
                result = self._evaluate(value)
                if result:
                    return result
            return result
        raise UnsafeCodeError(f"Boolean operator '{type(node.op).__name__}' is not allowed")

    def _evaluate_comparison(self, node: ast.Compare) -> bool:
        left = self._evaluate(node.left)
        for operation_node, comparator in zip(node.ops, node.comparators):
            right = self._evaluate(comparator)
            operation = self._get_operator(operation_node, self._COMPARISON_OPERATORS)
            if not operation(left, right):
                return False
            left = right
        return True

    def _evaluate_slice(self, node: ast.expr) -> Any:
        if isinstance(node, ast.Slice):
            return slice(
                self._evaluate(node.lower) if node.lower else None,
                self._evaluate(node.upper) if node.upper else None,
                self._evaluate(node.step) if node.step else None,
            )
        return self._evaluate(node)

    @staticmethod
    def _get_operator(node: ast.AST, operators: Dict[type, Callable[..., Any]]) -> Callable[..., Any]:
        operation = operators.get(type(node))
        if operation is None:
            raise UnsafeCodeError(f"Operator '{type(node).__name__}' is not allowed")
        return operation