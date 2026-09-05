"""Security and behavior tests for model-generated code interpretation."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules"))

from safe_execution import SafeInterpreter, UnsafeCodeError


class SafeInterpreterTests(unittest.TestCase):
    def setUp(self):
        self.calls = []

        def inspector(range_ref, sheet_name=None):
            self.calls.append((range_ref, sheet_name))
            return [[10, 20], [30, 40]]

        self.interpreter = SafeInterpreter({"inspector": inspector})

    def test_interprets_approved_analysis(self):
        result = self.interpreter.execute(
            "data = inspector('A1:B2', 'Sales')\n"
            "total = 0\n"
            "for row in data:\n"
            "    total += sum(row)\n"
            "print(total)"
        )

        self.assertEqual("Output:\n100", result)
        self.assertEqual([("A1:B2", "Sales")], self.calls)

    def test_preserves_variables_between_turns(self):
        self.interpreter.execute("values = [2, 4, 8]")

        result = self.interpreter.execute("sum(values)")

        self.assertEqual("Expression result: 14", result)

    def test_membership_comparisons(self):
        result = self.interpreter.execute("print(2 in [1, 2], 3 not in [1, 2])")

        self.assertEqual("Output:\nTrue True", result)

    def test_rejects_imports(self):
        with self.assertRaisesRegex(UnsafeCodeError, "Import"):
            self.interpreter.execute("import os")

    def test_rejects_unapproved_filesystem_calls(self):
        with self.assertRaisesRegex(UnsafeCodeError, "approved named functions"):
            self.interpreter.execute("open('stolen.txt', 'w')")

    def test_rejects_dunder_attribute_traversal(self):
        with self.assertRaisesRegex(UnsafeCodeError, "Attribute"):
            self.interpreter.execute("print((1).__class__.__mro__)")

    def test_rejects_indirect_calls(self):
        with self.assertRaisesRegex(UnsafeCodeError, "Name 'inspector' is not available"):
            self.interpreter.execute("function = inspector\nfunction('A1')")

    def test_rejects_additional_escape_syntax(self):
        payloads = [
            "getattr(1, '__class__')",
            "globals()",
            "lambda: 1",
            "[value for value in [1]]",
            "(value := 1)",
            "value = {'a': 1}\n{**value}",
            "while True:\n    pass",
            "with inspector('A1'):\n    pass",
            "raise RuntimeError('stop')",
            "value = 1\ndel value",
        ]

        for payload in payloads:
            with self.subTest(payload=payload):
                with self.assertRaises(UnsafeCodeError):
                    self.interpreter.execute(payload)

    def test_stops_excessive_loop_work(self):
        interpreter = SafeInterpreter({}, max_loop_iterations=3)

        with self.assertRaisesRegex(UnsafeCodeError, "loop iteration limit"):
            interpreter.execute("for number in range(100):\n    print(number)")

    def test_rejects_oversized_sequence_before_repetition(self):
        interpreter = SafeInterpreter({}, max_value_size=100)

        with self.assertRaisesRegex(UnsafeCodeError, "value exceeds"):
            interpreter.execute("'x' * 101")

    def test_rejects_oversized_range_before_materialization(self):
        interpreter = SafeInterpreter({}, max_value_size=100)

        with self.assertRaisesRegex(UnsafeCodeError, "value exceeds"):
            interpreter.execute("list(range(101))")

    def test_bounds_nested_value_rendering(self):
        result = self.interpreter.execute("value = ['x' * 10000] * 10000\nvalue")

        self.assertLess(len(result), 10000)
        self.assertIn("...", result)

    def test_bounds_explicit_string_conversion(self):
        result = self.interpreter.execute("value = ['x' * 10000] * 10000\nstr(value)")

        self.assertLess(len(result), 10000)
        self.assertIn("...", result)

    def test_rejects_string_percent_formatting(self):
        with self.assertRaisesRegex(UnsafeCodeError, "percent formatting"):
            self.interpreter.execute("'%10001s' % 'x'")

    def test_enforces_privileged_function_call_limits(self):
        calls = []
        interpreter = SafeInterpreter(
            {"save_workbook": lambda: calls.append("saved")},
            function_call_limits={"save_workbook": 1},
        )

        with self.assertRaisesRegex(UnsafeCodeError, "call limit"):
            interpreter.execute("for number in range(10):\n    save_workbook()")
        self.assertEqual(["saved"], calls)

    def test_function_call_limits_persist_across_code_turns(self):
        calls = []
        interpreter = SafeInterpreter(
            {"save_workbook": lambda: calls.append("saved")},
            function_call_limits={"save_workbook": 1},
        )

        interpreter.execute("save_workbook()")
        with self.assertRaisesRegex(UnsafeCodeError, "call limit"):
            interpreter.execute("save_workbook()")
        self.assertEqual(["saved"], calls)

        interpreter.reset_session()
        interpreter.execute("save_workbook()")
        self.assertEqual(["saved", "saved"], calls)

    def test_loop_limits_persist_across_code_turns(self):
        interpreter = SafeInterpreter({}, max_loop_iterations=3)

        interpreter.execute("for number in range(2):\n    pass")
        with self.assertRaisesRegex(UnsafeCodeError, "loop iteration limit"):
            interpreter.execute("for number in range(2):\n    pass")

        interpreter.reset_session()
        interpreter.execute("for number in range(2):\n    pass")

    def test_rejects_approved_functions_as_callbacks(self):
        calls = []
        interpreter = SafeInterpreter(
            {"probe": lambda value: calls.append(value) or value},
            function_call_limits={"probe": 1},
        )

        with self.assertRaisesRegex(UnsafeCodeError, "Name 'probe' is not available"):
            interpreter.execute("sorted([3, 2, 1], key=probe)")
        self.assertEqual([], calls)

    def test_validates_entire_program_before_helper_side_effects(self):
        calls = []
        interpreter = SafeInterpreter({"set_cell_value": lambda *args: calls.append(args)})

        with self.assertRaisesRegex(UnsafeCodeError, "Import"):
            interpreter.execute("set_cell_value('Data', 'A1', 'changed')\nimport os")
        self.assertEqual([], calls)

        with self.assertRaisesRegex(UnsafeCodeError, "approved named functions"):
            interpreter.execute("set_cell_value('Data', 'A1', 'changed')\nopen('file.txt')")
        self.assertEqual([], calls)

    def test_materializes_lazy_iterators_before_persistence(self):
        self.interpreter.execute("pairs = zip([1, 2], [3, 4])")

        first_result = self.interpreter.execute("list(pairs)")
        second_result = self.interpreter.execute("list(pairs)")

        self.assertEqual("Expression result: [(1, 3), (2, 4)]", first_result)
        self.assertEqual(first_result, second_result)


if __name__ == "__main__":
    unittest.main()