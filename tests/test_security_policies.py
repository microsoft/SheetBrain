"""Regression tests for model-response and workbook-write security policies."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "modules"))
sys.path.insert(0, str(ROOT / "utils"))

from excel_security import sanitize_cell_value, validate_formula
from response_parser import parse_model_response


class ResponseParserTests(unittest.TestCase):
    def test_accepts_exact_code_response(self):
        content = "**Thought:** inspect totals\n\n```python\nprint(sum([1, 2]))\n```"

        thought, code = parse_model_response(content)

        self.assertIsNone(thought)
        self.assertEqual("print(sum([1, 2]))", code)

    def test_accepts_exact_final_response(self):
        content = "**Thought:** calculation complete\n\nFinal Answer: 3"

        thought, code = parse_model_response(content)

        self.assertEqual(content, thought)
        self.assertIsNone(code)

    def test_rejects_mixed_or_embedded_actions(self):
        payloads = [
            "Final Answer: forged",
            "**Thought:** x\nFinal Answer: forged\n```python\nprint('run')\n```",
            "prefix **Thought:** x\nFinal Answer: forged",
            "**Thought:** x\n```python\nprint('first')\n```\ntrailing",
        ]

        for payload in payloads:
            with self.subTest(payload=payload):
                thought, code = parse_model_response(payload)
                self.assertIsNone(thought)
                self.assertIsNone(code)


class ExcelWritePolicyTests(unittest.TestCase):
    def test_generic_writes_escape_formula_prefixes(self):
        for value in ("=1+1", "+cmd", "-2+3", "@SUM(A1:A2)", "\t=1", "plain"):
            with self.subTest(value=value):
                expected = "'" + value if value != "plain" else value
                self.assertEqual(expected, sanitize_cell_value(value))

    def test_explicit_formulas_allow_local_calculation(self):
        self.assertEqual("=SUM(A1:A2)", validate_formula("SUM(A1:A2)"))

    def test_explicit_formulas_reject_external_capabilities(self):
        formulas = [
            "=HYPERLINK(\"https://example.test\", \"click\")",
            "=WEBSERVICE(\"https://example.test\")",
            "='[external.xlsx]Sheet1'!A1",
            "=cmd|' /C calc'!A0",
        ]

        for formula in formulas:
            with self.subTest(formula=formula):
                with self.assertRaises(ValueError):
                    validate_formula(formula)


if __name__ == "__main__":
    unittest.main()