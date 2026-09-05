# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Strict parser for model actions returned to the execution module."""

import re
from typing import Optional, Tuple


_CODE_RESPONSE = re.compile(
    r"\A\*\*Thought:\*\*\s*(?P<thought>.+?)\s*"
    r"```python\s*\n(?P<code>.*?)\n```\s*\Z",
    re.DOTALL,
)
_FINAL_RESPONSE = re.compile(
    r"\A\*\*Thought:\*\*\s*(?P<thought>.+?)\s*"
    r"Final Answer:\s*(?P<answer>.+?)\s*\Z",
    re.DOTALL,
)


def parse_model_response(content: str) -> Tuple[Optional[str], Optional[str]]:
    """Return a strict final response or one Python code block, never both."""
    has_final_marker = "Final Answer:" in content
    has_code_fence = "```" in content
    if has_final_marker == has_code_fence:
        return None, None
    if content.count("Final Answer:") > 1 or content.count("**Thought:**") != 1:
        return None, None

    final_match = _FINAL_RESPONSE.fullmatch(content)
    if final_match:
        return content.strip(), None

    code_match = _CODE_RESPONSE.fullmatch(content)
    if code_match:
        code = code_match.group("code").strip()
        if code:
            return None, code

    return None, None