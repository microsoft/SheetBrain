# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Execution module for multi-turn reasoning and constrained code interpretation."""

import re
import time
from typing import Dict, Any, Optional, Tuple

from openai import RateLimitError
from openai.types.chat import ChatCompletionMessage

from modules.response_parser import parse_model_response
from modules.safe_execution import SafeInterpreter
from utils.logger import setup_logger

logger = setup_logger(__name__)

ALLOWED_EXCEL_HELPERS = {
    "inspector",
    "inspector_attribute",
    "search",
    "save_workbook",
    "insert_rows",
    "insert_columns",
    "delete_rows",
    "delete_columns",
    "set_cell_value",
    "set_range_values",
    "copy_range",
    "apply_formatting",
    "create_chart",
    "add_formula",
}

HELPER_CALL_LIMITS = {
    "inspector": 20,
    "inspector_attribute": 20,
    "search": 5,
    "save_workbook": 1,
    "insert_rows": 20,
    "insert_columns": 20,
    "delete_rows": 20,
    "delete_columns": 20,
    "set_cell_value": 100,
    "set_range_values": 20,
    "copy_range": 20,
    "apply_formatting": 20,
    "create_chart": 20,
    "add_formula": 100,
}


class ExecutionModule:
    """
    Module responsible for multi-turn reasoning and code execution based on understanding context.
    Handles its own conversation flow internally and returns the final result.
    """

    def __init__(self, client, deployment: str, code_globals: dict, code_locals: dict,
                 excel_context_execution: str):
        """
        Initialize the ExecutionModule.

        Args:
            client: OpenAI client instance
            deployment: Model deployment name
            code_globals: Global variables for code execution
            code_locals: Local variables for code execution
            excel_context_execution: Excel context for execution
        """
        self.client = client
        self.deployment = deployment
        self.code_globals = code_globals
        self.code_locals = code_locals
        self.excel_context_execution = excel_context_execution
        self.conversation_history: list[Any] = []
        approved_helpers = {
            name: value for name, value in code_globals.items()
            if name in ALLOWED_EXCEL_HELPERS and callable(value)
        }
        self.interpreter = SafeInterpreter(
            approved_helpers,
            function_call_limits=HELPER_CALL_LIMITS,
        )

    def _get_system_prompt(self) -> dict:
        """Create the system prompt for the conversation."""

        system_content = """You are an expert Excel data analyst with access to a constrained language for Excel analysis.

**CODE EXECUTION ENVIRONMENT:**
    Code is interpreted by a deny-by-default engine. It supports literals, arithmetic, comparisons,
    variables, if statements, bounded for loops, indexing, and calls to the named functions below.
    Imports, attribute access, function/class definitions, comprehensions, while loops, file/network/process
    access, and calls to any unlisted function are rejected. Use lists and dictionaries for analysis.

    Available pure functions: `abs`, `all`, `any`, `bool`, `dict`, `enumerate`, `float`, `int`, `len`,
    `list`, `max`, `min`, `print`, `range`, `round`, `set`, `sorted`, `str`, `sum`, `tuple`, `zip`.

Available Excel Helper Functions:
- `inspector(range_ref, sheet_name=None)`: Read cell values from specified range
  - **Usage:** `data = inspector("A1:C3", "Sheet1")` or `value = inspector("B5")`
  - **Output:** List of lists format: `[['A1', 'B1', 'C1'], ['A2', 'B2', 'C2']]` or `[['single_value']]`

- `inspector_attribute(range_ref, attributes, sheet_name=None)`: Extract cell formatting and properties
  - **Usage:** `attrs = inspector_attribute("A1:B2", ["color", "font"], "Sheet1")`
  - **Attributes:** `["color", "font", "formula"]` - specify which properties to extract
  - **Output:** Dict with structure: `{"range": "A1:B2", "sheet": "Sheet1", "attributes": {"color": {"A1": "#FF0000"}, "font": {"B2": "name:Arial; size:12; bold:True"}}}`

- `search(value, sheet_name=None, case_sensitive=False, search_type='partial')`: Find cells containing specific values
  - **Usage:** `matches = search("Total", case_sensitive=True, search_type="whole")`
  - **Search types:** `"partial"` (default), `"whole"`, `"strip"`
  - **Output:** List of dicts: `[{"coordinate": "A5", "value": "Total Sales", "row": 5, "column": 1}]`

- `apply_formatting(sheet_name, range_ref, format_dict)`: Apply cell formatting (colors, fonts, borders)
  - **Usage:** `result = apply_formatting("Sheet1", "A1:C5", {"fill_color": "#FF0000", "bold": True})`
  - **Format Options:**
    - `fill_color`: Background color (hex: '#FF0000' or name: 'red')
    - `font_color`: Font color (hex: '#FF0000' or name: 'red')
    - `font_size`: Font size (int)
    - `font_name`: Font name (str)
    - `bold`: Bold text (bool)
    - `italic`: Italic text (bool)
    - `underline`: Underline text (bool)
    - `border`: Border style ('thin', 'medium', 'thick')
    - `alignment`: Text alignment ('left', 'center', 'right')
  - **Output:** String message confirming formatting applied to specified range

- `save_workbook()`: Save workbook to file with '_output' postfix
  - **Usage:** `filename = save_workbook()`
  - **Output:** Returns saved filename string: `"/path/to/original_output.xlsx"` and prints confirmation message

**RESPONSE FORMATS - MANDATORY COMPLIANCE:**

SYSTEM CONSTRAINT: Your response must contain EXACTLY one of these formats and NOTHING ELSE:

FORMAT A - Thinking + Code execution:
**Thought:** [Your reasoning and analysis here]

```python
# Your Python code here
```

FORMAT B - Thinking + Final answer:
**Thought:** [Your reasoning and analysis here]

Final Answer: [Your conclusive answer]

CRITICAL REQUIREMENTS:
- ALWAYS start with **Thought:** to explain your reasoning
- Follow with EITHER code execution OR final answer
- NO additional text, explanation, or commentary outside these formats
- NO preamble, postamble, or "how it works" sections
- VIOLATION WILL RESULT IN TASK FAILURE

**CRITICAL DECISION FRAMEWORK - When to use Code vs. Direct Analysis:**

** USE DIRECT ANALYSIS (Give Final Answer immediately) when:**
- The Sheet Content preview shows ALL necessary data for the question
- Simple calculations can be performed mentally from visible data
- The question asks for values that are directly visible in the preview
- Table structure is clear and hierarchical relationships are evident
- No complex aggregations, transformations, or editing operations are needed
- Data relationships (parent-child, subtotals) are obvious from the preview

** USE CODE when:**
- Data extends beyond what's shown in the preview
- Complex calculations, aggregations, or statistical analysis is required
- Data transformation, filtering, or manipulation is needed
- Need to edit/modify the Excel file
- Need to search across large datasets
- Verification of calculations through code is specifically requested

**IMPORTANT GUIDELINES:**
- **NO REDUNDANT CODE**: Don't write code to print data that's already visible in the Sheet Content
- Print intermediate results to show your thought process
- Use the helper functions for common operations
- **Identify hierarchical relationships** (e.g., "of which", "including", indented items)
- Use `save_workbook()` to save changes
- **ALWAYS call `save_workbook()` after making ANY changes to the Excel file**

### Multi-Table in One Sheet – Instructions
1. **Detect Multiple Tables**
   Recognize that a single sheet may contain several distinct tables separated by blank rows/columns or different header areas.
2. **Identify Boundaries**
   Clearly define the start and end range of each table to avoid mixing data.
3. **Check Relationships**
   Analyze whether tables are logically connected (e.g., raw data vs. summary, detail vs. KPIs).
4. **Follow Query Focus**
   If the query mentions multiple tables, address each one explicitly and compare where relevant.

### Complex Table – Instructions
1. **Identify Hierarchies**
   Detect multi-row column headers (top headers) and multi-level row headers (left headers) as hierarchical structures.
2. **Preserve Header Levels**
   Keep parent–child relationships intact when analyzing (e.g., Region → Product → Sales).
3. **Handle Subtotals**
   Recognize subtotal and total rows/columns, and clarify "of which" or aggregation relationships.
4. **Explain Hierarchy in Results**
   Clearly state how each level contributes to subtotals/totals in your explanation.

Start by exploring the data structure to understand what you're working with."""

        return {"role": "system", "content": system_content}

    def _create_initial_user_prompt(self, understanding_output: str, user_question: str) -> dict:
        """Create the initial user prompt for the conversation."""

        user_content = f"""**Sheet Content:**
{self.excel_context_execution}

**Understanding Context:**
{understanding_output}

**USER QUESTION:**
{user_question}

Please start by exploring the data structure and then work toward answering the question step by step.
"""

        return {"role": "user", "content": user_content}

    def run(self, understanding_output: str, user_question: str, max_turns: int = 20) -> Dict[str, Any]:
        """
        Run the execution module with understanding context and user question.

        Args:
            understanding_output: Output from UnderstandingModule
            user_question: Original user question
            max_turns: Maximum number of conversation turns

        Returns:
            Dictionary containing execution results and conversation history
        """
        logger.info("Starting multi-turn analysis")
        self.interpreter.reset_session()

        # Initialize conversation with system prompt and initial user prompt
        self.conversation_history = [self._get_system_prompt()]
        initial_prompt = self._create_initial_user_prompt(understanding_output, user_question)
        self.conversation_history.append(initial_prompt)

        execution_steps = []  # Track key execution steps

        for turn in range(max_turns):
            logger.info(f"Execution turn {turn + 1}")

            try:
                response_message = self._get_llm_response()
                self.conversation_history.append(response_message)

                # Parse response for code action or final answer
                if not isinstance(response_message.content, str):
                    raise ValueError("LLM response did not contain text")
                thought, code_action = self._parse_llm_response(response_message.content)

                if code_action is None:
                    # No code to execute, check if it's a final answer
                    if thought and "Final Answer:" in thought:
                        # Extract the final answer from the content
                        final_answer_match = re.search(r"Final Answer:\s*(.*?)$", thought, re.DOTALL)
                        if final_answer_match:
                            final_answer = final_answer_match.group(1).strip()
                        else:
                            final_answer = thought.replace("Final Answer:", "").strip()

                        logger.info("Final answer found")

                        return {
                            "success": True,
                            "answer": final_answer,
                            "total_turns": turn + 1,
                            "conversation_history": self._format_conversation_history(),
                            "execution_summary": self._generate_execution_summary(execution_steps, final_answer)
                        }
                    else:
                        # No valid action found, ask for clarification
                        logger.warning("No valid action found, asking for clarification")
                        reminder = (
                            "CRITICAL FORMAT VIOLATION: You must respond in EXACTLY one of these formats:\n\n"
                            "FORMAT A - Thinking + Code:\n"
                            "**Thought:** [Your reasoning here]\n\n"
                            "```python\n# Your code here\n```\n\n"
                            "FORMAT B - Thinking + Final Answer:\n"
                            "**Thought:** [Your reasoning here]\n\n"
                            "Final Answer: Your answer here\n\n"
                            "NO other text is allowed. Start with **Thought:** ALWAYS."
                        )
                        self.conversation_history.append({"role": "user", "content": reminder})
                        continue

                # Execute code action
                logger.info("Interpreting model-generated spreadsheet operation")

                try:
                    execution_result = self._execute_code(code_action)
                    observation = f"Code execution result:\n{execution_result}"
                    logger.info("Spreadsheet operation completed")

                    # Track this execution step
                    execution_steps.append({
                        "turn": turn + 1,
                        "code": code_action,
                        "result": execution_result,
                        "success": True
                    })

                    self.conversation_history.append({"role": "user", "content": observation})

                except Exception as e:
                    error_message = f"Code execution error: {type(e).__name__}"
                    logger.error("Model-generated spreadsheet operation failed: %s", type(e).__name__)

                    # Track this failed execution step
                    execution_steps.append({
                        "turn": turn + 1,
                        "code": code_action,
                        "result": error_message,
                        "success": False
                    })

                    self.conversation_history.append({"role": "user", "content": error_message})

            except Exception as e:
                logger.error(f"LLM Error: {str(e)}")
                return {
                    "success": False,
                    "answer": f"LLM communication error: {str(e)}",
                    "total_turns": turn + 1,
                    "conversation_history": self._format_conversation_history(),
                    "execution_summary": self._generate_execution_summary(execution_steps, None)
                }

        # Reached maximum turns without final answer
        logger.warning("Reached maximum turns without finding final answer")
        return {
            "success": False,
            "answer": "Unable to find a complete answer within the maximum number of turns.",
            "total_turns": max_turns,
            "conversation_history": self._format_conversation_history(),
            "execution_summary": self._generate_execution_summary(execution_steps, None)
        }

    def _execute_code(self, code: str) -> str:
        """Interpret generated code using only approved spreadsheet operations."""
        return self.interpreter.execute(code)

    def _get_llm_response(self, max_retries: int = 5,
                          base_delay: float = 1.0) -> ChatCompletionMessage:
        """Get response from OpenAI with retry logic."""
        last_exception = None

        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.deployment,
                    messages=self.conversation_history,
                )

                # Extract message
                choice = response.choices[0]
                message = choice.message

                return message

            except RateLimitError as e:
                last_exception = e
                logger.warning(f"Rate limit hit, attempt {attempt + 1}/{max_retries}: {str(e)}")

                # Extract wait time from error message if available
                wait_time = self._extract_wait_time_from_error(str(e))

                if attempt < max_retries - 1:
                    if wait_time:
                        delay = wait_time + 2
                        logger.info(f"Waiting {delay:.1f} seconds as suggested by API")
                    else:
                        delay = 10
                        logger.info(f"Waiting {delay:.1f} seconds (exponential backoff)")

                    time.sleep(delay)
                else:
                    logger.error(f"All {max_retries} attempts failed due to rate limiting")
                    break

            except Exception as e:
                last_exception = e
                logger.error(f"API error, attempt {attempt + 1}/{max_retries}: {str(e)}")

                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt) + 0.5
                    logger.info(f"Waiting {delay:.1f} seconds before retry")
                    time.sleep(delay)
                else:
                    logger.error(f"All {max_retries} attempts failed")
                    break

        if last_exception:
            raise last_exception
        raise RuntimeError("No LLM response received")

    def _parse_llm_response(self, content: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse LLM response for Final Answer or Code Action"""
        return parse_model_response(content)

    def _extract_wait_time_from_error(self, error_message: str) -> Optional[int]:
        """Extract wait time from rate limit error message."""
        try:
            # Look for patterns like "Try again in X seconds"
            match = re.search(r'try again in (\d+) seconds?', error_message.lower())
            if match:
                return int(match.group(1))

            # Look for other patterns like "Retry after X seconds"
            match = re.search(r'retry after (\d+) seconds?', error_message.lower())
            if match:
                return int(match.group(1))

            return None
        except:
            return None

    def _format_conversation_history(self) -> list:
        """Format conversation history for output."""
        formatted_history = []
        for msg in self.conversation_history:
            if hasattr(msg, 'dict'):
                formatted_history.append(msg.dict())
            elif isinstance(msg, dict):
                formatted_history.append(msg)
            else:
                # Convert other message types to dict format
                formatted_history.append({
                    "role": getattr(msg, 'role', 'unknown'),
                    "content": getattr(msg, 'content', str(msg))
                })
        return formatted_history

    def _generate_execution_summary(self, execution_steps: list, final_answer: Optional[str]) -> dict:
        """Generate a summary of the execution process."""
        successful_steps = [step for step in execution_steps if step["success"]]
        failed_steps = [step for step in execution_steps if not step["success"]]

        summary = {
            "total_code_executions": len(execution_steps),
            "successful_executions": len(successful_steps),
            "failed_executions": len(failed_steps),
            "execution_steps": execution_steps,
            "has_final_answer": final_answer is not None,
            "final_answer": final_answer
        }

        if execution_steps:
            summary["first_execution_turn"] = execution_steps[0]["turn"]
            summary["last_execution_turn"] = execution_steps[-1]["turn"]

        return summary