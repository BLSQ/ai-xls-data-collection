"""Parse the raw text response from Gemini into a Python dict.

Handles three formats: raw JSON, JSON inside a ```json code block,
and JSON embedded in free text (first ``{`` to last ``}``).
"""

import json
import re


def parse_structure_json(raw_text: str) -> dict:
    """Extract and parse JSON from Gemini's response text.

    Tries three strategies in order:

    1. The text starts with ``{`` — parse directly.
    2. A ````` ``json ... ``` ````` code block is present — extract and parse.
    3. Fall back to the substring from the first ``{`` to the last ``}``.

    Args:
        raw_text: The full text response from the Gemini API.

    Returns:
        The parsed structure proposal as a dict.

    Raises:
        ValueError:         If no JSON object can be found in the text.
        json.JSONDecodeError: If the extracted text is not valid JSON.
    """
    text = raw_text.strip()

    # Strategy 1: raw JSON.
    if text.startswith("{"):
        return json.loads(text)

    # Strategy 2: JSON inside a code block.
    match = re.search(r"```(?:json)?\s*\n?(.*?)```", text, re.DOTALL)
    if match:
        return json.loads(match.group(1).strip())

    # Strategy 3: first { to last }.
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return json.loads(text[start : end + 1])

    raise ValueError(f"Could not extract JSON from response:\n{text[:500]}")
