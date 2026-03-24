import json
from typing import Dict, Optional

import ollama


MODEL_NAME = "qwen2.5:3b-instruct"

ALLOWED_GROUPS = {
    "Executive Summary & Product Overview",
    "Scope Boundaries & Prerequisites",
    "Transition Operations & Governance",
    "Commercial & Risk Management",
    "Internal Presales Alignment",
    "Future Category 6",
    "Future Category 7",
}


def _extract_json_substring(raw_text: str) -> Optional[str]:
    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return None
    return raw_text[start : end + 1]


def _parse_classification_json(raw_text: str) -> Optional[Dict[str, str]]:
    try:
        parsed = json.loads(raw_text)
        if isinstance(parsed, dict):
            return parsed
    except Exception:
        pass

    json_substring = _extract_json_substring(raw_text)
    if not json_substring:
        return None

    try:
        parsed = json.loads(json_substring)
        if isinstance(parsed, dict):
            return parsed
    except Exception:
        return None

    return None


def classify_with_ollama(title: str, text: str) -> Dict[str, str]:
    prompt = f"""
Classificeer dit hoofdstuk in exact een van de 7 categorieen en geef JSON terug:
{{ "group": "...", "reason": "..." }}

Toegestane categorieen:
1. Executive Summary & Product Overview
2. Scope Boundaries & Prerequisites
3. Transition Operations & Governance
4. Commercial & Risk Management
5. Internal Presales Alignment
6. Future Category 6
7. Future Category 7

Geef alleen JSON terug, zonder extra tekst.

Hoofdstuktitel:
{title}

Hoofdstuktekst:
{text}
""".strip()

    try:
        response = ollama.chat(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            options={"temperature": 0},
        )

        # ollama >= 0.2.0 returns a ChatResponse object; older versions return a dict
        if isinstance(response, dict):
            raw_content = response.get("message", {}).get("content", "") or ""
        else:
            msg = getattr(response, "message", None)
            raw_content = getattr(msg, "content", "") or ""

        parsed = _parse_classification_json(raw_content)
        if not parsed:
            return {"group": "ERROR", "reason": "Failed parsing"}

        group = str(parsed.get("group", "")).strip()
        reason = str(parsed.get("reason", "")).strip() or "No reason provided"

        if group not in ALLOWED_GROUPS:
            return {"group": "ERROR", "reason": "Failed parsing"}

        return {"group": group, "reason": reason}
    except Exception:
        return {"group": "ERROR", "reason": "Failed parsing"}
