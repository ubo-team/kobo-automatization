"""Claude as the backup translator of the Excel and Word translation pages.

The pages translate with Gemini in numbered batches ([1] text, [2] text, …). When a Gemini call fails
(error, exhausted quota, blocked answer) or its answer leaves texts out, those texts go to Claude.
"""

import questionnaire_ai as qai

SYSTEM = """You translate the texts of survey documents (questionnaires, reports, tables).
Translate every numbered text and return it under the same number.
Keep unchanged: numbers, codes, ${...} references and "____". Keep square-bracket markers and translate the words inside them.
Do not add quotes, numbering or explanations."""

SCHEMA = {
    "type": "object",
    "properties": {
        "translations": {"type": "array", "items": {
            "type": "object",
            "properties": {"n": {"type": "integer"}, "text": {"type": "string"}},
            "required": ["n", "text"],
            "additionalProperties": False,
        }},
    },
    "required": ["translations"],
    "additionalProperties": False,
}


class ClaudeFallback:
    """Claude client from the app's secrets (without ANTHROPIC_API_KEY it stays off), with the cost and the
    number of texts Claude translated so far."""

    def __init__(self, secrets):
        self.client, self.model = None, qai.MODEL
        self.cost, self.texts = 0.0, 0
        try:
            self.client = qai.make_client(secrets["ANTHROPIC_API_KEY"])
            self.model = secrets.get("CLAUDE_MODEL", qai.MODEL)   # optional, e.g. claude-sonnet-5
        except Exception:
            pass

    def fill(self, texts, translations, from_name, to_name, gemini_error=None):
        """Translates with Claude the texts that have no entry in `translations` ({N: text}, N from 1) and
        adds them there. Returns an error message for the page, or None."""
        missing = [n for n in range(1, len(texts) + 1) if n not in translations]
        if not missing:
            return None
        if self.client is None:
            return f"Gemini: {gemini_error}" if gemini_error else None
        numbered = "\n".join(f"[{n}] {texts[n - 1]}" for n in missing)
        try:
            data, usage = qai._call_claude(self.client, SYSTEM, f"{from_name} to {to_name}.\n\n{numbered}", SCHEMA,
                                           max_tokens=32000, effort=qai.TRANSLATION_EFFORT, model=self.model)
        except qai.AIError as e:
            return f"Gemini: {gemini_error}; Claude: {e}" if gemini_error else f"Claude: {e}"
        self.cost += qai.estimate_cost(usage, self.model)
        for it in data.get("translations", []):
            n, text = it.get("n"), str(it.get("text") or "").strip()
            if n in missing and text:
                translations[n] = text
                self.texts += 1
        return None
