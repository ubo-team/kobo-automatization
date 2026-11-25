"""Claude as the translator of the Excel and Word translation pages.

The pages send the texts in numbered batches ([1] text, [2] text, …); Claude returns each translation under its
number. (The module keeps its old name: Claude was first only the backup of Gemini.)
"""

import threading

import questionnaire_ai as qai

BATCH_SIZE = 50     # texts per request
WORKERS = 4         # requests at the same time

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
    """Claude client from the app's secrets, with the cost, the tokens and the number of texts translated so
    far (the pages show them)."""

    def __init__(self, secrets):
        self.client, self.model = None, qai.MODEL
        self.cost, self.texts, self.tokens_in, self.tokens_out = 0.0, 0, 0, 0
        self.lock = threading.Lock()
        try:
            self.client = qai.make_client(secrets["ANTHROPIC_API_KEY"])
            self.model = secrets.get("CLAUDE_MODEL", qai.MODEL)   # optional, e.g. claude-sonnet-5
        except Exception:
            pass

    def fill(self, texts, translations, from_name, to_name, previous_error=None):
        """Translates with Claude the texts that have no entry in `translations` ({N: text}, N from 1) and
        adds them there. Returns an error message for the page, or None."""
        missing = [n for n in range(1, len(texts) + 1) if n not in translations]
        if not missing:
            return None
        if self.client is None:
            return "Mungon çelësi `ANTHROPIC_API_KEY` në secrets të aplikacionit."
        numbered = "\n".join(f"[{n}] {texts[n - 1]}" for n in missing)
        try:
            data, usage = qai._call_claude(self.client, SYSTEM, f"{from_name} to {to_name}.\n\n{numbered}", SCHEMA,
                                           max_tokens=32000, effort=qai.TRANSLATION_EFFORT, model=self.model)
        except qai.AIError as e:
            return f"Agjenti AI: {e}"
        with self.lock:
            self.cost += qai.estimate_cost(usage, self.model)
            self.tokens_in += (usage.input_tokens or 0) + (getattr(usage, "cache_read_input_tokens", 0) or 0)
            self.tokens_out += usage.output_tokens or 0
        for it in data.get("translations", []):
            n, text = it.get("n"), str(it.get("text") or "").strip()
            if n in missing and text:
                translations[n] = text
                with self.lock:
                    self.texts += 1
        return None

    def translate_all(self, texts, from_name, to_name, progress=None):
        """All `texts` in batches of BATCH_SIZE, WORKERS at a time, each batch tried a second time for the texts
        it left out. Returns ({index in texts: translation}, [errors]); progress(done, total) after each batch."""
        from concurrent.futures import ThreadPoolExecutor, as_completed

        batches = [list(range(i, min(i + BATCH_SIZE, len(texts)))) for i in range(0, len(texts), BATCH_SIZE)]

        def one(indexes):
            batch = [texts[i] for i in indexes]
            translations, errors = {}, []
            for _ in range(2):
                error = self.fill(batch, translations, from_name, to_name)
                if error:
                    errors.append(error)
                if len(translations) == len(batch) or error:
                    break
            return {indexes[n - 1]: t for n, t in translations.items()}, errors

        result, errors, done = {}, [], 0
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            for future in as_completed([pool.submit(one, b) for b in batches]):
                translated, batch_errors = future.result()
                result.update(translated)
                errors += batch_errors
                done += 1
                if progress:
                    progress(done, len(batches))
        return result, errors
