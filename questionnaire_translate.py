"""Languages of the Gjenero XLS form.

Works on the tagged line format of questionnaire_ai (texts of several languages joined with " || "):
- translates a questionnaire into the chosen languages with Gemini,
- merges translations that come as separate files (one file per language), also with Gemini,
- keeps or reorders the languages of a questionnaire that already holds them.
Only the texts change (question, option, note, hint, titles, scale labels); numbers and tags stay as they are.
"""

import json
import re
from concurrent.futures import ThreadPoolExecutor

import google.generativeai as genai

from questionnaire_ai import (AIError, LANG_SEP, TAG_RE, declared_languages, is_tag, strip_tags)

GEMINI_MODEL = "gemini-2.5-flash"      # the page can override it with GEMINI_MODEL in secrets
# USD per 1M tokens (input, output), used only for the cost estimate shown in the UI
GEMINI_PRICING = {"gemini-2.5-flash": (0.30, 2.50), "gemini-3.1-flash-lite": (0.25, 1.50),
                  "gemini-2.5-pro": (1.25, 10.00)}
BATCH_SIZE = 60           # texts per translation call
ALIGN_BATCH_SIZE = 120    # texts per call when matching against a translated file
WORKERS = 4

# Albanian and English names of languages -> the English name used in the form ([languages: …] line)
LANGUAGE_NAMES = {
    "shqip": "Albanian", "shqipe": "Albanian", "albanian": "Albanian",
    "anglisht": "English", "angleze": "English", "english": "English",
    "serbisht": "Serbian", "serbe": "Serbian", "serbian": "Serbian", "srpski": "Serbian",
    "kroatisht": "Croatian", "croatian": "Croatian", "hrvatski": "Croatian",
    "boshnjakisht": "Bosnian", "bosnian": "Bosnian", "bosanski": "Bosnian",
    "maqedonisht": "Macedonian", "macedonian": "Macedonian",
    "malazisht": "Montenegrin", "montenegrin": "Montenegrin",
    "turqisht": "Turkish", "turkish": "Turkish",
    "gjermanisht": "German", "german": "German",
    "frëngjisht": "French", "frengjisht": "French", "french": "French",
    "italisht": "Italian", "italian": "Italian",
    "greqisht": "Greek", "greek": "Greek",
    "spanjisht": "Spanish", "spanish": "Spanish",
    "rome": "Romani", "romisht": "Romani", "romani": "Romani",
}

QUESTION_TYPES = ("single", "multiple", "text", "string", "numeric", "decimal", "date", "time",
                  "matrix", "ranking", "scale", "note", "other")
NUMBER_RE = re.compile(r'^((?:[A-Z]+[_\-]?)*\d+(?:[._\-]?\d+|[a-zA-Z](?![a-z]))*)(?:[\.\):]\s*|\s+)')
SCALE_RE = re.compile(r'^(scale\s*)(\d+)(?:\((.*?)\))?(\s*-\s*)(\d+)(?:\((.*?)\))?(.*)$', flags=re.IGNORECASE | re.S)


def language_name(name):
    """English name of a language typed in Albanian or English (Kroatisht -> Croatian)."""
    name = name.strip()
    return LANGUAGE_NAMES.get(name.lower(), name[:1].upper() + name[1:])


def configure(api_key):
    genai.configure(api_key=api_key)


# ---------------------------------------------------------------------------
# Tagged line <-> texts
# ---------------------------------------------------------------------------

def _segments(text):
    return [p.strip() for p in text.split(LANG_SEP)]


def _parse(line):
    """Splits a tagged line into its question number, its tags and its translatable texts
    ("text", "hint", "scale_min", "scale_max"), each a list with one entry per language."""
    tags = [t for t in TAG_RE.findall(line) if is_tag(t)]
    words = [t.strip().lower() for t in tags]
    if any(w.startswith("languages:") or w.startswith("language:") for w in words):
        return None                                   # the [languages: …] line is rebuilt separately
    texts = {}
    text = strip_tags(line)
    number = None
    is_question = any(re.match(r'(' + '|'.join(QUESTION_TYPES) + r')\b', w) or w.startswith("scale")
                      for w in words)
    if text and is_question:
        m = NUMBER_RE.match(text)
        if m:
            number, text = m.group(1), text[m.end():]
    if text:
        segs = _segments(text)
        if number:   # a number repeated before the other languages' text
            segs = [re.sub(r'^' + re.escape(number) + r'[\.\):]?\s*', '', s) for s in segs]
        texts["text"] = segs
    for t in tags:
        m = re.match(r'\s*hint\s*:\s*(.+)', t, flags=re.IGNORECASE | re.S)
        if m:
            texts["hint"] = _segments(m.group(1))
        m = SCALE_RE.match(t.strip())
        if m:
            if m.group(3):
                texts["scale_min"] = _segments(m.group(3))
            if m.group(6):
                texts["scale_max"] = _segments(m.group(6))
    section_header = not any(w in ("section", "seksion") for w in words) \
        and re.match(r'^\s*(section|seksion)\b', line, flags=re.IGNORECASE) is not None
    return {"number": number, "tags": tags, "texts": texts, "section_header": section_header}


def _build(p, texts):
    """The line again, with `texts` ({slot: [one text per language]}) in place of the parsed texts."""
    def joined(slot):
        parts = [t.strip() for t in texts[slot]]
        if slot == "hint":      # a "]" inside a tag would end it
            parts = [t.replace("[", "(").replace("]", ")") for t in parts]
        elif slot != "text":    # scale labels sit inside "(…)"
            parts = [re.sub(r'[\[\]()]', '', t) for t in parts]
        return f" {LANG_SEP} ".join(parts)

    tags = []
    for t in p["tags"]:
        if re.match(r'\s*hint\s*:', t, flags=re.IGNORECASE) and "hint" in texts:
            tags.append(f"hint: {joined('hint')}")
            continue
        m = SCALE_RE.match(t.strip())
        if m and ("scale_min" in texts or "scale_max" in texts):
            low = f"({joined('scale_min')})" if "scale_min" in texts else ""
            high = f"({joined('scale_max')})" if "scale_max" in texts else ""
            tags.append(f"{m.group(1)}{m.group(2)}{low}{m.group(4)}{m.group(5)}{high}{m.group(7)}")
            continue
        tags.append(t)
    if p["section_header"]:
        tags.insert(0, "section")   # "Section 1 …" is found by its first word, which a translation may change
    text = joined("text") if "text" in texts else ""
    if p["number"]:
        text = f"{p['number']}. {text}"
    return " ".join([text] + [f"[{t}]" for t in tags]).strip()


def _rebuild(lines, languages, text_for):
    """Rebuilds every line with the texts text_for(parsed, slot, language index) returns, adding the
    [languages: …] line for more than one language."""
    out = [f"[languages: {', '.join(languages)}]"] if len(languages) > 1 else []
    for line in lines:
        p = _parse(line)
        if p is None:
            continue
        if not p["texts"] and not p["section_header"]:
            out.append(line)
            continue
        texts = {slot: [text_for(p, slot, k) for k in range(len(languages))] for slot in p["texts"]}
        out.append(_build(p, texts))
    return out


def _lang_key(name):
    return language_name(name).lower()


def select_languages(lines, languages):
    """Keeps the chosen languages of a questionnaire that already holds them, in the chosen order.
    Returns (lines, languages that are missing from the questionnaire)."""
    declared = declared_languages(lines)
    if not declared:
        if len(languages) == 1:
            return list(lines), []
        return _rebuild(lines, languages, lambda p, slot, k: _pick(p["texts"][slot], k)), []
    position = {_lang_key(l): i for i, l in enumerate(declared)}
    order = [position.get(_lang_key(l)) for l in languages]
    missing = [l for l, i in zip(languages, order) if i is None]
    return _rebuild(lines, languages, lambda p, slot, k: _pick(p["texts"][slot], order[k])), missing


def _pick(segs, k):
    if k is None:
        return segs[0]
    return segs[k] if k < len(segs) and segs[k] else segs[0]


# ---------------------------------------------------------------------------
# Gemini
# ---------------------------------------------------------------------------

def _gemini(prompt, model_name):
    model = genai.GenerativeModel(model_name)
    try:
        response = model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(temperature=0.1, max_output_tokens=65536,
                                                           response_mime_type="application/json"))
        data = json.loads(response.text)
    except json.JSONDecodeError:
        raise AIError("Përgjigja e Gemini nuk ishte në formatin e pritur. Provo përsëri.")
    except Exception as e:
        raise AIError(f"Gabim nga Gemini: {e}")
    usage = response.usage_metadata
    price_in, price_out = GEMINI_PRICING.get(model_name, GEMINI_PRICING["gemini-2.5-flash"])
    cost = ((getattr(usage, "prompt_token_count", 0) or 0) * price_in
            + (getattr(usage, "candidates_token_count", 0) or 0) * price_out) / 1_000_000
    return data, cost


RULES = """Rules:
- Formal, clear survey language that a respondent understands when the interviewer reads it out.
- Text in square brackets is an instruction marker for the interviewer ([READ ALOUD], [INTERVIEWER NOTE: …]): keep the square brackets and translate the words inside. Keep a "!" right after the opening bracket ([!SECTION] -> [!SEKSIONI]).
- Placeholders in parentheses such as (NAME) are translated and keep their parentheses.
- Keep unchanged: ${...} references, numbers, answer codes, "____" (four underscores), "||".
- Do not add numbering, quotes or explanations.
- Serbian, Croatian, Bosnian, Montenegrin and Macedonian are written in the Latin script."""


def _item_id(it, items):
    """The id of a returned item when it is one of the items asked for (Gemini may return it as a string)."""
    try:
        i = int(it.get("id")) if isinstance(it, dict) else None
    except (TypeError, ValueError):
        return None
    return i if i in dict(items) else None


def _translate_batch(items, languages, model_name):
    listing = "\n".join(json.dumps({"id": i, "text": t}, ensure_ascii=False) for i, t in items)
    prompt = f"""You translate the texts of a survey questionnaire for a KoboToolbox form.
Translate every item into each of these languages: {', '.join(languages)}.
When an item is already written in one of these languages, copy it unchanged for that language.

{RULES}

Return JSON: {{"items": [{{"id": <id>, "translations": {{{', '.join(f'"{l}": "..."' for l in languages)}}}}}]}} with one entry for every item.

Items:
{listing}"""
    data, cost = _gemini(prompt, model_name)
    result = {}
    for it in data.get("items", []) if isinstance(data, dict) else []:
        i = _item_id(it, items)
        tr = it.get("translations") if i is not None else None
        if isinstance(tr, dict):
            result[i] = {l: str(tr.get(l) or "").strip() for l in languages}
    return result, cost


def _align_batch(items, language, main_language, reference, model_name):
    listing = "\n".join(json.dumps({"id": i, "text": t}, ensure_ascii=False) for i, t in items)
    prompt = f"""Below is the {language} version of a survey questionnaire (the reference document) and a list of texts from its {main_language} version.
For every item, return the text of the reference document that corresponds to it, copied exactly as the reference writes it, without the question number and without answer codes. Keep square-bracket markers such as [READ ALOUD] as the reference writes them.
When the reference has no corresponding text, translate the item into {language} yourself and set "found" to false.

{RULES}

Return JSON: {{"items": [{{"id": <id>, "text": "...", "found": true}}]}} with one entry for every item.

<reference_document language="{language}">
{reference}
</reference_document>

Items ({main_language}):
{listing}"""
    data, cost = _gemini(prompt, model_name)
    result, not_found = {}, []
    for it in data.get("items", []) if isinstance(data, dict) else []:
        i = _item_id(it, items)
        if i is not None and str(it.get("text") or "").strip():
            result[i] = str(it["text"]).strip()
            if it.get("found") is False:
                not_found.append(i)
    return result, not_found, cost


def _batches(items, size):
    return [items[k:k + size] for k in range(0, len(items), size)]


def _run(fn, batches):
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(fn, batches))


def _source_texts(lines):
    """Unique first-language texts of all lines, as [(id, text)]."""
    seen = {}
    for line in lines:
        p = _parse(line)
        if p:
            for segs in p["texts"].values():
                if segs[0] and segs[0] not in seen:
                    seen[segs[0]] = len(seen)
    return [(i, t) for t, i in seen.items()]


def _keep_blanks(source, translated):
    """An option's "____" (it creates the "specify" field) must survive the translation."""
    if "_" in source and "_" not in translated:
        return f"{translated} ____"
    return translated


def translate(lines, languages, model_name=GEMINI_MODEL):
    """Translates a questionnaire (its first language) into `languages` with Gemini, retrying the texts that
    came back empty. Returns (lines, notes, cost)."""
    lines = _first_language(lines)
    items = _source_texts(lines)
    translations, cost = {}, 0.0
    for attempt in range(2):
        todo = [(i, t) for i, t in items
                if i not in translations or not all(translations[i].get(l) for l in languages)]
        if not todo:
            break
        for result, c in _run(lambda b: _translate_batch(b, languages, model_name), _batches(todo, BATCH_SIZE)):
            translations.update(result)
            cost += c
    by_text = {t: translations.get(i, {}) for i, t in items}
    missing = sum(1 for i, _ in items if not all(translations.get(i, {}).get(l) for l in languages))

    def text_for(p, slot, k):
        src = p["texts"][slot][0]
        return _keep_blanks(src, by_text.get(src, {}).get(languages[k]) or src)

    notes = [f"Gemini nuk ktheu përkthim për {missing} tekste; aty mbeti teksti origjinal."] if missing else []
    return _rebuild(lines, languages, text_for), notes, cost


def merge_translations(lines, languages, references, model_name=GEMINI_MODEL):
    """Adds the translations that come as separate files: `lines` hold the first language (the main file),
    `references` = {language: text of that language's file}. Returns (lines, notes, cost)."""
    lines = _first_language(lines)
    items = _source_texts(lines)
    texts = {lang: {} for lang in languages[1:]}
    notes, cost = [], 0.0
    for lang in languages[1:]:
        results = _run(lambda b: _align_batch(b, lang, languages[0], references[lang], model_name),
                       _batches(items, ALIGN_BATCH_SIZE))
        not_found = 0
        for result, nf, c in results:
            texts[lang].update(result)
            not_found += len(nf)
            cost += c
        if not_found:
            notes.append(f"{not_found} tekste nuk u gjetën në skedarin {lang} dhe i përktheu Gemini; kontrolloji.")
    by_text = {lang: {t: texts[lang].get(i) for i, t in items} for lang in languages[1:]}

    def text_for(p, slot, k):
        src = p["texts"][slot][0]
        if k == 0:
            return src
        return _keep_blanks(src, by_text[languages[k]].get(src) or "")

    return _rebuild(lines, languages, text_for), notes, cost


def _first_language(lines):
    """A multi-language questionnaire reduced to its first language."""
    if not declared_languages(lines):
        return list(lines)
    return _rebuild(lines, ["_"], lambda p, slot, k: p["texts"][slot][0])
