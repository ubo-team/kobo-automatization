"""Questionnaire design from the question bank (question_bank.py).

1. select_questionnaires(): Claude picks the past questionnaires closest to the new study from the catalog,
2. design(): Claude builds the new questionnaire from their questions, reusing the wording and the answer
   options and writing a new question only where the bank has none; every question records its source,
3. regenerate_question(): one question again (or a new one), with an extra instruction from the user,
4. to_docx(): the questionnaire as a Word document.
"""

import re
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

import demographics as demo
import questionnaire_ai as qai

SELECT_EFFORT = "low"
DESIGN_EFFORT = "high"        # the flow (screening, filters, funnels) needs real planning
REGENERATE_EFFORT = "medium"
MAX_SOURCES = 8               # past questionnaires whose questions are sent to the design call

QUESTION_TYPES = {
    "single": "Një përgjigje", "multiple": "Disa përgjigje", "scale": "Shkallë", "matrix": "Matricë",
    "ranking": "Renditje", "maxdiff": "MaxDiff", "text": "Tekst", "numeric": "Numër", "date": "Datë",
    "note": "Shënim",
}
OPTION_TYPES = {"single", "multiple", "scale", "matrix"}
ITEM_TYPES = {"matrix", "ranking", "maxdiff"}
LANGUAGES = {"Shqip": "Albanian", "English": "English", "Srpski": "Serbian (Latin script)"}

QUESTION_SCHEMA = {
    "type": "object",
    "properties": {
        "code": {"type": "string"},
        "section": {"type": "string"},
        "filter": {"type": "string"},
        "text": {"type": "string"},
        "type": {"type": "string", "enum": list(QUESTION_TYPES)},
        "options": {"type": "array", "items": {
            "type": "object", "properties": {"code": {"type": "string"}, "label": {"type": "string"}},
            "required": ["code", "label"], "additionalProperties": False}},
        "items": {"type": "array", "items": {"type": "string"}},
        "instruction": {"type": "string"},
        "source": {"type": "string"},
        "change": {"type": "string"},
    },
    "required": ["code", "section", "filter", "text", "type", "options", "items", "instruction", "source", "change"],
    "additionalProperties": False,
}
DESIGN_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "introduction": {"type": "string"},
        "questions": {"type": "array", "items": QUESTION_SCHEMA},
        # the instruction of each standard demographic question the user put at the start (who continues)
        "screening": {"type": "array", "items": {
            "type": "object", "properties": {"id": {"type": "string"}, "instruction": {"type": "string"}},
            "required": ["id", "instruction"], "additionalProperties": False}},
        # the standard demographic questions in the questionnaire's language, when it is not Albanian
        "translations": {"type": "array", "items": {
            "type": "object", "properties": {"id": {"type": "string"}, "text": {"type": "string"},
                                             "options": {"type": "array", "items": {"type": "string"}},
                                             "instruction": {"type": "string"}},
            "required": ["id", "text", "options", "instruction"], "additionalProperties": False}},
        "notes": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["title", "introduction", "questions", "screening", "translations", "notes"],
    "additionalProperties": False,
}
SELECT_SCHEMA = {
    "type": "object",
    "properties": {
        "questionnaires": {"type": "array", "items": {
            "type": "object", "properties": {"id": {"type": "string"}, "reason": {"type": "string"}},
            "required": ["id", "reason"], "additionalProperties": False}},
    },
    "required": ["questionnaires"],
    "additionalProperties": False,
}


# Data collection modes: the label on the page -> how it is described to Claude
MODES = {
    "CATI": "CATI (telephone interview: the interviewer reads everything out; no show cards or visual aids)",
    "CAPI": "CAPI (face-to-face interview on a tablet: the interviewer reads out and can show cards)",
    "CAWI": "CAWI (online, self-completed by the respondent: no interviewer)",
}


def usable(catalog):
    return [r for r in catalog if r.get("perdorshem") != "jo"]


def brief_text(brief):
    lines = [f"Objectives of the study: {brief['objectives']}"]
    for key, label in (("industry", "Industry"), ("survey_type", "Kind of study"), ("audience", "Target group"),
                       ("notes", "Other requirements")):
        if brief.get(key):
            lines.append(f"{label}: {brief[key]}")
    if brief.get("methods"):
        lines.append(f"Methodologies to include: {', '.join(brief['methods'])}")
    lines.append(f"Approximate number of questions, not counting the demographic questions: {brief['n_questions']}")
    lines.append(f"Language of the questionnaire: {LANGUAGES.get(brief['language'], brief['language'])}")
    if brief.get("mode"):
        lines.append(f"Data collection: {MODES.get(brief['mode'], brief['mode'])}")
    if "demographics" in brief:
        start = demo.with_required(brief.get("demographics_start") or [])
        end = [qid for qid in demo.with_required(brief["demographics"]) if qid not in start]
        lines.append("Standard demographic questions the user put at the start, to screen the target group "
                     f"(added by the program): {', '.join(start) or 'none'}")
        lines.append(f"Standard demographic questions ticked by the user (added by the program at the end): "
                     f"{', '.join(end) or 'none'}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 1. Similar questionnaires
# ---------------------------------------------------------------------------

SELECT_SYSTEM = f"""You help UBO Consulting, a market and social research company in Kosovo, build a new questionnaire from its bank of past questionnaires. From the catalog, pick the past questionnaires whose questions are the best material for the new study (at most {MAX_SOURCES}, best first):
- first those of the same kind of study in the same industry,
- then those that contain the requested methodologies (MaxDiff, price analysis, segmentation …), from any industry,
- then those covering the same topics or target group.
Prefer complete, recent questionnaires that were used (many submissions) over small or unused ones. Give the reason for each in one short sentence in Albanian."""


def catalog_text(catalog):
    return "\n".join(
        f"{r['id']} | {r['titulli']} | {r.get('industria', '')} | {r.get('lloji', '')} | "
        f"{r.get('metodologjite', '')} | {r.get('temat', '')} | {r.get('klienti', '')} | "
        f"{r.get('nr_pyetjeve', '')} pyetje | {r.get('dergesat', '')} dërgesa | {r.get('perditesuar', '')[:4]}"
        for r in usable(catalog))


def select_questionnaires(client, model, brief, catalog):
    """[(catalog row, reason)] of the questionnaires Claude picks, and the cost."""
    content = [
        {"type": "text", "text": "Catalog (id | title | industry | kind | methodologies | topics | client | "
                                 f"questions | submissions | year):\n{catalog_text(catalog)}",
         "cache_control": {"type": "ephemeral"}},
        {"type": "text", "text": brief_text(brief)},
    ]
    data, usage = qai._call_claude(client, SELECT_SYSTEM, content, SELECT_SCHEMA, max_tokens=8000,
                                   effort=SELECT_EFFORT, model=model)
    rows = {r["id"]: r for r in usable(catalog)}
    picked = [(rows[it["id"]], it["reason"]) for it in data.get("questionnaires", []) if it.get("id") in rows]
    return picked[:MAX_SOURCES], qai.estimate_cost(usage, model)


# ---------------------------------------------------------------------------
# 2. Design
# ---------------------------------------------------------------------------

def bank_context(rows, questions):
    """The questions of the chosen questionnaires, each with its reference [id#nr], and {reference: source
    description} for the editor."""
    parts, sources = [], {}
    for r in rows:
        lines = [f'<questionnaire id="{r["id"]}" title="{r["titulli"]}" industry="{r.get("industria", "")}" '
                 f'kind="{r.get("lloji", "")}" methodologies="{r.get("metodologjite", "")}">']
        group = None
        for q in questions.get(r["id"], []):
            if q["group"] != group and q["group"]:
                lines.append(f"## {q['group']}")
            group = q["group"]
            ref = f"{r['id']}#{q['nr']}"
            sources[ref] = f"{r['titulli']} — {q['name'] or 'pyetja ' + str(q['nr'])}: {q['text'][:120]}"
            options = "; ".join(f"{c}={label}" for c, label in q["options"])
            hint = f" (hint: {q['hint']})" if q["hint"] else ""
            matrix_rows = f" — rows: {' | '.join(q['items'])}" if q.get("items") else ""
            asked = f" — ASKED IF: {q['filter']}" if q.get("filter") else ""
            lines.append(f"[{ref}] {q['name'] or '-'} ({q['type']}) {q['text']}{hint}{matrix_rows}"
                         + (f" — {options}" if options else "") + asked)
        lines.append("</questionnaire>")
        parts.append("\n".join(lines))
    asked = [q for r in rows for q in questions.get(r["id"], []) if q["type"] != "note"]
    if asked:
        share = round(100 * sum(1 for q in asked if q.get("filter")) / len(asked))
        parts.insert(0, f"In these {len(rows)} questionnaires {share}% of the questions have a filter (ASKED IF): "
                        "that is the routing the new questionnaire should match.")
    return "\n\n".join(parts), sources


RULES = """Rules:
- Reuse the bank's questions wherever one fits the study. Keep their wording and their answer options with the same codes, changing only what the new study needs (brand, product, service, period, place, target group) and translating into the questionnaire's language when the bank question is in another. Set `source` to the bank reference of the question you used (e.g. "ubo:aBc123#14") and `change` to a short note in Albanian on what you changed ("" when nothing).
- Leave out the bank question's own number ("F2a.", "16.", "D4.") from `text`: the new questionnaire has its own codes.
- Write a new question only when no bank question covers a need of the study; then leave `source` empty and say in `change`, in Albanian, why it is new.
- The flow. A questionnaire is not a list of questions: it is a path through which every respondent is routed, and the bank shows how UBO builds it. In <bank> each question line is `[reference] variable (type) text … — ASKED IF: condition`, the condition in Kobo's notation over the variables of that questionnaire (e.g. `q2 = '1'`, `selected(q5, '3')`, `q7 >= 4`). Study how the past questionnaires route their respondents and build the new one the same way:
  - screening first: the qualifying questions (being a customer / user / voter …) with what happens to those who do not qualify ("END THE INTERVIEW" in `instruction`), and quotas when the target group needs them;
  - funnels from general to specific: awareness (unaided, then aided) → use / experience → evaluation and satisfaction → loyalty / recommendation → intentions; each step asked only of those to whom it applies (only the users of a service rate it, only those aware of a brand say whether they use it);
  - follow-ups on answers: "why" / "what should change" after a low (or high) rating, "which" after a "yes", "other, specify" where the bank has it, the reasons for not using after a "no";
  - batteries per brand / service / institution asked only for those the respondent knows or uses, as the bank does;
  - general and sensitive questions (voting, politics) late; group the questions into sections with a short title in `section` (in the questionnaire's language).
- `code`: a short variable for every question, unique in the questionnaire: S1, S2 … for screening, then Q1, Q2 …; notes get one too. Never use D codes: they belong to the standard demographic questions. Filters refer to these codes (and may refer to the D codes of the demographic questions the user put at the start).
- `filter`: when the question is asked, written in the questionnaire's language with the codes and the answer codes and labels, e.g. "Pyet nëse S3 = 1 (Po)", "Pyet nëse Q4 = 1-3 (të pakënaqur)", "Pyet nëse Q2 ≠ 98", "Pyet për çdo bankë të zgjedhur te Q5"; "" when everyone is asked. Every question that only makes sense for some respondents has a filter: a good questionnaire has many, as the bank's questionnaires do (their share is given at the top of <bank>). A filter can only refer to earlier questions.
- Typical routing to apply wherever it fits: a rating or satisfaction question is asked only of those who use / know what is rated; after a satisfaction or agreement scale, a "why" / "what should improve" for the low scores (and, when useful, what they like for the high ones); after NPS, the reason of the score (detractors 0-6, passives 7-8, promoters 9-10); "which / how often / since when" only after a "yes"; the reasons for not using or not doing after a "no"; questions about a main brand / provider only for those with more than one; questions on intentions (switching, buying) followed by the reasons for those who intend.
- Check the flow before answering: walk through the questionnaire as a respondent who qualifies and uses everything, as one who does not use the main service, and as one who is dissatisfied; every question each of them reaches must make sense for them. Add the missing filters and follow-ups, and remove the questions nobody reaches.
- Demographic questions are not yours to write: the program adds the company's standard demographic questions (listed in <standard_demographics>, their code is their id: D1, D2 …) that the user ticked, word for word, at the end, and those the user put at the start before your questions. Write no demographic question of your own (gender, age or year of birth, municipality, settlement, household size, marital status, children, ethnicity, education, occupation, employment, income, religion, diaspora …), not even from the bank, and do not count them in the number of questions.
- Methodologies, when asked for; prefer the bank's versions of them:
  MaxDiff: one `maxdiff` question, the attributes in `items`, the instruction on choosing the most and least important in each set;
  price analysis (Van Westendorp): the four price questions (too cheap, a bargain, getting expensive, too expensive) as `numeric`;
  price analysis (Gabor-Granger): purchase likelihood at a series of prices as `scale`;
  segmentation: a battery of attitude statements as a `matrix` with an agreement scale;
  NPS: likelihood to recommend 0-10 as `scale`.
- Types: `single` / `multiple` take `options`; `scale` takes the scale points as `options`; `matrix` takes the rows in `items` and the scale in `options`; `ranking` and `maxdiff` take `items`; `text`, `numeric`, `date` take neither; `note` is text read out or shown, not a question. Leave unused fields empty.
- Option codes: keep the bank's codes; new lists are coded 1, 2, 3 … and "Don't know" / "Refuses" get 98 / 99 when the bank uses them.
- `instruction`: interviewer or respondent instructions, in the questionnaire's language, written for the data collection mode:
  CATI: for the interviewer (READ OUT / DO NOT READ OUT the options, ONE ANSWER, PROBE …); never show cards; keep lists and scales short enough to be read out, and say how a long scale is read on the phone;
  CAPI: for the interviewer (READ OUT, SHOW CARD for long lists and scales, ONE ANSWER …);
  CAWI: for the respondent ("Select one answer", "Select all that apply" …); no interviewer instructions, no READ OUT or SHOW CARD; "Don't know" / "Refuses" only where a respondent needs them.
  Adapt the bank's instructions to the mode (a bank question from a CAPI study used in CATI loses its SHOW CARD, and so on), and adapt the introduction and the screening the same way (an introduction read by the interviewer for CATI / CAPI, a welcome text for CAWI).
- A matrix or battery counts as one question in the requested number."""

DEMOGRAPHICS_RULES = """Standard demographic questions:
- The ones the user put at the start screen the target group; the program asks them first. For each of them, give in `screening` its id and, in `instruction`, in the questionnaire's language, who continues and when the interview ends, from the target group and the objectives (e.g. "Vazhdo vetëm nëse është i/e punësuar në sektorin privat; përndryshe përfundo intervistën."); leave `instruction` empty when the target group puts no limit on that question. `screening` lists only those ids (empty when the user put none at the start). Your own screening questions, if the study needs any that are not demographic, go first among your questions as usual. When the target group is defined by a demographic characteristic that the user did not put at the start, say so in `notes`.
- The conditions of the standard questions (in brackets in <standard_demographics>: asked only for some answers, or once per child) are added by the program; do not repeat them.
- `translations`: only when the questionnaire's language is not Albanian, every standard question that is ticked or put at the start, translated: `text`, `options` (the labels only, in the same order and number as the standard list) and `instruction` (empty when the standard one is empty). Empty for an Albanian questionnaire."""


def demographics_block():
    return {"type": "text", "text": f"<standard_demographics>\n{demo.standard_text()}\n</standard_demographics>"}

DESIGN_SYSTEM = f"""You design survey questionnaires for UBO Consulting, a market and social research company in Kosovo. You build them from the company's question bank: the questions of past questionnaires given in <bank>. The company wants its tested questions, not generic ones.

{RULES}

{DEMOGRAPHICS_RULES}

Plan the questionnaire before writing it: the sections, the path of each kind of respondent through them, and the filters; then pick the bank questions for each step. Do not make it simpler than the past questionnaires of the same kind: match their depth (number of topics, follow-ups, batteries) within the requested number of questions.

Write the title, the introduction read to the respondent and all questions in the questionnaire's language. Put in `notes`, in Albanian, a short description of the flow (who is asked what, where the interview ends for those who do not qualify), then what the user should know: gaps the bank did not cover, choices you made, questions to check, and what the data collection mode asks for (CATI: the estimated interview length on the phone and the questions that are hard to ask by phone; CAPI: the show cards to print; CAWI: the estimated completion time, the questions to make mandatory and how they look on a mobile phone)."""

REGENERATE_SYSTEM = f"""You revise one question of a survey questionnaire for UBO Consulting, a market and social research company in Kosovo, following the user's instruction. You build from the company's question bank: the questions of past questionnaires given in <bank>.

{RULES}

Return only the one question. Keep it consistent with the rest of the questionnaire, which is given for context with the code and filter of each question, and write it in the questionnaire's language. A rewritten question keeps its code unless the user asks otherwise; give a new question a code not used yet. Set its filter so that it fits the flow: it is asked only of the respondents it applies to, and only refers to earlier questions."""


OLD_NUMBER = re.compile(r"^\s*(\(?[A-Za-z]{0,3}\d+[a-z]?(\.\d+)*[a-z]?\)?[.):]\s+)+")


def _bank_block(context):
    # the same bank text for the design and every regeneration that follows: read from the prompt cache
    return {"type": "text", "text": f"<bank>\n{context}\n</bank>", "cache_control": {"type": "ephemeral"}}


def _question(data, sources):
    """A question from Claude's answer, with an id for the editor's fields and its source described."""
    q = {k: data.get(k) for k in QUESTION_SCHEMA["properties"]}
    # the bank question's own number ("F2a.", "16.", "D4.") does not belong to the new questionnaire
    q["text"] = OLD_NUMBER.sub("", q.get("text") or "")
    q["options"] = [(str(o.get("code", "")), str(o.get("label", ""))) for o in data.get("options") or []]
    q["items"] = [str(i) for i in data.get("items") or []]
    q["source_text"] = sources.get(q["source"] or "", "")
    return q


def _demographic_options(qid, translation=None):
    """The options of a standard question, with the translated labels when there are as many as the codes."""
    options = list(demo.BY_ID[qid]["options"])
    labels = (translation or {}).get("options") or []
    if len(labels) == len(options):
        options = [(code, str(label)) for (code, _), label in zip(options, labels)]
    return options


def demographic_filter(qid, language="Shqip", translations=None):
    """When a standard question is asked, in the questionnaire's language and with the codes of the
    questionnaire ("Pyet nëse D4 = 2 (Rural)", "Pyet për çdo fëmijë sipas D10 …"); "" for everyone."""
    std, texts = demo.BY_ID[qid], demo.ROUTING.get(language, demo.ROUTING["Shqip"])
    if std["loop"]:
        return texts["loop"].format(code=std["loop"])
    if not std["ask_if"]:
        return ""
    parent, codes = std["ask_if"]
    labels = dict(_demographic_options(parent, (translations or {}).get(parent)))
    return texts["ask_if"].format(code=parent, answers=texts["or"].join(f"{c} ({labels.get(c, c)})" for c in codes))


def _demographic_question(qid, section, translation=None, screening="", language="Shqip", translations=None):
    """A standard demographic question for the editor, word for word (or in the questionnaire's language)."""
    std = demo.BY_ID[qid]
    text, instruction = std["text"], std["instruction"]
    if translation:
        text = translation.get("text") or text
        instruction = translation.get("instruction", instruction)
    return {"code": qid, "section": section, "filter": demographic_filter(qid, language, translations),
            "text": text, "type": std["type"], "options": _demographic_options(qid, translation), "items": [],
            "instruction": " ".join(x for x in (instruction, screening) if x), "source": "", "change": "",
            "source_text": "Pyetje demografike standarde (Paneli)", "demographic": qid}


def add_demographics(questions, data, brief):
    """The questionnaire's questions with the standard demographics: those the user put at the start first
    (with the AI agent's instruction on who continues), the other ticked ones (and those they depend on) at
    the end."""
    language = brief.get("language", "Shqip")
    translations = {t.get("id"): t for t in data.get("translations") or []} if language != "Shqip" else {}
    screening = {s.get("id"): s.get("instruction", "") for s in data.get("screening") or []}
    start_ids = demo.with_required(brief.get("demographics_start") or [])
    start = [_demographic_question(qid, demo.SCREENING_SECTION.get(language, demo.SCREENING_SECTION["Shqip"]),
                                   translations.get(qid), screening.get(qid, ""), language, translations)
             for qid in start_ids]
    end = [_demographic_question(qid, demo.SECTION.get(language, demo.SECTION["Shqip"]), translations.get(qid),
                                 "", language, translations)
           for qid in demo.with_required(brief.get("demographics") or []) if qid not in start_ids]
    return start + questions + end


def design(client, model, brief, context, sources):
    """The new questionnaire {"title", "introduction", "questions", "notes"}, and the cost."""
    content = [_bank_block(context), demographics_block(), {"type": "text", "text": brief_text(brief)}]
    data, usage = qai._call_claude(client, DESIGN_SYSTEM, content, DESIGN_SCHEMA, max_tokens=64000,
                                   effort=DESIGN_EFFORT, model=model)
    if not data.get("questions"):
        raise qai.AIError("Agjenti AI nuk ktheu pyetje. Provo përsëri ose shto më shumë detaje te objektivat.")
    questions = [_question(q, sources) for q in data["questions"]]
    if "demographics" in brief:
        questions = add_demographics(questions, data, brief)
    questionnaire = {"title": data.get("title", ""), "introduction": data.get("introduction", ""),
                     "notes": data.get("notes") or [], "questions": questions}
    return questionnaire, qai.estimate_cost(usage, model)


def outline(questionnaire, mark=None):
    lines = []
    for i, q in enumerate(questionnaire["questions"]):
        arrow = "  <-- this question" if i == mark else ""
        options = "; ".join(f"{c}={l}" for c, l in q["options"])
        asked = f" — ASKED IF: {q['filter']}" if q.get("filter") else ""
        lines.append(f"{i + 1}. {q.get('code') or '-'} [{q['section']}] ({q['type']}) {q['text']}"
                     + (f" — {options}" if options else "") + f"{asked}{arrow}")
    return "\n".join(lines)


def regenerate_question(client, model, brief, context, sources, questionnaire, index, instruction, new=False):
    """Question `index` written again following `instruction`, or with new=True a new question to insert
    after it (index -1: at the start). Returns (question, cost)."""
    if new:
        task = (f"Write a new question to insert after question {index + 1}." if index >= 0
                else "Write a new question to insert at the start.")
        mark = index if index >= 0 else None
    else:
        q = questionnaire["questions"][index]
        options = "; ".join(f"{c}={l}" for c, l in q["options"])
        task = (f"Rewrite question {index + 1}. It is now:\n{q.get('code') or ''} ({q['type']}) {q['text']}"
                + (f"\nFilter: {q['filter']}" if q.get("filter") else "")
                + (f"\nOptions: {options}" if options else "") + (f"\nItems: {'; '.join(q['items'])}" if q['items'] else "")
                + (f"\nInstruction: {q['instruction']}" if q['instruction'] else ""))
        mark = index
    content = [
        _bank_block(context),
        demographics_block(),
        {"type": "text", "text": f"{brief_text(brief)}\n\nThe questionnaire so far:\n{outline(questionnaire, mark)}"
                                 f"\n\n{task}\nThe user's instruction: {instruction or 'write a better version'}"},
    ]
    data, usage = qai._call_claude(client, REGENERATE_SYSTEM, content, QUESTION_SCHEMA, max_tokens=16000,
                                   effort=REGENERATE_EFFORT, model=model)
    return _question(data, sources), qai.estimate_cost(usage, model)


# ---------------------------------------------------------------------------
# 4. Word
# ---------------------------------------------------------------------------

DOCX_TEXTS = {
    "Shqip": {"q": "P", "filter": "FILTËR", "section": "Seksioni", "single": "Një përgjigje", "multiple": "Disa përgjigje të mundshme", "write": "Shëno",
              "rank": "Rendit sipas rëndësisë", "maxdiff": "MaxDiff: në çdo grup zgjidh më të rëndësishmin "
              "dhe më pak të rëndësishmin", "items": "Pohimet / elementet"},
    "English": {"q": "Q", "filter": "FILTER", "section": "Section", "single": "Single answer", "multiple": "Multiple answers possible", "write": "Write in",
                "rank": "Rank in order of importance", "maxdiff": "MaxDiff: in each set choose the most and the "
                "least important", "items": "Statements / items"},
    "Srpski": {"q": "P", "filter": "FILTER", "section": "Sekcija", "single": "Jedan odgovor", "multiple": "Moguće više odgovora", "write": "Upišite",
               "rank": "Rangirajte po važnosti", "maxdiff": "MaxDiff: u svakom setu izaberite najvažniji i "
               "najmanje važan", "items": "Tvrdnje / stavke"},
}
GREY = RGBColor(0x6B, 0x6F, 0x7B)
FILTER_COLOR = RGBColor(0x25, 0x60, 0xE6)


def question_labels(questions, numbering, prefix="P"):
    """{question code: the number it has in the Word document} ("P7", or "B3" with lettered sections), so that
    filters written with codes can be shown with the document's numbers."""
    labels, n, letter, k, section = {}, 0, 0, 0, None
    for q in questions:
        if q.get("section") and q["section"] != section:
            section, letter, k = q["section"], letter + 1, 0
        if q["type"] == "note":
            continue
        n, k = n + 1, k + 1
        label = f"{chr(64 + max(letter, 1))}{k}" if numbering == "sections" else f"{prefix}{n}"
        if q.get("code"):
            labels[q["code"]] = label
    return labels


def resolve_filter(text, labels):
    """The codes in a filter ("Pyet nëse S3 = 1") replaced by the document's question numbers."""
    if not labels:
        return text
    pattern = re.compile(r"\b(" + "|".join(re.escape(c) for c in sorted(labels, key=len, reverse=True)) + r")\b")
    return pattern.sub(lambda m: labels[m.group(1)], text)
# The answer-type hint in brackets is left out when the question's instruction already says the same
HINT_SAID = {
    "single": re.compile(r"një përgjigje|vetëm një|one answer|single answer|select one|jedan odgovor|samo jedan",
                         re.IGNORECASE),
    "multiple": re.compile(r"disa përgjigje|më shumë se një|shumë përgjigje|të gjitha që|multiple|all that apply|"
                           r"više odgovora|sve što", re.IGNORECASE),
    "ranking": re.compile(r"rendit|rank|rangir", re.IGNORECASE),
    "maxdiff": re.compile(r"më të rëndësishm|most important|najvažnij", re.IGNORECASE),
}


def _note(doc, text, italic=True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic, run.font.size, run.font.color.rgb = italic, Pt(9), GREY
    p.paragraph_format.space_after = Pt(2)
    return p


# Question numbering in Word: "auto" letters the sections (A1, A2 … B1) for long questionnaires
NUMBERING = {"auto": "Automatik", "plain": "1, 2, 3", "sections": "A1, A2 … sipas seksioneve"}
LONG_QUESTIONNAIRE = 30


class _Lists:
    """Word's own numbered lists (numbering.xml), so that deleting or moving a question or an option in Word
    renumbers the rest by itself."""

    def __init__(self, doc):
        self.numbering = doc.part.numbering_part.element
        ids = lambda tag, attr: [int(e.get(qn(attr))) for e in self.numbering.findall(qn(tag))]
        self.next_abstract = max(ids("w:abstractNum", "w:abstractNumId") or [0]) + 1
        self.next_num = max(ids("w:num", "w:numId") or [0]) + 1

    @staticmethod
    def level(ilvl, fmt, text, left, hanging, bold=False, size=None):
        rpr = ("<w:b/>" if bold else "") + (f'<w:sz w:val="{size}"/>' if size else "")
        return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>'
                + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "") + "</w:lvl>")

    def abstract(self, *levels):
        aid, self.next_abstract = self.next_abstract, self.next_abstract + 1
        el = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}">'
                       f'<w:multiLevelType w:val="multilevel"/>{"".join(levels)}</w:abstractNum>')
        first_num = self.numbering.find(qn("w:num"))       # abstractNum elements must come before num ones
        if first_num is not None:
            first_num.addprevious(el)
        else:
            self.numbering.append(el)
        return aid

    def num(self, aid, restart=False):
        nid, self.next_num = self.next_num, self.next_num + 1
        override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' if restart else ""
        self.numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/>'
                                        f"{override}</w:num>"))
        return nid

    @staticmethod
    def apply(paragraph, nid, ilvl=0):
        num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        num_pr.get_or_add_ilvl().val = ilvl
        num_pr.get_or_add_numId().val = nid


def to_docx(questionnaire, language="Shqip", with_sources=False, numbering="auto"):
    """The questionnaire as a Word document (bytes): sections, numbered questions, instructions, options
    a) b) c), matrices as tables. Questions, sections and options use Word's numbered lists. numbering: "plain"
    (1, 2, 3), "sections" (Seksioni A, questions A1, A2 …) or "auto" (sections for long questionnaires).
    with_sources adds the bank source of each question (for the internal review)."""
    t = DOCX_TEXTS.get(language, DOCX_TEXTS["Shqip"])
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)
    doc.add_heading(questionnaire.get("title") or "Pyetësori", level=0)
    if questionnaire.get("introduction"):
        p = doc.add_paragraph(questionnaire["introduction"])
        p.runs[0].italic = True

    questions = questionnaire["questions"]
    if numbering == "auto":
        n_asked = sum(q["type"] != "note" for q in questions)
        n_sections = len({q.get("section") for q in questions if q.get("section")})
        numbering = "sections" if n_asked > LONG_QUESTIONNAIRE and n_sections > 1 else "plain"
    labels = question_labels(questions, numbering, t["q"])
    lists = _Lists(doc)
    if numbering == "sections":
        # level 0: "Seksioni A:" on the section headings; level 1: A1., A2. … (restarts in every section)
        q_num = lists.num(lists.abstract(
            lists.level(0, "upperLetter", f"{t['section']} %1:", 0, 0),
            lists.level(1, "decimal", "%1%2.", 567, 567, bold=True)))
        q_level = 1
    else:
        q_num, q_level = lists.num(lists.abstract(lists.level(0, "decimal", "%1.", 454, 454, bold=True))), 0
    option_list = lists.abstract(lists.level(0, "lowerLetter", "%1)", 851, 397))
    indent = Pt(28) if q_level else Pt(23)

    section = None
    for q in questions:
        if q.get("section") and q["section"] != section:
            section = q["section"]
            heading = doc.add_heading(section, level=1)
            if numbering == "sections":
                lists.apply(heading, q_num, 0)
        if q.get("filter"):         # who is asked, with the numbers of the questions it refers to
            f = doc.add_paragraph()
            run = f.add_run(f"{t['filter']}: {resolve_filter(q['filter'], labels)}")
            run.bold, run.font.size, run.font.color.rgb = True, Pt(9), FILTER_COLOR
            f.paragraph_format.space_before, f.paragraph_format.space_after = Pt(8), Pt(0)
            f.paragraph_format.keep_with_next = True
        if q["type"] == "note":
            doc.add_paragraph().add_run(q["text"]).italic = True
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2) if q.get("filter") else Pt(8)
        p.paragraph_format.keep_with_next = True
        lists.apply(p, q_num, q_level)
        p.add_run(q["text"]).bold = True
        hint = {"single": t["single"], "multiple": t["multiple"], "ranking": t["rank"],
                "maxdiff": t["maxdiff"]}.get(q["type"], "")
        own = q.get("instruction") or ""
        if hint and HINT_SAID.get(q["type"]) and HINT_SAID[q["type"]].search(own):
            hint = ""       # the question's own instruction already says it ("NJË PËRGJIGJE")
        instruction = " ".join(x for x in (own, f"[{hint}]" if hint else "") if x)
        if instruction:
            _note(doc, instruction).paragraph_format.left_indent = indent

        if q["type"] == "matrix" and q["items"] and q["options"]:
            table = doc.add_table(rows=1, cols=len(q["options"]) + 1)
            table.style = "Table Grid"
            for cell, (code, label) in zip(table.rows[0].cells[1:], q["options"]):
                cell.text = label
            for item in q["items"]:
                table.add_row().cells[0].text = item
            continue
        # items (ranking, MaxDiff) and options: a) b) c), a new list for every question so it starts at a)
        entries = list(q["items"]) + [label for _, label in q["options"]]
        if entries:
            nid = lists.num(option_list, restart=True)
            for entry in entries:
                op = doc.add_paragraph(entry)
                op.paragraph_format.space_after = Pt(0)
                lists.apply(op, nid)
        if q["type"] in ("text", "numeric", "date"):
            doc.add_paragraph(f"{t['write']}: ____________________").paragraph_format.left_indent = indent
        if with_sources:
            _note(doc, f"Burimi: {q['source_text'] or 'pyetje e re'}" + (f" · {q['change']}" if q.get("change") else ""))

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def parse_options(text):
    """Editor lines "1. Label" (or "1 = Label", or only "Label") -> [(code, label)]; unnumbered lines get the
    next number."""
    options = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(\w{1,4})\s*[.)=:]\s+(.+)$", line)
        if m:
            options.append((m.group(1), m.group(2).strip()))
        else:
            options.append((str(len(options) + 1), line))
    return options


def options_text(options):
    return "\n".join(f"{c}. {l}" if c else l for c, l in options)
