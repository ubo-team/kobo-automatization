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

import questionnaire_ai as qai

SELECT_EFFORT = "low"
DESIGN_EFFORT = "medium"      # raise to "high" if the questionnaires come out thin or badly ordered
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
        "section": {"type": "string"},
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
    "required": ["section", "text", "type", "options", "items", "instruction", "source", "change"],
    "additionalProperties": False,
}
DESIGN_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "introduction": {"type": "string"},
        "questions": {"type": "array", "items": QUESTION_SCHEMA},
        "notes": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["title", "introduction", "questions", "notes"],
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
    lines.append(f"Approximate number of questions: {brief['n_questions']}")
    lines.append(f"Language of the questionnaire: {LANGUAGES.get(brief['language'], brief['language'])}")
    if brief.get("mode"):
        lines.append(f"Data collection: {MODES.get(brief['mode'], brief['mode'])}")
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
            rows = f" — rows: {' | '.join(q['items'])}" if q.get("items") else ""
            lines.append(f"[{ref}] ({q['type']}) {q['text']}{hint}{rows}" + (f" — {options}" if options else ""))
        lines.append("</questionnaire>")
        parts.append("\n".join(lines))
    return "\n\n".join(parts), sources


RULES = """Rules:
- Reuse the bank's questions wherever one fits the study. Keep their wording and their answer options with the same codes, changing only what the new study needs (brand, product, service, period, place, target group) and translating into the questionnaire's language when the bank question is in another. Set `source` to the bank reference of the question you used (e.g. "ubo:aBc123#14") and `change` to a short note in Albanian on what you changed ("" when nothing).
- Write a new question only when no bank question covers a need of the study; then leave `source` empty and say in `change`, in Albanian, why it is new.
- Usual order: introduction and screening first, then the main topics from general to specific, the methodology modules where they belong, demographics at the end. Give every question its `section` (a short section title in the questionnaire's language).
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

DESIGN_SYSTEM = f"""You design survey questionnaires for UBO Consulting, a market and social research company in Kosovo. You build them from the company's question bank: the questions of past questionnaires given in <bank>. The company wants its tested questions, not generic ones.

{RULES}

Write the title, the introduction read to the respondent and all questions in the questionnaire's language. Put in `notes`, in Albanian, what the user should know: gaps the bank did not cover, choices you made, questions to check, and what the data collection mode asks for (CATI: the estimated interview length on the phone and the questions that are hard to ask by phone; CAPI: the show cards to print; CAWI: the estimated completion time, the questions to make mandatory and how they look on a mobile phone)."""

REGENERATE_SYSTEM = f"""You revise one question of a survey questionnaire for UBO Consulting, a market and social research company in Kosovo, following the user's instruction. You build from the company's question bank: the questions of past questionnaires given in <bank>.

{RULES}

Return only the one question. Keep it consistent with the rest of the questionnaire, which is given for context, and write it in the questionnaire's language."""


def _bank_block(context):
    # the same bank text for the design and every regeneration that follows: read from the prompt cache
    return {"type": "text", "text": f"<bank>\n{context}\n</bank>", "cache_control": {"type": "ephemeral"}}


def _question(data, sources):
    """A question from Claude's answer, with an id for the editor's fields and its source described."""
    q = {k: data.get(k) for k in QUESTION_SCHEMA["properties"]}
    q["options"] = [(str(o.get("code", "")), str(o.get("label", ""))) for o in data.get("options") or []]
    q["items"] = [str(i) for i in data.get("items") or []]
    q["source_text"] = sources.get(q["source"] or "", "")
    return q


def design(client, model, brief, context, sources):
    """The new questionnaire {"title", "introduction", "questions", "notes"}, and the cost."""
    content = [_bank_block(context), {"type": "text", "text": brief_text(brief)}]
    data, usage = qai._call_claude(client, DESIGN_SYSTEM, content, DESIGN_SCHEMA, max_tokens=64000,
                                   effort=DESIGN_EFFORT, model=model)
    if not data.get("questions"):
        raise qai.AIError("Agjenti AI nuk ktheu pyetje. Provo përsëri ose shto më shumë detaje te objektivat.")
    questionnaire = {"title": data.get("title", ""), "introduction": data.get("introduction", ""),
                     "notes": data.get("notes") or [],
                     "questions": [_question(q, sources) for q in data["questions"]]}
    return questionnaire, qai.estimate_cost(usage, model)


def outline(questionnaire, mark=None):
    lines = []
    for i, q in enumerate(questionnaire["questions"]):
        arrow = "  <-- this question" if i == mark else ""
        lines.append(f"{i + 1}. [{q['section']}] ({q['type']}) {q['text']}{arrow}")
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
        task = (f"Rewrite question {index + 1}. It is now:\n({q['type']}) {q['text']}"
                + (f"\nOptions: {options}" if options else "") + (f"\nItems: {'; '.join(q['items'])}" if q['items'] else "")
                + (f"\nInstruction: {q['instruction']}" if q['instruction'] else ""))
        mark = index
    content = [
        _bank_block(context),
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
    "Shqip": {"section": "Seksioni", "single": "Një përgjigje", "multiple": "Disa përgjigje të mundshme", "write": "Shëno",
              "rank": "Rendit sipas rëndësisë", "maxdiff": "MaxDiff: në çdo grup zgjidh më të rëndësishmin "
              "dhe më pak të rëndësishmin", "items": "Pohimet / elementet"},
    "English": {"section": "Section", "single": "Single answer", "multiple": "Multiple answers possible", "write": "Write in",
                "rank": "Rank in order of importance", "maxdiff": "MaxDiff: in each set choose the most and the "
                "least important", "items": "Statements / items"},
    "Srpski": {"section": "Sekcija", "single": "Jedan odgovor", "multiple": "Moguće više odgovora", "write": "Upišite",
               "rank": "Rangirajte po važnosti", "maxdiff": "MaxDiff: u svakom setu izaberite najvažniji i "
               "najmanje važan", "items": "Tvrdnje / stavke"},
}
GREY = RGBColor(0x6B, 0x6F, 0x7B)
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
        if q["type"] == "note":
            doc.add_paragraph().add_run(q["text"]).italic = True
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_before, p.paragraph_format.keep_with_next = Pt(8), True
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
