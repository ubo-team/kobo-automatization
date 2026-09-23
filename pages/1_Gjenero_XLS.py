import streamlit as st
from docx2python import docx2python
import pandas as pd
import re
import tempfile
import os
from google.oauth2.service_account import Credentials
import gspread
import hashlib
from io import BytesIO
import questionnaire_ai as qai


st.set_page_config(page_title="Gjenero XLS", layout="centered")

logo_svg_path = "UBO-Logo.svg"
with st.sidebar:
    if os.path.exists(logo_svg_path):
        with open(logo_svg_path, "r", encoding="utf-8") as f:
            svg_logo = f.read()
        st.markdown(
            f'<div style="display:flex;justify-content:center;margin:15px 0;"><div style="width:150px;">{svg_logo}</div></div>',
            unsafe_allow_html=True
        )

st.title("Gjenero XLS")
st.markdown("Ngarko pyetësorin dhe gjenero formularin XLS për përdorim në Kobo Toolbox.")

WORKFLOW_GENERATE = "Gjenero pyetësorin"
WORKFLOW_CHECK = "Kontrollo pyetësorin"
workflow = st.radio("Mënyra e punës:", [WORKFLOW_GENERATE, WORKFLOW_CHECK], index=0)

SOURCE_TAGGED = "I formatuar me tag-e (.docx)"
SOURCE_PLAIN = "I paformatuar – AI e formaton (.docx, .xlsx, .pdf, .txt, .csv)"

uploaded_file = None
questionnaire_kind = SOURCE_TAGGED
if workflow == WORKFLOW_GENERATE:
    questionnaire_kind = st.radio("Pyetësori që do të ngarkosh:", [SOURCE_TAGGED, SOURCE_PLAIN], index=0)
    uploaded_file = st.file_uploader(
        "Zgjidh pyetësorin:",
        type=["docx"] if questionnaire_kind == SOURCE_TAGGED else qai.SOURCE_TYPES)

STRUCTURE_TAGS = {
    "group": "group", "end group": "end group", "end_group": "end group",
    "end repeat": "end repeat", "end_repeat": "end repeat",
    "section": "section", "seksion": "section",
}
STRUCTURE_TYPES = ("group", "end group", "repeat", "end repeat", "languages", "section")
# Answer types that can be followed by special-code options (e.g. Don't know [code: -98])
VALUE_TYPES = {"numeric": "integer", "decimal": "decimal", "date": "date", "time": "time", "text": "text", "string": "text"}
LANG_SEP = "||"
LANGUAGE_CODES = {
    "english": "en", "albanian": "sq", "shqip": "sq", "serbian": "sr", "srpski": "sr", "macedonian": "mk",
    "turkish": "tr", "bosnian": "bs", "croatian": "hr", "montenegrin": "cnr", "romani": "rom",
    "german": "de", "french": "fr", "italian": "it", "greek": "el", "spanish": "es",
}
# Texts the program adds itself; languages without an entry get the Albanian text
FIXED_TEXTS = {
    "gps": {"sq": "GPS", "en": "GPS", "sr": "GPS"},
    "enumerator": {"sq": "Anketuesi/ja", "en": "Enumerator", "sr": "Anketar/ka"},
    "full_name": {"sq": "Emri dhe mbiemri:", "en": "Full name:", "sr": "Ime i prezime:"},
    "phone": {"sq": "Numri i telefonit:", "en": "Phone number:", "sr": "Broj telefona:"},
    "enter_value": {"sq": "Shëno vlerën", "en": "Enter value", "sr": "Unesite vrednost"},
    "exclusive": {"sq": "Ky opsion nuk mund të zgjidhet bashkë me opsione të tjera.",
                  "en": "This option cannot be selected together with other options.",
                  "sr": "Ova opcija se ne može izabrati zajedno sa drugim opcijama."},
}
RANKING_LABELS_EN = [
    "First choice", "Second choice", "Third choice", "Fourth choice", "Fifth choice", "Sixth choice",
    "Seventh choice", "Eighth choice", "Ninth choice", "Tenth choice", "Eleventh choice", "Twelfth choice",
    "Thirteenth choice", "Fourteenth choice", "Fifteenth choice", "Sixteenth choice", "Seventeenth choice",
    "Eighteenth choice", "Nineteenth choice", "Twentieth choice", "Extra"
]
EXCLUSIVE_CODES = ("-97", "-98", "-99")
CODING_ORIGINAL = "Ruaj numërimin origjinal si në Word (A1, B2a, C1, …)"
CODING_VARIABLES = "Emrat e variablave nga pyetësori (p.sh. hh_study_child_confirm)"

def parse_languages(lines):
    """Languages declared with a [languages: English, Albanian] line; [] for a single-language questionnaire."""
    for line in lines:
        m = re.match(r'^\s*\[languages?:\s*(.+?)\]\s*$', line, flags=re.IGNORECASE)
        if m:
            langs = [l.strip() for l in re.split(r'\|\||,', m.group(1)) if l.strip()]
            return langs if len(langs) > 1 else []
    return []

def language_column(name):
    code = LANGUAGE_CODES.get(name.strip().lower())
    return f"{name.strip()} ({code})" if code else name.strip()

def option_tag(text, tag):
    """Value of an option tag such as [code: -98] or [name: x]; None when absent."""
    m = re.search(r'\[\s*' + tag + r'\s*:\s*([^\]]+?)\s*\]', text, flags=re.IGNORECASE)
    return m.group(1) if m else None

def has_flag(text, flag):
    return re.search(r'\[\s*' + flag + r'\s*\]', text, flags=re.IGNORECASE) is not None

def strip_option_tags(text):
    return re.sub(r'\s*\[\s*(code\s*:[^\]]*|skip\s*:[^\]]*|exclusive|value)\s*\]\s*', ' ', text, flags=re.IGNORECASE).strip()

def xls_name(value):
    """Makes a string usable as an XLSForm name / choice name."""
    value = re.sub(r'[^\w.\-]+', '_', value.strip()).strip('_')
    return value if re.match(r'^[A-Za-z_]', value) else f"_{value}"

def sanitize_name(label):
    return re.sub(r'\W+', '_', label.lower().strip())[:30]

def extract_tags(text):
    """Extract all bracketed tags like [random], [hint: ...], [single], [scale ...] etc."""
    return re.findall(r'\[(.*?)\]', text, flags=re.IGNORECASE)

def parse_question_tags(tags):
    """Classify tags into type, parameters, and hint."""
    q_type = None
    matrix_count = None
    hint = None
    parameters = None

    for raw_tag in tags:
        tag = raw_tag.strip().lower()

        # Randomization
        if tag == "random":
            parameters = "randomize=true"

        # Hint tag
        elif tag.startswith("hint:"):
            hint = raw_tag.split(":", 1)[1].strip()

        # Matrix type
        elif tag.startswith("matrix"):
            m = re.match(r"matrix\s+(single|multiple)\s+(\d+)", tag)
            if m:
                q_type = f"matrix {m.group(1)}"
                matrix_count = int(m.group(2))

        # Ranking type
        elif tag.startswith("ranking"):
            m = re.match(r"ranking\s+(\d+)", tag)
            if m:
                q_type = f"ranking {m.group(1)}"
                matrix_count = int(m.group(1))

        # Scale type
        elif tag.startswith("scale"):
            # Match on the raw tag so the min/max labels keep their original capitalization
            m = re.match(r"scale\s*(\d+)(?:\((.*?)\))?\s*-\s*(\d+)(?:\((.*?)\))?", raw_tag.strip(), flags=re.IGNORECASE)
            if m:
                start, min_label, end, max_label = m.groups()
                q_type = f"scale {start}-{end}"
                matrix_count = {
                    "start": int(start),
                    "end": int(end),
                    "min_label": min_label,
                    "max_label": max_label
                }

        # Structure lines: [group] / [end group], [repeat: ROST0] / [end repeat], [languages: English, Albanian]
        elif tag in STRUCTURE_TAGS:
            q_type = STRUCTURE_TAGS[tag]
        elif tag.startswith("repeat"):
            q_type = "repeat"
            matrix_count = raw_tag.split(":", 1)[1].strip() if ":" in raw_tag else None
        elif tag.startswith("languages:") or tag.startswith("language:"):
            q_type = "languages"

        # Generic question types
        elif tag in ["single", "multiple", "text", "string", "numeric", "decimal", "date", "time", "note", "other"]:
            q_type = tag

    return q_type, matrix_count, parameters, hint

def strip_type(text):
    return re.sub(r'\s*\[.*?\]\s*', '', text).strip()

def extract_question_number_and_text(line):
    match = re.match(r'^([A-Z]+\d+[a-zA-Z\.]*|\d+)[\.\)]?\s*(.+)', line.strip())
    if match:
        number = match.group(1)
        text = match.group(2)
        text = re.sub(r'[\|_]+', '', text).strip()
        text = re.sub(r'\s{2,}', ' ', text).strip()
        return number, text
    return None, line

def clean_label_prefix(text):
    text = re.sub(r'^[\(\[]?[a-zA-Z0-9]+[\.\)\]]\s*', '', text)
    text = re.sub(r'[?:]+', '', text)
    text = re.sub(r'[_\s]{2,}', '', text)
    return text.strip()

def has_random_tag(text):
    return "[random]" in text.lower()

def is_section_header(line):
    """Lines that start with 'Section' or 'Seksion' are section headers (they open a group)."""
    return re.match(r'^\s*(section|seksion)\b', line, flags=re.IGNORECASE) is not None

def line_type(line):
    """Type of a line for continuation loops: options/rows/note text stop at any tagged line or section header."""
    if is_section_header(line):
        return "section"
    return parse_question_tags(extract_tags(line))[0]

def find_orphan_lines(lines):
    """Lines that the main parser would silently drop: no tag, no recognized type,
    not a section header, and not consumed as an option/row/note continuation by a
    preceding tagged question. Mirrors the index advancement of generate_xlsform."""
    orphans = []
    i = 0
    while i < len(lines):
        line = lines[i]

        if is_section_header(line):
            i += 1
            continue

        tags = extract_tags(line)
        q_type, matrix_count, _, _ = parse_question_tags(tags)

        if not q_type:
            # Ignore visually-empty lines (just underscores, dashes, pipes, whitespace)
            if re.sub(r'[\s_\-\|]+', '', line):
                orphans.append(line)
            i += 1
            continue

        if q_type == "other":
            i += 1
            continue

        if q_type in VALUE_TYPES:
            # Special-code options (Don't know [code: -98]) may follow a numeric/date/time question
            i += 1
            while i < len(lines) and not line_type(lines[i]) \
                    and (option_tag(lines[i], "code") is not None or has_flag(lines[i], "value")):
                i += 1
            continue

        if q_type == "note" or q_type in ("single", "multiple") or q_type.startswith("ranking"):
            # Consume continuation/options until next tagged line
            i += 1
            while i < len(lines):
                next_type = line_type(lines[i])
                if next_type:
                    break
                i += 1
            continue

        if "matrix" in q_type:
            i += 1
            if isinstance(matrix_count, int):
                i = min(i + matrix_count, len(lines))
            while i < len(lines):
                next_type = line_type(lines[i])
                if next_type:
                    break
                i += 1
            continue

        # text, numeric, string, scale → single-line, no continuation
        i += 1

    return orphans

def load_anketuesit_choices():
    # Merr kredencialet nga st.secrets
    gcp_info = st.secrets["gcp_service_account"]
    
    # Deklaro scope të qartë për Google Sheets
    scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"]
    
    # Krijo kredencialet me scope
    credentials = Credentials.from_service_account_info(gcp_info, scopes=scopes)
    
    # Autorizo me gspread
    gc = gspread.authorize(credentials)
    
    # Hap dokumentin dhe worksheet-in
    sheet = gc.open("Sistemi i mbledhjes te te dhenave / Janar - Dhjetor 2025").worksheet("lists")

    ids = sheet.col_values(5)[3:]    # Kolona E
    names = sheet.col_values(6)[3:]  # Kolona F

    # Mbaj vetëm rreshtat që kanë të dyja vlerat jo bosh
    choices = []
    if ids is not None and names is not None:
        for id_, name in zip(ids, names):
            if id_ and name:
                choices.append({"list_name": "anketuesit_list","name": id_.strip(), "label": name.strip()})
    else:
        st.error("Gabim: Nuk u gjetën të dhëna në kolonat E dhe F.")

    return choices 

def generate_qname(qnum, q_index, coding_mode, var_name=None):
    """
    Returns (qname, updated_q_index)
    """

    # Variable names given in the questionnaire ([name: ...]) win when the user chose them
    if coding_mode == CODING_VARIABLES and var_name:
        return xls_name(var_name), q_index

    # Case 1: No number in Word → always P1, P2…
    if not qnum:
        return f"P{q_index}", q_index + 1

    qnum_clean = qnum.rstrip('.')

    # Case 2: D-questions → ALWAYS preserved
    if qnum_clean.upper().startswith("D"):
        return qnum_clean, q_index

    # Case 3: User wants original numbering (also the fallback for questions without a variable name)
    if coding_mode in (CODING_ORIGINAL, CODING_VARIABLES):
        return qnum_clean, q_index

    # Case 4: User wants Q1, Q2…
    if coding_mode == "Q1, Q2, Q3, ...":
        return f"Q{q_index}", q_index + 1

    # Case 5: Default → P1, P2…
    return f"P{q_index}", q_index + 1

def generate_xlsform(input_docx, output_xlsx, coding_mode, data_method=True, selected_questions=None, lines=None,
                     warnings=None, routing=None):
    """warnings: optional list that collects problems worth showing to the user (e.g. repeated answer codes).
    routing: optional dict filled with the routing captured by [ask if: ...] / [skip: ...] tags, for the AI filter
    step and its coverage tests: {"ask_if": {row: text}, "skips": [{source, code, target}], "refs": {id: row}}."""
    ranking_labels = [
        "Zgjedhja e parë", "Zgjedhja e dytë", "Zgjedhja e tretë",
        "Zgjedhja e katërt", "Zgjedhja e pestë", "Zgjedhja e gjashtë",
        "Zgjedhja e shtatë", "Zgjedhja e tetë", "Zgjedhja e nëntë",
        "Zgjedhja e dhjetë", "Zgjedhja e njëmbëdhjetë", "Zgjedhja e dymbëdhjetë",
        "Zgjedhja e trembëdhjetë", "Zgjedhja e katërmbëdhjetë", "Zgjedhja e pesëmbëdhjetë",
        "Zgjedhja e gjashtëmbëdhjetë", "Zgjedhja e shtatëmëdhjetë", "Zgjedhja e tetëmbëdhjetë",
        "Zgjedhja e nëntëmbëdhjetë", "Zgjedhja e njëzet", "Ekstra"
    ]

    # lines can be passed directly (AI-formatted questionnaire) instead of reading a .docx
    if lines is None:
        doc = docx2python(input_docx)
        lines = [line.strip() for line in doc.text.split('\n') if line.strip()]

    # Multi-language questionnaires: texts are separated with "||" in the declared language order.
    # Labels are kept as one entry per language and written as label::<language> columns at the end.
    languages = parse_languages(lines)
    lang_columns = [language_column(l) for l in languages]
    lang_codes = [LANGUAGE_CODES.get(l.strip().lower(), "") for l in languages] or ["sq"]
    n_lang = max(1, len(languages))

    def ml(text):
        """Splits a text into one entry per language; missing translations reuse the first language."""
        if n_lang == 1:
            return [text]
        parts = [p.strip() for p in text.split(LANG_SEP)]
        return parts[:n_lang] + [parts[0]] * (n_lang - len(parts))

    def fixed(key):
        return [FIXED_TEXTS[key].get(code, FIXED_TEXTS[key]["sq"]) for code in lang_codes]

    def clean_segment(text, qnum):
        """Question text of a non-first language: same cleaning as extract_question_number_and_text."""
        if qnum:
            text = re.sub(r'^' + re.escape(qnum) + r'[\.\)]?\s*', '', text.strip())
        text = re.sub(r'[\|_]+', '', text).strip()
        return re.sub(r'\s{2,}', ' ', text).strip()

    survey = []
    choices = []
    skipped_other_questions = []
    settings = [{'style': 'theme-grid no-text-transform'}]
    if lang_columns:
        settings[0]["default_language"] = lang_columns[0]

    survey.append({
       "type": "start",
        "name": "start"
    })

    survey.append({
       "type": "end",
        "name": "end"
    })

    if data_method:
        survey.append({
            "type": "geopoint",
            "name": "GPS",
            "label": fixed("gps"),
            "required": "true"
    })

    # Add Anketuesi_ja question
    survey.append({
       "type": "select_one anketuesit_list",
        "name": "Anketuesi_ja",
        "label": fixed("enumerator"),
        "required": "true",
        "appearance": "search"
    })
     # Add the dynamic choices
    try:
        anketuesit_choices = load_anketuesit_choices()
        choices.extend(anketuesit_choices)
    except Exception as e:
        raise RuntimeError(f"Gabim gjatë ngarkimit të listës së anketuesve: {e}")

    i = 0
    q_index = 1
    note_index = 1
    prev_i = -1
    # Open blocks: (kind, name) with kind "group" ([group] modules), "section" (section headers,
    # written as groups that close at the next section) or "repeat"
    structure_stack = []
    structure_index = 0
    qname_by_ref = {}      # question number / variable name → form name, for [repeat: ROST0]

    used_codes = {}   # list_name -> choice names already used

    def unique_code(list_name, code, qname, label):
        """Choice names must be unique within a list (Kobo rejects the form otherwise). A repeated code
        from the questionnaire gets the next free number and a warning for the user."""
        used = used_codes.setdefault(list_name, set())
        if code in used:
            numbers = [int(c) for c in used if re.match(r'^-?\d+$', c)]
            new_code = str(max(numbers + [0]) + 1)
            while new_code in used:
                new_code = str(int(new_code) + 1)
            if warnings is not None:
                warnings.append(f"Kodi '{code}' përsëritet te pyetja {qname}: opsioni '{label[0]}' mori kodin "
                                f"'{new_code}'. Kontrollo kodet në pyetësor.")
            code = new_code
        used.add(code)
        return code

    route = routing if routing is not None else {}
    route.update(ask_if={}, skips=[], refs={})

    def open_block(kind, label, repeat_count=None):
        nonlocal structure_index
        structure_index += 1
        name = f"{kind}_{structure_index}"
        row = {"type": "begin_repeat" if kind == "repeat" else "begin_group", "name": name, "label": label}
        if repeat_count:
            row["repeat_count"] = repeat_count
        survey.append(row)
        structure_stack.append((kind, name))
        return name

    def close_block():
        kind, name = structure_stack.pop()
        if survey and survey[-1].get("name") == name:
            survey.pop()   # nothing inside: drop the empty group instead of writing it
        else:
            survey.append({"type": "end_repeat" if kind == "repeat" else "end_group", "name": f"{name}_end"})

    def close_sections():
        while structure_stack and structure_stack[-1][0] == "section":
            close_block()

    while i < len(lines):
        if i == prev_i:
            raise ValueError(
                f"Formatimi i Word dokumentit nuk u njoh në linjën: '{lines[i]}'. "
                f"Pyetja mund të mos ketë tag të vlefshëm (p.sh. [single], [text]) ose numërim, "
                f"duke shkaktuar bllokim të procesit."
            )
        prev_i = i
        line = lines[i]

        # STEP 1: Extract all tags (like [single], [random], [hint:...])
        tags = extract_tags(line)

        # STEP 2: Parse those tags to understand type, hint, and randomization
        q_type, matrix_count, parameters, hint = parse_question_tags(tags)

        # Sections ([section] tag or a line starting with "Section" / "Seksion") become groups labelled
        # with the section title; a section ends at the next section or when its module/repeat ends
        if q_type == "section" or is_section_header(line):
            close_sections()
            label = ml(strip_type(line))
            if label[0].strip():
                name = open_block("section", label)
                if option_tag(line, "ask if"):
                    route["ask_if"][name] = option_tag(line, "ask if")
            i += 1
            continue

        # Structure lines: languages declaration, groups and repeats
        if q_type == "languages":
            i += 1
            continue
        if q_type in ("group", "repeat"):
            if q_type == "group":
                close_sections()   # a new module ends the previous module's last section
            repeat_count = None
            if q_type == "repeat" and matrix_count and matrix_count in qname_by_ref:
                repeat_count = f"${{{qname_by_ref[matrix_count]}}}"
            name = open_block(q_type, ml(strip_type(line)), repeat_count)
            if option_tag(line, "ask if"):
                route["ask_if"][name] = option_tag(line, "ask if")
            i += 1
            continue
        if q_type in ("end group", "end repeat"):
            kind = q_type.split()[1]
            close_sections()
            if structure_stack and structure_stack[-1][0] == kind:
                close_block()
            i += 1
            continue

        # STEP 3: Remove all tags from text so only question text remains (one entry per language)
        segments = ml(strip_type(line))
        full_line = segments[0]

        # STEP 4: Extract the question number and text
        qnum, label_text = extract_question_number_and_text(full_line)

            # Skip if q_type is "other"
        if q_type == "other":
            # Collect label for display
            if label_text:
                skipped_other_questions.append(label_text)
            i += 1
            continue

        if q_type:
            if selected_questions is not None and label_text in selected_questions:
                i += 1
                continue

            if qnum:
                qnum = re.sub(r'\.\.+', '.', qnum).rstrip('.')

            texts = [label_text] + [clean_segment(s, qnum) for s in segments[1:]]
            label = [f"{qnum}. {t}" for t in texts] if qnum else segments
            hint = ml(hint) if hint else None
            var_name = option_tag(line, "name")

            qname, q_index = generate_qname(qnum, q_index, coding_mode, var_name)

            qname = qname.rstrip('.')
            required = "true"
            if qnum:
                qname_by_ref[qnum] = qname
            if var_name:
                qname_by_ref[var_name] = qname

            first_row_index = len(survey)   # the question's first form row gets its routing

            def add_common_question(fields):
                if parameters:
                    fields["parameters"] = parameters
                if hint:
                    fields["hint"] = hint
                survey.append(fields)

            def collect_options(start_index):
                opts = []
                while start_index < len(lines):
                    next_type = line_type(lines[start_index])

                    if next_type:
                        break
                    opts.append(lines[start_index])
                    start_index += 1
                return opts, start_index

            def option_labels(opt):
                return [clean_label_prefix(s) for s in ml(strip_option_tags(opt))]

            if q_type == "note":
                note_parts = [segments]
                i += 1
                while i < len(lines):
                    next_type = line_type(lines[i])
                    if next_type:
                        break
                    note_parts.append(ml(lines[i].strip()))
                    i += 1

                note_label = ["\n".join(p[k].strip() for p in note_parts if p[k].strip()) for k in range(n_lang)]
                survey.append({
                    "type": "note",
                    "name": f"note{note_index}",
                    "label": note_label
                })
                note_index += 1

            elif q_type in ["single", "multiple"]:
                list_name = qname + "_list"
                qstyle = "select_one" if q_type == "single" else "select_multiple"
                question = {
                    "type": f"{qstyle} {list_name}",
                    "name": qname,
                    "label": label,
                    "required": required
                }
                add_common_question(question)

                i += 1
                options, i = collect_options(i)
                if not options:
                    # e.g. "[PRELOADED 15 options]": Kobo rejects a list question without choices
                    question["type"] = "text"
                    if warnings is not None:
                        warnings.append(f"Pyetja {qname} nuk ka opsione (listë e para-ngarkuar?) dhe u kodua si "
                                        f"tekst. Shtoni listën e opsioneve manualisht nëse duhet.")
                exclusive = []
                for idx, opt in enumerate(options, 1):
                    clean = option_labels(opt)
                    code = option_tag(opt, "code")
                    if code is not None:
                        name_value = xls_name(code) if not re.match(r'^-?\d+$', code) else code
                    else:
                        name_value = f"_{idx}" if q_type == "multiple" else str(idx)
                    name_value = unique_code(list_name, name_value, qname, clean)
                    if option_tag(opt, "skip"):
                        route["skips"].append({"source": qname, "code": name_value, "target": option_tag(opt, "skip")})
                    choices.append({
                        "list_name": list_name,
                        "name": name_value,
                        "label": clean
                    })
                    if has_flag(opt, "exclusive") or name_value in EXCLUSIVE_CODES:
                        exclusive.append(name_value)
                    if '_' in strip_option_tags(opt):
                        open_name = f"{qname}_{idx}"
                        relevant_expr = f"selected(${{{qname}}}, '{name_value}')" if q_type == "multiple" else f"${{{qname}}} = '{name_value}'"
                        survey.append({
                            "type": "text",
                            "name": open_name,
                            "label": clean,
                            "relevant": relevant_expr,
                            "required": "true"
                        })
                # "Don't know", "Refused", "No one" … cannot be combined with other answers
                if q_type == "multiple" and exclusive:
                    question["constraint"] = "count-selected(.) = 1 or (" + " and ".join(
                        f"not(selected(., '{n}'))" for n in exclusive) + ")"
                    question["constraint_message"] = fixed("exclusive")

            elif q_type in VALUE_TYPES:
                i += 1
                special = []
                while i < len(lines) and not line_type(lines[i]) \
                        and (option_tag(lines[i], "code") is not None or has_flag(lines[i], "value")):
                    special.append(lines[i])
                    i += 1

                if not special:
                    add_common_question({
                        "type": VALUE_TYPES[q_type],
                        "name": qname,
                        "label": label,
                        "required": required
                    })
                else:
                    # Value + special codes: a select_one "<name>_type" (value / Don't know / Refused …)
                    # and the value field itself, shown when the value option is chosen
                    list_name = f"{qname}_type_list"
                    add_common_question({
                        "type": f"select_one {list_name}",
                        "name": f"{qname}_type",
                        "label": label,
                        "required": required
                    })
                    codes = [option_tag(o, "code") for o in special]
                    value_opts = [o for o in special if has_flag(o, "value")]
                    if value_opts:
                        value_code = option_tag(value_opts[0], "code") or "1"
                        value_label = option_labels(value_opts[0])
                    else:
                        value_code = "1" if "1" not in codes else "value"
                        value_label = label
                        choices.append({"list_name": list_name, "name": unique_code(list_name, value_code, qname, label),
                                        "label": fixed("enter_value")})
                    for opt in special:
                        if value_opts and opt is value_opts[0]:
                            code = unique_code(list_name, value_code, qname, value_label)
                        else:
                            code = unique_code(list_name, option_tag(opt, "code") or value_code, qname, option_labels(opt))
                        if option_tag(opt, "skip"):
                            route["skips"].append({"source": f"{qname}_type", "code": code, "target": option_tag(opt, "skip")})
                        choices.append({"list_name": list_name, "name": code, "label": option_labels(opt)})
                    survey.append({
                        "type": VALUE_TYPES[q_type],
                        "name": qname,
                        "label": value_label,
                        "required": required,
                        "relevant": f"${{{qname}_type}} = '{value_code}'"
                    })

            elif q_type in ["text", "string"]:
                add_common_question({
                    "type": "text",
                    "name": qname,
                    "label": label,
                    "required": required
                })
                i += 1

            elif q_type.startswith("scale") and isinstance(matrix_count, dict):
                start = matrix_count["start"]
                end = matrix_count["end"]
                min_label = ml(matrix_count["min_label"]) if matrix_count.get("min_label") else None
                max_label = ml(matrix_count["max_label"]) if matrix_count.get("max_label") else None

                list_name = f"scale_{start}_{end}"
                question = {
                    "type": f"select_one {list_name}",
                    "name": qname,
                    "label": label,
                    "required": required,
                    "appearance": "likert"
                }
                add_common_question(question)

                if not any(c["list_name"] == list_name for c in choices):
                    for j in range(start, end + 1):
                        lbl = [f"{j} - {min_label[k]}" if j == start and min_label else
                               f"{j} - {max_label[k]}" if j == end and max_label else str(j)
                               for k in range(n_lang)]
                        choices.append({
                            "list_name": list_name,
                            "name": str(j),
                            "label": lbl
                        })
                i += 1

            elif "matrix" in q_type:
                style = "select_one" if "single" in q_type else "select_multiple"
                list_name = qname + "_matrix"
                i += 1

                columns = lines[i:i + matrix_count]
                i += matrix_count

                rows = []
                while i < len(lines):
                    next_type = line_type(lines[i])
                    if next_type:
                        break
                    rows.append(lines[i])
                    i += 1

                survey.append({"type": "begin_group", "name": f"{qname}_group", "appearance": "field-list", "required": "false"})
                survey.append({"type": f"{style} {list_name}", "name": f"{qname}_matrix_label", "label": label, "appearance": "label", "required": "false"})

                for idx, row in enumerate(rows, 1):
                    field = {
                        "type": f"{style} {list_name}",
                        "name": f"{qname}_{idx}",
                        "label": ml(row),
                        "appearance": "list-nolabel",
                        "required": "true"
                    }
                    if parameters:
                        field["parameters"] = parameters
                    survey.append(field)

                survey.append({"type": "end_group", "name": f"{qname}_group_end"})

                for j, col in enumerate(columns, 1):
                    code = option_tag(col, "code")
                    col_label = ml(strip_option_tags(col))
                    choices.append({"list_name": list_name,
                                    "name": unique_code(list_name, code if code is not None else str(j), qname, col_label),
                                    "label": col_label})

            elif q_type.startswith("ranking"):
                match = re.findall(r"\d+", q_type)
                if not match:
                    raise ValueError(
                        f"Tag-u [ranking] në linjën '{line}' nuk ka numër (p.sh. [ranking 5])."
                    )
                if match:
                    rank_count = int(match[0])
                    list_name = qname + "_list"

                    survey.append({"type": "begin_group", "name": f"{qname}_group", "appearance": "field-list"})
                    survey.append({"type": "note", "name": f"{qname}_label", "label": label})

                    for idx in range(1, rank_count + 1):
                        rank_name = f"{qname}_{idx}"
                        rank_label = [
                            (RANKING_LABELS_EN if code == "en" else ranking_labels)[min(idx, 21) - 1]
                            for code in lang_codes
                        ]
                        survey.append({
                            "type": f"select_one {list_name}",
                            "name": rank_name,
                            "label": rank_label,
                            "required": "true",
                            "appearance": "minimal",
                            "choice_filter": " and ".join([f"not(selected(${{{qname}_{j}}}, name))" for j in range(1, idx)])
                        })

                    survey.append({"type": "end_group", "name": f"{qname}_group_end"})

                    i += 1
                    options = []
                    while i < len(lines):
                        next_type = line_type(lines[i])
                        if next_type:
                            break
                        options.append(lines[i])
                        i += 1

                    for idx, opt in enumerate(options, 1):
                        code = option_tag(opt, "code")
                        opt_label = option_labels(opt)
                        choices.append({"list_name": list_name,
                                        "name": unique_code(list_name, code if code is not None else str(idx), qname, opt_label),
                                        "label": opt_label})

            else:
                raise ValueError(
                    f"Tipi i pyetjes '{q_type}' nuk u njoh në linjën: '{line}'. "
                    f"Kontrollo formatimin e tag-ut (p.sh. [scale 1-5], [matrix single 3])."
                )

            if len(survey) > first_row_index:
                first_row = survey[first_row_index]["name"]
                for ref in (qnum, var_name):
                    if ref:
                        route["refs"][ref] = first_row
                if option_tag(line, "ask if"):
                    route["ask_if"][first_row] = option_tag(line, "ask if")
        else:
            i += 1

    # Close sections/groups/repeats left open in the document
    while structure_stack:
        close_block()

    survey.append({"type": "text", "name": "emri_mbiemri", "label": fixed("full_name"), "required": "true"})
    survey.append({"type": "text", "name": "numri_telefonit", "label": fixed("phone"), "required": "true"})

    def localize(rows):
        """Writes per-language texts as label / label::<language> columns, keeping the column order."""
        out = []
        for row in rows:
            new = {}
            for key, value in row.items():
                if key in ("label", "hint", "constraint_message"):
                    values = value if isinstance(value, list) else [value] * n_lang
                    if lang_columns:
                        for col, v in zip(lang_columns, values):
                            new[f"{key}::{col}"] = v
                    else:
                        new[key] = values[0]
                else:
                    new[key] = value
            out.append(new)
        return out

    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        pd.DataFrame(localize(survey)).to_excel(writer, sheet_name="survey", index=False)
        if choices:
            pd.DataFrame(localize(choices)).to_excel(writer, sheet_name="choices", index=False)
        pd.DataFrame(settings).to_excel(writer, sheet_name="settings", index=False)

    return skipped_other_questions


def has_tags(lines):
    """True when the document is already formatted with question tags ([single], [text], ...)."""
    return any(parse_question_tags(extract_tags(line))[0] for line in lines)


# ---------------------------------------------------------------------------
# AI helpers (questionnaire_ai.py holds the Claude calls and the filter tests)
# ---------------------------------------------------------------------------

def get_claude_client():
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return None
    try:
        qai.MODEL = st.secrets.get("CLAUDE_MODEL", qai.MODEL)   # optional, e.g. claude-sonnet-5 (cheaper)
    except Exception:
        pass
    return qai.make_client(api_key)


def reset_state_for_file(prefix, file_key):
    """Clears the results of a workflow when a different file is uploaded."""
    if st.session_state.get(prefix + "file_key") != file_key:
        for k in [k for k in st.session_state if k.startswith(prefix)]:
            del st.session_state[k]
        st.session_state[prefix + "file_key"] = file_key


def show_test_report(report):
    if report["tested"] == 0:
        st.info("Formulari nuk ka asnjë filtër (kolona `relevant` është bosh).")
        return
    if not report["errors"] and not report["warnings"]:
        st.success(f"Të gjithë filtrat ({report['tested']}) kaluan testet automatike.")
    else:
        st.warning(f"U testuan {report['tested']} filtra: {report['errors']} gabime, "
                   f"{report['warnings']} paralajmërime.")
    def issues_table(issues):
        st.dataframe(pd.DataFrame([{
            "Rreshti": i["row"],
            "Pyetja": i["name"],
            "Niveli": qai.SEVERITY_LABELS[i["severity"]],
            "Filtri": i["relevant"],
            "Problemi": i["message"],
        } for i in issues]), hide_index=True)

    problems = [i for i in report["issues"] if i["severity"] != "info"]
    infos = [i for i in report["issues"] if i["severity"] == "info"]
    if problems:
        issues_table(problems)
    if infos:
        with st.expander(f"Shënime nga testet ({len(infos)})"):
            issues_table(infos)


def show_notes(title, notes):
    if notes:
        st.info(f"**{title}**\n\n" + "\n".join(f"- {n}" for n in notes))


def add_ai_filters(xlsx_path, source_blocks, routing=None):
    """Adds skip logic with Claude, tests it and gives Claude one round to fix failing filters.

    The form is always returned: if the AI step fails, it comes back without filters and an error message.
    """
    with open(xlsx_path, "rb") as f:
        data = f.read()
    result = {"xlsx": data, "applied": [], "rejected": [], "notes": [], "report": None, "cost": 0.0, "error": None}
    client = get_claude_client()
    if client is None:
        result["error"] = "Mungon çelësi `ANTHROPIC_API_KEY` – formulari u gjenerua pa filtra."
        return result
    sheets = qai.read_form(data)
    cost = 0.0
    try:
        with st.spinner("Claude po shton filtrat sipas pyetësorit..."):
            filters, notes, usage = qai.generate_filters(client, source_blocks, sheets, routing=routing)
            cost += qai.estimate_cost(usage)
            applied, rejected = qai.apply_filters(sheets, filters)
        with st.spinner("Po testohen filtrat..."):
            report = qai.test_form(sheets, routing)
            # errors, plus routing from the questionnaire that has no filter yet, get one repair round
            failed = [i for i in report["issues"] if i["severity"] == "error" or
                      (i.get("kind") == "routing" and i["severity"] == "warning")]
            if failed:
                fixes, fix_notes, usage = qai.generate_filters(client, source_blocks, sheets, failed_tests=failed,
                                                               routing=routing)
                cost += qai.estimate_cost(usage)
                fixed, fix_rejected = qai.apply_filters(sheets, fixes, combine_existing=False)
                fixed_targets = {f["target"] for f in fixed}
                applied = [a for a in applied if a["target"] not in fixed_targets] + fixed
                rejected += fix_rejected
                notes += fix_notes
                report = qai.test_form(sheets, routing)
    except qai.AIError as e:
        result["error"] = f"{e} Formulari u gjenerua pa filtra."
        result["cost"] = cost
        return result
    result.update(xlsx=qai.write_form(sheets), applied=applied, rejected=rejected, notes=notes,
                  report=report, cost=cost)
    return result


def show_filter_results(result):
    if result["error"]:
        st.warning(result["error"])
        return
    st.markdown("**Filtrat e shtuar nga AI**")
    if result["applied"]:
        st.dataframe(pd.DataFrame([{"Pyetja": a["target"], "Filtri": a["relevant"],
                                    "Bazuar në": a["instruction"]} for a in result["applied"]]), hide_index=True)
    else:
        st.info("AI nuk gjeti udhëzime filtrimi në pyetësor.")
    if result["rejected"]:
        st.warning("Këta filtra nuk u vendosën:")
        st.dataframe(pd.DataFrame([{"Pyetja": r["target"], "Filtri": r["relevant"], "Arsyeja": r["reason"]}
                                   for r in result["rejected"]]), hide_index=True)
    show_notes("Shënime nga AI për filtrat:", result["notes"])
    st.markdown("**Testimi i filtrave**")
    show_test_report(result["report"])


def render_filter_check():
    xls_file = st.file_uploader("Ngarko formularin XLS (.xlsx):", type=["xlsx"], key="chk_xlsx_upload")
    source_file = st.file_uploader(
        "Ngarko pyetësorin origjinal, me të cilin krahasohen filtrat (.docx, .xlsx, .pdf, .txt, .csv):",
        type=qai.SOURCE_TYPES, key="chk_source_upload")
    if not xls_file or not source_file:
        return
    xls_bytes = xls_file.getvalue()
    source_bytes = source_file.getvalue()
    reset_state_for_file("chkres_", hashlib.sha1(xls_bytes + b"|" + source_bytes).hexdigest())

    if st.button("Kontrollo filtrat"):
        try:
            sheets = qai.read_form(xls_bytes)
            source_blocks = qai.load_source(source_file.name, source_bytes)
            with st.spinner("Po testohen filtrat..."):
                report = qai.test_form(sheets)
            review, corrected, corrected_report, cost = None, None, None, 0.0
            client = get_claude_client()
            if client is None:
                st.warning("Mungon çelësi `ANTHROPIC_API_KEY` – u kryen vetëm testet automatike, pa kontrollin nga AI.")
            else:
                with st.spinner("Claude po kontrollon filtrat..."):
                    review, usage = qai.review_filters(client, sheets, report, source_blocks)
                    cost = qai.estimate_cost(usage)
                fixes = [i for i in review["issues"] if i["action"] in ("replace", "remove")]
                if fixes:
                    fixed_sheets = qai.read_form(xls_bytes)
                    qai.apply_suggestions(fixed_sheets, fixes)
                    corrected_report = qai.test_form(fixed_sheets)
                    corrected = qai.write_form(fixed_sheets)
        except qai.AIError as e:
            st.error(str(e))
            return
        st.session_state["chkres_report"] = report
        st.session_state["chkres_review"] = review
        st.session_state["chkres_corrected"] = corrected
        st.session_state["chkres_corrected_report"] = corrected_report
        st.session_state["chkres_cost"] = cost

    if "chkres_report" not in st.session_state:
        return

    st.markdown("**Testet automatike**")
    show_test_report(st.session_state["chkres_report"])

    review = st.session_state["chkres_review"]
    if review is not None:
        st.markdown("**Kontrolli nga AI**")
        st.write(review["summary"])
        if review["issues"]:
            st.dataframe(pd.DataFrame([{
                "Pyetja": i["name"],
                "Niveli": qai.SEVERITY_LABELS[i["severity"]],
                "Problemi": i["problem"],
                "Filtri i sugjeruar": i["suggested_relevant"] if i["action"] == "replace"
                else ("(hiq filtrin)" if i["action"] == "remove" else ""),
            } for i in review["issues"]]), hide_index=True)

    if st.session_state["chkres_corrected"]:
        st.markdown("**Formulari me korrigjimet e sugjeruara**")
        show_test_report(st.session_state["chkres_corrected_report"])
        st.download_button(
            label="Shkarko formularin e korrigjuar",
            data=st.session_state["chkres_corrected"],
            file_name=f"{os.path.splitext(xls_file.name)[0]}_korrigjuar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if st.session_state.get("chkres_cost"):
        st.caption(f"Kosto e përafërt e AI: ${st.session_state['chkres_cost']:.2f}")


if workflow == WORKFLOW_CHECK:
    render_filter_check()

if uploaded_file:
    uploaded_content = uploaded_file.getvalue()
    reset_state_for_file("gen_", hashlib.sha1(uploaded_content).hexdigest())
    base_name = os.path.splitext(uploaded_file.name)[0]

    try:
        source_blocks = qai.load_source(uploaded_file.name, uploaded_content)
    except qai.AIError as e:
        st.error(str(e))
        st.stop()

    # The user says which kind of questionnaire it is; the tags are only used for a sanity check
    doc_lines = None
    if uploaded_file.name.lower().endswith(".docx"):
        doc = docx2python(BytesIO(uploaded_content)).text
        doc_lines = [line.strip() for line in doc.split('\n') if line.strip()]

    lines = None
    if questionnaire_kind == SOURCE_TAGGED:
        if not doc_lines:
            st.error("Dokumenti nuk përmban tekst të lexueshëm.")
            st.stop()
        if not has_tags(doc_lines):
            st.error(
                "**Nuk u gjet asnjë tag në dokument** (p.sh. `[single]`, `[multiple]`, `[text]`, `[numeric]`). "
                "Nëse pyetësori nuk është i formatuar, zgjidh më lart opsionin "
                f"**'{SOURCE_PLAIN}'** që ta formatojë AI."
            )
            st.stop()
        lines = doc_lines

    if lines is None:
        if doc_lines and has_tags(doc_lines):
            st.info(f"Ky dokument duket se ka tag-e. Nëse është i formatuar tashmë, zgjidh më lart "
                    f"**'{SOURCE_TAGGED}'** për ta koduar drejtpërdrejt, pa kosto AI.")
        st.info("AI do ta formatojë pyetësorin në formatin e programit "
                "(llojet e pyetjeve, opsionet, seksionet).")
        if st.button("Formato pyetësorin me AI"):
            client = get_claude_client()
            if client is None:
                st.error("Mungon çelësi `ANTHROPIC_API_KEY` në secrets të aplikacionit.")
                st.stop()
            with st.spinner("Claude po lexon pyetësorin dhe po përcakton llojet e pyetjeve..."):
                try:
                    ai_lines, ai_notes, usage = qai.convert_questionnaire(client, source_blocks)
                except qai.AIError as e:
                    st.error(str(e))
                    st.stop()
            for k in [k for k in st.session_state if k.startswith("gen_") and k != "gen_file_key"]:
                del st.session_state[k]
            st.session_state["gen_tagged"] = "\n".join(ai_lines)
            st.session_state["gen_notes"] = ai_notes
            st.session_state["gen_format_cost"] = qai.estimate_cost(usage)
        if "gen_tagged" not in st.session_state:
            st.stop()

        st.markdown("**Pyetësori i formatuar nga AI.** Kontrollo llojet e pyetjeve dhe korrigjo nëse duhet:")
        tagged_text = st.text_area("Pyetësori i formatuar", key="gen_tagged", height=400,
                                   label_visibility="collapsed")
        lines = [line.strip() for line in tagged_text.split('\n') if line.strip()]
        show_notes("Shënime nga AI:", st.session_state.get("gen_notes"))
        st.download_button(
            label="Shkarko pyetësorin e formatuar (.docx)",
            data=qai.lines_to_docx(lines),
            file_name=f"{base_name}_formatuar.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    data_collection_method = st.selectbox(
    "Metoda e mbledhjes së të dhënave:",
    ["Face to face", "Telefon/Online"]
    )

    has_variable_names = any(option_tag(line, "name") for line in lines)
    coding_mode = st.radio(
    "Si të kodohen pyetjet që kanë numërim në Word?",
    options=[
        "P1, P2, P3, ...",
        "Q1, Q2, Q3, ...",
        CODING_ORIGINAL,
        CODING_VARIABLES
    ], index=3 if has_variable_names else 0)

    # Extract question numbers (e.g., 1, D1, 2a, Q1.2 etc.)
    question_options = []
    unnumbered_questions = []
    for line in lines:
        try:
            # Skip section headers entirely — they're treated as [other]
            if is_section_header(line):
                continue

            # Extract all tags from this line (e.g., [random][single][hint: ...])
            tags = extract_tags(line)
            q_type, _, _, _ = parse_question_tags(tags)

            # Only process lines that define a question type (skip [other] and [note] — they get filtered out / don't need numbers)
            if q_type and q_type not in ("other", "note") + STRUCTURE_TYPES:
                qnum, label_text = extract_question_number_and_text(strip_type(line))
                if label_text:
                    question_options.append(label_text)
                if not qnum and label_text:
                    unnumbered_questions.append(label_text)

        except ValueError as e:
            st.error(f"Gabim në rreshtin: **{line}**\n\n{str(e)}")
            st.stop()

    is_original_mode = coding_mode == CODING_ORIGINAL
    block_generation = False
    if unnumbered_questions:
        if is_original_mode:
            st.error(
                "**Janë gjetur pyetje pa numërim në dokumentin Word.**\n\n"
                "Ju keni zgjedhur modalitetin **'Ruaj numërimin origjinal'**, por pyetjet e mëposhtme nuk kanë numër "
                "dhe do të marrin emra automatikë (P1, P2, ...), duke krijuar një përzierje me numërimin origjinal.\n\n"
                "Ju lutemi shtoni numra në dokumentin Word për këto pyetje, ose zgjidhni një modalitet tjetër kodimi:"
            )
            for q in unnumbered_questions:
                st.markdown(f"- {q}")
            block_generation = True
        else:
            st.warning(
                "**Janë gjetur pyetje pa numërim.** Këto do të marrin emra automatikë (P{n} ose Q{n}):"
            )
            for q in unnumbered_questions:
                st.markdown(f"- {q}")

    orphan_lines = find_orphan_lines(lines)
    if orphan_lines:
        st.warning(
            "**Janë gjetur paragrafë pa tag dhe pa lidhje me ndonjë pyetje.** "
            "Këto rreshta do të injorohen plotësisht (nuk do të shfaqen në formular). "
            "Nëse janë pyetje, shto një tag (p.sh. `[single]`, `[text]`); nëse janë seksione, fillojini me `Section` ose `Seksion`, ose shtoni tag-un `[section]`:"
        )
        for o in orphan_lines:
            st.markdown(f"- {o}")

    st.session_state["question_lines"] = lines
    selected_questions = st.multiselect(
        "Zgjidh pyetjet që NUK dëshiron të kodosh:",
        options=question_options,
        default=None
    )
    st.session_state["selected_questions"] = selected_questions


    if data_collection_method:
        generate_button = st.button("Gjenero formularin XLS", disabled=block_generation)
        if generate_button:
            generated_name = f"{base_name}_gjeneruar.xlsx"
            temp_xlsx_path = os.path.join(tempfile.gettempdir(), generated_name)
            error = None
            generation_warnings = []
            routing = {}   # [ask if] / [skip] notes captured while generating, used by the filter step
            with st.spinner("Po përpunon dokumentin..."):
                data_method = data_collection_method == "Face to face"
                try:
                    skipped = generate_xlsform(None, temp_xlsx_path, coding_mode, data_method,
                                               st.session_state.get("selected_questions", None), lines=lines,
                                               warnings=generation_warnings, routing=routing)
                except Exception as e:
                    error = str(e)

            if error:
                st.error(f"Gabimi: {error}")
            else:
                filter_result = add_ai_filters(temp_xlsx_path, source_blocks, routing)
                st.session_state["gen_xlsx_data"] = filter_result["xlsx"]
                st.session_state["gen_xlsx_name"] = generated_name
                st.session_state["gen_skipped_other_questions"] = skipped
                st.session_state["gen_warnings"] = generation_warnings
                st.session_state["gen_filter_result"] = filter_result

        if st.session_state.get("gen_xlsx_data"):
            st.success("Formulari XLS u gjenerua me sukses!")
            st.download_button(
                label="Shkarko formularin XLS",
                data=st.session_state["gen_xlsx_data"],
                file_name=st.session_state["gen_xlsx_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if st.session_state.get("gen_warnings"):
                st.warning("**Kontrollo këto në formular:**\n\n"
                           + "\n".join(f"- {w}" for w in st.session_state["gen_warnings"]))
            if st.session_state.get("gen_skipped_other_questions"):
                st.info("Pyetjet me tag-un [other] që u anashkaluan:")
                for q in st.session_state["gen_skipped_other_questions"]:
                    st.markdown(f"- {q}")
            show_filter_results(st.session_state["gen_filter_result"])

    total_cost = st.session_state.get("gen_format_cost", 0.0)
    if st.session_state.get("gen_filter_result"):
        total_cost += st.session_state["gen_filter_result"]["cost"]
    if total_cost:
        st.caption(f"Kosto e përafërt e AI: ${total_cost:.2f}")
