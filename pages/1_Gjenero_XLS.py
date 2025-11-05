import streamlit as st
from docx2python import docx2python
import pandas as pd
import re
import tempfile
import os
from PIL import Image
from google.oauth2.service_account import Credentials
import gspread
from io import BytesIO


st.set_page_config(page_title="Gjenero XLS", layout="centered")

logo_path = "logo.png"
with st.sidebar:
    if os.path.exists(logo_path):
        st.image(Image.open(logo_path), width=150)

st.title("Gjenero XLS")
st.markdown("Ngarko dokumentin `.docx` dhe gjenero formularin XLS për përdorim në Kobo Toolbox.")

uploaded_file = st.file_uploader("Zgjidh një dokument `.docx` të formatuar:", type=["docx"])

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
            m = re.match(r"scale\s*(\d+)(?:\((.*?)\))?\s*-\s*(\d+)(?:\((.*?)\))?", tag)
            if m:
                start, min_label, end, max_label = m.groups()
                q_type = f"scale {start}-{end}"
                matrix_count = {
                    "start": int(start),
                    "end": int(end),
                    "min_label": min_label,
                    "max_label": max_label
                }

        # Generic question types
        elif tag in ["single", "multiple", "text", "string", "numeric", "note", "other"]:
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
  
def generate_xlsform(input_docx, output_xlsx, data_method=True, selected_questions=None):
    ranking_labels = [
        "Zgjedhja e parë", "Zgjedhja e dytë", "Zgjedhja e tretë",
        "Zgjedhja e katërt", "Zgjedhja e pestë", "Zgjedhja e gjashtë",
        "Zgjedhja e shtatë", "Zgjedhja e tetë", "Zgjedhja e nëntë",
        "Zgjedhja e dhjetë", "Zgjedhja e njëmbëdhjetë", "Zgjedhja e dymbëdhjetë",
        "Zgjedhja e trembëdhjetë", "Zgjedhja e katërmbëdhjetë", "Zgjedhja e pesëmbëdhjetë",
        "Zgjedhja e gjashtëmbëdhjetë", "Zgjedhja e shtatëmëdhjetë", "Zgjedhja e tetëmbëdhjetë",
        "Zgjedhja e nëntëmbëdhjetë", "Zgjedhja e njëzet", "Ekstra"
    ]

    doc = docx2python(input_docx)
    lines = [line.strip() for line in doc.text.split('\n') if line.strip()]

    survey = []
    choices = []
    skipped_other_questions = []
    settings = [{'style': 'theme-grid no-text-transform'}]

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
            "label": "GPS",
            "required": "true"
    })
           
    # Add Anketuesi_ja question
    survey.append({
       "type": "select_one anketuesit_list",
        "name": "Anketuesi_ja",
        "label": "Anketuesi/ja",
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

    while i < len(lines):
        line = lines[i]

        if line.lower().startswith("[note]"):
            label = line[6:].strip()
            survey.append({
                "type": "note",
                "name": f"note{note_index}",
                "label": label
            })
            note_index += 1
            i += 1
            continue

        # STEP 1: Extract all tags (like [single], [random], [hint:...])
        tags = extract_tags(line)

        # STEP 2: Parse those tags to understand type, hint, and randomization
        q_type, matrix_count, parameters, hint = parse_question_tags(tags)

        # STEP 3: Remove all tags from text so only question text remains
        full_line = strip_type(line)

        # STEP 4: Extract the question number and text
        qnum, label_text = extract_question_number_and_text(full_line)
       
            # Skip if q_type is "other"
        if q_type == "other":
            # Collect label for display
            full_line = strip_type(line)
            _, label_text = extract_question_number_and_text(full_line)
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

            label = f"{qnum}. {label_text}" if qnum else full_line

            if qnum:
                if qnum.upper().startswith("D"):
                    qname = qnum
                elif re.match(r'Q[\d\w\.]+', qnum, re.IGNORECASE):
                    qname = re.sub(r'^[Qq]', 'P', qnum)
                elif qnum.isdigit():
                    qname = f"P{q_index}"
                    q_index += 1
                else:
                    qname = f"P{q_index}"
                    q_index += 1
            else:
                qname = f"P{q_index}"
                q_index += 1

            qname = qname.rstrip('.')
            required = "yes"

            def add_common_question(fields):
                if parameters:
                    fields["parameters"] = parameters
                if hint:
                    fields["hint"] = hint
                survey.append(fields)

            def collect_options(start_index):
                opts = []
                while start_index < len(lines):
                    tags = extract_tags(lines[start_index])
                    next_type, _, _, _ = parse_question_tags(tags)

                    if next_type:
                        break
                    opts.append(lines[start_index])
                    start_index += 1
                return opts, start_index

            if q_type in ["single", "multiple"]:
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
                for idx, opt in enumerate(options, 1):
                    clean = clean_label_prefix(opt)
                    name_value = f"_{idx}" if q_type == "multiple" else str(idx)
                    choices.append({
                        "list_name": list_name,
                        "name": name_value,
                        "label": clean
                    })
                    if '_' in opt:
                        open_name = f"{qname}_{idx}"
                        relevant_expr = f"selected(${{{qname}}}, '{name_value}')" if q_type == "multiple" else f"${{{qname}}} = '{name_value}'"
                        survey.append({
                            "type": "text",
                            "name": open_name,
                            "label": f"{clean}",
                            "relevant": relevant_expr,
                            "required": "yes"
                        })

            elif q_type == "numeric":
                add_common_question({
                    "type": "integer",
                    "name": qname,
                    "label": label,
                    "required": required
                })
                i += 1

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
                min_label = matrix_count.get("min_label")
                max_label = matrix_count.get("max_label")

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
                        lbl = f"{j} - {min_label}" if j == start and min_label else \
                              f"{j} - {max_label}" if j == end and max_label else str(j)
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
                    tags = extract_tags(lines[i])
                    next_type, _, _, _ = parse_question_tags(tags)
                    if next_type:
                        break
                    rows.append(lines[i])
                    i += 1

                survey.append({"type": "begin_group", "name": f"{qname}_group", "appearance": "field-list", "required": "no"})
                survey.append({"type": f"{style} {list_name}", "name": f"{qname}_matrix_label", "label": label, "appearance": "label", "required": "no"})

                for idx, row in enumerate(rows, 1):
                    field = {
                        "type": f"{style} {list_name}",
                        "name": f"{qname}_{idx}",
                        "label": row,
                        "appearance": "list-nolabel",
                        "required": "yes"
                    }
                    if parameters:
                        field["parameters"] = parameters
                    survey.append(field)

                survey.append({"type": "end_group", "name": f"{qname}_group_end"})

                for j, col in enumerate(columns, 1):
                    choices.append({"list_name": list_name, "name": str(j), "label": col})

            elif q_type.startswith("ranking"):
                match = re.findall(r"\d+", q_type)
                if match:
                    rank_count = int(match[0])
                    list_name = qname + "_list"

                    survey.append({"type": "begin_group", "name": f"{qname}_group", "appearance": "field-list"})
                    survey.append({"type": "note", "name": f"{qname}_label", "label": label})

                    for idx in range(1, rank_count + 1):
                        rank_name = f"{qname}_{idx}"
                        constraint = " and ".join([f"${rank_name} != ${qname}_{j}" for j in range(1, idx)]) if idx > 1 else ""
                        constraint_msg = "Opsioni i njejtë nuk mund të zgjedhet më shumë se një herë"
                        survey.append({
                            "type": f"select_one {list_name}",
                            "name": rank_name,
                            "label": ranking_labels[idx - 1] if idx <= 20 else ranking_labels[-1],
                            "required": "yes",
                            "appearance": "minimal",
                            "choice_filter": " and ".join([f"not(selected(${{{qname}_{j}}}, name))" for j in range(1, idx)])
                        })

                    survey.append({"type": "end_group", "name": f"{qname}_group_end"})

                    i += 1
                    options = []
                    while i < len(lines):
                        tags = extract_tags(lines[i])
                        next_type, _, _, _ = parse_question_tags(tags)
                        if next_type:
                            break
                        options.append(lines[i])
                        i += 1

                    for idx, opt in enumerate(options, 1):
                        clean = clean_label_prefix(opt)
                        choices.append({"list_name": list_name, "name": str(idx), "label": clean})
                        
            elif q_type is None:
                raise ValueError(f"Formatimi i Word dokumentit nuk është valid në këtë linjë: '{line}'")
        else:
            i += 1

    survey.append({"type": "text", "name": "emri_mbiemri", "label": "Emri dhe mbiemri:", "required": "yes"})
    survey.append({"type": "text", "name": "numri_telefonit", "label": "Numri i telefonit:", "required": "yes"})

    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        pd.DataFrame(survey).to_excel(writer, sheet_name="survey", index=False)
        if choices:
            pd.DataFrame(choices).to_excel(writer, sheet_name="choices", index=False)
        pd.DataFrame(settings).to_excel(writer, sheet_name="settings", index=False)

    return skipped_other_questions


def process_uploaded_docx(uploaded_bytesio, filename, data_method, selected_questions):
    base_name = os.path.splitext(filename)[0]
    generated_name = f"{base_name}_gjeneruar.xlsx"
    temp_xlsx_path = os.path.join(tempfile.gettempdir(), generated_name)

    try:
        uploaded_bytesio.seek(0)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_bytesio.read())
            tmp.flush()
            skipped = generate_xlsform(tmp.name, temp_xlsx_path, data_method, selected_questions)
        return temp_xlsx_path, generated_name, None, skipped
    except Exception as e:
        return None, None, str(e), None

if uploaded_file:
    uploaded_content = uploaded_file.read()
    uploaded_bytesio = BytesIO(uploaded_content)
    uploaded_bytesio.seek(0)  # rifillon stream-in që të përdoret prapë

    data_collection_method = st.selectbox(
    "Metoda e mbledhjes së të dhënave:",
    ["Face to face", "Telefon/Online"] 
    )

    doc = docx2python(uploaded_bytesio).text
    lines = [line.strip() for line in doc.split('\n') if line.strip()]

    # Extract question numbers (e.g., 1, D1, 2a, Q1.2 etc.)
    question_options = []
    for line in lines:
        try:
            # Extract all tags from this line (e.g., [random][single][hint: ...])
            tags = extract_tags(line)
            q_type, _, _, _ = parse_question_tags(tags)

            # Only process lines that define a question type
            if q_type:
                _, label_text = extract_question_number_and_text(strip_type(line))
                if label_text:
                    question_options.append(label_text)

        except ValueError as e:
            st.error(f"Gabim në rreshtin: **{line}**\n\n{str(e)}")
            st.stop()


    st.session_state["question_lines"] = lines
    selected_questions = st.multiselect(
        "Zgjidh pyetjet që NUK dëshiron të kodosh:",
        options=question_options,
        default=None
    )
    st.session_state["selected_questions"] = selected_questions


    if data_collection_method:
        generate_button = st.button("Gjenero formularin XLS")
        if generate_button:
            with st.spinner("Po përpunon dokumentin..."):
                data_method = data_collection_method == "Face to face"
                uploaded_bytesio.seek(0)
                xlsx_path, generated_file_name, error, skipped = process_uploaded_docx(uploaded_bytesio, uploaded_file.name, data_method, st.session_state.get("selected_questions", None))
        
                if error:
                    st.error(f"Gabimi: {error}")
                else:
                    with open(xlsx_path, "rb") as f:
                        st.session_state["xlsx_data"] = f.read()
                        st.session_state["xlsx_name"] = generated_file_name
                        st.session_state["xlsx_ready"] = True
                        st.session_state["skipped_other_questions"] = skipped
        if st.session_state.get("xlsx_ready", False):
            st.success("Formulari XLS u gjenerua me sukses!")
            st.download_button(
                label="Shkarko formularin XLS",
                data=st.session_state["xlsx_data"],
                file_name=st.session_state["xlsx_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            if st.session_state.get("skipped_other_questions"):
                st.info("Pyetjet me tag-un [other] që u anashkaluan:")
                for q in st.session_state["skipped_other_questions"]:
                    st.markdown(f"- {q}")


