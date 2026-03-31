import pandas as pd
import streamlit as st
from io import BytesIO
import re
import os
import google.generativeai as genai



st.set_page_config(page_title="Perkthim Excel Files A", layout="centered")


logo_svg_path = "UBO-Logo.svg"

with st.sidebar:
    if os.path.exists(logo_svg_path):
        with open(logo_svg_path, "r", encoding="utf-8") as f:
            svg_logo = f.read()
        st.markdown(
            f'<div style="display:flex;justify-content:center;margin:15px 0;"><div style="width:150px;">{svg_logo}</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("""
        <style>
        [data-testid="stSidebar"] img {
            display: block;
            margin-left: auto;
            margin-right: auto;
            margin-top: 15px;
            margin-bottom: 15px;
        }
        </style>
    """, unsafe_allow_html=True)
    
GEMINI_TRANSLATION_API_KEY = st.secrets["GEMINI_TRANSLATION_API_KEY"]
genai.configure(api_key=GEMINI_TRANSLATION_API_KEY)

GEMINI_PRICING = {
    "models/gemini-2.5-flash": {
        "inputPer1MTokens":  0.30,
        "outputPer1MTokens": 2.50,
    },
    "models/gemini-3.1-flash-lite-preview": {
        "inputPer1MTokens":  0.25,
        "outputPer1MTokens": 1.50,
    },
}


def calculate_gemini_cost(prompt_tokens: int, completion_tokens: int, model: str) -> float:
    pricing = GEMINI_PRICING.get(model, GEMINI_PRICING["models/gemini-2.5-flash"])
    input_rate  = pricing["inputPer1MTokens"]
    output_rate = pricing["outputPer1MTokens"]
    input_cost  = ((prompt_tokens     or 0) / 1_000_000) * input_rate
    output_cost = ((completion_tokens or 0) / 1_000_000) * output_rate
    return input_cost + output_cost

LANGUAGE_OPTIONS_UI = {
    "Gjuha Shqipe": "sq",
    "Gjuha Angleze": "en",
    "Gjuha Serbe": "sr",
    "Gjuha Maqedonase": "mk",
    "Gjuha Boshnjake": "bs"

}

def adjust_question_code(text, from_lang, to_lang):
    match = re.match(r'^(Q\d+[a-zA-Z]?|P\d+[a-zA-Z]?)(.*)', str(text))
    if match:
        code = match.group(1)
        rest = match.group(2)
        if from_lang == "en" and to_lang in ["sq", "sr", "mk"]:
            code = code.replace("Q", "P")
        elif from_lang == "sq" and to_lang == "en":
            code = code.replace("P", "Q")
        elif from_lang == "en" and to_lang == "sr":
            code = code.replace("Q", "P")
        return code, rest
    else:
        return '', text

LANG_NAMES = {
    "sq": "Albanian",
    "en": "English",
    "sr": "Serbian",
    "mk": "Macedonian",
    "bs": "Bosnian"
}

MODEL_NAME = "gemini-2.5-flash"
gemini_model = genai.GenerativeModel(MODEL_NAME)


def translate_text(text, from_lang, to_lang):
    """Returns (translated_text, input_tokens, output_tokens)."""
    if pd.isna(text) or not str(text).strip():
        return text, 0, 0

    code, remaining_text = adjust_question_code(text, from_lang, to_lang)

    if not remaining_text.strip():
        return code + remaining_text, 0, 0

    from_name = LANG_NAMES.get(from_lang, from_lang)
    to_name = LANG_NAMES.get(to_lang, to_lang)

    prompt = (
        f"Translate the following text from {from_name} to {to_name}. "
        f"Return ONLY the translated text, nothing else.\n\n"
        f"{remaining_text}"
    )

    try:
        response = gemini_model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(temperature=0.1, max_output_tokens=1024),
        )
        in_tok = getattr(response.usage_metadata, "prompt_token_count", 0) or 0
        out_tok = getattr(response.usage_metadata, "candidates_token_count", 0) or 0
        translated_text = response.text.strip()
        return code + translated_text, in_tok, out_tok
    except Exception:
        return text, 0, 0


def translate_dataframe(df, source_col, target_col, from_lang, to_lang):
    total_in, total_out = 0, 0
    results = []
    for val in df[source_col]:
        translated, in_tok, out_tok = translate_text(val, from_lang, to_lang)
        results.append(translated)
        total_in += in_tok
        total_out += out_tok
    df[target_col] = results
    return df, total_in, total_out

st.title("Fillo me Përkthimin e Pyetësorëve")

uploaded_file = st.file_uploader("Ngarko Excel-in", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    all_sheets = {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in sheet_names}

    if "translated_sheets" not in st.session_state:
        st.session_state.translated_sheets = {sheet: df.copy() for sheet, df in all_sheets.items()}

    if "translation_blocks" not in st.session_state:
        st.session_state.translation_blocks = [0]

    for block_id in st.session_state.translation_blocks:
        st.markdown(f"---\n### Blloku {block_id + 1}")

        selected_sheet = st.selectbox(
            f"Zgjidh një faqe për përkthim (Blloku {block_id + 1})",
            sheet_names,
            key=f"sheet_select_{block_id}"
        )
        df = st.session_state.translated_sheets[selected_sheet]
        st.write(f"Pamje paraprake për {selected_sheet} (Blloku {block_id + 1}):", df.head())

        columns = df.columns.tolist()
        source_col = st.selectbox(f"Kolona burimore (Blloku {block_id + 1})", columns, key=f"source_col_{block_id}")
        from_lang_label = st.selectbox(f"Gjuha burimore (Blloku {block_id + 1})", list(LANGUAGE_OPTIONS_UI.keys()), key=f"from_lang_{block_id}")
        from_lang = LANGUAGE_OPTIONS_UI[from_lang_label]
        multiple_targets = st.multiselect(f"Kolonat ku dëshiron të përkthehet (Blloku {block_id + 1})", columns, key=f"multi_target_{block_id}")

        target_languages = []
        for target_col in multiple_targets:
            lang_label = st.selectbox(f"Gjuha për kolonën: {target_col} (Blloku {block_id + 1})", list(LANGUAGE_OPTIONS_UI.keys()), key=f"{target_col}_lang_{block_id}")
            target_languages.append((target_col, LANGUAGE_OPTIONS_UI[lang_label]))

        if st.button(f"Fillo Përkthimin për {selected_sheet} (Blloku {block_id + 1})", key=f"translate_btn_{block_id}"):
            block_in_tokens, block_out_tokens = 0, 0
            with st.spinner("Duke përkthyer... Ju lutemi prisni"):
                for target_col, to_lang in target_languages:
                    df, in_tok, out_tok = translate_dataframe(df, source_col, target_col, from_lang=from_lang, to_lang=to_lang)
                    block_in_tokens += in_tok
                    block_out_tokens += out_tok

            st.session_state.translated_sheets[selected_sheet] = df.copy()
            st.success(f"Përkthimi për {selected_sheet} u krye me sukses në Bllokun {block_id + 1}!")
            st.write(df.head())

            model_id = f"models/{MODEL_NAME}"
            block_cost = calculate_gemini_cost(block_in_tokens, block_out_tokens, model_id)
            st.info(
                f"**Kostoja e Bllokut {block_id + 1}:**  \n"
                f"Input tokens: **{block_in_tokens:,}** | Output tokens: **{block_out_tokens:,}**  \n"
                f"Kostoja: **${block_cost:.4f}**"
            )

        if block_id == len(st.session_state.translation_blocks) - 1:
            add_block = st.button("Shto bllok përkthimi të ri", key=f"add_block_{block_id}")
            if add_block:
                st.session_state.translation_blocks.append(len(st.session_state.translation_blocks))

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet in sheet_names:
            st.session_state.translated_sheets.get(sheet, all_sheets[sheet]).to_excel(writer, sheet_name=sheet, index=False)

    st.download_button(
        label="Shkarko Excel-in me të gjitha përkthimet",
        data=output.getvalue(),
        file_name=uploaded_file.name
    )
