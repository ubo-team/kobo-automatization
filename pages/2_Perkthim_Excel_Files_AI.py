import pandas as pd
import requests
import streamlit as st
from io import BytesIO
import re
from PIL import Image
import os

import streamlit as st, os
st.write("Loaded secret keys:", list(st.secrets.keys()))
if "AZURE_TRANSLATOR_KEY" not in st.secrets:
    st.error("AZURE_TRANSLATOR_KEY is missing from st.secrets. Open Manage app → Settings → Secrets.")
    st.stop()


st.set_page_config(page_title="Perkthim Excel Files A", layout="centered")


logo_path = "logo.png"  

with st.sidebar:
    if os.path.exists(logo_path):
        st.image(Image.open(logo_path), width=150)

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
    
AZURE_TRANSLATOR_KEY = st.secrets["AZURE_TRANSLATOR_KEY"]
AZURE_TRANSLATOR_ENDPOINT = st.secrets["AZURE_TRANSLATOR_ENDPOINT"]
AZURE_TRANSLATOR_REGION = st.secrets["AZURE_TRANSLATOR_REGION"]

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

def translate_text(text, from_lang, to_lang):
    if pd.isna(text) or not str(text).strip():
        return text

    code, remaining_text = adjust_question_code(text, from_lang, to_lang)

    path = "/translate?api-version=3.0"
    params = f"&from={from_lang}&to={to_lang}"
    url = AZURE_TRANSLATOR_ENDPOINT + path + params

    headers = {
        'Ocp-Apim-Subscription-Key': AZURE_TRANSLATOR_KEY,
        'Ocp-Apim-Subscription-Region': AZURE_TRANSLATOR_REGION,
        'Content-type': 'application/json'
    }

    body = [{"text": str(remaining_text)}]
    response = requests.post(url, headers=headers, json=body)

    if response.status_code != 200:
        return text

    result = response.json()

    try:
        translated_text = result[0]["translations"][0]["text"]
        return code + translated_text
    except (KeyError, IndexError, TypeError):
        return text

def translate_dataframe(df, source_col, target_col, from_lang, to_lang):
    df[target_col] = df[source_col].apply(lambda x: translate_text(x, from_lang, to_lang))
    return df

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
            with st.spinner("Duke përkthyer... Ju lutemi prisni"):
                for target_col, to_lang in target_languages:
                    df = translate_dataframe(df, source_col, target_col, from_lang=from_lang, to_lang=to_lang)

            st.session_state.translated_sheets[selected_sheet] = df.copy()
            st.success(f"Përkthimi për {selected_sheet} u krye me sukses në Bllokun {block_id + 1}!")
            st.write(df.head())

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
