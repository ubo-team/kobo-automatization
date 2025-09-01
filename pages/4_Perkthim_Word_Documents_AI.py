import pandas as pd
import requests
import streamlit as st
from io import BytesIO
import re
from PIL import Image
import os
import docx
from docx import Document
from docx.shared import Pt

st.set_page_config(page_title="Përkthe Word Dokumente me AI", layout="centered")


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

def split_multiline_text(text):
    return [line.strip() for line in text.split("\n") if line.strip()]

def batch_translate_lines(lines, from_lang, to_lang):
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_TRANSLATOR_KEY,
        "Ocp-Apim-Subscription-Region": AZURE_REGION,
        "Content-type": "application/json"
    }
    params = {"api-version": "3.0", "from": from_lang, "to": [to_lang]}
    body = [{"text": line} for line in lines]

    try:
        response = requests.post(AZURE_TRANSLATOR_ENDPOINT + "/translate", params=params, headers=headers, json=body)
        response.raise_for_status()
        return [item["translations"][0]["text"] for item in response.json()]
    except Exception as e:
        print("Translation error:", e)
        return lines

def translate_docx_in_place(doc, from_lang, to_lang):
    for para in doc.paragraphs:
        for run in para.runs:
            original_text = run.text.strip()
            if original_text:
                translated_lines = batch_translate_lines(split_multiline_text(original_text), from_lang, to_lang)
                run.text = "\n".join(translated_lines)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        original_text = run.text.strip()
                        if original_text:
                            translated_lines = batch_translate_lines(split_multiline_text(original_text), from_lang, to_lang)
                            run.text = "\n".join(translated_lines)

    return doc


st.title("Fillo me Përkthimin e Pyetësorëve")

uploaded_file = st.file_uploader("Ngarko dokumentin (vetëm Word)", type=["docx"])

if uploaded_file:
    from_lang_label = st.selectbox("Gjuha Burimore", list(LANGUAGE_OPTIONS_UI.keys()), key="word_lang_from")
    to_lang_label = st.selectbox("Gjuha për Përkthim", list(LANGUAGE_OPTIONS_UI.keys()), key="word_lang_to")
    from_lang = LANGUAGE_OPTIONS_UI[from_lang_label]
    to_lang = LANGUAGE_OPTIONS_UI[to_lang_label]

    if st.button("Përkthe Word Dokumentin"):
        with st.spinner("Duke përkthyer dokumentin..."):
            doc = Document(uploaded_file)
            translated_doc = translate_docx_in_place(doc, from_lang, to_lang)

            output = BytesIO()
            translated_doc.save(output)
            output.seek(0)

            st.success("Përkthimi përfundoi me sukses!")
            st.download_button(
                label="Shkarko dokumentin e përkthyer (Word)",
                data=output,
                file_name=f"translated_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
