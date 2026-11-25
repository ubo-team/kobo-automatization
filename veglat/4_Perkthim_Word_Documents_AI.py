import streamlit as st
import ubo_ui
from io import BytesIO
import os
import re
import time
from docx import Document
from translation_fallback import ClaudeFallback

st.set_page_config(layout="centered")


LANGUAGE_OPTIONS_UI = {
    "Gjuha Shqipe": "sq",
    "Gjuha Angleze": "en",
    "Gjuha Serbe": "sr",
    "Gjuha Maqedonase": "mk",
    "Gjuha Boshnjake": "bs"
}

LANG_NAMES = {
    "sq": "Albanian",
    "en": "English",
    "sr": "Serbian (Latin script)",
    "mk": "Macedonian (Latin script)",
    "bs": "Bosnian"
}

claude = ClaudeFallback(st.secrets)   # the translator (Claude)


def translate_docx_in_place(doc, from_lang, to_lang):
    tokens_in, tokens_out = claude.tokens_in, claude.tokens_out

    # Collect full paragraph text (not individual runs) for better translation
    # Each entry: (paragraph, full_text)
    para_entries = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            para_entries.append(("para", para, text))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if text:
                        para_entries.append(("cell", para, text))

    if not para_entries:
        return doc, 0, 0, []

    progress = st.progress(0, text="Duke përkthyer... 0%")
    translated, errors = claude.translate_all(
        [text for _, _, text in para_entries], LANG_NAMES.get(from_lang, from_lang), LANG_NAMES.get(to_lang, to_lang),
        progress=lambda done, total: progress.progress(done / total, text=f"Duke përkthyer... {done}/{total} grupe"))
    for k, (_, para, _) in enumerate(para_entries):
        if k in translated:
            # all the translated text in the first run, the other runs emptied
            if para.runs:
                para.runs[0].text = translated[k]
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.text = translated[k]

    progress.empty()
    return doc, claude.tokens_in - tokens_in, claude.tokens_out - tokens_out, errors


with ubo_ui.card("Ngarkoni dokumentin Word", step=1):
    uploaded_file = st.file_uploader("Ngarko dokumentin (vetëm Word)", type=["docx"], label_visibility="collapsed")

if uploaded_file:
    with ubo_ui.card("Gjuhët dhe përkthimi", step=2):
        from_lang_label = st.selectbox("Gjuha Burimore", list(LANGUAGE_OPTIONS_UI.keys()), key="word_lang_from")
        to_lang_label = st.selectbox("Gjuha për Përkthim", list(LANGUAGE_OPTIONS_UI.keys()), key="word_lang_to")
        from_lang = LANGUAGE_OPTIONS_UI[from_lang_label]
        to_lang = LANGUAGE_OPTIONS_UI[to_lang_label]

        if st.button("Përkthe dokumentin Word", type="primary"):
            st.session_state["word_cost_before"] = claude.cost
            doc = Document(uploaded_file)
            translated_doc, total_in, total_out, errors = translate_docx_in_place(doc, from_lang, to_lang)

            output = BytesIO()
            translated_doc.save(output)
            output.seek(0)

            if errors:
                st.error(f"Ka pasur {len(errors)} gabime. Gabimi i parë: {errors[0]}")
            else:
                st.success("Përkthimi përfundoi me sukses!")

            cost_before = st.session_state.get("word_cost_before", 0.0)
            st.info(
                f"**Kostoja:**  \n"
                f"Tokenë hyrës: **{total_in:,}** | Tokenë dalës: **{total_out:,}**  \n"
                f"Kostoja: **${claude.cost - cost_before:.4f}**"
            )

            st.download_button(
                label="Shkarko dokumentin e përkthyer (Word)",
                data=output,
                file_name=f"translated_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
