import pandas as pd
import streamlit as st
import ubo_ui
from io import BytesIO
import re
import os
import time
from translation_fallback import ClaudeFallback



st.set_page_config(layout="centered")


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
    "sr": "Serbian (Latin script)",
    "mk": "Macedonian (Latin script)",
    "bs": "Bosnian"
}

claude = ClaudeFallback(st.secrets)   # the translator (Claude)




def translate_dataframe(df, source_col, target_col, from_lang, to_lang):
    tokens_in, tokens_out = claude.tokens_in, claude.tokens_out
    errors = []
    results = list(df[source_col].values)

    # Collect texts that need translation
    to_translate = []
    for i, val in enumerate(results):
        if pd.isna(val) or not str(val).strip() or str(val).strip().lower() == "none":
            continue
        code, remaining = adjust_question_code(str(val), from_lang, to_lang)
        if not remaining.strip():
            results[i] = code + remaining
            continue
        to_translate.append((i, code, remaining.strip()))

    if not to_translate:
        df[target_col] = results
        return df, 0, 0, []

    from_name = LANG_NAMES.get(from_lang, from_lang)
    to_name = LANG_NAMES.get(to_lang, to_lang)

    progress = st.progress(0, text="Duke përkthyer... 0%")
    translated, errors = claude.translate_all(
        [text for _, _, text in to_translate], from_name, to_name,
        progress=lambda done, total: progress.progress(done / total, text=f"Duke përkthyer... {done}/{total} grupe"))
    for k, (idx, code, _) in enumerate(to_translate):
        if k in translated:
            results[idx] = code + translated[k]

    progress.empty()
    df[target_col] = results
    return df, claude.tokens_in - tokens_in, claude.tokens_out - tokens_out, errors

with ubo_ui.card("Ngarkoni dokumentin Excel", step=1):
    uploaded_file = st.file_uploader("Ngarko Excel-in", type=["xlsx"], label_visibility="collapsed")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    all_sheets = {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in sheet_names}

    if "translated_sheets" not in st.session_state:
        st.session_state.translated_sheets = {sheet: df.copy() for sheet, df in all_sheets.items()}

    if "translation_blocks" not in st.session_state:
        st.session_state.translation_blocks = [0]

    for block_id in st.session_state.translation_blocks:
        with ubo_ui.card(f"Blloku {block_id + 1} i përkthimit", step=block_id + 2, key=f"blloku_{block_id}"):

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
            multiple_targets = st.multiselect(f"Kolonat ku dëshiron të përkthehet (Blloku {block_id + 1})", columns, key=f"multi_target_{block_id}", placeholder="Zgjidhni kolonat")

            target_languages = []
            for target_col in multiple_targets:
                lang_label = st.selectbox(f"Gjuha për kolonën: {target_col} (Blloku {block_id + 1})", list(LANGUAGE_OPTIONS_UI.keys()), key=f"{target_col}_lang_{block_id}")
                target_languages.append((target_col, LANGUAGE_OPTIONS_UI[lang_label]))

            if st.button(f"Fillo Përkthimin për {selected_sheet} (Blloku {block_id + 1})", key=f"translate_btn_{block_id}", type="primary"):
                block_in_tokens, block_out_tokens = 0, 0
                cost_before = claude.cost
                all_errors = []
                for target_col, to_lang in target_languages:
                    df, in_tok, out_tok, errors = translate_dataframe(df, source_col, target_col, from_lang=from_lang, to_lang=to_lang)
                    block_in_tokens += in_tok
                    block_out_tokens += out_tok
                    all_errors.extend(errors)

                if all_errors:
                    st.error(f"Ka pasur {len(all_errors)} gabime. Gabimi i parë: {all_errors[0]}")
                else:
                    st.session_state.translated_sheets[selected_sheet] = df.copy()
                    st.success(f"Përkthimi për {selected_sheet} u krye me sukses në Bllokun {block_id + 1}!")

                st.write(df.head())

                block_cost = claude.cost - cost_before
                st.info(
                    f"**Kostoja e Bllokut {block_id + 1}:**  \n"
                    f"Tokenë hyrës: **{block_in_tokens:,}** | Tokenë dalës: **{block_out_tokens:,}**  \n"
                    f"Kostoja: **${block_cost:.4f}**"
                )

            if block_id == len(st.session_state.translation_blocks) - 1:
                add_block = st.button("Shto bllok përkthimi të ri", key=f"add_block_{block_id}")
                if add_block:
                    st.session_state.translation_blocks.append(len(st.session_state.translation_blocks))

    with ubo_ui.card("Shkarkoni rezultatin", step=len(st.session_state.translation_blocks) + 2):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet in sheet_names:
                st.session_state.translated_sheets.get(sheet, all_sheets[sheet]).to_excel(writer, sheet_name=sheet, index=False)

        st.download_button(
            label="Shkarko Excel-in me të gjitha përkthimet",
            data=output.getvalue(),
            file_name=uploaded_file.name
        )
