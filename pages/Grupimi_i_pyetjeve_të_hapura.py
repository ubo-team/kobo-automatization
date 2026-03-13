import os
import io
import re
import time
from datetime import datetime
from collections import Counter

import streamlit as st
import pandas as pd
import google.generativeai as genai

# -------------------------------
# Page Configuration
# -------------------------------
st.set_page_config(page_title="Survey Response Categorizer", layout="centered")

# -------------------------------
# Sidebar Logo
# -------------------------------
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

# ── Gemini API setup ─────────────────────────────────────────────────────────
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

# ── Pricing ──────────────────────────────────────────────────────────────────
GEMINI_PRICING = {
    "models/gemini-2.5-pro": {
        "inputPer1MTokens_low": 1.25,
        "outputPer1MTokens_low": 10.00,
        "inputPer1MTokens_high": 2.50,
        "outputPer1MTokens_high": 15.00,
        "tier_threshold": 200_000,
    },
    "models/gemini-2.5-flash": {
        "inputPer1MTokens": 0.30,
        "outputPer1MTokens": 2.50,
    },
}


def calculate_gemini_cost(prompt_tokens: int, completion_tokens: int, model: str) -> float:
    pricing = GEMINI_PRICING.get(model, GEMINI_PRICING["models/gemini-2.5-pro"])

    if "tier_threshold" in pricing:
        high = (prompt_tokens or 0) > pricing["tier_threshold"]
        input_rate = pricing["inputPer1MTokens_high"] if high else pricing["inputPer1MTokens_low"]
        output_rate = pricing["outputPer1MTokens_high"] if high else pricing["outputPer1MTokens_low"]
    else:
        input_rate = pricing["inputPer1MTokens"]
        output_rate = pricing["outputPer1MTokens"]

    input_cost = ((prompt_tokens or 0) / 1_000_000) * input_rate
    output_cost = ((completion_tokens or 0) / 1_000_000) * output_rate
    return round(input_cost + output_cost, 8)


# ── Page header ──────────────────────────────────────────────────────────────
st.title("Grupimi i pyetjeve të hapura")
st.markdown(
    "Ngarko një dokument Excel me përgjigje të hapura. "
    "Aplikacioni do t'i kategorizojë automatikisht duke përdorur Gemini API."
)

# ── Default prompt ───────────────────────────────────────────────────────────
DEFAULT_PROMPT = """You are a survey response categorizer. Your task is to assign ONE category to each survey response in a batch.

Question: {question_label}

Available categories:
{categories}

Rules:
1. Choose the single best-matching category from the list above for each response.
2. If the response clearly represents a NEW, distinct theme that appears frequently (not covered by existing categories), output: NEW: <short category name>
3. If the response is empty output: 999
4. If the response is irrelevant, or unclassifiable, output: Other
5. The output should be all in English, even if the answers are in other languages.
6. Output ONLY the category names — no explanation, no punctuation, no extra text.

Responses (one per line, numbered):
{responses}

Output one category per line in the same order (numbered to match), e.g.:
1. Category
2. Category
..."""

# ── Session state ────────────────────────────────────────────────────────────
if "question_categories" not in st.session_state:
    st.session_state.question_categories = {}

if "prompt_template" not in st.session_state or "{response}" in st.session_state.prompt_template:
    st.session_state.prompt_template = DEFAULT_PROMPT

# ── Progress helpers ─────────────────────────────────────────────────────────
PROGRESS_DIR = "progress_files"
os.makedirs(PROGRESS_DIR, exist_ok=True)


def get_progress_file_name(uploaded_file_name: str) -> str:
    base_name = os.path.splitext(uploaded_file_name)[0]
    safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", base_name)
    return os.path.join(PROGRESS_DIR, f"{safe_name}_categorized_progress.xlsx")


def save_progress(result_df: pd.DataFrame, progress_file: str):
    """
    Safe save using temp file then atomic replace.
    """
    temp_file = progress_file.replace(".xlsx", "_tmp.xlsx")
    result_df.to_excel(temp_file, index=False, engine="openpyxl")
    os.replace(temp_file, progress_file)


def is_missing_label(value) -> bool:
    """
    Treat empty, NaN, and Error as missing so failed rows retry later.
    """
    if pd.isna(value):
        return True
    value = str(value).strip()
    return value == "" or value.lower() in ["nan", "error"]


def build_question_checkpoint_name(progress_file: str, col: str) -> str:
    safe_col = re.sub(r"[^A-Za-z0-9_\-]", "_", str(col))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(progress_file))[0]
    return f"{base}_{safe_col}_{timestamp}.xlsx"


# ── Settings ─────────────────────────────────────────────────────────────────
with st.expander("Cilësimet", expanded=False):
    col_model, col_batch = st.columns(2)

    with col_model:
        model_name = st.selectbox(
            "Model",
            ["gemini-2.5-flash", "gemini-2.5-pro"],
            index=0
        )

    with col_batch:
        batch_size = st.number_input(
            "Batch size (rreshta per thirrje)",
            min_value=5,
            max_value=100,
            value=20,
            step=5
        )

    st.divider()
    st.subheader("Prompt Template")
    st.caption("Placeholders: `{question_label}`, `{categories}`, `{responses}`")

    st.session_state.prompt_template = st.text_area(
        "Edit prompt",
        value=st.session_state.prompt_template,
        height=320,
        label_visibility="collapsed",
    )

    if st.button("Rivendos prompt-in fillestar"):
        st.session_state.prompt_template = DEFAULT_PROMPT
        st.rerun()

st.markdown("---")

# ── Step 1: Upload file ──────────────────────────────────────────────────────
st.header("1. Ngarko dokumentin Excel")
uploaded_file = st.file_uploader(
    "Dokument Excel me Response ID + kolona me përgjigje të hapura",
    type=["xlsx", "xls"]
)

df = None
question_cols = []
progress_file = None

if uploaded_file:
    original_df = pd.read_excel(uploaded_file)
    progress_file = get_progress_file_name(uploaded_file.name)

    col_a, col_b = st.columns([3, 1])

    with col_a:
        use_saved_progress = False
        if os.path.exists(progress_file):
            use_saved_progress = st.checkbox(
                "Përdor progresin e ruajtur më parë",
                value=True,
                help="Nëse zgjidhet, aplikacioni do të vazhdojë nga versioni i ruajtur dhe nuk do të dërgojë sërish rreshtat e kategorizuar."
            )

    with col_b:
        if os.path.exists(progress_file):
            if st.button("Fshi progresin"):
                os.remove(progress_file)
                st.success("Progresi i ruajtur u fshi.")
                st.rerun()

    if use_saved_progress and os.path.exists(progress_file):
        try:
            df = pd.read_excel(progress_file)
            st.info("U ngarkua versioni i ruajtur më parë.")
        except Exception as e:
            st.warning(f"Nuk u lexua dot skedari i progresit. Po përdoret skedari origjinal. Gabimi: {e}")
            df = original_df.copy()
    else:
        df = original_df.copy()

    st.success(f"U ngarkuan **{len(df)} rreshta** dhe **{len(df.columns)} kolona**")
    st.dataframe(df.head(5), use_container_width=True)

    st.caption(f"Skedari i autosave: `{progress_file}`")

    if progress_file and os.path.exists(progress_file):
        with open(progress_file, "rb") as f:
            st.download_button(
                label="Shkarko progresin aktual",
                data=f.read(),
                file_name=os.path.basename(progress_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_current_progress_top"
            )

        st.info(
            "Nëse aplikacioni ndalet, ngarko këtë file të ruajtur për të vazhduar pa ri-kategorizuar rreshtat e përfunduar."
        )

    id_col = st.selectbox("Zgjidh kolonën e Response ID", df.columns.tolist(), index=0)
    question_cols = st.multiselect(
        "Zgjidh kolonat me përgjigje të hapura për kategorizim",
        [c for c in df.columns if c != id_col],
    )

# ── Step 2: Define categories per question ──────────────────────────────────
if df is not None and question_cols:
    st.header("2. Përcakto kategoritë për çdo pyetje")
    st.caption("Shkruaj një kategori për rresht. Modeli gjithashtu do të detektojë tema të reja automatikisht.")

    for col in question_cols:
        if col not in st.session_state.question_categories:
            st.session_state.question_categories[col] = "Positive\nNegative\nNeutral\nOther"

        with st.expander(f"Kategoritë për **{col}**", expanded=True):
            st.session_state.question_categories[col] = st.text_area(
                f"Categories for {col}",
                value=st.session_state.question_categories[col],
                height=140,
                key=f"cats_{col}",
                label_visibility="collapsed",
            )

# ── Step 3: Run categorization ───────────────────────────────────────────────
if df is not None and question_cols:
    st.header("3. Ekzekuto kategorizimin")

    new_cat_threshold = st.slider(
        "Frekuenca minimale për të promovuar një kategori të re",
        min_value=2,
        max_value=20,
        value=3,
        help="Nëse një etiketë 'NEW: X' shfaqet kaq herë, X shtohet si kategori zyrtare dhe përgjigjet ri-vlerësohen.",
    )

    run_btn = st.button("Kategorizo përgjigjet", type="primary")

    if run_btn:
        model_id = f"models/{model_name}"
        gemini_model = genai.GenerativeModel(model_name)
        result_df = df.copy()
        token_counts = {"input": 0, "output": 0}
        MAX_RETRIES = 3

        save_progress(result_df, progress_file)

        def call_gemini_batch(prompt_text: str) -> tuple[str, int, int]:
            for attempt in range(MAX_RETRIES):
                try:
                    response = gemini_model.generate_content(
                        prompt_text,
                        generation_config=genai.types.GenerationConfig(
                            max_output_tokens=4096,
                            temperature=0,
                        ),
                        request_options={"timeout": 120},
                    )

                    usage = getattr(response, "usage_metadata", None)
                    in_tok = getattr(usage, "prompt_token_count", 0) if usage else 0
                    out_tok = getattr(usage, "candidates_token_count", 0) if usage else 0

                    text = getattr(response, "text", "")
                    if text is None:
                        text = ""

                    return text.strip(), in_tok, out_tok

                except Exception as e:
                    if attempt < MAX_RETRIES - 1:
                        wait = 2 ** attempt
                        st.toast(f"Retry {attempt + 1}/{MAX_RETRIES} pas {wait}s: {e}")
                        time.sleep(wait)
                    else:
                        raise e

        def parse_batch_response(text: str, expected_count: int) -> list[str]:
            lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
            results = []

            for line in lines:
                m = re.match(r"^\d+[\.\)\-:]\s*(.+)$", line)
                if m:
                    results.append(m.group(1).strip())
                elif not re.match(r"^\d+$", line):
                    results.append(line.strip())

            while len(results) < expected_count:
                results.append("Error")

            return results[:expected_count]

        def categorize_column(
            col: str,
            categories: list[str],
            responses: pd.Series,
            result_df: pd.DataFrame,
            grouped_col_name: str,
        ) -> list[str]:
            cats_str = "\n".join(f"- {c}" for c in categories)
            results = [""] * len(responses)
            non_empty_indices = []

            if grouped_col_name not in result_df.columns:
                result_df[grouped_col_name] = ""

            for i, resp in enumerate(responses):
                existing_val = result_df.at[i, grouped_col_name] if grouped_col_name in result_df.columns else ""

                if not is_missing_label(existing_val):
                    results[i] = str(existing_val).strip()
                    continue

                if pd.isna(resp) or str(resp).strip() == "":
                    results[i] = "999"
                    result_df.at[i, grouped_col_name] = "999"
                else:
                    non_empty_indices.append(i)

            save_progress(result_df, progress_file)

            if not non_empty_indices:
                return results

            total_to_process = len(non_empty_indices)
            num_batches = (total_to_process + batch_size - 1) // batch_size
            skipped = len(responses) - total_to_process

            prog = st.progress(
                0,
                text=f"Duke kategorizuar **{col}** ({total_to_process} për t'u procesuar, {skipped} të kapërcyera)...",
            )

            for batch_idx in range(num_batches):
                start = batch_idx * batch_size
                end = min(start + batch_size, total_to_process)
                batch_indices = non_empty_indices[start:end]

                numbered_responses = [
                    f"{j + 1}. {str(responses.iloc[idx])}"
                    for j, idx in enumerate(batch_indices)
                ]

                prompt = st.session_state.prompt_template.format(
                    question_label=col,
                    categories=cats_str,
                    responses="\n".join(numbered_responses),
                )

                try:
                    text, in_tok, out_tok = call_gemini_batch(prompt)
                    token_counts["input"] += in_tok
                    token_counts["output"] += out_tok
                    batch_labels = parse_batch_response(text, len(batch_indices))
                except Exception as e:
                    st.warning(f"Gabim API në batch {batch_idx + 1}: {e}")
                    batch_labels = ["Error"] * len(batch_indices)

                for j, idx in enumerate(batch_indices):
                    label = batch_labels[j]
                    results[idx] = label
                    result_df.at[idx, grouped_col_name] = label

                save_progress(result_df, progress_file)

                prog.progress(
                    end / total_to_process,
                    text=f"Duke kategorizuar **{col}** ({end}/{total_to_process})"
                )

            prog.empty()
            return results

        for col in question_cols:
            base_cats = [
                c.strip()
                for c in st.session_state.question_categories[col].splitlines()
                if c.strip()
            ]
            grouped_col = f"{col}_grouped"

            with st.spinner(f"Duke procesuar **{col}**..."):
                labels = categorize_column(
                    col=col,
                    categories=base_cats,
                    responses=result_df[col],
                    result_df=result_df,
                    grouped_col_name=grouped_col,
                )

            new_labels = [
                l for l in labels
                if isinstance(l, str) and l.lower().startswith("new:")
            ]
            new_counts = Counter(
                re.sub(r"(?i)^new:\s*", "", l).strip()
                for l in new_labels
            )
            promoted = [
                cat for cat, cnt in new_counts.items()
                if cnt >= new_cat_threshold
            ]

            if promoted:
                st.info(
                    f"Kategori të reja të detektuara për **{col}**: "
                    f"{', '.join(promoted)} — duke ri-ekzekutuar vetëm rreshtat NEW."
                )

                updated_cats = base_cats + promoted
                new_indices = [
                    i for i, l in enumerate(labels)
                    if isinstance(l, str) and l.lower().startswith("new:")
                ]

                if new_indices:
                    partial_series = pd.Series([None] * len(result_df[col]), dtype=object)

                    for i in new_indices:
                        partial_series.iloc[i] = result_df[col].iloc[i]
                        result_df.at[i, grouped_col] = ""

                    partial_labels = categorize_column(
                        col=col,
                        categories=updated_cats,
                        responses=partial_series,
                        result_df=result_df,
                        grouped_col_name=grouped_col,
                    )

                    for i in new_indices:
                        labels[i] = partial_labels[i]

            def clean_label(label):
                if not isinstance(label, str):
                    return label
                m = re.match(r"(?i)^new:\s*(.+)$", label)
                return m.group(1).strip() if m else label

            cleaned_labels = [clean_label(l) for l in labels]
            result_df[grouped_col] = cleaned_labels

            save_progress(result_df, progress_file)

            st.success(f"Përfundoi: **{col}** → **{grouped_col}**")

            if os.path.exists(progress_file):
                with open(progress_file, "rb") as f:
                    st.download_button(
                        label=f"Shkarko checkpoint pas pyetjes: {col}",
                        data=f.read(),
                        file_name=build_question_checkpoint_name(progress_file, col),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_progress_{col}"
                    )

            st.subheader(f"Shpërndarja e kategorive — {col}")
            dist = result_df[grouped_col].value_counts(dropna=False).reset_index()
            dist.columns = ["Kategoria", "Numri"]
            dist["Përqindja"] = (
                (dist["Numri"] / dist["Numri"].sum() * 100).round(1).astype(str) + "%"
            )
            st.dataframe(dist, use_container_width=True, hide_index=True)

            st.dataframe(
                result_df[[id_col, col, grouped_col]].head(20),
                use_container_width=True,
            )

        total_cost = calculate_gemini_cost(
            token_counts["input"],
            token_counts["output"],
            model_id
        )

        st.markdown("---")
        st.header("Përmbledhje")

        cost_col1, cost_col2, cost_col3 = st.columns(3)
        cost_col1.metric("Input tokens", f"{token_counts['input']:,}")
        cost_col2.metric("Output tokens", f"{token_counts['output']:,}")
        cost_col3.metric("Kostoja totale", f"${total_cost:.6f}")

        st.success(f"Progresi u ruajt te: {progress_file}")

        if os.path.exists(progress_file):
            with open(progress_file, "rb") as f:
                st.download_button(
                    label="Shkarko versionin me autosave",
                    data=f.read(),
                    file_name=os.path.basename(progress_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_autosave_final"
                )

        output = io.BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="Shkarko Excel-in e kategorizuar",
            data=output,
            file_name="categorized_responses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
