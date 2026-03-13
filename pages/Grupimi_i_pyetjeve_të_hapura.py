import os
import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import re
from collections import Counter

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
        "inputPer1MTokens_low":   1.25,
        "outputPer1MTokens_low":  10.00,
        "inputPer1MTokens_high":  2.50,
        "outputPer1MTokens_high": 15.00,
        "tier_threshold": 200_000,
    },
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
    pricing = GEMINI_PRICING.get(model, GEMINI_PRICING["models/gemini-2.5-pro"])
    if "tier_threshold" in pricing:
        high = (prompt_tokens or 0) > pricing["tier_threshold"]
        input_rate  = pricing["inputPer1MTokens_high"]  if high else pricing["inputPer1MTokens_low"]
        output_rate = pricing["outputPer1MTokens_high"] if high else pricing["outputPer1MTokens_low"]
    else:
        input_rate  = pricing["inputPer1MTokens"]
        output_rate = pricing["outputPer1MTokens"]
    input_cost  = ((prompt_tokens     or 0) / 1_000_000) * input_rate
    output_cost = ((completion_tokens or 0) / 1_000_000) * output_rate
    return round(input_cost + output_cost, 8)


# ── Page header ──────────────────────────────────────────────────────────────
st.title("Grupimi i pyetjeve të hapura")
st.markdown("Ngarko një dokument Excel me përgjigje të hapura. Aplikacioni do t'i kategorizojë automatikisht duke përdorur Gemini API.")

# ── Default prompt ────────────────────────────────────────────────────────────
DEFAULT_PROMPT = """You are a survey response categorizer. Your task is to assign ONE category to each survey response in a batch.

Question: {question_label}

Available categories:
{categories}

Rules:
1. You MUST choose from the categories listed above. Do NOT invent new category names or rephrase existing ones.
2. If a response does not clearly fit any category, assign it to "Other".
3. If the response is empty, output: 999
4. ONLY use "NEW: <short category name>" if the response represents a genuinely distinct theme that NONE of the existing categories can cover. Be very conservative — prefer "Other" over creating new categories.
5. The output should be all in English, even if the answers are in other languages.
6. Output ONLY the category names — no explanation, no punctuation, no extra text.
7. Use the EXACT category names as listed above (case-sensitive).

Responses (one per line, numbered):
{responses}

Output one category per line in the same order (numbered to match), e.g.:
1. Category
2. Category
..."""

# ── Session state ─────────────────────────────────────────────────────────────
if "question_categories" not in st.session_state:
    st.session_state.question_categories = {}
if "question_labels" not in st.session_state:
    st.session_state.question_labels = {}
if "question_followup" not in st.session_state:
    st.session_state.question_followup = {}
if "prompt_template" not in st.session_state or "{response}" in st.session_state.prompt_template:
    st.session_state.prompt_template = DEFAULT_PROMPT
if "results" not in st.session_state:
    st.session_state.results = None

# ── Settings ─────────────────────────────────────────────────────────────────
with st.expander("Cilësimet", expanded=False):
    col_model, col_batch = st.columns(2)
    with col_model:
        model_name = st.selectbox("Model", ["gemini-3.1-flash-lite-preview", "gemini-2.5-flash", "gemini-2.5-pro"], index=0)
    with col_batch:
        batch_size = st.number_input("Batch size (rreshta per thirrje)", min_value=5, max_value=100, value=20, step=5)

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

# ── Step 1: Upload file ───────────────────────────────────────────────────────
st.header("1. Ngarko dokumentin Excel")
uploaded_file = st.file_uploader("Dokument Excel me Response ID + kolona me përgjigje të hapura", type=["xlsx", "xls"])

df = None
question_cols = []

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"U ngarkuan **{len(df)} rreshta** dhe **{len(df.columns)} kolona**")
    st.dataframe(df.head(5), use_container_width=True)

    id_col = st.selectbox("Zgjidh kolonën e Response ID", df.columns.tolist(), index=0)
    question_cols = st.multiselect(
        "Zgjidh kolonat me përgjigje të hapura për kategorizim",
        [c for c in df.columns if c != id_col],
    )

# ── Step 2: Define categories per question ───────────────────────────────────
if df is not None and question_cols:
    st.header("2. Përcakto kategoritë për çdo pyetje")
    st.caption("Shkruaj një kategori për rresht. Modeli gjithashtu do të detektojë tema të reja automatikisht.")

    for col in question_cols:
        if col not in st.session_state.question_categories:
            st.session_state.question_categories[col] = "Positive\nNegative\nNeutral\nOther"
        if col not in st.session_state.question_labels:
            st.session_state.question_labels[col] = col

        with st.expander(f"Kategoritë për **{col}**", expanded=True):
            st.session_state.question_labels[col] = st.text_input(
                "Etiketa e pyetjes (konteksti për modelin)",
                value=st.session_state.question_labels[col],
                key=f"label_{col}",
                help="Shkruaj pyetjen e plotë që u është bërë të anketuarve, p.sh. 'Çfarë mendoni për shërbimin tonë?'",
            )

            # Follow-up question toggle
            other_cols = [c for c in df.columns if c != col and c != id_col]
            is_followup = st.checkbox(
                "Kjo pyetje është vazhdim (follow-up) i një pyetjeje tjetër",
                key=f"followup_check_{col}",
                value=col in st.session_state.question_followup,
            )
            if is_followup and other_cols:
                default_idx = 0
                if col in st.session_state.question_followup:
                    prev = st.session_state.question_followup[col]["column"]
                    if prev in other_cols:
                        default_idx = other_cols.index(prev)
                parent_col = st.selectbox(
                    "Zgjidh kolonën e pyetjes paraprake",
                    other_cols,
                    index=default_idx,
                    key=f"followup_col_{col}",
                )
                parent_label = st.text_input(
                    "Etiketa e pyetjes paraprake",
                    value=st.session_state.question_followup.get(col, {}).get("label", parent_col),
                    key=f"followup_label_{col}",
                    help="P.sh. 'Which is the most important organization providing safety environment for everyone in Kosovo?'",
                )
                st.session_state.question_followup[col] = {
                    "column": parent_col,
                    "label": parent_label,
                }
            elif col in st.session_state.question_followup:
                del st.session_state.question_followup[col]

            st.session_state.question_categories[col] = st.text_area(
                f"Categories for {col}",
                value=st.session_state.question_categories[col],
                height=140,
                key=f"cats_{col}",
                label_visibility="collapsed",
            )

# ── Step 3: Run categorization ────────────────────────────────────────────────
if df is not None and question_cols:
    st.header("3. Ekzekuto kategorizimin")

    col_thresh, col_maxcat = st.columns(2)
    with col_thresh:
        new_cat_threshold = st.slider(
            "Frekuenca minimale për kategori të re",
            min_value=2, max_value=20, value=3,
            help="Nëse një etiketë 'NEW: X' shfaqet kaq herë, X shtohet si kategori zyrtare dhe përgjigjet ri-vlerësohen.",
        )
    with col_maxcat:
        max_categories = st.number_input(
            "Numri maksimal i kategorive",
            min_value=5, max_value=50, value=20, step=1,
            help="Kategoritë me frekuencë të ulët do të bashkohen në 'Other' për të mbajtur numrin brenda kufirit.",
        )

    run_btn = st.button("Kategorizo përgjigjet", type="primary")

    if run_btn:
        model_id = f"models/{model_name}"
        gemini_model = genai.GenerativeModel(model_name)
        result_df = df.copy()
        token_counts = {"input": 0, "output": 0}

        import time

        MAX_RETRIES = 3

        def call_gemini_batch(prompt_text: str) -> tuple[str, int, int]:
            """Returns (text, input_tokens, output_tokens) with retry."""
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
                    in_tok = response.usage_metadata.prompt_token_count
                    out_tok = response.usage_metadata.candidates_token_count
                    return response.text.strip(), in_tok, out_tok
                except Exception as e:
                    if attempt < MAX_RETRIES - 1:
                        wait = 2 ** attempt
                        st.toast(f"Retry {attempt+1}/{MAX_RETRIES} pas {wait}s: {e}")
                        time.sleep(wait)
                    else:
                        raise e

        def parse_batch_response(text: str, expected_count: int) -> list[str]:
            """Parse numbered lines from model output. Handles multi-word categories."""
            lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
            results = []
            for line in lines:
                # Match lines starting with a number (e.g. "1. Category name here")
                m = re.match(r"^\d+[\.\)\-:]\s*(.+)$", line)
                if m:
                    results.append(m.group(1).strip())
                elif not re.match(r"^\d+$", line):
                    # Non-numbered, non-empty line — include as-is (fallback)
                    results.append(line.strip())
            # Pad or truncate to match expected count
            while len(results) < expected_count:
                results.append("Error")
            return results[:expected_count]

        def categorize_column(col: str, categories: list[str], responses: pd.Series) -> list[str]:
            cats_str = "\n".join(f"- {c}" for c in categories)

            # Pre-fill results: mark nulls/empty as 999 immediately
            results = [""] * len(responses)
            non_empty_indices = []
            for i, resp in enumerate(responses):
                if pd.isna(resp) or str(resp).strip() == "":
                    results[i] = "999"
                else:
                    non_empty_indices.append(i)

            if not non_empty_indices:
                return results

            # Only send non-empty responses to Gemini in batches
            total_to_process = len(non_empty_indices)
            num_batches = (total_to_process + batch_size - 1) // batch_size
            skipped = len(responses) - total_to_process
            prog = st.progress(0, text=f"Duke kategorizuar **{col}** ({total_to_process} përgjigje, {skipped} bosh të kapërcyera)…")

            for batch_idx in range(num_batches):
                start = batch_idx * batch_size
                end = min(start + batch_size, total_to_process)
                batch_indices = non_empty_indices[start:end]

                followup_info = st.session_state.question_followup.get(col)
                numbered_responses = []
                for j, idx in enumerate(batch_indices):
                    resp_text = str(responses.iloc[idx])
                    if followup_info:
                        parent_answer = str(df[followup_info["column"]].iloc[idx])
                        if pd.isna(df[followup_info["column"]].iloc[idx]) or parent_answer.strip() == "":
                            parent_answer = "(no answer)"
                        numbered_responses.append(f"{j+1}. [Previous answer: {parent_answer}] {resp_text}")
                    else:
                        numbered_responses.append(f"{j+1}. {resp_text}")

                question_label = st.session_state.question_labels.get(col, col)
                if followup_info:
                    question_label = f"{question_label}\n(This is a follow-up to: \"{followup_info['label']}\" — each response includes the respondent's previous answer in [brackets] for context.)"

                prompt = st.session_state.prompt_template.format(
                    question_label=question_label,
                    categories=cats_str,
                    responses="\n".join(numbered_responses),
                )

                try:
                    text, in_tok, out_tok = call_gemini_batch(prompt)
                    token_counts["input"] += in_tok
                    token_counts["output"] += out_tok
                    batch_labels = parse_batch_response(text, len(batch_indices))
                except Exception as e:
                    st.warning(f"Gabim API në batch {batch_idx+1}: {e}")
                    batch_labels = ["Error"] * len(batch_indices)

                # Map results back to original positions
                for j, idx in enumerate(batch_indices):
                    results[idx] = batch_labels[j]

                prog.progress(end / total_to_process, text=f"Duke kategorizuar **{col}** ({end}/{total_to_process})")

            prog.empty()
            return results

        for col in question_cols:
            base_cats = [c.strip() for c in st.session_state.question_categories[col].splitlines() if c.strip()]

            with st.spinner(f"Duke procesuar **{col}**…"):
                labels = categorize_column(col, base_cats, df[col])

            # Detect high-frequency NEW categories
            new_labels = [l for l in labels if l.lower().startswith("new:")]
            new_counts = Counter(re.sub(r"(?i)^new:\s*", "", l).strip() for l in new_labels)
            promoted = [cat for cat, cnt in new_counts.items() if cnt >= new_cat_threshold]

            if promoted:
                st.info(f"Kategori të reja të detektuara për **{col}**: {', '.join(promoted)} — duke ri-ekzekutuar me listën e përditësuar…")
                updated_cats = base_cats + promoted
                # Only re-categorize responses that were tagged as NEW:
                new_indices = [i for i, l in enumerate(labels) if l.lower().startswith("new:")]
                if new_indices:
                    # Build a series with only the NEW-tagged responses, rest as NaN
                    partial_series = pd.Series([None] * len(df[col]), dtype=object)
                    for i in new_indices:
                        partial_series.iloc[i] = df[col].iloc[i]
                    partial_labels = categorize_column(col, updated_cats, partial_series)
                    # Merge: only replace labels that were NEW:
                    for i in new_indices:
                        labels[i] = partial_labels[i]

            # Clean up any remaining "NEW: X" labels
            def clean_label(l):
                m = re.match(r"(?i)^new:\s*(.+)$", l)
                return m.group(1).strip() if m else l

            labels = [clean_label(l) for l in labels]

            # Consolidate: keep top (max_categories - 1) categories, merge rest into "Other"
            label_counts = Counter(l for l in labels if l not in ("999", "Error"))
            if len(label_counts) > max_categories:
                top_cats = {cat for cat, _ in label_counts.most_common(max_categories - 1)}
                merged_count = sum(cnt for cat, cnt in label_counts.items() if cat not in top_cats)
                st.info(f"**{col}**: {len(label_counts)} kategori u gjetën → duke bashkuar {len(label_counts) - len(top_cats)} kategori me frekuencë të ulët ({merged_count} përgjigje) në 'Other'")
                labels = [l if l in top_cats or l in ("999", "Error") else "Other" for l in labels]

            result_df[f"{col}_grouped"] = labels

        # ── Cost calculation ─────────────────────────────────────────────────
        total_cost = calculate_gemini_cost(token_counts["input"], token_counts["output"], model_id)

        # Store results in session state so they persist across reruns
        output = io.BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        cols_suffix = "_".join(question_cols)
        st.session_state.results = {
            "result_df": result_df,
            "question_cols": list(question_cols),
            "id_col": id_col,
            "token_counts": dict(token_counts),
            "total_cost": total_cost,
            "excel_bytes": output.getvalue(),
            "file_name": f"categorized_responses_{cols_suffix}.xlsx",
        }
        st.rerun()

# ── Display results (persisted in session state) ────────────────────────────
if st.session_state.results is not None:
    res = st.session_state.results
    result_df = res["result_df"]

    st.markdown("---")
    for col in res["question_cols"]:
        grouped_col = f"{col}_grouped"
        if grouped_col not in result_df.columns:
            continue
        st.success(f"Përfundoi: **{col}** → **{grouped_col}**")
        st.subheader(f"Shpërndarja e kategorive — {col}")
        dist = result_df[grouped_col].value_counts().reset_index()
        dist.columns = ["Kategoria", "Numri"]
        dist["Përqindja"] = (dist["Numri"] / dist["Numri"].sum() * 100).round(1).astype(str) + "%"
        st.dataframe(dist, use_container_width=True, hide_index=True)

        st.dataframe(
            result_df[[res["id_col"], col, grouped_col]].head(20),
            use_container_width=True,
        )

    st.markdown("---")
    st.header("Përmbledhje")

    cost_col1, cost_col2, cost_col3 = st.columns(3)
    cost_col1.metric("Input tokens", f"{res['token_counts']['input']:,}")
    cost_col2.metric("Output tokens", f"{res['token_counts']['output']:,}")
    cost_col3.metric("Kostoja totale", f"${res['total_cost']:.6f}")

    st.download_button(
        label="Shkarko Excel-in e kategorizuar",
        data=res["excel_bytes"],
        file_name=res.get("file_name", "categorized_responses.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
