import os
import streamlit as st
import pandas as pd
import google.generativeai as genai
import io
import re
from collections import Counter, OrderedDict

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
    "models/gemini-3.1-flash-lite": {
        "inputPer1MTokens":  0.25,
        "outputPer1MTokens": 1.50,
    },
}


_MULTI_RESP_SEP_RE = re.compile(r"\s*[,;|/]\s*|\n+| and | & ", re.IGNORECASE)


def split_multi_response(series: pd.Series, n: int) -> list[pd.Series]:
    """Split each cell into n positions. Handles mixed delimiters per-cell (comma, semicolon,
    pipe, slash, newline, ' and ', ' & '). If no structured separator is found but the cell
    has exactly n whitespace-separated tokens, falls back to whitespace split.
    Missing pieces are filled with empty strings."""
    columns = [[] for _ in range(n)]
    for cell in series:
        if pd.isna(cell) or not str(cell).strip():
            pieces = [""] * n
        else:
            text = str(cell).strip()
            parts = [p.strip() for p in _MULTI_RESP_SEP_RE.split(text) if p.strip()]
            if len(parts) == 1:
                words = text.split()
                if len(words) == n:
                    parts = words
            pieces = parts[:n] + [""] * max(0, n - len(parts))
        for i in range(n):
            columns[i].append(pieces[i])
    return [pd.Series(c, dtype=object) for c in columns]


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
DEFAULT_PROMPT = """You are a survey response categorizer. Your ONLY task is to assign exactly ONE category from the provided list to each survey response.

Question: {question_label}

Available categories (use these EXACT names — copy-paste, do not rephrase):
{categories}

CRITICAL RULES FOR CONSISTENCY:
1. You MUST copy-paste category names EXACTLY as listed above. Do NOT paraphrase, abbreviate, reword, or create synonyms. For example, if the category is "Water supply", NEVER write "Water", "Water issues", "Water supply problems", or any variation.
2. Two responses that express the same idea MUST receive the same category, even if they use different words. For example, "water is bad", "we need clean water", and "water supply is poor" should ALL get the same water-related category.
3. DO NOT INVENT CONTRADICTING OR SPECIFIC ATTRIBUTES. Three distinct guardrails:
   a) **No contradictions.** Never place a response into a category whose stated attribute is the OPPOSITE of what the response says. Example: "Policia është e dobët" → "Policia e mirë" ❌, "Spital i keq" → "Spital i mirë" ❌.
   b) **No invented place/person names.** A category that names a specific location, organization, or person may ONLY be assigned when the response refers to that same place/person — either explicitly, OR implicitly through the question's subject (e.g., if the question is about a known mayor of X, then generic responses about "the mayor" / "kryetar komune" refer to that same place by context and CAN be assigned to the place-specific category). Only invent attributes the question itself does not establish.
   c) **Silence on sentiment is OK.** If the response is neutral about a topic (no positive/negative wording) and the only topical categories carry sentiment, you MAY still assign the closest topical category — DO NOT default to "Other" just because the response lacks sentiment. Only the contradiction rule (3a) and the place/person rule (3b) force you away from a topical match.
4. Match by underlying meaning, not by literal words. A category named with a specific trait (e.g., "I korruptuar / Nuk më pëlqen") covers semantically equivalent negative judgments (weak, incapable, dishonest, failure, no character, etc.) unless the slash-naming clearly restricts it. Likewise a general positive category covers specific positive traits (modest, polite, accessible, educated) when no more specific positive category fits.
5. When two categories both fit, pick the one whose topic and any stated attributes (place, sentiment, specifics) the response actually contains. Never pick a category that adds attributes the response contradicts or names a different place/person.
6. Use "Other" only when no category shares the response's topic, OR when every topical category is ruled out by 3a / 3b. Do not over-use "Other".
7. If the response is empty, output: 999
8. Use "NEW: <short category name>" whenever you see a clear, recurring theme that none of the existing categories cover. Don't be conservative — downstream code only promotes a NEW label to a real category if it appears at least N times, so over-proposing is safe and under-proposing forces responses into "Other".
9. The output must be in {language}, even if the answers are in other languages.
10. Output ONLY the category name per line — no explanation, no punctuation, no extra text.

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
if "question_multi_response" not in st.session_state:
    st.session_state.question_multi_response = {}
if "prompt_template" not in st.session_state or "{response}" in st.session_state.prompt_template:
    st.session_state.prompt_template = DEFAULT_PROMPT
if "language" not in st.session_state:
    st.session_state.language = "English"
if "results" not in st.session_state:
    st.session_state.results = None

# ── Settings ─────────────────────────────────────────────────────────────────
with st.expander("Konfigurimet", expanded=False):
    col_model, col_batch, col_lang = st.columns(3)
    with col_model:
        model_name = st.selectbox("Model", ["gemini-3.1-flash-lite", "gemini-2.5-flash", "gemini-2.5-pro"], index=0)
    with col_batch:
        batch_size = st.number_input("Batch size (rreshta per thirrje)", min_value=5, max_value=100, value=20, step=5)
    with col_lang:
        st.session_state.language = st.selectbox("Gjuha e output-it", ["English", "Albanian"], index=["English", "Albanian"].index(st.session_state.language))

    st.divider()
    st.subheader("Prompt Template")
    st.caption("Placeholders: `{question_label}`, `{categories}`, `{responses}`, `{language}`")
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

            # Multi-response toggle (one cell contains multiple answers)
            is_multi = st.checkbox(
                "Kjo pyetje ka shumë përgjigje në një qelizë (p.sh. 'opsioni 1, opsioni 2, ...')",
                key=f"multi_check_{col}",
                value=col in st.session_state.question_multi_response,
                help="Nëse i njëjti respondent ka dhënë disa përgjigje në një qelizë (të ndara me presje, hapësirë, pikëpresje, etj.), aktivizoje këtë. Çdo qelizë do të ndahet automatikisht në N pjesë dhe do të krijohen N kolona të kategorizuara.",
            )
            if is_multi:
                n_resp = st.number_input(
                    "Sa përgjigje pritet të ketë në çdo rresht?",
                    min_value=2, max_value=20,
                    value=st.session_state.question_multi_response.get(col, {}).get("n", 3),
                    step=1,
                    key=f"multi_n_{col}",
                    help="Pjesët e munguara (kur respondenti dha më pak përgjigje) do të shënohen si 999.",
                )
                st.session_state.question_multi_response[col] = {"n": int(n_resp)}
            elif col in st.session_state.question_multi_response:
                del st.session_state.question_multi_response[col]

            # Suggest categories button
            if st.button("Sugjero kategoritë me AI", key=f"suggest_{col}"):
                with st.spinner("Duke analizuar përgjigjet…"):
                    multi_info_suggest = st.session_state.question_multi_response.get(col)
                    if multi_info_suggest:
                        n_resp = multi_info_suggest["n"]
                        sub_series_list = split_multi_response(df[col], n_resp)
                        flat = []
                        for s in sub_series_list:
                            for v in s:
                                if v and str(v).strip():
                                    flat.append(str(v).strip())
                        sample_responses = pd.Series(flat, dtype=object)
                    else:
                        sample_responses = df[col].dropna().astype(str)
                        sample_responses = sample_responses[sample_responses.str.strip() != ""]
                    sample_size = max(1, int(0.8 * len(sample_responses)))
                    sample = sample_responses.sample(sample_size, random_state=42).tolist() if len(sample_responses) else []
                    numbered = "\n".join(f"{i+1}. {r}" for i, r in enumerate(sample))

                    q_label = st.session_state.question_labels.get(col, col)
                    lang = st.session_state.language
                    if lang == "Albanian":
                        lang_instruction = "in Albanian. If the responses are in Albanian, first understand them in their original language, then produce category names in Albanian."
                        other_label = "Tjetër"
                    else:
                        lang_instruction = "in English."
                        other_label = "Other"
                    suggest_prompt = f"""You are a survey analyst. Your task is to suggest categories that will minimize "Other" assignments by covering the most frequent response patterns.

Question: {q_label}

Sample responses:
{numbered}

STEP 1 — Frequency analysis (internal, do not output):
Read every response. Group near-identical or semantically equivalent answers together. Count each group. Rank groups from most to least frequent. Note the top patterns that together account for at least 80% of responses.

STEP 2 — Generate categories:
Create categories ONLY from the top patterns identified in Step 1. Do NOT invent categories for rare or unique responses — those belong in "{other_label}".

Rules:
1. Output one category name per line, nothing else.
2. Between 5 and 15 categories total, {lang_instruction}
3. Categories must be short (2–5 words) and specific — name the actual thing people said, not a vague umbrella.
4. Order categories by estimated frequency, most common first.
5. NEVER use generic labels like "Positive", "Negative", "Other issues", or "Miscellaneous" except for "{other_label}".
6. "{other_label}" MUST be the last line and should represent fewer than 20% of responses — if it would be higher, add more categories.
7. Do not create a category unless at least 2 responses clearly belong to it."""

                    try:
                        suggest_model = genai.GenerativeModel(model_name)
                        resp = suggest_model.generate_content(
                            suggest_prompt,
                            generation_config=genai.types.GenerationConfig(temperature=0.3, max_output_tokens=1024),
                        )
                        suggested = resp.text.strip()
                        # Clean numbered prefixes if model adds them
                        lines = []
                        for line in suggested.splitlines():
                            line = line.strip()
                            if line:
                                m = re.match(r"^\d+[\.\)\-:]\s*(.+)$", line)
                                lines.append(m.group(1).strip() if m else line)
                        suggested_cats = "\n".join(lines)
                        st.session_state.question_categories[col] = suggested_cats
                        st.session_state[f"cats_{col}"] = suggested_cats
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gabim: {e}")

            cats_key = f"cats_{col}"
            if cats_key not in st.session_state:
                st.session_state[cats_key] = st.session_state.question_categories[col]
            st.session_state.question_categories[col] = st.text_area(
                f"Categories for {col}",
                height=140,
                key=cats_key,
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

        def llm_split_cells(series: pd.Series, n: int, question_label: str, split_batch_size: int = 30) -> list[pd.Series]:
            """Use Gemini to split each non-empty cell into up to n distinct answers, using semantic
            judgment (so 'Eurostore malvesa' splits into two but 'Auto Star Mitrovica' stays as one).
            Returns a list of n pd.Series aligned with the input. Identical cells are deduplicated."""
            cells = list(series)

            unique_cells = OrderedDict()
            for cell in cells:
                if pd.isna(cell):
                    continue
                key = str(cell).strip()
                if key and key not in unique_cells:
                    unique_cells[key] = None

            unique_list = list(unique_cells.keys())
            total = len(unique_list)
            split_map: dict[str, list[str]] = {}

            if total > 0:
                num_batches = (total + split_batch_size - 1) // split_batch_size
                prog = st.progress(0, text=f"Duke ndarë me AI ({total} qeliza unike)…")
                for batch_idx in range(num_batches):
                    start = batch_idx * split_batch_size
                    end = min(start + split_batch_size, total)
                    batch = unique_list[start:end]
                    numbered = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batch))

                    split_prompt = f"""You are splitting survey responses into individual answers.

The question asked was: "{question_label}"
Each respondent was expected to give up to {n} distinct answers in ONE cell. Respondents are inconsistent — some use commas, some semicolons, some just spaces, some use "and"/"&", some use newlines. CRITICAL: a single brand/place/person/concept that happens to be multiple words (e.g. "Auto Star Mitrovica", "Coca Cola", "New York") must stay as ONE answer.

Use semantic judgment to decide what is one answer vs two. Examples:
- "Eurostore malvesa" → two separate brand names → ["Eurostore", "malvesa"]
- "Auto Star Mitrovica" → one company name → ["Auto Star Mitrovica"]
- "uji, rryma, rrugen" → three things → ["uji", "rryma", "rrugen"]
- "water and electricity" → two things → ["water", "electricity"]
- "New York and Los Angeles" → two cities → ["New York", "Los Angeles"]

Output rules:
- One line per input, in the same numbered order as the input.
- Format: "<number>. piece1 | piece2 | piece3"  (use " | " — space-pipe-space — between pieces).
- If the respondent gave only one item, output just that item with no pipe.
- Maximum {n} pieces per line. If you find more, keep the {n} most prominent.
- Do NOT add explanations, headers, or extra text. Only the numbered lines.

Inputs:
{numbered}"""

                    try:
                        text, in_tok, out_tok = call_gemini_batch(split_prompt)
                        token_counts["input"] += in_tok
                        token_counts["output"] += out_tok
                        for line in text.splitlines():
                            line = line.strip()
                            if not line:
                                continue
                            m = re.match(r"^(\d+)[\.\)\-:]\s*(.+)$", line)
                            if not m:
                                continue
                            idx_in_batch = int(m.group(1)) - 1
                            if 0 <= idx_in_batch < len(batch):
                                parts = [p.strip() for p in m.group(2).split("|") if p.strip()]
                                if parts:
                                    split_map[batch[idx_in_batch]] = parts[:n]
                    except Exception as e:
                        st.warning(f"Gabim gjatë ndarjes me AI në batch {batch_idx+1}: {e}")

                    prog.progress(end / total, text=f"Duke ndarë me AI ({end}/{total} qeliza)")
                prog.empty()

            columns = [[] for _ in range(n)]
            for cell in cells:
                if pd.isna(cell) or not str(cell).strip():
                    pieces = [""] * n
                else:
                    key = str(cell).strip()
                    parts = split_map.get(key, [key])
                    pieces = parts[:n] + [""] * max(0, n - len(parts))
                for i in range(n):
                    columns[i].append(pieces[i])
            return [pd.Series(c, dtype=object) for c in columns]

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

            # --- Deduplication: categorize each unique response text only once ---
            followup_info = st.session_state.question_followup.get(col)

            # Build a key for each response (includes parent answer for follow-ups)
            def make_key(idx):
                resp_text = str(responses.iloc[idx]).strip()
                if followup_info:
                    parent_val = df[followup_info["column"]].iloc[idx]
                    if pd.isna(parent_val) or str(parent_val).strip() == "":
                        parent_answer = "(no answer)"
                    else:
                        parent_answer = str(parent_val).strip()
                    return f"[{parent_answer}] {resp_text}"
                return resp_text

            # Map each unique key to the list of row indices that share it
            unique_keys = OrderedDict()
            for idx in non_empty_indices:
                key = make_key(idx)
                if key not in unique_keys:
                    unique_keys[key] = {"idx": idx, "rows": []}
                unique_keys[key]["rows"].append(idx)

            unique_list = list(unique_keys.items())  # [(key, {"idx": ..., "rows": [...]}), ...]
            total_unique = len(unique_list)
            total_original = len(non_empty_indices)
            deduped = total_original - total_unique
            skipped = len(responses) - total_original

            num_batches = (total_unique + batch_size - 1) // batch_size
            prog = st.progress(0, text=f"Duke kategorizuar **{col}** ({total_unique} unik nga {total_original} përgjigje, {deduped} dublikatë, {skipped} bosh)…")

            unique_labels = [""] * total_unique

            for batch_idx in range(num_batches):
                start = batch_idx * batch_size
                end = min(start + batch_size, total_unique)
                batch_items = unique_list[start:end]

                numbered_responses = []
                for j, (key, info) in enumerate(batch_items):
                    idx = info["idx"]
                    resp_text = str(responses.iloc[idx])
                    if followup_info:
                        parent_val = df[followup_info["column"]].iloc[idx]
                        if pd.isna(parent_val) or str(parent_val).strip() == "":
                            parent_answer = "(no answer)"
                        else:
                            parent_answer = str(parent_val)
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
                    language=st.session_state.language,
                )

                try:
                    text, in_tok, out_tok = call_gemini_batch(prompt)
                    token_counts["input"] += in_tok
                    token_counts["output"] += out_tok
                    batch_labels = parse_batch_response(text, len(batch_items))
                except Exception as e:
                    st.warning(f"Gabim API në batch {batch_idx+1}: {e}")
                    batch_labels = ["Error"] * len(batch_items)

                for j in range(len(batch_items)):
                    unique_labels[start + j] = batch_labels[j]

                prog.progress(end / total_unique, text=f"Duke kategorizuar **{col}** ({end}/{total_unique} unik)")

            # --- Map labels back: every duplicate row gets the same category ---
            for i, (key, info) in enumerate(unique_list):
                label = unique_labels[i]
                for row_idx in info["rows"]:
                    results[row_idx] = label

            prog.empty()
            return results

        # Expand selected columns into "virtual units": multi-response columns become N sub-units.
        # Multi-response splitting uses Gemini so brand/place names with spaces stay intact
        # (e.g. "Eurostore malvesa" → ["Eurostore","malvesa"]; "Auto Star Mitrovica" → one item).
        virtual_units = []
        for col in question_cols:
            multi_info = st.session_state.question_multi_response.get(col)
            if multi_info:
                n_resp = multi_info["n"]
                q_label_for_split = st.session_state.question_labels.get(col, col)
                with st.spinner(f"Duke ndarë **{col}** në deri në {n_resp} përgjigje me AI…"):
                    sub_series_list = llm_split_cells(df[col], n_resp, q_label_for_split)
                for i, sub in enumerate(sub_series_list):
                    virtual_units.append({
                        "source_col": col,
                        "display_col": f"{col} (përgjigja {i+1}/{n_resp})",
                        "output_col": f"{col}_{i+1}_grouped",
                        "series": sub,
                    })
            else:
                virtual_units.append({
                    "source_col": col,
                    "display_col": col,
                    "output_col": f"{col}_grouped",
                    "series": df[col],
                })

        def clean_label(l):
            m = re.match(r"(?i)^new:\s*(.+)$", l)
            return m.group(1).strip() if m else l

        def normalize_label(l):
            if l in ("999", "Error"):
                return l
            return l.rstrip(".")

        for unit in virtual_units:
            source_col = unit["source_col"]
            display_col = unit["display_col"]
            output_col = unit["output_col"]
            series = unit["series"]

            base_cats = [c.strip() for c in st.session_state.question_categories[source_col].splitlines() if c.strip()]

            with st.spinner(f"Duke procesuar **{display_col}**…"):
                labels = categorize_column(source_col, base_cats, series)

            # Detect high-frequency NEW categories
            new_labels = [l for l in labels if l.lower().startswith("new:")]
            new_counts = Counter(re.sub(r"(?i)^new:\s*", "", l).strip() for l in new_labels)
            promoted = [cat for cat, cnt in new_counts.items() if cnt >= new_cat_threshold]

            if promoted:
                st.info(f"Kategori të reja të detektuara për **{display_col}**: {', '.join(promoted)} — duke ri-ekzekutuar me listën e përditësuar…")
                updated_cats = base_cats + promoted
                new_indices = [i for i, l in enumerate(labels) if l.lower().startswith("new:")]
                if new_indices:
                    partial_series = pd.Series([None] * len(series), dtype=object)
                    for i in new_indices:
                        partial_series.iloc[i] = series.iloc[i]
                    partial_labels = categorize_column(source_col, updated_cats, partial_series)
                    for i in new_indices:
                        labels[i] = partial_labels[i]

            labels = [clean_label(l) for l in labels]
            labels = [normalize_label(l) for l in labels]

            # Canonical mapping: for each lowercased name, keep the first seen form
            canonical = {}
            for l in labels:
                if l in ("999", "Error"):
                    continue
                key = l.lower()
                if key not in canonical:
                    canonical[key] = l
            labels = [canonical.get(l.lower(), l) if l not in ("999", "Error") else l for l in labels]

            # Consolidate: keep top (max_categories - 1) categories, merge rest into "Other"
            label_counts = Counter(l for l in labels if l not in ("999", "Error"))
            if len(label_counts) > max_categories:
                top_cats = {cat for cat, _ in label_counts.most_common(max_categories - 1)}
                merged_count = sum(cnt for cat, cnt in label_counts.items() if cat not in top_cats)
                st.info(f"**{display_col}**: {len(label_counts)} kategori u gjetën → duke bashkuar {len(label_counts) - len(top_cats)} kategori me frekuencë të ulët ({merged_count} përgjigje) në 'Other'")
                labels = [l if l in top_cats or l in ("999", "Error") else "Other" for l in labels]

            result_df[output_col] = labels

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
            "grouped_units": [
                {"source_col": u["source_col"], "display_col": u["display_col"], "output_col": u["output_col"]}
                for u in virtual_units
            ],
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
    grouped_units = res.get("grouped_units")
    if not grouped_units:
        grouped_units = [
            {"source_col": col, "display_col": col, "output_col": f"{col}_grouped"}
            for col in res["question_cols"]
        ]
    for unit in grouped_units:
        source_col = unit["source_col"]
        display_col = unit["display_col"]
        grouped_col = unit["output_col"]
        if grouped_col not in result_df.columns:
            continue
        st.success(f"Përfundoi: **{display_col}** → **{grouped_col}**")
        st.subheader(f"Shpërndarja e kategorive — {display_col}")
        dist = result_df[grouped_col].value_counts().reset_index()
        dist.columns = ["Kategoria", "Numri"]
        dist["Përqindja"] = (dist["Numri"] / dist["Numri"].sum() * 100).round(1).astype(str) + "%"
        st.dataframe(dist, use_container_width=True, hide_index=True)

        st.dataframe(
            result_df[[res["id_col"], source_col, grouped_col]].head(20),
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
