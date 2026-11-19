import datetime
import re
from io import BytesIO

import pandas as pd
import streamlit as st

import question_bank as qb
import questionnaire_ai as qai
import ubo_ui

# The bank's table is wide: the side menu starts closed (the arrow at the top left opens it)
st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

SECRETS_EXAMPLE = """QUESTION_BANK_SHEET = "https://docs.google.com/spreadsheets/d/..."

[kobo_accounts.ubo]
server = "https://kf.kobotoolbox.org"
token = "..."

[kobo_accounts.ubo_eu]
server = "https://eu.kobotoolbox.org"
token = "..."

[panel_sources.ks]
surveys = "https://admin.uboconsulting-ks.com/QuestionBank/Surveys?secret=..."
questions = "https://admin.uboconsulting-ks.com/QuestionBank/Questions?secret=..."
"""


PAGE_SIZE = 15          # questionnaires per page of the table
# hidden at first, so the table fits without scrolling sideways; the "Kolonat" button shows them
HIDDEN_COLUMNS = ("metodologjite", "temat", "klienti", "perditesuar", "kopjet", "id")

# The table's toolbar (columns, download, search, full screen) is always shown, with Albanian names next to the
# icons; Streamlit shows only the icons, and only while the mouse is over the table
TOOLBAR_CSS = """<style>
[class*="st-key-bank_table"] { margin-top: 52px; }
[class*="st-key-bank_table"] .stElementToolbar {
    opacity: 1 !important; visibility: visible !important; transform: none !important; top: -50px !important;
}
[class*="st-key-bank_table"] [data-testid="stElementToolbarButtonContainer"] {
    gap: 6px; padding: 4px; background: transparent; box-shadow: none; border: 0;
}
[class*="st-key-bank_table"] [data-testid="stElementToolbarButton"] button {
    display: inline-flex; align-items: center; gap: 6px; height: 34px; width: auto; padding: 0 12px;
    border: 1px solid var(--ubo-line-2, #d0d4de); border-radius: 8px; background: var(--ubo-surface, #fff);
    color: var(--ubo-ink, inherit); font: 600 13px Inter, system-ui, sans-serif; white-space: nowrap;
}
[class*="st-key-bank_table"] [data-testid="stElementToolbarButton"] button:hover {
    border-color: var(--ubo-sel-border, #2560E6); background: var(--ubo-hover-bg, #eef3fe);
}
[class*="st-key-bank_table"] button[aria-label="Show/hide columns"]::after { content: "Fsheh / shfaq kolonat"; }
/* the grid's own search and CSV download see only the 15 rows of the page: "Kërko" above and the
   download buttons below work on the whole bank */
[class*="st-key-bank_table"] [data-testid="stElementToolbarButton"]:has(button[aria-label="Search"]),
[class*="st-key-bank_table"] [data-testid="stElementToolbarButton"]:has(button[aria-label="Download as CSV"]) { display: none; }
[class*="st-key-bank_table"] button[aria-label="Fullscreen"]::after { content: "Tabela në ekran të plotë"; }
[class*="st-key-bank_table"] button[aria-label="Close fullscreen"]::after { content: "Mbyll ekranin e plotë"; }
.st-key-bank_js { display: none; }
/* the filter row: no frame around the form, the button as tall as the fields */
.st-key-bank_filter_go button { height: 40px; }
/* Streamlit's English "Press Enter to submit form" inside the search field */
[data-testid="stForm"] [data-testid="InputInstructions"] { display: none; }
/* the update button: a tile of the same size as the numbers, the last update written inside */
.st-key-bank_update button {
    height: 98px; border-radius: 12px; flex-direction: column; gap: 6px; font-size: 16px; font-weight: 600;
}
.st-key-bank_update button p { font-size: 18px; font-weight: 700; }
.st-key-bank_update button::after { font-size: 12.5px; font-weight: 400; opacity: .85; white-space: nowrap; }
/* the numbers at the top as tiles */
[data-testid="stMetric"] {
    background: var(--ubo-surface, #fff); border: 1px solid var(--ubo-line-2, #d0d4de) !important;
    border-radius: 12px; padding: 12px 16px; box-shadow: 0 1px 2px rgba(20, 24, 40, .06);
    height: 98px; box-sizing: border-box;
    transition: border-color .15s, box-shadow .15s;
}
[data-testid="stMetric"]:hover {
    border-color: var(--ubo-sel-border, #2560E6) !important; box-shadow: 0 4px 14px rgba(20, 24, 40, .10);
}
/* the download of the whole bank, in green next to the blue update button */
.st-key-bank_download_all button {
    background: var(--ubo-surface, #fff) !important; border: 1px solid #2560E6 !important; color: #2560E6 !important;
}
.st-key-bank_download_all button:hover { background: #eef3fe !important; }
.st-key-bank_download button, .st-key-bank_download_q button {
    background: #0B8F76 !important; border-color: #0B8F76 !important; color: #fff !important;
}
.st-key-bank_download button:hover, .st-key-bank_download_q button:hover { background: #097a65 !important; border-color: #097a65 !important; }
[class*="st-key-bank_page_pick"] input[role="combobox"] { text-align: center; }
.bank-note { color: var(--ubo-muted, #6b7280); font-size: 14px; margin: 0; line-height: 1.35; }
.bank-right { text-align: right; }
[class*="st-key-bank_"] button p { white-space: nowrap; }
/* the two tabs as large buttons: the open one in green */
.st-key-bank_tab [role="tablist"] { gap: 10px; border-bottom: 0; margin-bottom: 6px; }
.st-key-bank_tab [data-testid="stTab"] {
    height: 44px; padding: 0 26px; border: 1px solid var(--ubo-line-2, #d0d4de); border-radius: 10px;
    background: var(--ubo-surface, #fff); transition: border-color .15s, background-color .15s;
}
.st-key-bank_tab [data-testid="stTab"] p { font: 600 15px Inter, system-ui, sans-serif; }
.st-key-bank_tab [data-testid="stTab"]:hover { border-color: #0B8F76; }
.st-key-bank_tab [data-testid="stTab"][aria-selected="true"] { background: #0B8F76; border-color: #0B8F76; }
.st-key-bank_tab [data-testid="stTab"][aria-selected="true"] p { color: #fff; }
.st-key-bank_tab .react-aria-SelectionIndicator, .st-key-bank_tab [role="tablist"]::after { display: none; }
/* the toolbar of the questionnaires table sits on the row of the tabs, on the right */
.st-key-bank_table { margin-top: 0 !important; }
.st-key-bank_table .stElementToolbar { top: -73px !important; }
/* thin scroll bars on this page: the page and the long tables */
* { scrollbar-width: thin; scrollbar-color: color-mix(in srgb, var(--ubo-muted, #888) 55%, transparent) transparent; }
</style>"""
st.markdown(TOOLBAR_CSS, unsafe_allow_html=True)

# initial_sidebar_state works only when the page is loaded from scratch; opened from the side menu, the menu
# stays open. So, once per visit, a small script closes it (the arrow at the top left opens it again).
COLLAPSE_SIDEBAR_JS = """<script>
(function () {
  const D = window.parent.document;
  let tries = 0;
  const timer = setInterval(() => {
    const bar = D.querySelector('[data-testid="stSidebar"]');
    const btn = D.querySelector('[data-testid="stSidebarCollapseButton"] button');
    if (bar && bar.getAttribute("aria-expanded") === "true" && btn) { btn.click(); clearInterval(timer); }
    else if ((bar && bar.getAttribute("aria-expanded") === "false") || ++tries > 50) clearInterval(timer);
  }, 100);
})();
</script><!-- %s -->"""
if st.session_state.get("ubo_page_changed", True):
    with st.container(key="bank_js"):
        st.iframe(COLLAPSE_SIDEBAR_JS % datetime.datetime.now().timestamp(), height=1)


def claude_client():
    try:
        return qai.make_client(st.secrets["ANTHROPIC_API_KEY"]), st.secrets.get("CLAUDE_MODEL", qai.MODEL)
    except Exception:
        return None, None


def local_time(utc_text):
    """ "YYYY-MM-DD HH:MM" in UTC -> "DD.MM.YYYY, ora HH:MM" in Kosovo time."""
    try:
        from zoneinfo import ZoneInfo
        t = datetime.datetime.strptime(utc_text, "%Y-%m-%d %H:%M").replace(tzinfo=datetime.timezone.utc)
        return t.astimezone(ZoneInfo("Europe/Belgrade")).strftime("%d.%m.%Y, ora %H:%M")
    except Exception:
        return f"{utc_text} (UTC)"


def fold(text):
    """Lower case without ë / ç, so "kenaqshmeria" also finds "kënaqshmëria"."""
    return text.lower().replace("ë", "e").replace("ç", "c")


@st.cache_resource(ttl=qb.BANK_CACHE_SECONDS, max_entries=2, show_spinner=False)
def question_index(bank_version, _questions):
    """(questionnaire id, question, searchable text) for every question of the bank but the notes: the text,
    the options and the rows, folded. Built once per version of the bank (~0.2 s), shared by all sessions."""
    rows = []
    for qid, qs in _questions.items():
        for q in qs:
            if q["type"] != "note":
                extra = " ".join(label for _, label in q["options"]) + " " + " ".join(q.get("items") or [])
                rows.append((qid, q, fold(f"{q['text']} {extra}")))
    return rows


def page_slice(total, state_key):
    """The page of a paged table kept in state_key: (page number, number of pages, first row, end row)."""
    n_pages = max(1, -(-total // PAGE_SIZE))
    page_no = min(st.session_state.get(state_key, 1), n_pages)
    return page_no, n_pages, (page_no - 1) * PAGE_SIZE, min(page_no * PAGE_SIZE, total)


def zebra(frame):
    """Every second row lightly shaded, so long tables read more easily (a see-through colour: it works in
    both the light and the dark theme)."""
    shade = "background-color: rgba(37, 96, 230, 0.06)"
    return frame.style.apply(lambda d: pd.DataFrame([[shade if i % 2 else ""] * d.shape[1] for i in range(len(d))],
                                                    index=d.index, columns=d.columns), axis=None)


def used_in(titles):
    """The questionnaires a question was used in, short: the first one (the "Pyetësorë" column says how many;
    clicking the question lists them all)."""
    return titles[0] if titles else ""


def table_bar(total, state_key, nouns, page_no, n_pages, hint, downloads):
    """One row under a paged table: the hint on the left, the download buttons, and on the right ‹ Para, the
    page list ("Faqja 1 (pyetësorët 1–15) nga 53 (789 pyetësorë)") and Pas ›.
    nouns: (plural for a range, e.g. "pyetësorët"; plural for the total, e.g. "pyetësorë").
    downloads: [(container key, download_button keyword arguments)]."""
    widths = ([2.0, 1.45] if len(downloads) == 1 else [1.45] + [1.35] * len(downloads)) + [0.7, 2.6, 0.7]
    hint_col, *dl_cols, prev, pick, nxt = st.columns(widths, vertical_alignment="center")
    hint_col.markdown(f'<p class="bank-note">{hint}</p>', unsafe_allow_html=True)
    for col, (key, kwargs) in zip(dl_cols, downloads):
        with col.container(key=key):
            st.download_button(width="stretch", **kwargs)

    def label(n):
        first, last = (n - 1) * PAGE_SIZE + 1, min(n * PAGE_SIZE, total)
        return f"Faqja {n} ({nouns[0]} {first}–{last}) nga {n_pages} ({total} {nouns[1]})"

    if prev.button("‹ Para", width="stretch", disabled=page_no <= 1, key=f"{state_key}_prev"):
        st.session_state[state_key] = page_no - 1
        st.rerun()
    new_page = pick.selectbox("Faqja", range(1, n_pages + 1), index=page_no - 1, label_visibility="collapsed",
                              key=f"bank_page_pick_{state_key}_{page_no}_{n_pages}_{total}", format_func=label)
    if new_page != page_no:
        st.session_state[state_key] = new_page
        st.rerun()
    if nxt.button("Pas ›", width="stretch", disabled=page_no >= n_pages, key=f"{state_key}_next"):
        st.session_state[state_key] = page_no + 1
        st.rerun()


@st.dialog("Përditëso bankën e pyetjeve")
def confirm_update():
    st.markdown(
        "Programi merr nga Kobo dhe nga Paneli pyetësorët e rinj dhe ata që kanë ndryshuar, dhe i kategorizon me AI "
        "(industria, lloji i studimit, temat).\n\n"
        "- **Kosto:** rreth 1–8 cent për çdo pyetësor të ri, zakonisht më pak se 1 $ gjithsej. Pyetësorët e "
        "kategorizuar më parë nuk kushtojnë.\n"
        "- **Koha:** disa minuta. Mos e mbyllni faqen derisa të përfundojë.\n\n"
        "Banka përditësohet vetë edhe çdo muaj, më datën 1.")
    yes, no = st.columns(2)
    if yes.button("Po, përditëso", type="primary", icon=":material/refresh:", width="stretch"):
        st.session_state["bank_run_update"] = True
        st.rerun()
    if no.button("Anulo", width="stretch"):
        st.rerun()


with ubo_ui.card("", key="banka"):     # the page header already says what it is
    try:
        qb.kobo_accounts(st.secrets)
    except qb.BankError as e:
        st.warning(str(e))
        st.markdown("Shtoni te secrets e aplikacionit lidhjen e Google Sheet-it të bankës dhe një seksion për "
                    "çdo llogari (tokeni gjendet te Kobo: **Account Settings → Security**). Google Sheet-i "
                    "ndahet (Editor) me e-mailin e `gcp_service_account`.")
        st.code(SECRETS_EXAMPLE, language="toml")
        st.stop()

    try:
        catalog, questions = qb.cached_bank()
    except qb.BankError as e:
        st.warning(str(e))
        catalog, questions = [], {}
    # tests, training forms and applications (marked "jo" by the AI agent) are left out everywhere
    usable_ids = {r["id"] for r in catalog if r.get("perdorshem") != "jo"}
    df = pd.DataFrame([r for r in catalog if r["id"] in usable_ids], columns=qb.CATALOG_COLUMNS)
    questions = {qid: qs for qid, qs in questions.items() if qid in usable_ids}
    for col in ("nr_pyetjeve", "dergesat", "kopjet"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    last = qb.cached_last_sync()
    last_text = f"Përditësuar: {local_time(last)}" if last else "Përditësuar: —"
    # written here, outside the button's column, so the button stays level with the tiles
    st.markdown(f'<style>.st-key-bank_update button::after {{ content: "{last_text}"; }}</style>',
                unsafe_allow_html=True)
    m1, m2, m3, m4, m5, upd = st.columns([1, 1, 1, 1, 1, 1.4], vertical_alignment="center")
    m1.metric("Pyetësorë", f"{len(df):,}")
    m2.metric("Pyetje", f"{sum(q['type'] != 'note' for v in questions.values() for q in v):,}")
    m3.metric("Pyetje unike", f"{len({q['unique'] for v in questions.values() for q in v if q.get('unique')}):,}",
              help="Pyetjet me tekst të njëjtë ose shumë të ngjashëm numërohen një herë, p.sh. \"Ju lutem komentoni\" "
                   "dhe \"Ju lutemi komentoni:\", ose e njëjta pyetje për politikanë apo periudha të ndryshme.")
    m4.metric("Industri", df["industria"].replace("", pd.NA).nunique())
    m5.metric("Lloje studimi", df["lloji"].replace("", pd.NA).nunique())
    def bank_excel(catalog_rows):
        """The questionnaires of catalog_rows and their questions as Excel, as a callable for the download button:
        it is built only when the button is clicked (the whole bank takes ~15 s, which every click on the page
        would otherwise wait for)."""
        def build():
            ids = set(catalog_rows["id"])
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                catalog_rows.drop(columns=["perdorshem"]).to_excel(writer, sheet_name="Pyetesoret", index=False)
                pd.DataFrame([{"pyetesori_id": qid, **q, "options": "; ".join(f"{c}={l}" for c, l in q["options"]),
                               "items": " | ".join(q.get("items") or [])}
                              for qid, qs in questions.items() if qid in ids for q in qs]
                             ).to_excel(writer, sheet_name="Pyetjet", index=False)
            return output.getvalue()
        return build

    with upd:
        # The update costs (Claude categorizes the new questionnaires): it starts only after the confirmation
        # a tile as tall as the numbers beside it, with the time of the last update inside, under the label
        if st.button("Përditëso bankën", type="primary", icon=":material/refresh:", width="stretch",
                     key="bank_update", help=last_text):
            confirm_update()
    clicked = st.session_state.pop("bank_run_update", False)

    if clicked:
        client, model = claude_client()
        if client is None:
            st.warning("Mungon `ANTHROPIC_API_KEY`: pyetësorët përditësohen pa u kategorizuar.")
        bar = st.progress(0.0, text="Po lidhet me KoboToolbox dhe Panelin...")
        try:
            st.session_state["bank_stats"] = qb.sync(
                st.secrets, client, model, False,
                progress=lambda frac, text: bar.progress(min(frac, 1.0), text=text))
            ok = True
        except qb.BankError as e:
            st.error(str(e))
            ok = False
        finally:
            bar.empty()
            qb.cached_bank.clear()
            qb.cached_last_sync.clear()
        if ok:
            st.rerun()      # new numbers and time of the last update

    stats = st.session_state.get("bank_stats")
    if stats:
        st.success(f"Banka u përditësua: {stats['categorized']} pyetësorë të rinj u kategorizuan "
                   f"(kostoja ${stats['cost']:.2f}).")
        if stats["failed"]:
            st.warning(f"{len(stats['failed'])} pyetësorë nuk u lexuan nga Kobo (mbetën si ishin në bankë).")
        for err in stats.get("panel_failed") or []:
            st.warning(f"{err} Pyetësorët e Panelit mbetën në bankë si ishin.")
        if stats.get("panel"):
            st.caption("Nga Paneli: " + ", ".join(f"{n} pyetësorë ({name})" for name, n in stats["panel"].items()))
        if stats["errors"]:
            with st.expander(f"{len(stats['errors'])} pyetësorë nuk u kategorizuan (provoni përsëri)"):
                for err in stats["errors"]:
                    st.write(err)
    if df.empty:
        st.info("Banka është ende bosh: shtypni **Përditëso bankën**.")
        st.stop()

    # The filters apply when "Filtro" is pressed (or Enter in the search field), not at every choice
    with st.form("bank_filters_form", border=False):
        f1, f2, f3, f4, go = st.columns([1, 1, 1, 1, 0.55], vertical_alignment="bottom")
        industry = f1.multiselect("Industria", sorted(x for x in df["industria"].unique() if x),
                                  placeholder="Të gjitha", key="bank_f_industry")
        kind = f2.multiselect("Lloji", sorted(x for x in df["lloji"].unique() if x), placeholder="Të gjitha",
                              key="bank_f_kind")
        method = f3.multiselect("Metodologjia", qb.METHODS, placeholder="Të gjitha", key="bank_f_method")
        search = f4.text_input("Kërko", placeholder="titulli, klienti, tema...", key="bank_f_search")
        with go.container(key="bank_filter_go"):
            st.form_submit_button("Filtro", icon=":material/filter_alt:", type="primary", width="stretch")
    view = df
    if industry:
        view = view[view["industria"].isin(industry)]
    if kind:
        view = view[view["lloji"].isin(kind)]
    if method:
        view = view[view["metodologjite"].apply(lambda s: any(m in s for m in method))]
    if search:
        text = (view["titulli"] + " " + view["klienti"] + " " + view["temat"] + " " + view["pershkrimi"]).str.lower()
        view = view[text.str.contains(search.lower(), regex=False)]

    columns = {"titulli": "Titulli", "industria": "Industria", "lloji": "Lloji", "metodologjite": "Metodologjitë",
               "temat": "Temat", "klienti": "Klienti", "nr_pyetjeve": "Pyetje", "dergesat": "Dërgesa",
               "perditesuar": "Përditësuar", "kopjet": "Kopje", "id": "ID"}
    filters = (tuple(industry), tuple(kind), tuple(method), search)
    if st.session_state.get("bank_filters") != filters:      # other filters: back to the first page
        st.session_state["bank_filters"] = filters
        st.session_state["bank_page"] = 1
        st.session_state["bank_q_page"] = 1
    xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # stateful tabs: a search or a page change keeps the open tab (plain tabs go back to the first one)
    tab_forms, tab_questions = st.tabs(["Pyetësorët", "Pyetjet"], key="bank_tab", on_change="rerun")

    with tab_forms:
        # A page of PAGE_SIZE questionnaires at a time, the table as tall as its rows: no long scrolling inside it
        page_no, n_pages, start, end = page_slice(len(view), "bank_page")
        page_view = view.iloc[start:end]
        with st.container(key="bank_table"):
            # a new key per page and filter, so a row picked on one page is not taken for a row of another
            picked = st.dataframe(zebra(page_view[list(columns)].rename(columns=columns)), hide_index=True,
                                  width="stretch",
                                  column_order=[label for col, label in columns.items() if col not in HIDDEN_COLUMNS],
                                  height="content", on_select="rerun", selection_mode="single-row",
                                  key=f"bank_rows_{page_no}_{hash(tuple(page_view['id']))}")
        # with filters, the results (green) and the whole bank (blue); without filters, the whole bank (green)
        filtered = len(view) < len(df)
        downloads = []
        if filtered:
            downloads.append(("bank_download", dict(
                label=f"Shkarko rezultatet ({len(view)})", data=bank_excel(view), disabled=view.empty,
                file_name="banka_e_pyetjeve_filtruar.xlsx", mime=xlsx, icon=":material/filter_alt:",
                help="Excel me pyetësorët që kanë mbetur pas filtrave dhe kërkimit, bashkë me pyetjet e tyre.")))
        downloads.append(("bank_download_all" if filtered else "bank_download", dict(
            label="Shkarko gjithë bankën", data=bank_excel(df), file_name="banka_e_pyetjeve.xlsx", mime=xlsx,
            icon=":material/download:", help="Excel me të gjithë pyetësorët e bankës, bashkë me pyetjet e tyre.")))
        table_bar(len(view), "bank_page", ("pyetësorët", "pyetësorë"), page_no, n_pages,
                  "Klikoni një rresht për të parë pyetjet e pyetësorit.", downloads)

        rows = picked.selection.rows if picked else []
        if rows:
            row = page_view.iloc[rows[0]]
            st.markdown(f"**{row['titulli']}** — {row['pershkrimi']}")
            st.caption(f"Burimet: {row['burimet']}")
            st.dataframe(pd.DataFrame([{"Nr": q["nr"], "Emri": q["name"], "Lloji": q["type"], "Grupi": q["group"],
                                        "Pyetja": q["text"], "Rreshtat": " | ".join(q.get("items") or []),
                                        "Opsionet": "; ".join(f"{c}={l}" for c, l in q["options"])}
                                       for q in questions.get(row["id"], [])]), hide_index=True, width="stretch")

    with tab_questions:
        # Search in the questions themselves; the same question used in several questionnaires is one row
        search_col, button_col = st.columns([5, 1], vertical_alignment="bottom")
        # the button only submits the field (Enter does the same); the search runs on every rerun
        button_col.button("Kërko", icon=":material/search:", type="primary", width="stretch", key="bank_q_go")
        q_search = search_col.text_input("Kërko në pyetje", key="bank_q_search",
                                 placeholder="Shkruani një fjalë, p.sh. mosha: dalin pyetjet, opsionet e tyre "
                                             "dhe pyetësorët ku janë përdorur",
                                 help="Kërkon te teksti i pyetjeve, opsionet dhe rreshtat, në fillim të fjalëve: \"mosh\" gjen "
                                      "\"mosha\" dhe \"moshën\". Shkronjat ë dhe ç "
                                      "mund të shkruhen edhe e dhe c. Me disa fjalë gjenden pyetjet që i kanë të "
                                      "gjitha. Filtrat lart (industria, lloji, metodologjia, kërkimi) vlejnë edhe "
                                      "këtu.")
        words = fold(q_search).split()
        if st.session_state.get("bank_q_words") != words:
            st.session_state["bank_q_words"] = words
            st.session_state["bank_q_page"] = 1
        if words:
            view_ids = set(view["id"])
            titles = dict(zip(df["id"], df["titulli"]))
            index = question_index((qb.cached_last_sync(), len(questions)), questions)
            # each word at the start of a word of the question: "mosha" finds "Mosha juaj", not "bashkëmoshatarët"
            patterns = [re.compile(r"(?<!\w)" + re.escape(w)) for w in words]
            groups = {}
            for qid, q, text in index:
                if qid in view_ids and all(pat.search(text) for pat in patterns):
                    g = groups.setdefault(q.get("unique") or fold(q["text"]).strip(), {"q": q, "uses": []})
                    g["uses"].append((qid, q))
            found = sorted(groups.values(), key=lambda g: -len({qid for qid, _ in g["uses"]}))
            if not found:
                st.info("Asnjë pyetje nuk u gjet. Provoni një fjalë tjetër ose hiqni filtrat lart.")
            else:
                result = pd.DataFrame([{
                    "Pyetja": g["q"]["text"],
                    "Opsionet": "; ".join(label for _, label in g["q"]["options"]),
                    "Lloji": g["q"]["type"],
                    "Pyetësorë": len({qid for qid, _ in g["uses"]}),
                    "Përdorur p.sh. te": used_in([titles.get(qid, qid) for qid, _ in g["uses"]]),
                } for g in found])
                page_no, n_pages, start, end = page_slice(len(found), "bank_q_page")
                with st.container(key="bank_table_q"):
                    # fixed widths so the table fits without scrolling sideways; long texts end in "…", the
                    # whole text shows when the cell is hovered, and the full list when the question is clicked
                    picked_q = st.dataframe(zebra(result.iloc[start:end]), hide_index=True, width="stretch",
                                            column_config={"Pyetja": st.column_config.TextColumn(width="large"),
                                                           "Opsionet": st.column_config.TextColumn(width="medium"),
                                                           "Lloji": st.column_config.TextColumn(width="small"),
                                                           "Pyetësorë": st.column_config.NumberColumn(width="small"),
                                                           "Përdorur p.sh. te": st.column_config.TextColumn(width="medium")},
                                            height="content", on_select="rerun", selection_mode="single-row",
                                            key=f"bank_q_rows_{page_no}_{hash(tuple(words))}_{len(found)}")
                def found_excel():
                    out = BytesIO()
                    pd.DataFrame([{
                        "pyetja": q["text"], "pyetja_unike": q.get("unique", ""), "lloji": q["type"],
                        "nr": q["nr"], "emri": q["name"], "pyetesori": titles.get(qid, qid), "pyetesori_id": qid,
                        "opsionet": "; ".join(f"{c}={l}" for c, l in q["options"]),
                        "rreshtat": " | ".join(q.get("items") or []),
                    } for g in found for qid, q in g["uses"]]).to_excel(out, sheet_name="Pyetjet", index=False)
                    return out.getvalue()

                table_bar(len(found), "bank_q_page", ("pyetjet", "pyetje"), page_no, n_pages,
                          "Klikoni një pyetje për të parë ku është përdorur.",
                          [("bank_download_q", dict(
                              label=f"Shkarko pyetjet ({len(found)})", data=found_excel,
                              file_name="pyetjet_e_gjetura.xlsx", mime=xlsx, icon=":material/download:",
                              help="Excel me çdo përdorim të pyetjeve të gjetura: pyetësori, teksti dhe opsionet."))])

                q_rows = picked_q.selection.rows if picked_q else []
                if q_rows:
                    # every questionnaire that used it, with its own wording and options
                    g = found[start + q_rows[0]]
                    st.markdown(f"**{g['q']['text']}** — përdorur në {len({qid for qid, _ in g['uses']})} "
                                f"pyetësorë")
                    info = df.set_index("id")
                    st.dataframe(pd.DataFrame([{
                        "Pyetësori": titles.get(qid, qid),
                        "Industria": info.at[qid, "industria"] if qid in info.index else "",
                        "Lloji i studimit": info.at[qid, "lloji"] if qid in info.index else "",
                        "Nr": q["nr"], "Emri": q["name"], "Pyetja": q["text"],
                        "Opsionet": "; ".join(f"{c}={l}" for c, l in q["options"]),
                    } for qid, q in g["uses"]]), hide_index=True, width="stretch")
