import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

import question_bank as qb
import questionnaire_ai as qai
import ubo_ui

st.set_page_config(layout="wide")

SECRETS_EXAMPLE = """QUESTION_BANK_SHEET = "https://docs.google.com/spreadsheets/d/..."

[kobo_accounts.ubo]
server = "https://kf.kobotoolbox.org"
token = "..."

[kobo_accounts.ubo_eu]
server = "https://eu.kobotoolbox.org"
token = "..."
"""


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

    m1, m2, m3, m4, m5, upd = st.columns([1, 1, 1, 1, 1, 1.4], vertical_alignment="center")
    m1.metric("Pyetësorë", f"{len(df):,}")
    m2.metric("Pyetje", f"{sum(q['type'] != 'note' for v in questions.values() for q in v):,}")
    m3.metric("Pyetje unike", f"{len({q['unique'] for v in questions.values() for q in v if q.get('unique')}):,}",
              help="Pyetjet me tekst të njëjtë ose shumë të ngjashëm numërohen një herë, p.sh. \"Ju lutem komentoni\" "
                   "dhe \"Ju lutemi komentoni:\", ose e njëjta pyetje për politikanë apo periudha të ndryshme.")
    m4.metric("Industri", df["industria"].replace("", pd.NA).nunique())
    m5.metric("Lloje studimi", df["lloji"].replace("", pd.NA).nunique())
    with upd:
        clicked = st.button("Përditëso bankën", type="primary", icon=":material/refresh:", width="stretch")
        last = qb.cached_last_sync()
        st.caption(f"Përditësimi i fundit: {local_time(last)}" if last else "Përditësimi i fundit: —")

    if clicked:
        client, model = claude_client()
        if client is None:
            st.warning("Mungon `ANTHROPIC_API_KEY`: pyetësorët përditësohen pa u kategorizuar.")
        bar = st.progress(0.0, text="Po lidhet me KoboToolbox...")
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
        if stats["errors"]:
            with st.expander(f"{len(stats['errors'])} pyetësorë nuk u kategorizuan (provoni përsëri)"):
                for err in stats["errors"]:
                    st.write(err)
    if df.empty:
        st.info("Banka është ende bosh: shtypni **Përditëso bankën**.")
        st.stop()

    f1, f2, f3, f4 = st.columns(4)
    industry = f1.multiselect("Industria", sorted(x for x in df["industria"].unique() if x), placeholder="Të gjitha")
    kind = f2.multiselect("Lloji", sorted(x for x in df["lloji"].unique() if x), placeholder="Të gjitha")
    method = f3.multiselect("Metodologjia", qb.METHODS, placeholder="Të gjitha")
    search = f4.text_input("Kërko", placeholder="titulli, klienti, tema...")
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
    picked = st.dataframe(view[list(columns)].rename(columns=columns), hide_index=True, width="stretch",
                          on_select="rerun", selection_mode="single-row", key="bank_table")
    rows = picked.selection.rows if picked else []
    if rows:
        row = view.iloc[rows[0]]
        st.markdown(f"**{row['titulli']}** — {row['pershkrimi']}")
        st.caption(f"Burimet: {row['burimet']}")
        st.dataframe(pd.DataFrame([{"Nr": q["nr"], "Emri": q["name"], "Lloji": q["type"], "Grupi": q["group"],
                                    "Pyetja": q["text"], "Rreshtat": " | ".join(q.get("items") or []),
                                    "Opsionet": "; ".join(f"{c}={l}" for c, l in q["options"])}
                                   for q in questions.get(row["id"], [])]), hide_index=True, width="stretch")
    else:
        st.caption("Klikoni një rresht për të parë pyetjet e pyetësorit.")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.drop(columns=["perdorshem"]).to_excel(writer, sheet_name="Pyetesoret", index=False)
        pd.DataFrame([{"pyetesori_id": qid, **q, "options": "; ".join(f"{c}={l}" for c, l in q["options"]),
                       "items": " | ".join(q.get("items") or [])}
                      for qid, qs in questions.items() for q in qs]).to_excel(writer, sheet_name="Pyetjet", index=False)
    st.download_button("Shkarko bankën (Excel)", output.getvalue(), file_name="banka_e_pyetjeve.xlsx")
