import html
import json
import re

import streamlit as st

import question_bank as qb
import questionnaire_ai as qai
import questionnaire_design as qd
import ubo_ui

st.set_page_config(layout="wide")
ss = st.session_state

try:
    client = qai.make_client(st.secrets["ANTHROPIC_API_KEY"])
    model = st.secrets.get("CLAUDE_MODEL", qai.MODEL)
except Exception:
    st.error("Mungon çelësi `ANTHROPIC_API_KEY` në secrets të aplikacionit.")
    st.stop()

try:
    catalog, bank_questions = qb.cached_bank()
except qb.BankError as e:
    st.warning(f"{e} Banka plotësohet te vegla **Banka e pyetjeve**.")
    st.stop()
usable = qd.usable(catalog)
if not usable:
    st.info("Banka e pyetjeve është bosh: plotësojeni fillimisht te vegla **Banka e pyetjeve**.")
    st.stop()
rows_by_id = {r["id"]: r for r in catalog}


def add_cost(cost):
    ss["qd_cost"] = ss.get("qd_cost", 0.0) + cost


def run(fn, *args, **kwargs):
    """A Claude call that the user can stop; errors are shown on the page. Returns (result, ok)."""
    try:
        return ubo_ui.interruptible(fn, *args, **kwargs), True
    except qai.Cancelled:
        return None, False
    except qai.AIError as e:
        st.error(str(e))
        return None, False


QUESTION_TEXT = re.compile(r'"text"\s*:\s*"((?:[^"\\]|\\.)*)"')


def show_design_progress(bar, text, target):
    """The progress bar while Claude writes the questionnaire: the step, and which question it is writing.
    text is the JSON answer streamed so far; every question starts with its "section" key."""
    n = text.count('"section"')
    if n == 0:
        label = ("Agjenti AI po shkruan titullin dhe hyrjen..." if '"title"' in text else
                 "Agjenti AI po lexon pyetjet e bankës dhe po planifikon rrjedhën e pyetësorit (disa minuta)...")
        bar.progress(0.03 if text else 0.0, text=label)
        return
    if '"notes"' in text:
        bar.progress(0.98, text=f"Agjenti AI po shkruan shënimet ({n} pyetje gati)...")
        return
    texts = QUESTION_TEXT.findall(text)
    current = ""
    if len(texts) == n:     # the current question's text is complete
        try:
            current = json.loads(f'"{texts[-1]}"')
        except ValueError:
            current = texts[-1]
    if len(current) > 90:
        current = current[:90] + "…"
    bar.progress(min(0.05 + 0.9 * n / max(target, 1), 0.95),
                 text=f"Agjenti AI po shkruan pyetjen {n} nga ~{target}" + (f": {current}" if current else "..."))


# ---------------------------------------------------------------------------
# 1. The study
# ---------------------------------------------------------------------------

with ubo_ui.card("Studimi", step=1, subtitle="Pyetësori ndërtohet nga pyetjet e pyetësorëve të kaluar në bankë."):
    objectives = st.text_area("Objektivat e studimit", key="qd_objectives", height=130,
                              placeholder="P.sh. Matja e kënaqshmërisë së klientëve të bankave me shërbimet në "
                                          "degë dhe ato digjitale, arsyet e ndërrimit të bankës, NPS...")
    c1, c2, c3 = st.columns(3)
    # both are required: they decide which past questionnaires Claude picks as material
    industry = c1.selectbox("Industria", qb.INDUSTRIES, index=None, key="qd_industry",
                            placeholder="Zgjidhni industrinë")
    survey_type = c2.selectbox("Lloji i studimit", qb.SURVEY_TYPES, index=None, key="qd_type",
                               placeholder="Zgjidhni llojin")
    methods = c3.multiselect("Metodologji e veçantë", qb.METHODS, key="qd_methods", placeholder="Asnjë")
    c4, c5, c6, c7 = st.columns(4)
    n_questions = c4.number_input("Numri i përafërt i pyetjeve", 5, 200, 30, step=5, key="qd_n")
    language = c5.selectbox("Gjuha e pyetësorit", list(qd.LANGUAGES), key="qd_language")
    mode = c6.selectbox("Mbledhja e të dhënave", list(qd.MODES), index=None, key="qd_mode",
                        placeholder="CATI, CAPI apo CAWI", help="Udhëzimet e pyetjeve, hyrja dhe shënimet "
                        "përshtaten sipas mënyrës: telefon (CATI), ballë për ballë (CAPI) apo online (CAWI).")
    audience = c7.text_input("Grupi i synuar", key="qd_audience", placeholder="P.sh. klientë të bankave 18+")
    notes = st.text_area("Kërkesa të tjera (opsionale)", key="qd_notes", height=70,
                         placeholder="P.sh. intervistë telefonike deri në 15 minuta; pa pyetje për të ardhurat")
    brief = {"objectives": objectives.strip(), "industry": industry or "", "survey_type": survey_type or "",
             "methods": methods, "n_questions": n_questions, "language": language, "audience": audience.strip(),
             "notes": notes.strip(), "mode": mode or ""}
    missing = [name for name, value in (("objektivat", brief["objectives"]), ("industrinë", industry),
                                        ("llojin e studimit", survey_type), ("mbledhjen e të dhënave", mode))
               if not value]
    if missing:
        st.caption("Plotësoni " + ", ".join(missing) + " për të vazhduar.")

    if st.button("Gjej pyetësorët e ngjashëm", type="primary", disabled=bool(missing)):
        with st.spinner("Agjenti AI po kërkon pyetësorët më të përshtatshëm..."):
            result, ok = run(qd.select_questionnaires, client, model, brief, catalog)
        if ok:
            picked, cost = result
            add_cost(cost)
            ss["qd_reasons"] = {r["id"]: reason for r, reason in picked}
            ss["qd_selected"] = [r["id"] for r, _ in picked]
            if not picked:
                st.warning("Agjenti AI nuk gjeti pyetësorë të ngjashëm; zgjidhni vetë më poshtë.")

# ---------------------------------------------------------------------------
# 2. Source questionnaires
# ---------------------------------------------------------------------------

with ubo_ui.card("Pyetësorët burim nga banka", step=2,
                 subtitle=f"Pyetjet e tyre (deri në {qd.MAX_SOURCES} pyetësorë) janë materiali i pyetësorit të ri."):
    def label(qid):
        r = rows_by_id[qid]
        return f"{r['titulli']} · {r.get('industria') or '–'} · {r.get('lloji') or '–'} · {r.get('nr_pyetjeve')} pyetje"

    selected = st.multiselect("Pyetësorët", [r["id"] for r in usable], key="qd_selected", format_func=label,
                              max_selections=qd.MAX_SOURCES, placeholder="Zgjidhni pyetësorët ose përdorni hapin 1")
    reasons = ss.get("qd_reasons", {})
    for qid in selected:
        r = rows_by_id[qid]
        why = f" — _{reasons[qid]}_" if qid in reasons else ""
        st.markdown(f"- **{r['titulli']}**{why}  \n  <small>{r.get('pershkrimi', '')}</small>", unsafe_allow_html=True)

    # A prompt for the design: enough on its own when the questionnaires were picked by hand without step 1
    prompt = st.text_area(
        "Udhëzim për Agjentin AI" + ("" if brief["objectives"] else " (ose plotësoni hapin 1)"), key="qd_prompt",
        height=90, placeholder="P.sh. Pyetësor CATI për kënaqshmërinë e vizitorëve të qendrave tregtare, rreth 25 "
                               "pyetje; merr pyetjet për njohjen dhe përdorimin nga PRISHTINA MALL SURVEY 2026")
    brief = dict(brief)
    if prompt.strip():
        if brief["objectives"]:
            brief["notes"] = "\n".join(x for x in (brief["notes"], prompt.strip()) if x)
        else:
            brief["objectives"] = prompt.strip()
    if selected and not brief["objectives"]:
        st.caption("Shkruani udhëzimin për Agjentin AI ose plotësoni objektivat te hapi 1.")

    if st.button("Dizajno pyetësorin", type="primary", disabled=not (selected and brief["objectives"])):
        context, sources = qd.bank_context([rows_by_id[i] for i in selected], bank_questions)
        bar = st.progress(0.0, text="Agjenti AI po lexon pyetjet e bankës dhe po planifikon rrjedhën e pyetësorit (disa minuta)...")
        try:
            result, ok = run(qd.design, client, model, brief, context, sources,
                             on_text=lambda text: show_design_progress(bar, text, brief["n_questions"]))
        finally:
            bar.empty()
        if ok:
            questionnaire, cost = result
            add_cost(cost)
            for q in questionnaire["questions"]:
                ss["qd_next"] = ss.get("qd_next", 0) + 1
                q["uid"] = ss["qd_next"]
            ss.update(qd_q=questionnaire, qd_context=context, qd_sources=sources, qd_brief=brief,
                      qd_title=questionnaire["title"], qd_intro=questionnaire["introduction"])

# ---------------------------------------------------------------------------
# 3. Editor
# ---------------------------------------------------------------------------

questionnaire = ss.get("qd_q")
if questionnaire:
    def new_uid():
        ss["qd_next"] = ss.get("qd_next", 0) + 1
        return ss["qd_next"]

    def regenerate(index, instruction, new=False):
        with st.spinner("Agjenti AI po shkruan pyetjen..."):
            result, ok = run(qd.regenerate_question, client, model, ss["qd_brief"], ss["qd_context"],
                             ss["qd_sources"], questionnaire, index, instruction, new)
        if not ok:
            return False
        q, cost = result
        add_cost(cost)
        q["uid"] = new_uid()        # new fields for the new text
        codes = {x.get("code") for x in questionnaire["questions"]}
        if not new and questionnaire["questions"][index].get("code"):
            q["code"] = questionnaire["questions"][index]["code"]       # filters elsewhere refer to it
        elif not q.get("code") or q["code"] in codes:
            n = 1
            while f"Q{n}" in codes:
                n += 1
            q["code"] = f"Q{n}"
        if new:
            questionnaire["questions"].insert(index + 1, q)
        else:
            questionnaire["questions"][index] = q
        return True

    qs = questionnaire["questions"]
    from_bank = sum(1 for q in qs if q.get("source"))
    with ubo_ui.card("Pyetësori", step=3,
                     subtitle=f"{len(qs)} pyetje · {from_bank} nga banka · {len(qs) - from_bank} të reja. "
                              "Çdo pyetje mund të ndryshohet këtu ose të rigjenerohet me një udhëzim shtesë."):
        with st.expander(f"Titulli dhe hyrja: {questionnaire['title']}"):
            questionnaire["title"] = st.text_input("Titulli", key="qd_title")
            questionnaire["introduction"] = st.text_area("Hyrja", key="qd_intro", height=110)
        if questionnaire["notes"]:
            with st.expander("Shënimet e Agjentit AI për këtë pyetësor"):
                for note in questionnaire["notes"]:
                    st.markdown(f"- {note}")

        # Each question is a card with a coloured edge: its number beside the text, the options under it, the
        # regeneration at the bottom; the section is a heading above the first question of the section.
        st.markdown("""<style>
[class*="st-key-qd_box_"] { border-left: 4px solid #22C6A0 !important; background: var(--ubo-surface-2, #F6F7F9);
                            padding: 14px 16px 10px !important; margin-top: 2px; }
.qd-num { display: grid; place-items: center; width: 34px; height: 34px; border-radius: 10px; margin-top: 2px;
          background: #22C6A0; color: #fff; font: 700 15px "Space Grotesk", system-ui, sans-serif; }
.qd-code { text-align: center; margin-top: 4px; font: 600 11px "Inter", system-ui, sans-serif; color: #0B8F76; }
.qd-filter { font: 600 13px "Inter", system-ui, sans-serif; color: #2560E6; margin: 0 0 6px; }
.qd-sec { margin: 22px 0 4px; font: 700 13px "Inter", system-ui, sans-serif; letter-spacing: .1em;
          text-transform: uppercase; color: #0B8F76; border-bottom: 2px solid #22C6A022; padding-bottom: 6px; }
[class*="st-key-qd_box_"] [data-testid="stTextArea"] textarea { font-size: 15px; }
[class*="st-key-qd_q_"] textarea { font-weight: 600; }
</style>""", unsafe_allow_html=True)

        action = None
        section = None
        for i, q in enumerate(qs):
            k = lambda field: f"qd_{q['uid']}_{field}"
            defaults = {"text": q["text"], "type": q["type"], "section": q["section"],
                        "options": qd.options_text(q["options"]), "items": "\n".join(q["items"]),
                        "instruction": q["instruction"], "prompt": "", "code": q.get("code") or "",
                        "filter": q.get("filter") or ""}
            for field, value in defaults.items():
                if k(field) not in ss:
                    ss[k(field)] = value
            if q["section"] and q["section"] != section:
                section = q["section"]
                st.markdown(f'<div class="qd-sec">{html.escape(section)}</div>', unsafe_allow_html=True)

            with st.container(border=True, key=f"qd_box_{q['uid']}"):
                num, body, tools = st.columns([0.45, 10, 1.6], gap="small")
                num.markdown(f'<div class="qd-num">{i + 1}</div><div class="qd-code">{html.escape(ss[k("code")])}</div>',
                             unsafe_allow_html=True)
                with body:
                    q["type"] = ss[k("type")]       # its selectbox comes later, under "Më shumë"
                    if ss[k("filter")].strip():     # who is asked; edited under "Më shumë"
                        st.markdown(f'<div class="qd-filter">↳ {html.escape(ss[k("filter")])}</div>',
                                    unsafe_allow_html=True)
                    with st.container(key=f"qd_q_{q['uid']}"):
                        q["text"] = st.text_area("Pyetja", key=k("text"), height=68, label_visibility="collapsed")
                    if q["type"] in qd.ITEM_TYPES:
                        lines = len(ss[k("items")].splitlines())
                        q["items"] = [l.strip() for l in st.text_area(
                            "Elementet / pohimet (një për rresht)", key=k("items"),
                            height=min(max(68, 24 * lines + 26), 320)).splitlines() if l.strip()]
                    if q["type"] in qd.OPTION_TYPES:
                        lines = len(ss[k("options")].splitlines())
                        q["options"] = qd.parse_options(st.text_area(
                            "Opsionet (një për rresht)", key=k("options"),
                            height=min(max(68, 24 * lines + 26), 320)))
                    with st.expander("Më shumë: filtri, lloji, seksioni, udhëzimi, burimi"):
                        q["filter"] = st.text_input(
                            "Filtri (kujt i bëhet pyetja)", key=k("filter"),
                            placeholder="Bosh: të gjithëve. P.sh. Pyet nëse S3 = 1 (Po)",
                            help="Me kodet e pyetjeve (S1, Q4 …). Në Word kodet zëvendësohen me numrat e pyetjeve.")
                        c_type, c_sec, c_code = st.columns([2, 2, 1])
                        q["type"] = c_type.selectbox("Lloji", list(qd.QUESTION_TYPES), key=k("type"),
                                                     format_func=qd.QUESTION_TYPES.get)
                        q["section"] = c_sec.text_input("Seksioni", key=k("section"))
                        q["code"] = c_code.text_input("Kodi", key=k("code"))
                        q["instruction"] = st.text_input("Udhëzimi (për anketuesin / të anketuarin)",
                                                         key=k("instruction"))
                        if q.get("source"):
                            st.caption(f"📚 Nga banka: {q.get('source_text') or q['source']}"
                                       + (f" · Ndryshimi: {q['change']}" if q.get("change") else ""))
                        else:
                            st.caption("✨ Pyetje e re" + (f" · {q['change']}" if q.get("change") else ""))
                    p_col, r_col, a_col = st.columns([5, 1.4, 1.4], vertical_alignment="bottom")
                    prompt = p_col.text_input("Udhëzim për Agjentin AI", key=k("prompt"), label_visibility="collapsed",
                                              placeholder="Udhëzim për Agjentin AI, p.sh. bëje me shkallë 1–5")
                    if r_col.button("Rigjenero", icon=":material/refresh:", key=k("regen"), width="stretch"):
                        action = ("regenerate", i, prompt)
                    if a_col.button("Shto pas saj", icon=":material/add:", key=k("add"), width="stretch",
                                    disabled=not prompt.strip(), help="Shkruani te udhëzimi çfarë pyetje duhet shtuar"):
                        action = ("add", i, prompt)
                with tools:
                    b_up, b_down, b_del = st.columns(3, gap="small")
                    if b_up.button("", icon=":material/arrow_upward:", key=k("up"), disabled=i == 0, help="Lart"):
                        action = ("up", i)
                    if b_down.button("", icon=":material/arrow_downward:", key=k("down"), disabled=i == len(qs) - 1,
                                     help="Poshtë"):
                        action = ("down", i)
                    if b_del.button("", icon=":material/delete:", key=k("del"), help="Fshije pyetjen"):
                        action = ("delete", i)

        if action:
            kind, i = action[0], action[1]
            changed = True
            if kind == "up":
                qs[i - 1], qs[i] = qs[i], qs[i - 1]
            elif kind == "down":
                qs[i + 1], qs[i] = qs[i], qs[i + 1]
            elif kind == "delete":
                qs.pop(i)
            else:
                changed = regenerate(i, action[2], new=kind == "add")
            if changed:
                st.rerun()

    # -----------------------------------------------------------------------
    # 4. Word
    # -----------------------------------------------------------------------

    with ubo_ui.card("Shkarko pyetësorin", step=4):
        numbering = st.radio("Numërimi i pyetjeve", list(qd.NUMBERING), format_func=qd.NUMBERING.get,
                             horizontal=True, key="qd_numbering",
                             help=f"Automatik: 1, 2, 3; kur pyetësori ka më shumë se {qd.LONG_QUESTIONNAIRE} pyetje "
                                  "dhe disa seksione, seksionet marrin shkronja (Seksioni A, B …) dhe pyetjet A1, A2 …")
        with_sources = st.checkbox("Shto burimin e çdo pyetjeje (për rishikimin e brendshëm)")
        language = ss["qd_brief"]["language"]
        name = re.sub(r"[^\w\- ]+", "", questionnaire["title"] or "pyetesori").strip().replace(" ", "_") or "pyetesori"
        st.download_button("Shkarko në Word", qd.to_docx(questionnaire, language, with_sources, numbering),
                           file_name=f"{name}.docx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.caption(f"Kostoja e Agjentit AI në këtë sesion: ${ss.get('qd_cost', 0.0):.3f}")
