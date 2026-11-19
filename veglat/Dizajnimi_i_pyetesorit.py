import html
import json
import re

import streamlit as st

import demographics as demo
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


def set_demographics(ids):
    """Ticks exactly these standard demographic questions (the buttons above the checkboxes); with none,
    nothing stays at the start either."""
    for q in demo.STANDARD:
        ss[f"qd_demo_{q['id']}"] = q["id"] in ids
    if not ids:
        ss["qd_demo_start"] = []


def tick_start():
    """The questions put at the start are in the questionnaire too, with those they depend on: ticks them."""
    for qid in demo.with_required(ss.get("qd_demo_start") or []):
        ss[f"qd_demo_{qid}"] = True


# the badge beside a step's title: optional / required
st.markdown("""<style>
.qd-badge { display: inline-block; margin-left: 10px; padding: 2px 10px; border-radius: 999px; vertical-align: 3px;
            font: 600 12px "Inter", system-ui, sans-serif; letter-spacing: .02em; }
.qd-badge.opt { background: #E8ECF3; color: #4B5563; }
.qd-badge.req { background: #FDE8E8; color: #B42318; }
.qd-demo-head { margin: 14px 0 12px !important; padding-bottom: 8px; border-bottom: 2px solid #22C6A033;
                font: 700 12px "Inter", system-ui, sans-serif; letter-spacing: .08em; text-transform: uppercase;
                color: #0B8F76; }
[class*="st-key-qd_demo_group_"] { gap: 10px !important; }
.st-key-qd_demo_tools { margin-top: -66px; margin-bottom: 18px; }
.st-key-qd_demo_tools button { min-height: 34px; padding: 4px 14px; border-radius: 999px; cursor: pointer;
                               border: 1px solid #22C6A0; color: #0B8F76; background: #fff; }
.st-key-qd_demo_tools button p { font-size: 13px; font-weight: 600; }
.st-key-qd_demo_tools button:hover { background: #E8FAF5; border-color: #0B8F76; color: #0B8F76; }
.st-key-qd_demo_tools button[kind="primary"], .st-key-qd_demo_tools button[data-testid="stBaseButton-primary"] {
    background: #0B8F76; border-color: #0B8F76; color: #fff; box-shadow: 0 2px 8px rgba(11, 143, 118, .3); }
.st-key-qd_demo_tools button[data-testid="stBaseButton-primary"] p { color: #fff; }
[class*="st-key-qd_demo_group_"] [data-testid="stCheckbox"] label { width: 100%; padding: 4px 5px !important; }
[class*="st-key-qd_demo_group_"] [data-testid="stCheckbox"] p { font-size: 12.5px; white-space: nowrap; }
[class*="st-key-qd_demo_group_"] [data-testid="stCheckbox"], [class*="st-key-qd_demo_group_"] [class*="st-key-qd_demo_D"] { width: 100% !important; }
.st-key-qd_design_go { margin: 8px 0 18px; }
.st-key-qd_design_go button { min-height: 64px; border-radius: 14px; border: none;
                              background: linear-gradient(90deg, #0B8F76, #22C6A0);
                              box-shadow: 0 6px 18px rgba(34, 198, 160, .35); }
.st-key-qd_design_go button p { font-size: 20px; font-weight: 700; }
.st-key-qd_design_go button [data-testid="stIconMaterial"] { font-size: 26px; }
.st-key-qd_design_go button:hover:not(:disabled) { filter: brightness(1.06); box-shadow: 0 8px 22px rgba(34, 198, 160, .45); }
.st-key-qd_design_go button:disabled { background: linear-gradient(90deg, #0B8F76, #22C6A0); opacity: .65;
                                       box-shadow: none; }
.st-key-qd_design_go button:disabled p, .st-key-qd_design_go button:disabled [data-testid="stIconMaterial"] {
    color: #fff; }
</style>""", unsafe_allow_html=True)
OPTIONAL = '<span class="qd-badge opt">Opsional</span>'
REQUIRED = '<span class="qd-badge req">I detyrueshëm</span>'


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

with ubo_ui.card(f"Studimi {OPTIONAL}", step=1, key="studimi",
                 subtitle="Plotësojeni që Agjenti AI t'i gjejë vetë pyetësorët e ngjashëm në bankë dhe ta njohë "
                          "studimin. Mund ta kaloni: zgjidhni pyetësorët dhe shkruani udhëzimin te hapi 2."):
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
    n_questions = c4.number_input("Numri i përafërt i pyetjeve", 5, 200, 30, step=5, key="qd_n",
                                  help="Pa pyetjet demografike: ato zgjidhen te hapi 3 dhe shtohen veç.")
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
        st.caption("Që Agjenti AI të gjejë pyetësorët e ngjashëm, plotësoni " + ", ".join(missing) + ".")

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

with ubo_ui.card(f"Pyetësorët burim nga banka {REQUIRED}", step=2, key="burimet",
                 subtitle=f"Zgjidhni të paktën një pyetësor (deri në {qd.MAX_SOURCES}): pyetjet e tyre janë "
                          "materiali i pyetësorit të ri."):
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
    if not selected:
        st.caption("Zgjidhni të paktën një pyetësor burim.")
    elif not brief["objectives"]:
        st.caption("Shkruani udhëzimin për Agjentin AI ose plotësoni objektivat te hapi 1.")

# ---------------------------------------------------------------------------
# 3. The standard demographic questions (as in the Panel): added word for word, not counted
# ---------------------------------------------------------------------------

for q in demo.STANDARD:
    ss.setdefault(f"qd_demo_{q['id']}", q["default"])
ss.setdefault("qd_demo_start", [])
ticked = set(demo.with_required([q["id"] for q in demo.STANDARD if ss[f"qd_demo_{q['id']}"]]
                                + ss["qd_demo_start"]))
n_ticked = len(ticked)
with ubo_ui.card(f"Pyetjet demografike {OPTIONAL}", step=3, key="demografike",
                 subtitle=f"{n_ticked} të zgjedhura. Shtohen ashtu siç janë në Panel, në fund të pyetësorit, dhe "
                          "nuk hyjnë te numri i pyetjeve."):
    # top right, beside the title: the one that matches the ticks now is filled (active)
    all_ids = [q["id"] for q in demo.STANDARD]
    with st.container(horizontal=True, horizontal_alignment="right", gap="small", key="qd_demo_tools"):
        for label, ids, icon, tip in (
                ("Standardet", demo.DEFAULT_IDS, ":material/star:",
                 "Gjinia, mosha, komuna, vendbanimi (lagjja / fshati), statusi martesor, anëtarët e familjes, "
                 "të ardhurat personale dhe familjare, shkollimi, punësimi"),
                ("Kliko të gjitha", all_ids, ":material/done_all:", "Të gjitha pyetjet demografike"),
                ("Largo të gjitha", [], ":material/remove_done:", "Asnjë pyetje demografike")):
            st.button(label, icon=icon, help=tip, on_click=set_demographics, args=(ids,),
                      type="primary" if ticked == set(ids) else "secondary", key=f"qd_demo_set_{len(ids)}")
    def depth(q):
        """How many questions this one depends on, one under the other (the child's age: 2)."""
        return 1 + depth(demo.BY_ID[q["requires"]]) if q["requires"] else 0

    def demo_checkbox(q, where=st):
        options = "; ".join(label for _, label in q["options"])
        where.checkbox(
            q["label"] + ("" if q["verified"] else " *"),
            key=f"qd_demo_{q['id']}", width="stretch",
            help="  \n".join(x for x in (
                q["text"], f"Opsionet: {options}" if options else "", q["instruction"], demo.condition_text(q),
                "" if q["verified"] else "*Opsionet janë propozuar; konfirmojini me Panelin.*") if x))

    # one column per topic; a question that depends on another one is indented under it
    for col, (group, ids) in zip(st.columns([0.95, 0.95, 1.05, 1.05], gap="large"), demo.GROUPS.items()):
        with col.container(key=f"qd_demo_group_{ids[0]}"):
            st.markdown(f'<p class="qd-demo-head">{group}</p>', unsafe_allow_html=True)
            for qid in ids:
                demo_checkbox(demo.BY_ID[qid])
    st.markdown("<style>" + " ".join(
        f'[class*="st-key-qd_demo_group_"] .st-key-qd_demo_{q["id"]} {{ margin-left: {8 * depth(q)}px; '
        f'width: calc(100% - {8 * depth(q)}px) !important; box-sizing: border-box; padding-left: 4px; '
        f'border-left: 2px solid rgba(34, 198, 160, .45); }}'
        for q in demo.STANDARD if q["requires"]) + "</style>", unsafe_allow_html=True)
    st.caption("\\* Opsionet janë propozuar dhe duhen konfirmuar me Panelin. Pyetjet e lidhura shtojnë vetë "
               "pyetjen nga varen dhe kushtin: lagjja vetëm për urbanët, fshati vetëm për ruralët, mosha dhe gjinia "
               "pyeten për secilin fëmijë (fëmija 1, fëmija 2 …) sipas numrit të fëmijëve.")
    st.multiselect(
        "Në fillim të pyetësorit (pyetje kualifikuese)", [q["id"] for q in demo.STANDARD], key="qd_demo_start",
        format_func=lambda qid: demo.BY_ID[qid]["label"], on_change=tick_start,
        placeholder="Asnjë: të gjitha dalin në fund",
        help="Kur grupi i synuar zgjidhet sipas një pyetjeje demografike (p.sh. vetëm të punësuarit në sektorin "
             "privat), zgjidheni këtu: ajo del në fillim, dhe Agjenti AI shkruan kush vazhdon intervistën.")
    demographics = [q["id"] for q in demo.STANDARD if ss[f"qd_demo_{q['id']}"]]
    brief.update(demographics=demo.with_required(demographics + ss["qd_demo_start"]),
                 demographics_start=list(ss["qd_demo_start"]))

# The design: a big button under the steps
ready = bool(selected and brief["objectives"])
with st.container(key="qd_design_go"):
    clicked = st.button("Dizajno pyetësorin", type="primary", icon=":material/auto_awesome:", width="stretch",
                        disabled=not ready)
    if not ready:
        st.caption("Për të dizajnuar pyetësorin: zgjidhni pyetësorët burim dhe shkruani udhëzimin te hapi 2 "
                   "(ose objektivat te hapi 1).")
if clicked:
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
    n_demo = sum(1 for q in qs if q.get("demographic"))
    from_bank = sum(1 for q in qs if q.get("source"))
    with ubo_ui.card("Pyetësori", step=4,
                     subtitle=f"{len(qs) - n_demo} pyetje · {from_bank} nga banka · {len(qs) - n_demo - from_bank} "
                              f"të reja · {n_demo} demografike. "
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
                        if q.get("demographic"):
                            st.caption("👤 Pyetje demografike standarde (Paneli)")
                        elif q.get("source"):
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

    with ubo_ui.card("Shkarko pyetësorin", step=5):
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
