"""Question bank: the questionnaires of past studies, taken from the KoboToolbox accounts.

- fetch_forms(): every questionnaire (survey and template) of one account, with its questions and options,
- dedupe(): merges copies, the same questionnaire in two accounts or a copy with small changes (a new wave),
- categorize(): Claude sets the industry, the kind of study, the methodologies and the topics,
- the bank is kept in a private Google Sheet ("Pyetesoret" and "Pyetjet"), never in the repository (it is
  public). Categories can be corrected by hand in the sheet: a sync keeps them.

Secrets:
    QUESTION_BANK_SHEET = "<the Google Sheet's URL>"   # shared (Editor) with the gcp_service_account e-mail
    [kobo_accounts.ubo]                                # one section per account; the name is free
    server = "https://kf.kobotoolbox.org"              # or https://eu.kobotoolbox.org
    token = "<API token: Account Settings -> Security>"

Runs from the "Banka e pyetjeve" page, or from a terminal: py -m question_bank
"""

import datetime
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

import questionnaire_ai as qai

ASSET_TYPES = ("survey", "template")
GROUP_START = {"begin_group", "begin group", "begin_repeat", "begin repeat", "begin_kobomatrix"}
GROUP_END = {"end_group", "end group", "end_repeat", "end repeat", "end_kobomatrix"}
# Rows of a form that are not questions: metadata, calculations, hidden values
SKIP_TYPES = {"start", "end", "today", "deviceid", "username", "phonenumber", "simserial", "subscriberid", "imei",
              "audit", "calculate", "hidden", "background-audio", "xml-external", "start-geopoint"}
MIN_QUESTIONS = 3        # forms with fewer questions are tests or empty drafts
SIMILARITY = 0.9         # share of common questions above which two forms are copies of one questionnaire
WORKERS = 8              # parallel requests to Kobo
CATEGORIZE_WORKERS = 4   # parallel requests to Claude
CATEGORY_EFFORT = "low"  # reading a form and picking categories needs little reasoning
CHECKPOINT = 20          # categorized forms between two saves of the catalog
CELL_LIMIT = 49000       # Google Sheets holds at most 50 000 characters in a cell

# Fieldwork questions about the interviewer (name, code, ID; the options are the team's names): not study
# questions, so they stay out of the bank. "ANKETUESI: SHËNO GJININË" is a real question and stays.
STAFF_RE = re.compile(r"anketues|enumerator|interviewer|intervistues|anketar|mbikëqyr|mbikqyr|supervisor|kontrollor",
                      re.IGNORECASE)
IDENTITY_RE = re.compile(r"\b(emri|emër|emer|mbiemri|name|ime|prezime|kodi|code|šifra|sifra|id|numri)\b",
                         re.IGNORECASE)

INDUSTRIES = [
    "Bankat dhe financat", "Sigurimet", "Telekomunikacioni", "Retail dhe FMCG", "Energjia dhe shërbimet publike",
    "Shëndetësia", "Arsimi", "Administrata publike", "Politika dhe zgjedhjet", "Media dhe komunikimi",
    "Turizmi dhe hotelieria", "Transporti", "Ndërtimi dhe patundshmëritë", "Teknologjia dhe shërbimet digjitale",
    "Bujqësia", "Zhvillimi ndërkombëtar dhe donatorët", "Shoqëria civile dhe OJQ", "Tregu i punës dhe biznesi",
    "Siguria dhe drejtësia", "Tjetër",
]
SURVEY_TYPES = [
    "Kënaqshmëria e klientëve", "Sondazh publik / opinion", "Testim përdorshmërie (usability)",
    "Vlerësim ndikimi (impact assessment)", "Studim bazë (baseline / endline)", "Studim tregu (usage & attitude)",
    "Shëndeti i markës (brand health)", "Testim produkti / koncepti", "Testim reklame / komunikimi",
    "Mystery shopping", "Kënaqshmëria e punonjësve", "Vlerësim i nevojave", "Monitorim dhe vlerësim (M&E)",
    "Sondazh zgjedhor / exit poll", "Anketë e bizneseve", "Regjistrim / listim", "Tjetër",
]
METHODS = [
    "MaxDiff", "Segmentim", "Analizë çmimi (Van Westendorp)", "Analizë çmimi (Gabor-Granger)", "Conjoint",
    "NPS", "CSAT / CES", "TURF", "Brand funnel", "Driver analysis", "Studim i përsëritur / trend",
]

CATALOG_SHEET, QUESTIONS_SHEET, INFO_SHEET = "Pyetesoret", "Pyetjet", "Info"
INFO_COLUMNS = ["kyci", "vlera"]
CATALOG_COLUMNS = ["id", "titulli", "perdorshem", "industria", "lloji", "metodologjite", "temat", "klienti",
                   "pershkrimi", "nr_pyetjeve", "gjuhet", "perditesuar", "dergesat", "kopjet", "burimet",
                   "kategorizuar"]
CATEGORY_COLUMNS = ["perdorshem", "industria", "lloji", "metodologjite", "temat", "klienti", "pershkrimi",
                    "kategorizuar"]
QUESTION_COLUMNS = ["pyetesori_id", "nr", "emri", "lloji", "grupi", "teksti", "opsionet", "udhezimi",
                    "perkthimet"]


class BankError(Exception):
    """Error with a message that can be shown to the user as-is."""


# ---------------------------------------------------------------------------
# KoboToolbox
# ---------------------------------------------------------------------------

def kobo_accounts(secrets):
    """{name: (server, token)} from the [kobo_accounts.<name>] sections of the secrets."""
    try:
        sections = secrets["kobo_accounts"]
    except Exception:
        raise BankError("Mungojnë llogaritë e KoboToolbox te secrets (`[kobo_accounts.<emri>]` me `server` dhe `token`).")
    accounts = {}
    for name, acc in dict(sections).items():
        server = str(acc.get("server") or "https://kf.kobotoolbox.org").rstrip("/")
        if not acc.get("token"):
            raise BankError(f"Llogarisë `{name}` i mungon `token` te secrets.")
        accounts[name] = (server, acc["token"])
    if not accounts:        # an empty sync would wipe the bank
        raise BankError("Nuk ka asnjë llogari KoboToolbox te secrets (`[kobo_accounts.<emri>]`).")
    return accounts


def _get(url, token, params=None):
    for attempt in range(3):
        try:
            r = requests.get(url, headers={"Authorization": f"Token {token}"}, params=params, timeout=60)
        except requests.RequestException:
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code in (401, 403):
            raise BankError("Tokeni nuk pranohet nga KoboToolbox (i pavlefshëm ose pa leje).")
        if r.status_code in (429, 500, 502, 503, 504) and attempt < 2:
            time.sleep(2 * (attempt + 1))
            continue
        r.raise_for_status()
        return r.json()


def list_assets(server, token):
    """Questionnaires and templates of the account (without their content)."""
    url, params, assets = f"{server}/api/v2/assets/", {"format": "json", "limit": 200}, []
    while url:
        page = _get(url, token, params)
        assets += [a for a in page.get("results", []) if a.get("asset_type") in ASSET_TYPES]
        url, params = page.get("next"), None
    return assets


def _clean(text):
    # labels may carry HTML: line breaks become spaces, inline tags (<span style=…>, <b>) disappear
    text = re.sub(r"<[^>]+>", "", re.sub(r"<br\s*/?>|</?(p|div)\b[^>]*>", " ", str(text), flags=re.I))
    return re.sub(r"\s+", " ", text.replace("**", "")).strip()


def _texts(value, languages):
    """{language: text} of a label / hint, which Kobo gives as one entry per translation (or as a string)."""
    if isinstance(value, list):
        return {lang: _clean(v) for lang, v in zip(languages, value) if v and _clean(v)}
    if isinstance(value, str) and _clean(value):
        return {languages[0]: _clean(value)}
    return {}


def _staff_question(label, list_name, name):
    """True for a question asking who the interviewer is (see STAFF_RE)."""
    if list_name and STAFF_RE.search(str(list_name)):
        return True
    texts = list(label.values()) + [str(name or "").replace("_", " ")]
    return any(STAFF_RE.search(t) and IDENTITY_RE.search(t) for t in texts)


def parse_content(content):
    """(languages, questions) of an asset's content: groups, metadata and calculations are left out."""
    languages = [l or "default" for l in (content.get("translations") or [None])]
    lists = {}
    for c in content.get("choices") or []:
        lists.setdefault(c.get("list_name"), []).append(
            {"code": str(c.get("name") or c.get("$autovalue") or ""), "label": _texts(c.get("label"), languages)})
    questions, groups = [], []
    for row in content.get("survey") or []:
        rtype = str(row.get("type") or "").strip()
        if rtype in GROUP_START:
            label = _texts(row.get("label"), languages)
            groups.append(next(iter(label.values()), row.get("name") or ""))
            continue
        if rtype in GROUP_END:
            if groups:
                groups.pop()
            continue
        base, _, list_name = rtype.partition(" ")
        label = _texts(row.get("label"), languages)
        if base in SKIP_TYPES or not label:
            continue
        list_name = (row.get("select_from_list_name") or row.get("kobo--rank-items")
                     or row.get("kobo--score-choices") or list_name.strip())
        if _staff_question(label, list_name, row.get("name")):
            continue
        questions.append({
            "name": str(row.get("name") or row.get("$autoname") or ""),
            "type": base,
            "group": " / ".join(g for g in groups if g),
            "label": label,
            "hint": _texts(row.get("hint"), languages),
            "options": lists.get(list_name, []) if list_name else [],
        })
    return languages, questions


def fetch_forms(account, server, token, progress=None):
    """Every questionnaire of one account as a dict with its questions. Returns (forms, failed asset uids)."""
    assets = list_assets(server, token)
    forms, failed = [], []

    def one(asset):
        detail = _get(f"{server}/api/v2/assets/{asset['uid']}/", token, {"format": "json"})
        languages, questions = parse_content(detail.get("content") or {})
        return {
            "id": f"{account}:{asset['uid']}",
            "title": _clean(asset.get("name") or "") or asset["uid"],
            "date_modified": (asset.get("date_modified") or "")[:10],
            "submissions": int(asset.get("deployment__submission_count") or 0),
            "languages": languages,
            "questions": questions,
        }

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(one, a): a for a in assets}
        for done, future in enumerate(as_completed(futures), 1):
            try:
                forms.append(future.result())
            except BankError:
                raise
            except Exception:
                failed.append(f"{account}:{futures[future]['uid']}")
            if progress:
                progress(done / max(len(assets), 1), f"{account}: {done}/{len(assets)} pyetësorë")
    return forms, failed


# ---------------------------------------------------------------------------
# Duplicates
# ---------------------------------------------------------------------------

def _main_text(texts):
    return next(iter(texts.values()), "")


def _signature(form):
    """The questions of a form as normalized texts (all languages), to compare forms."""
    sig = set()
    for q in form["questions"]:
        for text in q["label"].values():
            norm = re.sub(r"\W+", " ", text.lower()).strip()
            if norm:
                sig.add(norm)
    return sig


def dedupe(forms):
    """One form per questionnaire: copies in other accounts and near-identical versions are merged into the
    most recently changed one, whose "sources" lists the ids of all copies."""
    forms = sorted((f for f in forms if len(f["questions"]) >= MIN_QUESTIONS),
                   key=lambda f: (f["date_modified"], f["submissions"]), reverse=True)
    kept = []
    for form in forms:
        sig = _signature(form)
        for k in kept:
            union = len(sig | k["sig"])
            if union and len(sig & k["sig"]) / union >= SIMILARITY:
                k["form"]["sources"].append(form["id"])
                k["form"]["submissions"] += form["submissions"]
                break
        else:
            kept.append({"form": {**form, "sources": [form["id"]]}, "sig": sig})
    return [k["form"] for k in kept]


# ---------------------------------------------------------------------------
# Categories (Claude)
# ---------------------------------------------------------------------------

CATEGORY_SYSTEM = f"""You catalogue the questionnaires of UBO Consulting, a market and social research company in Kosovo, for a question bank from which new questionnaires are built.

For the questionnaire you receive, return:
- usable: false only for forms that are not real study questionnaires: tests, demos, trainings, empty or abandoned drafts, copies made to try something out. Real questionnaires of any size are usable.
- industry: the sector the study is about (for a citizens' poll on public services: "Administrata publike").
- survey_type: the kind of study.
- methods: the special methodologies the questionnaire contains (a MaxDiff exercise, price questions of Van Westendorp or Gabor-Granger, a segmentation battery, NPS …); an empty list when there are none.
- topics: 3 to 8 short topics in Albanian (e.g. "kënaqshmëria me shërbimin", "kanalet digjitale", "demografia").
- client: the client or organisation named in the title or text, "" when not clear.
- summary: one or two sentences in Albanian on what the questionnaire measures and among whom.

Industries: {", ".join(INDUSTRIES)}.
Kinds of study: {", ".join(SURVEY_TYPES)}."""

CATEGORY_SCHEMA = {
    "type": "object",
    "properties": {
        "usable": {"type": "boolean"},
        "industry": {"type": "string", "enum": INDUSTRIES},
        "survey_type": {"type": "string", "enum": SURVEY_TYPES},
        "methods": {"type": "array", "items": {"type": "string", "enum": METHODS}},
        "topics": {"type": "array", "items": {"type": "string"}},
        "client": {"type": "string"},
        "summary": {"type": "string"},
    },
    "required": ["usable", "industry", "survey_type", "methods", "topics", "client", "summary"],
    "additionalProperties": False,
}


def form_text(form, max_questions=250, max_options=12):
    """The form as text for the categorization. Long forms are cut: their first 250 questions are enough
    to tell what the study is about."""
    lines = [f"Title: {form['title']}", f"Languages: {', '.join(form['languages'])}",
             f"Submissions: {form['submissions']}", "Questions:"]
    for q in form["questions"][:max_questions]:
        options = "; ".join(_main_text(o["label"]) for o in q["options"][:max_options])
        group = f"[{q['group']}] " if q["group"] else ""
        lines.append(f"- {group}({q['type']}) {_main_text(q['label'])[:300]}" + (f" | {options}" if options else ""))
    if len(form["questions"]) > max_questions:
        lines.append(f"… and {len(form['questions']) - max_questions} more questions")
    return "\n".join(lines)


def categorize(client, model, form):
    """The category columns of one form, and the cost."""
    data, usage = qai._call_claude(client, CATEGORY_SYSTEM, form_text(form), CATEGORY_SCHEMA, max_tokens=4000,
                                   effort=CATEGORY_EFFORT, model=model)
    row = {
        "perdorshem": "po" if data.get("usable") else "jo",
        "industria": data.get("industry", ""),
        "lloji": data.get("survey_type", ""),
        "metodologjite": "; ".join(data.get("methods") or []),
        "temat": "; ".join(data.get("topics") or []),
        "klienti": data.get("client", ""),
        "pershkrimi": data.get("summary", ""),
        "kategorizuar": f"{datetime.date.today()} {model}",
    }
    return row, qai.estimate_cost(usage, model)


# ---------------------------------------------------------------------------
# Google Sheet
# ---------------------------------------------------------------------------

def open_sheet(secrets):
    import gspread
    from google.oauth2.service_account import Credentials

    try:
        ref = str(secrets["QUESTION_BANK_SHEET"]).strip()
        info = dict(secrets["gcp_service_account"])
    except Exception:
        raise BankError("Mungon `QUESTION_BANK_SHEET` (lidhja e Google Sheet-it të bankës) ose "
                        "`gcp_service_account` te secrets.")
    credentials = Credentials.from_service_account_info(info, scopes=[
        "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    gc = gspread.authorize(credentials)
    try:
        return gc.open_by_url(ref) if ref.startswith("http") else gc.open_by_key(ref)
    except gspread.SpreadsheetNotFound:
        raise BankError(f"Google Sheet-i i bankës nuk u gjet. Ndajeni (Editor) me {info.get('client_email')}.")


def _worksheet(sh, name, columns):
    import gspread
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(name, rows=2, cols=len(columns))


def _read(sh, name, columns, upto=None):
    """Rows of a worksheet as dicts; with upto, only the columns up to that one (header name)."""
    ws = _worksheet(sh, name, columns)
    if upto:
        import gspread.utils
        values = ws.get(f"A:{gspread.utils.rowcol_to_a1(1, columns.index(upto) + 1)[:-1]}")
    else:
        values = ws.get_all_values()
    if not values:
        return []
    header = values[0]
    return [dict(zip(header, row)) for row in values[1:]]


def _write(sh, name, columns, rows, chunk=5000):
    ws = _worksheet(sh, name, columns)
    values = [columns] + [[str(r.get(c, ""))[:CELL_LIMIT] for c in columns] for r in rows]
    ws.clear()
    ws.resize(rows=max(len(values), 2), cols=len(columns))
    for start in range(0, len(values), chunk):
        ws.update(values[start:start + chunk], f"A{start + 1}", value_input_option="RAW")


def _question_rows(form):
    rows = []
    for nr, q in enumerate(form["questions"], 1):
        rows.append({
            "pyetesori_id": form["id"], "nr": nr, "emri": q["name"], "lloji": q["type"], "grupi": q["group"],
            "teksti": _main_text(q["label"]),
            "opsionet": json.dumps([[o["code"], _main_text(o["label"])] for o in q["options"]], ensure_ascii=False)
            if q["options"] else "",
            "udhezimi": _main_text(q["hint"]),
            "perkthimet": json.dumps({"label": q["label"], "hint": q["hint"],
                                      "options": [o["label"] for o in q["options"]]}, ensure_ascii=False)
            if len(form["languages"]) > 1 else "",
        })
    return rows


def _cached_loader():
    import streamlit as st

    # ubo_ui's page script hides its "Ju lutem prisni pak" window while this spinner shows (matched on its text)
    @st.cache_data(ttl=600, show_spinner="Po lexohet banka e pyetjeve...")
    def cached_bank():
        return load_bank(st.secrets)

    @st.cache_data(ttl=600, show_spinner=False)
    def cached_last_sync():
        return last_sync(st.secrets)
    return cached_bank, cached_last_sync


# load_bank / last_sync for the pages, kept for 10 minutes and shared by them (.clear() after a sync)
cached_bank, cached_last_sync = _cached_loader()


def last_sync(secrets):
    """The time of the last finished sync (UTC, "YYYY-MM-DD HH:MM") from the Info sheet, "" when unknown."""
    try:
        rows = _read(open_sheet(secrets), INFO_SHEET, INFO_COLUMNS)
    except Exception:
        return ""
    return next((r["vlera"] for r in rows if r.get("kyci") == "perditesimi_i_fundit"), "")


def _save_last_sync(sh):
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M")
    _write(sh, INFO_SHEET, INFO_COLUMNS, [{"kyci": "perditesimi_i_fundit", "vlera": now}])


def load_bank(secrets):
    """(catalog rows, {questionnaire id: [question dicts]}) from the sheet. Question dicts carry "nr", "name",
    "type", "group", "text", "options" ([[code, label]]) and "hint"."""
    sh = open_sheet(secrets)
    catalog = _read(sh, CATALOG_SHEET, CATALOG_COLUMNS)
    questions = {}
    # without "perkthimet" (the other languages), the heaviest column, which the pages do not use
    for r in _read(sh, QUESTIONS_SHEET, QUESTION_COLUMNS, upto="udhezimi"):
        try:
            options = json.loads(r.get("opsionet") or "[]")
        except json.JSONDecodeError:
            options = []
        questions.setdefault(r["pyetesori_id"], []).append({
            "nr": r.get("nr", ""), "name": r.get("emri", ""), "type": r.get("lloji", ""),
            "group": r.get("grupi", ""), "text": r.get("teksti", ""), "options": options,
            "hint": r.get("udhezimi", ""),
        })
    return catalog, questions


# ---------------------------------------------------------------------------
# Sync
# ---------------------------------------------------------------------------

def sync(secrets, client=None, model=None, recategorize=False, progress=None):
    """Brings the bank up to date with the Kobo accounts and categorizes the questionnaires that have no
    categories yet (all of them with recategorize). Categories already in the sheet, also those corrected by
    hand, are kept. Without a Claude client only the questionnaires are updated. Returns a summary dict."""
    progress = progress or (lambda frac, text: None)
    stats = {"accounts": {}, "failed": [], "cost": 0.0, "categorized": 0, "errors": []}
    forms = []
    for name, (server, token) in kobo_accounts(secrets).items():
        try:
            account_forms, failed = fetch_forms(name, server, token, progress)
        except BankError as e:
            raise BankError(f"Llogaria `{name}`: {e}")
        except requests.RequestException as e:
            raise BankError(f"Llogaria `{name}`: KoboToolbox nuk u arrit ({e}).")
        stats["accounts"][name] = len(account_forms)
        stats["failed"] += failed
        forms += account_forms
    stats["forms"] = len(forms)

    progress(1.0, "Po krahasohen pyetësorët për dyfishime...")
    unique = dedupe(forms)
    stats["unique"] = len(unique)

    sh = open_sheet(secrets)
    old_catalog = _read(sh, CATALOG_SHEET, CATALOG_COLUMNS)
    by_source = {}
    for row in old_catalog:
        for source in (row.get("burimet") or row.get("id", "")).split("; "):
            by_source.setdefault(source, row)

    catalog, question_rows = [], []
    for form in unique:
        old = next((by_source[s] for s in form["sources"] if s in by_source), {})
        row = {
            "id": form["id"], "titulli": form["title"], "nr_pyetjeve": len(form["questions"]),
            "gjuhet": ", ".join(form["languages"]), "perditesuar": form["date_modified"],
            "dergesat": form["submissions"], "kopjet": len(form["sources"]) - 1,
            "burimet": "; ".join(form["sources"]),
            **{c: ("" if recategorize else old.get(c, "")) for c in CATEGORY_COLUMNS},
        }
        catalog.append(row)
        question_rows += _question_rows(form)

    # Questionnaires Kobo did not return this time (a failed request) stay in the bank as they were
    failed = set(stats["failed"])
    kept = [r for r in old_catalog
            if failed & set((r.get("burimet") or "").split("; ")) and r["id"] not in {c["id"] for c in catalog}]
    if kept:
        kept_ids = {r["id"] for r in kept}
        catalog += kept
        question_rows += [r for r in _read(sh, QUESTIONS_SHEET, QUESTION_COLUMNS) if r["pyetesori_id"] in kept_ids]

    progress(1.0, "Po ruhen pyetjet në Google Sheet...")
    _write(sh, QUESTIONS_SHEET, QUESTION_COLUMNS, question_rows)
    _write(sh, CATALOG_SHEET, CATALOG_COLUMNS, catalog)
    stats["questions"] = len(question_rows)

    todo = [f for f, row in zip(unique, catalog) if not row["kategorizuar"]]
    stats["todo"] = len(todo)
    if client is None or not todo:
        _save_last_sync(sh)
        return stats
    rows_by_id = {r["id"]: r for r in catalog}
    with ThreadPoolExecutor(max_workers=CATEGORIZE_WORKERS) as pool:
        futures = {pool.submit(categorize, client, model, f): f for f in todo}
        for done, future in enumerate(as_completed(futures), 1):
            form = futures[future]
            try:
                categories, cost = future.result()
                rows_by_id[form["id"]].update(categories)
                stats["cost"] += cost
                stats["categorized"] += 1
            except qai.AIError as e:
                stats["errors"].append(f"{form['title']}: {e}")
            progress(done / len(todo), f"Kategorizimi: {done}/{len(todo)} (${stats['cost']:.2f})")
            if done % CHECKPOINT == 0:      # a stopped page loses at most the last few categories
                _write(sh, CATALOG_SHEET, CATALOG_COLUMNS, catalog)
    _write(sh, CATALOG_SHEET, CATALOG_COLUMNS, catalog)
    _save_last_sync(sh)
    return stats


if __name__ == "__main__":
    # py -m question_bank [--recategorize] [--quiet]
    # --quiet is for the monthly GitHub Action: the repo is public, so its log may only hold counts, never
    # questionnaire titles, account names or error texts. Exits with 1 when something failed.
    import os
    import sys
    import tomllib

    quiet = "--quiet" in sys.argv
    if os.environ.get("UBO_SECRETS"):
        secrets = tomllib.loads(os.environ["UBO_SECRETS"])
    else:
        with open(".streamlit/secrets.toml", "rb") as f:
            secrets = tomllib.load(f)
    claude = qai.make_client(secrets["ANTHROPIC_API_KEY"]) if secrets.get("ANTHROPIC_API_KEY") else None
    progress = None if quiet else (lambda frac, text: print(f"\r{text[:100]:<100}", end="", flush=True))
    try:
        result = sync(secrets, claude, secrets.get("CLAUDE_MODEL", qai.MODEL), "--recategorize" in sys.argv,
                      progress=progress)
    except Exception as e:
        print("Përditësimi dështoi:", type(e).__name__ if quiet else e)
        sys.exit(1)
    if quiet:
        print(f"Llogari: {len(result['accounts'])} | pyetësorë nga Kobo: {result['forms']} | unikë: {result['unique']}"
              f" | pyetje: {result['questions']} | të rinj për kategorizim: {result['todo']}"
              f" | u kategorizuan: {result['categorized']} | kosto: ${result['cost']:.2f}"
              f" | dështuan nga Kobo: {len(result['failed'])} | gabime Agjenti AI: {len(result['errors'])}")
    else:
        print("\n", json.dumps(result, ensure_ascii=False, indent=2))
    if result["failed"] or result["errors"]:
        sys.exit(1)
