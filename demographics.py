"""The standard demographic questions of UBO Consulting (as in the Panel), for "Dizajnimi i pyetësorit".

They are not written by the AI agent: the user ticks the ones the study needs, and the program adds them,
word for word, at the end of the questionnaire. Those the user puts at the start (to screen the target group,
e.g. the employment status for a study among private-sector employees) go to the start instead. Questions that
depend on another one carry the condition (the village only for rural residents) or are asked once per child.

`verified`: True when the wording and the options come from the bank's "Standard of Demographics"
questionnaire; False when the options were proposed (the bank has no standard version of them) and should be
checked against the Panel. To change a question, change it here.
"""

COMMON_OPTIONS = {
    "yes_no": [("1", "Po"), ("2", "Jo")],
    "income": [("1", "Nuk kam të ardhura"), ("2", "Deri 264 euro"), ("3", "264-450 euro"), ("4", "451-600 euro"),
               ("5", "601-750 euro"), ("6", "751-900 euro"), ("7", "901-1200 euro"), ("8", "1201-1500 euro"),
               ("9", "1501-2000 euro"), ("10", "2001 – 2500 euro"), ("11", "2501 – 3000 euro")],
}

MUNICIPALITIES = [
    "Prishtinë", "Prizren", "Ferizaj", "Pejë", "Gjakovë", "Gjilan", "Mitrovicë Jugore", "Mitrovicë Veriore",
    "Podujevë", "Vushtrri", "Suharekë", "Rahovec", "Drenas", "Lipjan", "Malishevë", "Kamenicë", "Viti", "Deçan",
    "Istog", "Klinë", "Skenderaj", "Dragash", "Fushë Kosovë", "Kaçanik", "Shtime", "Obiliq", "Leposaviq",
    "Graçanicë", "Han i Elezit", "Zveçan", "Shtërpcë", "Novobërdë", "Zubin Potok", "Junik", "Mamushë", "Ranillug",
    "Kllokot", "Partesh",
]
MONTHS = ["Janar", "Shkurt", "Mars", "Prill", "Maj", "Qershor", "Korrik", "Gusht", "Shtator", "Tetor", "Nëntor",
          "Dhjetor"]

# id, label (short, for the checkboxes), text, type (questionnaire_design.QUESTION_TYPES), options, instruction,
# ask_if: (id, [codes]): asked only when that question got one of these answers;
# loop: id: asked once for each child, as many times as the answer to that numeric question;
# ticked by default, verified
STANDARD = [
    dict(id="D1", label="Gjinia", text="Cila është gjinia juaj?", type="single",
         options=[("1", "Femër"), ("2", "Mashkull"), ("3", "Tjetër, ju lutem specifikoni"),
                  ("99", "Refuzoj të përgjigjem")], default=True, verified=True),
    dict(id="D2", label="Mosha", text="Sa vjeç jeni?", type="numeric", default=True, verified=False),
    dict(id="D3", label="Komuna", text="Në cilën komunë jetoni?", type="single",
         options=[(str(i), m) for i, m in enumerate(MUNICIPALITIES, 1)], default=True, verified=True),
    dict(id="D4", label="Vendbanimi", text="Në çfarë vendbanimi jetoni?", type="single",
         options=[("1", "Urban"), ("2", "Rural")], default=True, verified=True),
    dict(id="D5", label="Lagjja (urban)", text="Emri i lagjës.", type="text", ask_if=("D4", ["1"]),
         default=True, verified=True),
    dict(id="D6", label="Fshati (rural)", text="Emri i fshatit?", type="text", ask_if=("D4", ["2"]),
         default=True, verified=True),
    dict(id="D7", label="Anëtarët e familjes", text="Sa anëtarë jeni në familjen tuaj, duke përfshirë edhe veten?",
         type="numeric", default=True, verified=True),
    dict(id="D8", label="Statusi martesor", text="Cili është statusi juaj martesor?", type="single",
         options=[("1", "Beqar/e"), ("2", "I/E martuar"), ("3", "I/E divorcuar"), ("4", "I/E ndarë"),
                  ("5", "I/E ve"), ("6", "Bashkëjetesë")], default=True, verified=True),
    dict(id="D9", label="Fëmijë nën 18 vjeç", text="A keni fëmijë nën moshën 18 vjeçare?", type="single",
         options=COMMON_OPTIONS["yes_no"], verified=False),
    dict(id="D10", label="Numri i fëmijëve", text="Sa fëmijë i keni?", type="numeric", ask_if=("D9", ["1"]),
         verified=False),
    dict(id="D11", label="Mosha e fëmijës", text="Mosha e fëmijës.", type="numeric", loop="D10",
         verified=False),
    dict(id="D12", label="Gjinia e fëmijës", text="Gjinia e fëmijës.", type="single",
         options=[("1", "Femër"), ("2", "Mashkull")], loop="D10", verified=False),
    dict(id="D13", label="Lloji i banimit", text="Në çfarë lloj banimi jetoni?", type="single",
         options=[("1", "Shtëpi private"), ("2", "Banesë (apartament) në pronësi"), ("3", "Banesë me qira"),
                  ("4", "Te prindërit / familja"), ("5", "Tjetër, ju lutem specifikoni")], verified=False),
    dict(id="D14", label="Grupi etnik", text="Cilit grup etnik i takoni?", type="single",
         options=[("1", "Shqiptar"), ("2", "Serb"), ("3", "Boshnjak"), ("4", "Goran"), ("5", "Turk"), ("6", "Rom"),
                  ("7", "Ashkali"), ("8", "Egjiptas"), ("9", "Tjetër, ju lutem specifikoni")], verified=True),
    dict(id="D15", label="Shkollimi", text="Cili është niveli më i lartë i shkollimit që keni përfunduar?",
         type="single",
         options=[("1", "Pa edukim/shkollim të përfunduar"), ("2", "Shkollë fillore (1 - 5 vjet)"),
                  ("3", "Shkollë e mesme e ulët (6 - 9 vjet)"), ("4", "Shkollë e mesme e lartë (10 - 12 vjet)"),
                  ("5", "Shkollë profesionale (10 - 12 vjet)"), ("6", "Edukim universitar (Bachelor, Master)"),
                  ("7", "Edukim post-universitar (Doktorraturë)")], default=True, verified=True),
    dict(id="D16", label="Dega e studimeve", text="Në cilën degë/drejtim jeni diplomuar?", type="text",
         ask_if=("D15", ["6", "7"]), verified=False),
    dict(id="D17", label="Profesioni", text="Cili është profesioni juaj?", type="text", verified=False),
    dict(id="D18", label="Punësimi", text="Cili është statusi juaj i punësimit?", type="single",
         options=[("1", "I/E papunësuar – duke kërkuar punë"), ("2", "I/E papunësuar – duke mos kërkuar punë"),
                  ("3", "I/E vetëpunësuar (Pronar/e biznesi)"), ("4", "I/E punësuar në sektorin publik"),
                  ("5", "I/E punësuar në sektorin privat"), ("6", "I/E punësuar kohë pas kohe"),
                  ("7", "Pensionist/e"),
                  ("8", "Punë e papaguar (kujdesi për shtëpinë, më të moshuarit ose fëmijët)"),
                  ("9", "Praktikant/e"), ("10", "Student/e"), ("11", "Tjetër, ju lutem specifikoni")],
         default=True, verified=True),
    dict(id="D19", label="Universiteti", text="A studioni në universitet publik apo privat?",
         type="single", options=[("1", "Universitet publik"), ("2", "Universitet privat")],
         ask_if=("D18", ["10"]), verified=False),
    dict(id="D20", label="Sektori i punës",
         text="Cila është kategoria ose veprimtaria kryesore e organizatës/kompanisë suaj?", type="single",
         options=[("1", "Administratë publike"), ("2", "Arsim"), ("3", "Shëndetësi"), ("4", "Tregti"),
                  ("5", "Prodhim / industri"), ("6", "Ndërtim"), ("7", "Bujqësi"),
                  ("8", "Banka, financa dhe sigurime"), ("9", "Telekomunikacion dhe TIK"),
                  ("10", "Hoteleri dhe gastronomi"), ("11", "Transport dhe logjistikë"),
                  ("12", "OJQ / organizatë ndërkombëtare"), ("13", "Tjetër, ju lutem specifikoni")],
         ask_if=("D18", ["3", "4", "5", "6"]), verified=False),
    dict(id="D21", label="Të ardhurat personale", text="Sa janë të ardhurat tuaja mujore personale (Neto)?",
         type="single", options=COMMON_OPTIONS["income"] + [("12", "3000+ euro"), ("99", "Refuzoj të përgjigjem")],
         default=True, verified=True),
    dict(id="D22", label="Të ardhurat familjare", text="Sa janë të ardhurat mujore të familjes tuaj (Neto)?",
         type="single", options=COMMON_OPTIONS["income"] + [("12", "3001 - 4000 euro"), ("13", "4001 – 5000 euro"),
                                                           ("14", "5000+"), ("99", "Refuzoj të përgjigjem")],
         default=True, verified=True),
    dict(id="D23", label="Përkatësia fetare", text="Cila është përkatësia juaj fetare?", type="single",
         options=[("1", "Musliman"), ("2", "Katolik"), ("3", "Ortodoks"), ("4", "Ateist"), ("5", "Agnostik"),
                  ("6", "Tjetër, ju lutem specifikoni"), ("99", "Refuzoj të përgjigjem")], verified=True),
    dict(id="D24", label="Familjarë jashtë", text="A keni anëtarë të familjes në diasporë?", type="single",
         options=COMMON_OPTIONS["yes_no"], verified=False),
    dict(id="D25", label="Remitencat", text="A merrni para nga familjarët në diasporë?", type="single",
         options=COMMON_OPTIONS["yes_no"], ask_if=("D24", ["1"]), verified=False),
    dict(id="D26", label="Shuma në vit", text="Sa para në vit i merrni nga ta (në Euro)?",
         type="numeric", ask_if=("D25", ["1"]), verified=False),
    dict(id="D27", label="Viti i lindjes", text="Viti i lindjes.", type="numeric", verified=False),
    dict(id="D28", label="Muaji i lindjes", text="Muaji i lindjes.", type="single",
         options=[(str(i), m) for i, m in enumerate(MONTHS, 1)], verified=False),
    dict(id="D29", label="Dita e lindjes", text="Dita e lindjes.", type="numeric", verified=False),
]
for _q in STANDARD:
    _q.setdefault("options", [])
    _q.setdefault("instruction", "")
    _q.setdefault("ask_if", None)
    _q.setdefault("loop", None)
    _q.setdefault("default", False)
    # the question it depends on: ticking this one also adds that one
    _q["requires"] = _q["ask_if"][0] if _q["ask_if"] else _q["loop"]

BY_ID = {q["id"]: q for q in STANDARD}

# The groups on the page (one column each); a question that depends on another one comes right under it
GROUPS = {
    "Personale": ["D1", "D2", "D27", "D28", "D29", "D8", "D14", "D23"],
    "Vendbanimi": ["D3", "D4", "D5", "D6", "D13"],
    "Familja": ["D7", "D9", "D10", "D11", "D12", "D24", "D25", "D26"],
    "Arsimi dhe puna": ["D15", "D16", "D17", "D18", "D19", "D20", "D21", "D22"],
}
assert sorted(sum(GROUPS.values(), [])) == sorted(BY_ID), "every standard question in one group"
DEFAULT_IDS = [q["id"] for q in STANDARD if q["default"]]

# The section title of the demographics, and of the screening questions moved to the start
SECTION = {"Shqip": "Të dhënat demografike", "English": "Demographics", "Srpski": "Demografski podaci"}
SCREENING_SECTION = {"Shqip": "Pyetjet kualifikuese", "English": "Screening", "Srpski": "Pitanja za kvalifikaciju"}

# The filters of the questions that depend on another one, in the questionnaire's language ({code}: that question)
ROUTING = {
    "Shqip": {"ask_if": "Pyet nëse {code} = {answers}", "or": " ose ",
              "loop": "Pyet për çdo fëmijë, aq herë sa përgjigjja te {code} (fëmija 1, fëmija 2 …)"},
    "English": {"ask_if": "Ask if {code} = {answers}", "or": " or ",
                "loop": "Ask for each child, as many times as the answer to {code} (child 1, child 2 …)"},
    "Srpski": {"ask_if": "Pitati ako {code} = {answers}", "or": " ili ",
               "loop": "Pitati za svako dete, onoliko puta koliko je odgovor na {code} (dete 1, dete 2 …)"},
}


def with_required(ids):
    """The ids with the questions they depend on (D11 needs D10 and D9), in the standard order."""
    wanted = set()
    for qid in ids:
        while qid and qid not in wanted:
            wanted.add(qid)
            qid = BY_ID[qid]["requires"]
    return [q["id"] for q in STANDARD if q["id"] in wanted]


def condition_text(q):
    """The condition of a standard question in words, in Albanian: "Pyetet nëse Vendbanimi = Rural"."""
    if q.get("ask_if"):
        parent = BY_ID[q["ask_if"][0]]
        labels = dict(parent["options"])
        return f"Pyetet nëse {parent['label']} = " + " ose ".join(labels.get(c, c) for c in q["ask_if"][1])
    if q.get("loop"):
        return f"Pyetet për secilin fëmijë, aq herë sa përgjigjja te «{BY_ID[q['loop']]['label']}»"
    return ""


def standard_text():
    """The list for the AI agent: id, type, text, options and condition of every standard question."""
    lines = []
    for q in STANDARD:
        options = "; ".join(f"{c}={label}" for c, label in q["options"])
        notes = " ".join(x for x in (q["instruction"], condition_text(q)) if x)
        lines.append(f"{q['id']} ({q['type']}) {q['text']}" + (f" — {options}" if options else "")
                     + (f" [{notes}]" if notes else ""))
    return "\n".join(lines)
