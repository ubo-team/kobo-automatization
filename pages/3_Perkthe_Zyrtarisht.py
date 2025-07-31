import pandas as pd
import re
import streamlit as st
from docx import Document
from collections import defaultdict
from difflib import get_close_matches
import os

QUESTION_PATTERN = re.compile(
    r"\[ *(single|multiple|open|numeric|matrix(?: [a-z0-9]*)*|multiple matrix|scale) *\]",
    re.IGNORECASE,
)

def clean_question_text(text):
    return QUESTION_PATTERN.sub("", text).strip()

def get_numbering(para):
    for run in para.runs:
        if run.text.strip():
            return run.text.strip()
    return ""

def is_list_option(para):
    return para.style and para.style.name.lower().startswith("list")

def extract_matrix_table(table, data, parent_qid, parent_qtext):
    rows = table.rows
    if not rows or len(rows) < 2:
        return
    headers = [cell.text.strip() for cell in rows[0].cells[1:] if cell.text.strip()]
    data.append({"Question ID": parent_qid, "Question Text": parent_qtext, "Option ID": None, "Option Text": None})
    for row_index, row in enumerate(rows[1:], start=1):
        cells = row.cells
        subquestion_text = cells[0].text.strip()
        if not subquestion_text:
            continue
        sub_qid = f"{parent_qid}_{row_index}"
        data.append({"Question ID": sub_qid, "Question Text": subquestion_text, "Option ID": None, "Option Text": None})
        for col_index, option_text in enumerate(headers, start=1):
            if option_text:
                option_qid = f"{sub_qid}.{col_index}"
                data.append({"Question ID": sub_qid, "Question Text": "", "Option ID": option_qid, "Option Text": option_text})

def extract_from_docx_to_excel(docx_file):
    doc = Document(docx_file)
    data, question_counter, table_index = [], 0, 0
    paragraphs, para_index = list(doc.paragraphs), 0
    skip_options, last_matrix_qid, last_matrix_qtext = False, None, None

    while para_index < len(paragraphs):
        para = paragraphs[para_index]
        text = para.text.strip()

        if not text:
            para_index += 1
            continue

     
        if QUESTION_PATTERN.search(text) or re.match(r"^Q\d+\.", text):
            question_counter += 1
            qid, cleaned = f"Q{question_counter}", clean_question_text(text)

            if "matrix" in text.lower():  
                data.append({"Question ID": qid, "Question Text": cleaned, "Option ID": None, "Option Text": None})
                if table_index < len(doc.tables):
                    last_matrix_qid, last_matrix_qtext = qid, cleaned
                    extract_matrix_table(doc.tables[table_index], data, qid, cleaned)
                    table_index += 1
                skip_options = True

            elif "scale" in text.lower(): 
                data.append({"Question ID": qid, "Question Text": cleaned, "Option ID": None, "Option Text": None})
                skip_options = True  

            else: 
                data.append({"Question ID": qid, "Question Text": cleaned, "Option ID": None, "Option Text": None})
                skip_options = False

      
        elif is_list_option(para) and not skip_options and question_counter:
            prefix = get_numbering(para)
            full_option = f"{prefix} {text}" if prefix and not text.startswith(prefix) else text
            qid = f"Q{question_counter}"
            option_count = len([d for d in data if d['Question ID'] == qid and d['Option ID']])
            data.append({"Question ID": qid, "Question Text": None, "Option ID": f"{qid}_option_{option_count+1}", "Option Text": full_option})

        para_index += 1

   
    while table_index < len(doc.tables):
        table = doc.tables[table_index]
        first_col = [row.cells[0].text.strip() for row in table.rows[1:] if row.cells and row.cells[0].text.strip()]
        if any(re.match(r"\d+(\.\d+)?", cell) for cell in first_col):
            parent_qid = last_matrix_qid or f"Q{question_counter+1}"
            extract_matrix_table(table, data, parent_qid, last_matrix_qtext or "Matrix Question")
        table_index += 1

    return pd.DataFrame(data)


st.title("Përkthimi i dokumenteve zyrtare")
mode = st.radio("Zgjidh mënyrën:", ["Ngarko DOCX", "Ngarko XLSForm"])

if mode == "Ngarko DOCX":
    docx_file = st.file_uploader("Ngarko dokumentin Word (.docx)", type=["docx"])
    if docx_file:
        extracted_df = extract_from_docx_to_excel(docx_file)
        st.success("Dokumenti DOCX u ekstraktua me sukses!")
        st.dataframe(extracted_df.head())
        extracted_df.to_excel("cleaned_output.xlsx", index=False)
        st.download_button("Shkarko Excelin e ekstraktuar", data=open("cleaned_output.xlsx", "rb").read(), file_name="cleaned_output.xlsx")

elif mode == "Ngarko XLSForm":
    original_file = st.file_uploader("Ngarko XLSForm-in origjinal (Excelin)", type=["xlsx"])
    translated_file = st.file_uploader("Ngarko Excelin me përkthimin e pastruar", type=["xlsx"])

    if original_file and translated_file:
        import os
        from collections import defaultdict
        from difflib import get_close_matches

        xls = pd.ExcelFile(original_file)
        survey_df = xls.parse("survey")
        choices_df = xls.parse("choices")
        settings_df = xls.parse("settings")
        translated_df = pd.read_excel(translated_file)

        survey_df.columns = survey_df.columns.str.strip()
        choices_df.columns = choices_df.columns.str.strip()
        settings_df.columns = settings_df.columns.str.strip()

        label_columns = [col for col in survey_df.columns if col.startswith("label::")]
        from_label = st.selectbox("Zgjidh kolonën në Excel prej nga do të përkthehet:", label_columns).strip()
        to_label = st.selectbox("Zgjidh kolonën në Excel ku do të vendoset përkthimi:", label_columns).strip()

        def detect_language(label_col):
            if "albanian" in label_col.lower():
                return "al"
            elif "serbian" in label_col.lower():
                return "sr"
            elif "english" in label_col.lower():
                return "en"
            return None

        from_lang = detect_language(from_label)
        to_lang = detect_language(to_label)

        if not from_lang or not to_lang:
            st.warning("Nuk u detektua gjuha nga kolonat e përzgjedhura. Sigurohu që kolonat përmbajnë emrin e gjuhës.")
            st.stop()

        def clean_label(val):
            if pd.isna(val): return ""
            text = str(val).strip().lower()
            text = re.sub(r"[\u2013\u2014-]", "-", text)
            text = re.sub(r"\s+", " ", text)
            return text

        def capitalize_first(text):
            return text[0].upper() + text[1:] if text else text

        question_translations = translated_df[translated_df["Option ID"].isna()].set_index("Question ID")["Question Text"].to_dict()
        option_translations = translated_df[translated_df["Option ID"].notna()].set_index("Option ID")["Option Text"].to_dict()

        def build_translation_dictionaries():
            manual_al_to_sr, manual_sr_to_al = {}, {}
            manual_al_to_en, manual_en_to_al = {}, {}
            manual_sr_to_en, manual_en_to_sr = {}, {}

            core_pairs = [
                ("GPS", "GPS", "GPS"),
                ("Anketuesi_ja", "Anketar/e", "Enumerator"),
                ("A pranoni të merrni pjesë në anketë?", "Da li se slažete da učestvujete u anketi?", "Do you agree to participate in the survey?"),
                ("Arsyet e refuzimit", "Razlozi za odbijanje", "Reasons for refusal"),
                ("Tjetër, specifiko", "Drugo, navedite", "Other, specify"),
                ("Po", "Da", "Yes"),
                ("Jo", "Ne", "No"),
                ("Tjetër. Çka?", "Drugo. Šta?", "Other. What?"),
                ("Tjetër, ju lutem specifikoni", "Drugo, navedite", "Other, please specify")
            ]

            likert_pairs = [
                ("1-Aspak i kënaqur", "1-Uopšte nisam zadovoljan/na", "1-Not at all satisfied"),
                ("5-Plotësisht i kënaqur", "5-Potpuno zadovoljan/zadovoljna", "5-Completely satisfied"),

                ("1-Aspak nuk pajtohem", "1-Uopšte se ne slažem", "1-Strongly disagree"),
                ("5-Plotësisht pajtohem", "5-Potpuno se slažem", "5-Strongly agree"),

                ("1– Aspak efektive", "1 – Uopšte efektivno", "1-Not effective at all"),
                ("5– Plotësisht efektive", "5 – Potpuno efektivno", "5-Completely effective"),

                ("1-Aspak e sigurtë", "1-Uopšte nije bezbedno", "1-Not safe at all"),
                ("5-Plotësisht e sigurtë", "5-Potpuno je bezbedno", "5-Completely safe"),

                ("1-Aspak meritore", "1-Nimalo zaslužne", "1-Not deserving at all"),
                ("5-Plotësisht meritore", "5-Potpuno zaslužne", "5-Completely deserving"),

                ("Shumë negative", "Veoma negativno", "Very negative"),
                ("Shumë pozitive", "Veoma pozitivno", "Very positive"),

                ("1 – aspak i mirë", "1 – uopšte nije dobar", "1-Not good at all"),
                ("5 – shumë i mirë", "5 – veoma dobar", "5-Very good"),

                ("88 – Refuzoj të përgjigjem", "Odbijam odgovoriti", "88-Refuse to answer"),
                ("Refuzoj të përgjigjem", "Odbijam odgovoriti", "Refuse to answer")
            ]
        
            demographic_pairs = [
                ("D1. (GJINIA)", "D1. (ROD/POL)", "D1. (GENDER)"),
                ("D2. (MOSHA) (vjet)", "D2. (STAROST) (godine)", "D2. (AGE) (years)"),
                ("D3. (STATUSI MARTESOR)  Aktualisht Ju jeni...", "D3. (BRAČNO STANJE) Trenutno vi ste…", "D3. (MARITAL STATUS) Currently you are..."),
                ("D4.  (EDUKIMI)  Sa vite shkollë i keni kryer?", "D4. (OBRAZOVANJE) Koliko godina škole ste završili?", "D4. (EDUCATION) How many years of schooling have you completed?"),
                ("D5.  (PËRKATËSIA ETNIKE)  Cili është nacionaliteti Juaj/cilit grup i takoni?", "D5. (ETNIČKA PRIPADNOST) Koja je vaša etnička pripadnost/kojoj grupi pripadate?", "D5. (ETHNICITY) What is your nationality/which group do you belong to?"),
                ("Tjetër. Cili?", "Drugo. Koja?", "Other. Which?"),
                ("D6. (FAMILJA)  Sa anëtarë i ka familja Juaj?", "D6. (PORODICA) Koliko članova ima vaša porodica?", "D6. (FAMILY) How many members are in your family?"),
                ("D8. (TË ARDHURAT PERSONALE) A mund të na tregoni se sa kanë qenë të ardhurat personale në muajin e fundit?", 
                "D8. (LIČNI PRIHODI) Da li nam možete reći koliki su bili vaši lični prihodi u zadnjem mesecu?", 
                "D8. (PERSONAL INCOME) Can you tell us what your personal income was last month?"),
                ("D9.  (TË ARDHURAT FAMILJARE) A mund të na tregoni se sa kanë qenë të ardhurat familjare në muajin e fundit?", 
                "D9. (PORODIČNI PRIHODI) Da li nam možete reći koliki je bio vaš porodični prihod u zadnjem mesecu?", 
                "D9. (HOUSEHOLD INCOME) Can you tell us what your household income was last month?"),
                ("D10. Komuna", "Opstina", "D10. Municipality"),
                ("D11.    VENDBANIMI", "D11. PREBIVALIŠTE", "D11. Residence"),
                ("Emri i lagjes", "Naziv komšiluka", "Neighborhood name"),
                ("Emri i fshatit", "Ime sela", "Village name"),
                ("Emri dhe mbiemri", "Ime i prezime", "Full name"),
                ("Numri i telefonit", "Broj telefona", "Phone number")
            ]

            extra_pairs = [("Mashkull", "Muško", "Male"), ("Femër", "Žensko", "Female")]
            reason_pairs = [
                ("Mungesa e kohës", "Nedostatak vremena", "Lack of time"),
                ("Jo i interesuar", "Nije zainteresovan", "Not interested"),
                ("Mbrojtja e të dhënave, përdorimi i të drejtës së privatësisë", 
                "Zaštita podataka, korišćenje politike privatnosti", 
                "Data protection, use of privacy rights"),
                ("Nuk beson në sondazhe", "Ne veruje u ankete", "Does not believe in surveys"),
                ("Të tjera (nuk di të përgjigjet, kushtet e motit, frikë nga pyetjet)", 
                "Ostalo (ne zna da odgovori, vremenski uslovi, strah od pitanja)", 
                "Other (don’t know how to answer, weather conditions, fear of questions)"),
                ("Problemet e shëndetit", "Zdravstveni problemi", "Health problems"),
                ("Moshë më e vjetër", "Starije godine", "Older age"),
                ("Nuk i pëlqen subjekti i kërkimit", "Ne voli temu istraživanja", "Does not like research topic"),
                ("Ka pasur një përvojë të keqe me sondazhet", 
                "Imao/la je loše iskustvo sa anketama", 
                "Had a bad experience with surveys"),
                ("Asnjë arsye", "Nema razloga", "No reason")
            ]

            demographic_pairs_answers = [
                ("Mashkull", "Muško", "Male"),
                ("Femër", "Žensko", "Female"),
                ("I/ e martuar", "Oženjen/Udata", "Married"),
                ("I/ e pamartuar", "Neoženjen/Neudata", "Single"),
                ("I/ e ndarë", "Razveden/a", "Divorced"),
                ("I/e vej", "Udovac/udovica", "Widowed"),
                ("Disa vite të shkollës fillore", "Nekoliko godina osnovne škole", "Some years of primary school"),
                ("Shkolla fillore", "Osnovna škola", "Primary school"),
                ("Disa vite të shkollës së mesme", "Nekoliko godina srednje škole", "Some years of secondary school"),
                ("Shkolla e mesme", "Srednja škola", "Secondary school"),
                ("Student", "Student", "Student"),
                ("Fakultet", "Fakultet", "University"),
                ("Magjistraturë ose Doktoraturë", "Magistratura ili", "Masters or Doctorate"),
                ("Shqiptar", "Albanska", "Albanian"),
                ("Serb", "Srpska", "Serbian"),
                ("Boshnjak", "Bosanska", "Bosniak"),
                ("Goran", "Goranska", "Gorani"),
                ("Turk", "Turska", "Turkish"),
                ("Rom", "Romska", "Roma"),
                ("Ashkali", "Aškalijska", "Ashkali"),
                ("Egjiptas", "Egipatska", "Egyptian"),
                ("Tjetër. Cili?", "Drugo. Koja?", "Other. Which?"),
                ("DK/PP", "Ne znam/Bez odgovora", "Don't know/No answer"),
            ]

            income_pairs = [
                ("Deri 150 euro", "Do 150 evra", "Up to 150 euros"),
                ("151-300 euro", "151-300 evra", "151-300 euros"),
                ("301-450 euro", "301-450 evra", "301-450 euros"),
                ("451-600 euro", "451-600 evra", "451-600 euros"),
                ("601-750 euro", "601-750 evra", "601-750 euros"),
                ("751-900 euro", "751-900 evra", "751-900 euros"),
                ("Mbi 900 euro", "Preko 900 evra", "Over 900 euros"),
                ("Nuk kam realizuar fare të ardhura", "Nisam ostvario/la nikakav prihod.", "I had no income"),
                ("Refuzon/PP", "Odbija/BO", "Refused/No answer")
            ]

            frequency_pairs = [
                ("Asnjëherë", "Nikad", "Never"),
                ("Rallë", "Retko", "Rarely"),
                ("Ndonjëherë", "Ponekad", "Sometimes"),
                ("Shpesh", "Često", "Often"),
                ("Gjithmonë", "Uvek", "Always")
            ]

            awareness_pairs = [
                ("Shumë i informuar", "Veoma informisani", "Very informed"),
                ("Deri diku i informuar", "Donekle informisani", "Somewhat informed"),
                ("Deri diku jo i informuar", "Donekle ne informisani", "Somewhat uninformed"),
                ("Aspak i informuar", "Potpuno ne informisani", "Not at all informed")
            ]

            satisfaction_pairs = [
                ("Shumë të kënaqur", "Veoma zadovoljni", "Very satisfied"),
                ("Deri diku i kënaqur", "Donekle zadovoljni", "Somewhat satisfied"),
                ("Deri diku jo i kënaqur", "Donekle nezadovoljni", "Somewhat dissatisfied"),
                ("Aspak i kënaqur", "Potpuno nezadovoljni", "Not at all satisfied"),
                ("Shumë i/e kënaqur", "Veoma zadovoljni", "Very satisfied"),
                ("I/e kënaqur", "Zadovoljni", "Satisfied"),
                ("I/e pakënaqur", "Nezadovoljni", "Dissatisfied"),
                ("Shumë i/e pakënaqur", "Veoma nezadovoljni", "Very dissatisfied"),
                ("Nuk e di/refuzoj të përgjigjem (mos e lexo)", 
                "Ne znam/Odbijam odgovoriti (nemojte čitati)", 
                "Don't know/Refuse to answer (do not read)")
            ]
            employment_pairs = [
                ("I papunësuar – duke kërkuar punë", "Nezaposlen/a – tražim posao", "Unemployed – seeking work"),
                ("I papunësuar – duke mos kërkuar punë", "Nezaposlen/a – ne tražim posao", "Unemployed – not seeking work"),
                ("I punësuar në sektorin publik", "Zaposlen/a u javnom sektoru", "Employed in public sector"),
                ("I punësuar në sektorin privat", "Zaposlen/a u privatnom sektoru", "Employed in private sector"),
                ("I punësuar kohë pas kohe", "Zaposlen/a s vremena na vreme", "Employed occasionally"),
                ("Pensionist", "Penzioner", "Pensioner"),
                ("Amvise", "Domaćica", "Housewife"),
                ("Student/ nxënës", "Student/učenik", "Student/Pupil"),
                ("Tjetër. Çka?", "Drugo. Šta?", "Other. What?")
            ]

            voting_pairs = [
                ("Gjithsesi do të votoja", "Svakako bih glasao/ala", "Definitely would vote"),
                ("Ndoshta do të votoja", "Možda bih glasao/ala", "Might vote"),
                ("Me gjasë nuk do të votoja", "Verovatno ne bih glasao/ala", "Probably would not vote"),
                ("Definitivisht nuk do të votoja", "Definitivno ne bih glasao/ala", "Definitely would not vote")
            ]

            all_pairs = core_pairs + voting_pairs + employment_pairs + satisfaction_pairs + awareness_pairs + frequency_pairs + likert_pairs + demographic_pairs + extra_pairs + reason_pairs + income_pairs + demographic_pairs_answers

            for al, sr, en in all_pairs:
                manual_al_to_sr[clean_label(al)] = sr
                manual_sr_to_al[clean_label(sr)] = capitalize_first(al)
                manual_al_to_en[clean_label(al)] = en
                manual_en_to_al[clean_label(en)] = capitalize_first(al)
                manual_sr_to_en[clean_label(sr)] = en
                manual_en_to_sr[clean_label(en)] = sr

            return {
                ("al", "sr"): manual_al_to_sr,
                ("sr", "al"): manual_sr_to_al,
                ("al", "en"): manual_al_to_en,
                ("en", "al"): manual_en_to_al,
                ("sr", "en"): manual_sr_to_en,
                ("en", "sr"): manual_en_to_sr,
            }

        manual_translations = build_translation_dictionaries()

        def fuzzy_lookup(word, dictionary):
            if not word: return ""
            if word in dictionary: return dictionary[word]
            matches = get_close_matches(word, dictionary.keys(), n=1, cutoff=0.9)
            return dictionary[matches[0]] if matches else ""

        unmatched_terms = []

        def apply_manual(text):
            normalized = clean_label(text)
            translation_dict = manual_translations.get((from_lang, to_lang), {})
            translation = fuzzy_lookup(normalized, translation_dict)
            if translation:
                return capitalize_first(translation)
            if normalized:
                unmatched_terms.append(text)
            return ""

        def map_name_to_qid(name):
            if pd.isna(name): return None
            name = str(name).strip().lower()
            if re.match(r"^p\d+_\d+$", name): return f"Q{name[1:]}"
            if re.match(r"^p\d+$", name): return f"Q{name[1:]}"
            return None

        def guess_qid_from_label(label_text):
            label_text = clean_label(label_text)
            match = re.match(r"(\d+(\.\d+)?)", label_text)
            return f"Q{match.group(1).replace('.', '_')}" if match else None

        def get_question_id(row):
            return map_name_to_qid(row.get("name")) or guess_qid_from_label(row.get(from_label))

        def translate_question_auto(row):
            if isinstance(row.get("type"), str) and row["type"].lower().startswith("begin_group"):
                return row.get(from_label, "")
            qid = get_question_id(row)
            if qid and qid in question_translations:
                return capitalize_first(question_translations[qid])
            return ""

        survey_df[to_label] = survey_df.apply(translate_question_auto, axis=1)
        survey_df[to_label] = survey_df[to_label].where(survey_df[to_label] != "", survey_df[from_label].apply(apply_manual))

        list_qid_map = defaultdict(list)
        for _, row in survey_df.iterrows():
            if isinstance(row.get("type"), str) and ("select_one" in row["type"] or "select_multiple" in row["type"]):
                qid = get_question_id(row)
                if qid:
                    list_qid_map[row["type"].split()[1]].append(qid)

        def build_option_id(row):
            list_name, raw_name = row.get("list_name"), row.get("name")
            qids = list_qid_map.get(list_name)
            if pd.isna(raw_name) or not qids: return None
            match = re.search(r"(\d+)", str(raw_name))
            if not match: return None
            num = match.group(1)
            for qid in qids:
                candidate = f"{qid}_option_{num}"
                if candidate in option_translations: return candidate
            return None

        choices_df["Option ID"] = choices_df.apply(build_option_id, axis=1)

        if from_label in choices_df.columns and to_label in choices_df.columns:
            choices_df[to_label] = choices_df["Option ID"].map(option_translations).fillna("")
            choices_df[to_label] = choices_df[to_label].where(choices_df[to_label] != "", choices_df[from_label].apply(apply_manual))
        else:
            st.warning(f"⚠️ Kolona '{from_label}' ose '{to_label}' nuk ekziston në fletën 'choices'. U anashkalua përkthimi i choices.")

        choices_df.drop(columns=["Option ID"], inplace=True)

        output_file = "translated_output.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            survey_df.to_excel(writer, sheet_name="survey", index=False)
            choices_df.to_excel(writer, sheet_name="choices", index=False)
            settings_df.to_excel(writer, sheet_name="settings", index=False)

        base_name = os.path.splitext(original_file.name)[0]
        translated_file_name = f"{base_name}_perkthyer.xlsx"

        st.success("Përkthimi u përfundua me sukses!")
        st.download_button("Shkarko Excelin e Përkthyer", data=open(output_file, "rb").read(), file_name=translated_file_name)
