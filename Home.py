import streamlit as st
from PIL import Image
import os

st.set_page_config(
    page_title="Platforma për Menaxhimin e Pyetësorëve",
    layout="wide",
)

st.markdown("""
<style>

/* Remove sidebar */
section[data-testid="stSidebar"] {
    display: none !important;
}

div[data-testid="stAppViewContainer"] > .main {
    margin-left: 0 !important;
}sa

</style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------
# SIDEBAR LOGO
# ---------------------------------------------------------------
logo_path = "logo.png"

with st.sidebar:
    if os.path.exists(logo_path):
        st.image(Image.open(logo_path), width=150)

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


# ---------------------------------------------------------------
# CSS PER KARTAT KLIKUESE
# ---------------------------------------------------------------
st.markdown("""
<style>

.card-button {
    display: block;
    background-color: #ffffff;
    text-decoration: none !important;
    border-radius: 14px;
    padding: 30px;
    height: 230px;
    color: #000000;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    transition: all 0.15s ease;
    border: 1px solid #eee;
}

.card-button:hover {
    transform: translateY(-6px);
    box-shadow: 0 4px 14px rgba(0,0,0,0.12);
    border: 2px solid #d0d0ff;
}

.card-title {
    font-size: 20px;
    font-weight: 600;
    margin: 0;   /* reset margins */
    line-height: 1.2;
}

.card-desc {
    font-size: 15px;
    color: #444;
    margin-top: 20px;
    min-height: 70px;
}

.arrow {
    font-size: 40px;
    font-weight: bold;
    color: #0054a3;
    text-align: right;
    margin-top: 10px;
}
            
.card-header {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 8px;
}

.card-header svg {
    width: 28px;
    height: 28px;
    flex-shrink: 0;
    fill: #0054a3 !important;
}

</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------
# FUNKSIONI I KARTAVE (CLICKABLE)
# ---------------------------------------------------------------
def card(title, description, page, icon_path):
    svg_content = load_svg(icon_path)

    st.markdown(
        f"""
<a href="/{page}" target="_self" class="card-button">
    <div class="card-header">
        <div class="svg-icon">{svg_content}</div>
        <div class="card-title">{title}</div>
    </div>
    <div class="card-desc">{description}</div>
    <div class="arrow">→</div>
</a>
""",
        unsafe_allow_html=True
    )


def render_svg(svg_string):
    """Renders an SVG string as HTML."""
    b64 = svg_string.encode("utf-8").decode("utf-8")
    html = f"<div>{b64}</div>"
    st.markdown(html, unsafe_allow_html=True)

def load_svg(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


# ---------------------------------------------------------------
# LAYOUT (3×2 KARTA)
# ---------------------------------------------------------------
st.title("Platforma për Menaxhimin e Pyetësorëve")
st.markdown("Zgjidh një nga veglat për të vazhduar")
st.markdown("---")

# ROW 1
col1, col2, col3 = st.columns(3)

with col1:
    card(
        "Gjenero XLS për KoboToolbox",
        "Krijo dhe menaxho dokumenta Excel për pyetësorët dhe anketat në mënyrë të automatizuar.",
        "Gjenero_XLS",
        "icons/survey-xmark.svg"
    )

with col2:
    card(
        "Përkthim Excel Files AI",
        "Përkthe Excel automatikisht duke përdorur inteligjencë artificiale për rezultate të shpejta.",
        "Perkthim_Excel_Files_AI",
        "icons/file-excel.svg"
    )

with col3:
    card(
        "Përkthim Word Documents AI",
        "Përkthe dokumente Word shpejt dhe saktë me teknologji të avancuar AI.",
        "Perkthim_Word_Documents_AI",
        "icons/file-word.svg"
    )


st.markdown("<div class='row-spacer'></div>", unsafe_allow_html=True)

# ROW 2
col4, col5, col6 = st.columns(3)


with col4:
    card(
        "Përkthim Zyrtar",
        "Përkthime të verifikuara dhe zyrtare për dokumentet që kërkojnë saktësi të plotë.",
        "Perkthe_Zyrtarisht",
        "icons/language-exchange.svg"
    )

with col5:
    card(
        "MaxDiff Analysis",
        "Analizo të dhënat me metodën MaxDiff për rezultate të thelluara.",
        "MaxDiff_Analysis",
        "icons/analyse.svg"
    )

with col6:
    card(
        "Dokumentimi",
        "Qasje e plotë në manuale, udhëzime dhe resurse të platformës.",
        "Dokumentimi",
        "icons/info.svg"
    )