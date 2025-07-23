import streamlit as st
from PIL import Image
import os

# -------------------------------
# Page Configuration
# -------------------------------
st.set_page_config(page_title="Dokumentimi", layout="centered")

# -------------------------------
# Sidebar Logo
# -------------------------------
logo_path = "logo.png"  # Make sure the logo is in the root directory

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


st.title("Dokumentimi")

st.markdown("""
Kjo faqe do të shërbejë si qendër dokumentimi për të gjitha funksionalitetet e aplikacionit.

Dokumentet përkatëse dhe udhëzuesit do të ngarkohen së shpejti si Word Documents.

Nëse diçka nuk është e qartë ose kërkoni sqarime shtesë,
ju lutemi drejtohuni tek departamenti përkatës për udhëzime të mëtejshme.
""")
