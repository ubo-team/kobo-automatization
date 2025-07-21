import streamlit as st
from PIL import Image
import os

st.set_page_config(page_title="Përkthim Zyrtar", layout="centered")
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

st.title("Përkthe pyetësorët zyrtarisht")
st.write("Kjo pjesë është në ndërtim e sipër. Së shpejti do të mund të përktheni dokumente në mënyrë zyrtare.")
