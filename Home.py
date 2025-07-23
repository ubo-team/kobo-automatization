import streamlit as st
from PIL import Image
import os

st.set_page_config(
    page_title="Platforma për Automatizimin e Kodimit të Pyetësorëve",
    layout="centered",
)

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

st.title("Platforma për Menaxhimin e Pyetësorëve")
st.markdown("---")
st.markdown("Zgjidh një nga veglat më poshtë për të vazhduar:")

col1, col3 = st.columns(2)

with col1:
    st.subheader("Gjenero XLS")
    if st.button("Shko tek Vegla", key="xls"):
        st.switch_page("pages/1_Gjenero_XLS.py")

with col3:
    st.subheader("Përkthim Zyrtar")
    if st.button("Shko tek Vegla", key="zyrtar"):
        st.switch_page("pages/3_Perkthe_Zyrtarisht.py")

col2, col4 = st.columns(2)

with col2:
    st.subheader("Përkthim Excel Files AI")
    if st.button("Shko tek Vegla", key="ai"):
        st.switch_page("pages/2_Perkthim_Excel_Files_AI.py")

with col4:
    st.subheader("Përkthim Word Documents AI")
    if st.button("Shko tek Vegla", key="word"):
        st.switch_page("pages/4_Perkthim_Word_Documents_AI.py")
