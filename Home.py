import streamlit as st
from PIL import Image
import os


st.set_page_config(
    page_title="Vegla Pyetësorësh",
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

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("Gjenero XLS"):
        st.switch_page("pages/1_Gjenero_XLS.py")

with col2:
    if st.button("Përkthe me AI"):
        st.switch_page("pages/2_Perkthe_me_AI.py")

with col3:
    if st.button("Përkthim Zyrtar"):
        st.switch_page("pages/3_Perkthe_Zyrtarisht.py")
