import importlib
import os

import streamlit as st

import ubo_ui

# Pas një push-i, Streamlit Cloud e lexon përsëri këtë skedar, por mund të mbajë në memorie ubo_ui.py-në e
# vjetër (faqja atëherë thërret funksione që s'ekzistojnë). Kur skedari në disk është më i ri, ringarkohet.
if os.path.getmtime(ubo_ui.__file__) > getattr(ubo_ui, "LOADED_AT", 0):
    ubo_ui = importlib.reload(ubo_ui)

# ---------------------------------------------------------------
# NAVIGIMI I PLATFORMËS
# Veglat, emrat dhe grupet vijnë nga ubo_ui.TOOLS (të njëjtat si te kryefaqja).
# url_path mban lidhjet e vjetra, që linqet ekzistuese të vazhdojnë të punojnë.
# ---------------------------------------------------------------
home = st.Page("kryefaqja.py", title="Kryefaqja", default=True)
tool_pages = {tool.url_path: st.Page(tool.file, title=tool.title, url_path=tool.url_path) for tool in ubo_ui.TOOLS}

# Menyja e Streamlit-it fshihet dhe ndërtohet këtu vetë, vetëm për veglat: kryefaqja nuk ka menu anësore
# fare (më parë krijohej e pastaj fshihej me CSS, prandaj dukej për një çast kur hapej faqja).
page = st.navigation([home, *tool_pages.values()], position="hidden")
tool = next((t for t in ubo_ui.TOOLS if t.url_path == page.url_path), None)

# Për çdo faqe: ngjyrat e temës (e çelët / e errët), butoni i temës dhe dritarja "Ju lutem prisni"
st.markdown(ubo_ui.base_css(), unsafe_allow_html=True)
with st.container(key="ubo_js"):
    st.iframe(ubo_ui.page_script(), height=1)

if tool is not None:
    with st.sidebar:
        # logo lart; klikimi mbi të të kthen gjithmonë në kryefaqe
        st.markdown(f'<a class="ubo-home-logo" href="./" target="_self" title="Kryefaqja">'
                    f'{ubo_ui.logo_imgs()}</a>', unsafe_allow_html=True)
        # menyja me ngjyra: çdo vegël me ngjyrën e grupit të saj, vegla aktive e theksuar
        st.markdown(ubo_ui.sidebar_css(tool), unsafe_allow_html=True)
        st.page_link(home, label="Kryefaqja", icon=ubo_ui.NAV_ICONS[""], width="stretch")
        for section, (color, dark) in ubo_ui.SECTIONS.items():
            st.markdown(f'<div class="ubo-nav-sec" style="--c:{color};--cd:{dark}"><i style="background:{color}"></i>'
                        f'{section}</div>', unsafe_allow_html=True)
            for t in ubo_ui.TOOLS:
                if t.section == section:
                    st.page_link(tool_pages[t.url_path], label=t.title, icon=ubo_ui.NAV_ICONS.get(t.url_path),
                                 width="stretch")
    # kreu i njëjtë për çdo vegël, bashkë me stilin e faqes (zëvendëson titujt e veçantë të faqeve)
    st.markdown(ubo_ui.tool_header(tool), unsafe_allow_html=True)

page.run()
