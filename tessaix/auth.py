"""Pantalla de acceso por contraseña."""
import streamlit as st

from .assets import make_ailerons_png
from .config import PASSWORD
from .styles import LOGIN_CSS


def check_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        tessaix_b64 = make_ailerons_png("TESSAIX", color=(255, 255, 255), size=100, width=700, height=130)
        st.markdown(LOGIN_CSS, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="login-wrap">
          <img src="data:image/png;base64,{tessaix_b64}" style="width:380px">
          <div class="tagline">Propuestas comerciales inteligentes</div>
        </div>""", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            pwd = st.text_input("pwd", type="password", placeholder="Contraseña de acceso",
                                 label_visibility="collapsed")
            if st.button("Entrar →", use_container_width=True):
                if pwd == PASSWORD:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        st.stop()
