"""Punto de entrada de TESSAIX.

Toda la lógica vive en el paquete `tessaix/`, organizado por responsabilidad
(config, estilos, autenticación, generación de contenido, construcción del
.pptx, y cada pantalla del wizard). Este archivo sólo importa esas piezas y
las conecta.
"""
import streamlit as st

from tessaix.auth import check_auth
from tessaix.config import (
    BUSINESS_LINE_HC_FINANCE, FAVICON, STEP_CLIENTE, STEP_CONFIG, STEP_CONTEXTO,
    STEP_EQUIPO, STEP_GENERATE, STEP_REVIEW, STEP_SERVICIOS,
)
from tessaix.styles import inject_css, render_header, render_steps
from tessaix.ui_generate import step_generate
from tessaix.ui_home import render_home
from tessaix.ui_review import step_review
from tessaix.ui_wizard import step_cliente, step_config, step_contexto, step_equipo, step_servicios

STEP_RENDERERS = {
    STEP_CONFIG: step_config,
    STEP_CLIENTE: step_cliente,
    STEP_SERVICIOS: step_servicios,
    STEP_EQUIPO: step_equipo,
    STEP_CONTEXTO: step_contexto,
    STEP_REVIEW: step_review,
    STEP_GENERATE: step_generate,
}

DEFAULT_SESSION_STATE = [
    ("screen", "home"),
    ("step", 0),
    ("form", {}),
    ("content", None),
    ("pptx", None),
]


def main():
    st.set_page_config(
        page_title="TESSAIX",
        page_icon=str(FAVICON) if FAVICON.exists() else "T",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    check_auth()
    inject_css()

    for k, v in DEFAULT_SESSION_STATE:
        if k not in st.session_state:
            st.session_state[k] = v

    if st.session_state.screen == "home":
        render_home()
        return

    if st.session_state.screen == "finance":
        render_header("FINANCE")
        st.markdown("""
        <div style="text-align:center;padding:80px 0">
          <div style="font-size:64px;margin-bottom:20px">🏗️</div>
          <div style="font-size:1.2rem;font-weight:700;color:#202031;margin-bottom:12px">En obras</div>
          <div style="color:#888;font-size:.9rem;line-height:1.8">
            Tessera Finance estará disponible muy pronto.</div>
        </div>""", unsafe_allow_html=True)
        if st.button("← Volver"):
            st.session_state.screen = "home"
            st.rerun()
        return

    business_line = st.session_state.form.get("business_line")
    render_header("HUMAN CAPITAL + FINANCE" if business_line == BUSINESS_LINE_HC_FINANCE else "HUMAN CAPITAL")
    step = st.session_state.step
    render_steps(step)

    STEP_RENDERERS[step]()


if __name__ == "__main__":
    main()
