"""Pantalla inicial: selección de línea de negocio."""
import base64

import streamlit as st

from .assets import make_ailerons_png
from .config import _find


def render_home():
    # Logos en Ailerons PIL — tamaño grande y legible
    tessera_hdr = make_ailerons_png("TESSERA", color=(255, 255, 255), size=42, width=380, height=62)
    tessaix_hdr = make_ailerons_png("TESSAIX", color=(255, 255, 255), size=42, width=380, height=62)
    hc_logo = make_ailerons_png("TESSERA", color=(255, 255, 255), size=36, width=320, height=54)
    fi_logo = make_ailerons_png("TESSERA", color=(240, 232, 222), size=36, width=320, height=54)

    # Logos externos
    cd_b64, lt_b64 = "", ""
    try:
        with open(str(_find("costdown_logo.png")), 'rb') as f:
            cd_b64 = base64.b64encode(f.read()).decode()
    except Exception:
        pass
    try:
        with open(str(_find("logo_LT_.svg")), 'rb') as f:
            lt_b64 = base64.b64encode(f.read()).decode()
    except Exception:
        pass

    cd_img = f'<img src="data:image/png;base64,{cd_b64}" style="height:44px;margin-bottom:12px;opacity:.4;filter:brightness(0) invert(1)">' if cd_b64 else '<div style="font-size:2rem;margin-bottom:12px;opacity:.4">💰</div>'
    lt_img = f'<img src="data:image/svg+xml;base64,{lt_b64}" style="height:36px;margin-bottom:12px;opacity:.4;filter:brightness(0) invert(1)">' if lt_b64 else '<div style="font-size:2rem;margin-bottom:12px;opacity:.4">📊</div>'

    # CSS global
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800;900&display=swap');
    [data-testid="stAppViewContainer"]{background:#202031!important}
    [data-testid="stHeader"]{display:none}
    .stApp{background:#202031!important;min-height:100vh}
    .block-container{padding-top:0!important;max-width:1100px!important}
    .neb{position:fixed;border-radius:50%;filter:blur(100px);pointer-events:none;z-index:0}
    .neb1{width:650px;height:550px;top:-80px;right:-100px;background:rgba(88,117,121,0.45);animation:orb1 5s ease-in-out infinite alternate}
    .neb2{width:550px;height:450px;bottom:-60px;left:-80px;background:rgba(88,117,121,0.3);animation:orb2 6s ease-in-out infinite alternate}
    .neb3{width:400px;height:350px;bottom:120px;right:100px;background:rgba(251,224,160,0.2);animation:orb3 7s ease-in-out infinite alternate}
    .neb4{width:300px;height:260px;top:160px;left:80px;background:rgba(251,224,160,0.13);animation:orb4 4s ease-in-out infinite alternate}
    @keyframes orb1{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-80px,60px) scale(1.2)}}
    @keyframes orb2{0%{transform:translate(0,0) scale(1)}100%{transform:translate(70px,-70px) scale(1.25)}}
    @keyframes orb3{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-60px,50px) scale(0.85)}}
    @keyframes orb4{0%{transform:translate(0,0) scale(1)}100%{transform:translate(50px,60px) scale(1.15)}}
    /* Cards */
    .card{background:rgba(255,255,255,0.04);border:1.5px solid rgba(255,255,255,0.1);
      border-radius:16px;padding:36px 24px;text-align:center;min-height:240px;
      transition:border-color .3s,transform .3s,background .3s;position:relative;overflow:hidden}
    .card:hover{border-color:rgba(88,117,121,0.7);transform:translateY(-4px);background:rgba(88,117,121,0.08)}
    .card-active{border-color:rgba(88,117,121,0.6)!important;cursor:pointer}
    .card-inactive{opacity:.7;cursor:default}
    .card-label{font-size:.75rem;font-weight:700;letter-spacing:.16em;text-transform:uppercase;
                font-family:Raleway,sans-serif;margin:10px 0 8px}
    .card-desc{font-size:.78rem;line-height:1.8;font-family:Raleway,sans-serif}
    .card-badge{display:inline-block;margin-top:12px;font-size:.65rem;font-weight:700;
      letter-spacing:.1em;text-transform:uppercase;
      background:rgba(251,224,160,0.15);color:#FBE0A0;border-radius:4px;padding:3px 10px}
    .stImage img{border:none!important}
    [data-testid="stImage"]{background:transparent!important}
    div.stButton>button{
      background:rgba(88,117,121,0.15)!important;color:#587579!important;
      border:1px solid rgba(88,117,121,0.4)!important;
      font-family:Raleway,sans-serif!important;font-weight:700!important;
      font-size:.7rem!important;letter-spacing:.12em!important;
      border-radius:6px!important;padding:8px 0!important;
      width:100%!important;margin-top:8px!important}
    div.stButton>button:hover{background:#587579!important;color:white!important}
    </style>
    <div class="neb neb1"></div><div class="neb neb2"></div>
    <div class="neb neb3"></div><div class="neb neb4"></div>
    """, unsafe_allow_html=True)

    # ── Header ────────────────────────────────────────────────────────────
    st.markdown("<div style='height:56px'></div>", unsafe_allow_html=True)

    _, hcol, _ = st.columns([1, 4, 1])
    with hcol:
        ha, hb, hc = st.columns([3, 1, 3])
        with ha:
            st.image(f"data:image/png;base64,{tessera_hdr}", use_container_width=True)
        with hb:
            st.markdown("<div style='text-align:center;color:rgba(255,255,255,0.2);font-size:2rem;line-height:1.8'>|</div>", unsafe_allow_html=True)
        with hc:
            st.image(f"data:image/png;base64,{tessaix_hdr}", use_container_width=True)

    st.markdown("""
    <div style='text-align:center;color:rgba(255,255,255,0.4);font-size:.78rem;
                letter-spacing:.22em;text-transform:uppercase;font-family:Raleway,sans-serif;
                margin:20px 0 48px'>
      Propuestas comerciales inteligentes
    </div>""", unsafe_allow_html=True)

    # ── Cards ─────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="card card-active">
          <img src="data:image/png;base64,{hc_logo}" style="height:44px;margin-bottom:10px">
          <div class="card-label" style="color:#587579">Human Capital</div>
          <div class="card-desc" style="color:rgba(240,232,222,0.82)">
            Recruitment · HR Advisory<br>Outsourcing · RPO
          </div>
        </div>""", unsafe_allow_html=True)
        st.markdown("""
        <style>
        div[data-testid="stButton"]:has(button#btn_hc_label) button {
          background: rgba(88,117,121,0.15) !important;
          border: 1px solid rgba(88,117,121,0.4) !important;
          color: #587579 !important;
          font-size: .7rem !important;
          letter-spacing: .12em !important;
          font-family: Raleway, sans-serif !important;
          font-weight: 700 !important;
          border-radius: 6px !important;
          padding: 8px 0 !important;
          width: 100% !important;
          margin-top: 8px !important;
          cursor: pointer !important;
        }
        </style>""", unsafe_allow_html=True)
        if st.button("Entrar →", key="btn_hc", use_container_width=True):
            st.session_state.screen = "hc"
            st.rerun()

    with c2:
        st.markdown(f"""
        <div class="card card-inactive">
          <img src="data:image/png;base64,{fi_logo}" style="height:44px;margin-bottom:10px;opacity:.55">
          <div class="card-label" style="color:rgba(240,232,222,0.55)">Finance</div>
          <div class="card-desc" style="color:rgba(240,232,222,0.42)">
            CFO · M&amp;A<br>Due Diligence
          </div>
          <div class="card-badge">🏗️ En obras</div>
        </div>""", unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="card card-inactive">
          {cd_img}
          <div class="card-label" style="color:rgba(240,232,222,0.65)">Cost Down</div>
          <div class="card-desc" style="color:rgba(240,232,222,0.45)">
            Reducción de costes<br>operativos
          </div>
          <div class="card-badge">🏗️ En obras</div>
        </div>""", unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="card card-inactive">
          {lt_img}
          <div class="card-label" style="color:rgba(240,232,222,0.65)">LT Impulsa</div>
          <div class="card-desc" style="color:rgba(240,232,222,0.45)">
            Asesoría fiscal<br>laboral y contable
          </div>
          <div class="card-badge">🏗️ En obras</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='text-align:center;color:#FBE0A0;font-size:.75rem;
                letter-spacing:.15em;font-family:Raleway,sans-serif;margin-top:36px;
                font-weight:500;opacity:0.8'>
      Better decisions, together.
    </div>""", unsafe_allow_html=True)
