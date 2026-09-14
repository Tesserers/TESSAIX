"""CSS compartido, cabecera de sección y barra de pasos del wizard."""
import streamlit as st

from .assets import make_ailerons_png
from .config import STEPS

LOGIN_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Raleway',sans-serif!important}
[data-testid="stHeader"]{display:none}
[data-testid="stAppViewContainer"]{background:#202031!important}
.stApp{background:#202031!important;min-height:100vh}
.block-container{padding-top:0!important;position:relative;z-index:10}
.login-wrap{max-width:420px;margin:0 auto;padding-top:16vh;text-align:center;position:relative;z-index:10}
.tagline{color:#ffffff;font-size:.68rem;letter-spacing:.22em;text-transform:uppercase;
          margin-top:14px;margin-bottom:52px;font-family:'Raleway',sans-serif;font-weight:400;opacity:0.55}
.neb{position:fixed;border-radius:50%;filter:blur(80px);pointer-events:none;z-index:1}
.neb1{width:600px;height:500px;top:-80px;right:-100px;background:rgba(88,117,121,0.5);
       animation:orb1 5s ease-in-out infinite alternate}
.neb2{width:500px;height:400px;bottom:-60px;left:-80px;background:rgba(88,117,121,0.35);
       animation:orb2 6s ease-in-out infinite alternate}
.neb3{width:380px;height:320px;bottom:80px;right:60px;background:rgba(251,224,160,0.22);
       animation:orb3 7s ease-in-out infinite alternate}
.neb4{width:280px;height:240px;top:120px;left:80px;background:rgba(251,224,160,0.14);
       animation:orb4 4s ease-in-out infinite alternate}
@keyframes orb1{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-80px,60px) scale(1.2)}}
@keyframes orb2{0%{transform:translate(0,0) scale(1)}100%{transform:translate(70px,-70px) scale(1.25)}}
@keyframes orb3{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-60px,50px) scale(0.85)}}
@keyframes orb4{0%{transform:translate(0,0) scale(1)}100%{transform:translate(50px,60px) scale(1.15)}}
div.stButton>button{background:transparent!important;color:#587579!important;
  border:1.5px solid #587579!important;font-family:'Raleway',sans-serif!important;
  font-weight:700!important;border-radius:6px!important;width:100%!important;
  padding:14px!important;font-size:.95rem!important;margin-top:10px;letter-spacing:.08em!important;
  position:relative;z-index:10}
div.stButton>button:hover{background:#587579!important;color:white!important}
.stTextInput input{background:rgba(42,45,69,0.85)!important;border:1.5px solid #3a3d55!important;
  color:white!important;border-radius:6px!important;font-family:'Raleway',sans-serif!important;
  padding:14px!important;position:relative;z-index:10}
.stTextInput input::placeholder{color:#666!important}
label{color:#666!important;font-size:0!important;line-height:0!important}
</style>
<div class="neb neb1"></div><div class="neb neb2"></div>
<div class="neb neb3"></div><div class="neb neb4"></div>
"""


def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800;900&family=Manrope:wght@400;700&display=swap');
    html,body,[class*="css"]{font-family:'Raleway',sans-serif!important}
    [data-testid="stHeader"]{display:none}
    [data-testid="stAppViewContainer"]{background:#fafaf8}
    .hdr{background:#202031;padding:14px 28px;display:flex;align-items:center;
          justify-content:space-between;margin:-1rem -1rem 2rem -1rem}
    .hb{background:#58757922;color:#587579;border-radius:4px;
         padding:3px 12px;font-size:.68rem;font-weight:700;letter-spacing:.12em}
    div.stButton>button{background:#202031!important;color:white!important;
      border:none!important;font-family:'Raleway',sans-serif!important;
      font-weight:700!important;border-radius:6px!important;width:100%!important}
    div.stButton>button:hover{background:#587579!important}
    .stTextInput input,.stTextArea textarea{font-family:'Raleway',sans-serif!important;
      border:1.5px solid #e0ddd8!important;border-radius:6px!important}
    label{font-size:.72rem!important;font-weight:700!important;color:#999!important;
          text-transform:uppercase!important;letter-spacing:.08em!important}
    .sum{background:white;border:1.5px solid #e8e5e0;border-radius:10px;padding:20px 24px;margin-bottom:20px}
    .sr{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #f0ede8;font-size:.83rem}
    .sk{color:#bbb;font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em}
    .sv{color:#202031;font-weight:600}
    .ok{background:#58757910;border:1px solid #58757930;border-radius:8px;padding:14px 18px;margin-bottom:16px}
    .warn{background:#f9a82510;border:1px solid #f9a82550;border-radius:8px;padding:14px 18px;margin-bottom:16px}
    .err{background:#e5484d10;border:1px solid #e5484d50;border-radius:8px;padding:14px 18px;margin-bottom:16px}
    .section-divider{border:none;border-top:1px solid #e8e5e0;margin:24px 0}
    </style>""", unsafe_allow_html=True)


def render_header(section="HUMAN CAPITAL"):
    tessera_hdr = make_ailerons_png("TESSERA", color=(255, 255, 255), size=22, width=200, height=36)
    tessaix_hdr = make_ailerons_png("TESSAIX", color=(251, 224, 160), size=22, width=180, height=36)
    st.markdown(f"""
    <div class="hdr">
      <div style="display:flex;align-items:center;gap:12px">
        <img src="data:image/png;base64,{tessera_hdr}" style="height:20px;opacity:0.85">
        <span style="color:#3a3d55;font-size:.9rem">|</span>
        <img src="data:image/png;base64,{tessaix_hdr}" style="height:20px">
      </div>
      <span class="hb">{section}</span>
    </div>""", unsafe_allow_html=True)


def render_steps(current):
    cols = st.columns(len(STEPS))
    for i, (col, label) in enumerate(zip(cols, STEPS)):
        col_c = "#587579" if i <= current else "#e0ddd8"
        t_c = "#587579" if i == current else ("#aaa" if i < current else "#ccc")
        w = "700" if i == current else "400"
        with col:
            st.markdown(f'<div style="height:2px;background:{col_c};border-radius:1px;margin-bottom:5px"></div>'
                        f'<div style="font-size:.6rem;font-weight:{w};color:{t_c};text-align:center;'
                        f'text-transform:uppercase;letter-spacing:.06em">{label}</div>',
                        unsafe_allow_html=True)
