import streamlit as st
import anthropic
import shutil, os, re, io, json, base64, zipfile
from pathlib import Path
import tempfile
from PIL import Image, ImageDraw, ImageFont

HERE     = Path(__file__).parent
def _find(f):
    for p in [HERE/f, HERE/"assets"/f]:
        if p.exists(): return p
    return HERE/f

PLANTILLA = _find("plantilla_2.pptx")
AILERONS  = _find("Ailerons-Typeface.otf")
LOGO_B    = _find("logo_blanco.png")
FAVICON   = _find("logo1.png")
PASSWORD  = "Sales2026@"

SERVICES_HC = {
    "headhunting": "Headhunting",
    "outsourcing":  "Outsourcing Time & Material",
    "rpo":          "RPO / Equipo embebido",
    "salary":       "Consultoría Salarial",
}

# ─── AILERONS PNG ─────────────────────────────────────────────────────────────
def make_ailerons_png(text, color=(255,255,255), size=64, width=520, height=90):
    img = Image.new("RGBA", (width, height), (0,0,0,0))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(str(AILERONS), size)
    except:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0,0), text, font=font)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    draw.text(((width-tw)//2, (height-th)//2 - bbox[1]), text, font=font, fill=color)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode()

def logo_b64():
    try: return base64.b64encode(LOGO_B.read_bytes()).decode()
    except: return ""

# ─── XML REPLACE — CONSOLIDATES RUNS ──────────────────────────────────────────
def replace_text_in_xml(xml: str, old: str, new: str) -> str:
    """
    Replace text that may be split across multiple <a:t> runs in a paragraph.
    Strategy: for each <a:p>, consolidate all <a:t> text, check if old text
    is present, then replace by putting all text in the first run and clearing others.
    """
    new_esc = new.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    old_esc = old.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    def process_para(m):
        para = m.group(0)
        # Get all <a:t> contents in this paragraph
        runs_text = re.findall(r'<a:t[^>]*>(.*?)</a:t>', para, re.DOTALL)
        full_text = ''.join(runs_text)

        # Check both escaped and unescaped versions
        if old_esc in full_text:
            new_full = full_text.replace(old_esc, new_esc)
        elif old in full_text:
            new_full = full_text.replace(old, new_esc)
        else:
            return para  # no match, return unchanged

        # Put the new text in the first <a:t> tag, clear the rest
        count = [0]
        def replace_first(m2):
            count[0] += 1
            if count[0] == 1:
                return m2.group(1) + new_full + m2.group(2)
            else:
                return m2.group(1) + '' + m2.group(2)

        result = re.sub(r'(<a:t[^>]*>)(.*?)(</a:t>)',
                        lambda m2: (m2.group(1) + (new_full if count[0] == 0 else '') + m2.group(3)) if (count.__setitem__(0, count[0]+1) or True) and count[0] == 1 else m2.group(1) + '' + m2.group(3),
                        para, flags=re.DOTALL)

        # Simpler approach: replace all a:t content
        count2 = [0]
        def repl(m2):
            count2[0] += 1
            if count2[0] == 1:
                return f'{m2.group(1)}{new_full}{m2.group(2)}'
            else:
                return f'{m2.group(1)}{m2.group(2)}'
        result = re.sub(r'(<a:t[^>]*>)([^<]*)(</a:t>)', repl, para)
        return result

    return re.sub(r'<a:p\b[^>]*>.*?</a:p>', process_para, xml, flags=re.DOTALL)

# ─── AUTH ─────────────────────────────────────────────────────────────────────
def check_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        tessaix_b64 = make_ailerons_png("TESSAIX", color=(255,255,255), size=100, width=700, height=130)
        st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800&display=swap');
        html,body,[class*="css"]{{font-family:'Raleway',sans-serif!important}}
        [data-testid="stHeader"]{{display:none}}
        [data-testid="stAppViewContainer"]{{background:#202031!important}}
        .stApp{{background:#202031!important;min-height:100vh}}
        .block-container{{padding-top:0!important;position:relative;z-index:10}}
        .login-wrap{{max-width:420px;margin:0 auto;padding-top:16vh;text-align:center;position:relative;z-index:10}}
        .tagline{{color:#ffffff;font-size:.68rem;letter-spacing:.22em;text-transform:uppercase;
                  margin-top:14px;margin-bottom:52px;font-family:'Raleway',sans-serif;font-weight:400;opacity:0.55}}
        .neb{{position:fixed;border-radius:50%;filter:blur(80px);pointer-events:none;z-index:1}}
        .neb1{{width:600px;height:500px;top:-80px;right:-100px;background:rgba(88,117,121,0.5);
               animation:orb1 5s ease-in-out infinite alternate}}
        .neb2{{width:500px;height:400px;bottom:-60px;left:-80px;background:rgba(88,117,121,0.35);
               animation:orb2 6s ease-in-out infinite alternate}}
        .neb3{{width:380px;height:320px;bottom:80px;right:60px;background:rgba(251,224,160,0.22);
               animation:orb3 7s ease-in-out infinite alternate}}
        .neb4{{width:280px;height:240px;top:120px;left:80px;background:rgba(251,224,160,0.14);
               animation:orb4 4s ease-in-out infinite alternate}}
        @keyframes orb1{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(-80px,60px) scale(1.2)}}}}
        @keyframes orb2{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(70px,-70px) scale(1.25)}}}}
        @keyframes orb3{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(-60px,50px) scale(0.85)}}}}
        @keyframes orb4{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(50px,60px) scale(1.15)}}}}
        div.stButton>button{{background:transparent!important;color:#587579!important;
          border:1.5px solid #587579!important;font-family:'Raleway',sans-serif!important;
          font-weight:700!important;border-radius:6px!important;width:100%!important;
          padding:14px!important;font-size:.95rem!important;margin-top:10px;letter-spacing:.08em!important;
          position:relative;z-index:10}}
        div.stButton>button:hover{{background:#587579!important;color:white!important}}
        .stTextInput input{{background:rgba(42,45,69,0.85)!important;border:1.5px solid #3a3d55!important;
          color:white!important;border-radius:6px!important;font-family:'Raleway',sans-serif!important;
          padding:14px!important;position:relative;z-index:10}}
        .stTextInput input::placeholder{{color:#666!important}}
        label{{color:#666!important;font-size:0!important;line-height:0!important}}
        </style>
        <div class="neb neb1"></div><div class="neb neb2"></div>
        <div class="neb neb3"></div><div class="neb neb4"></div>
        <div class="login-wrap">
          <img src="data:image/png;base64,{tessaix_b64}" style="width:380px">
          <div class="tagline">Propuestas comerciales inteligentes</div>
        </div>""", unsafe_allow_html=True)
        col1,col2,col3 = st.columns([1,2,1])
        with col2:
            pwd = st.text_input("pwd", type="password", placeholder="Contraseña de acceso",
                                label_visibility="collapsed")
            if st.button("Entrar →", use_container_width=True):
                if pwd == PASSWORD:
                    st.session_state.authenticated = True; st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        st.stop()

# ─── CSS INTERIOR ─────────────────────────────────────────────────────────────
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
    .section-divider{border:none;border-top:1px solid #e8e5e0;margin:24px 0}
    </style>""", unsafe_allow_html=True)

def render_header(section="HUMAN CAPITAL"):
    tessera_hdr = make_ailerons_png("TESSERA", color=(255,255,255), size=22, width=200, height=36)
    tessaix_hdr = make_ailerons_png("TESSAIX", color=(251,224,160), size=22, width=180, height=36)
    st.markdown(f"""
    <div class="hdr">
      <div style="display:flex;align-items:center;gap:12px">
        <img src="data:image/png;base64,{tessera_hdr}" style="height:20px;opacity:0.85">
        <span style="color:#3a3d55;font-size:.9rem">|</span>
        <img src="data:image/png;base64,{tessaix_hdr}" style="height:20px">
      </div>
      <span class="hb">{section}</span>
    </div>""", unsafe_allow_html=True)

STEPS = ["Config","Cliente","Servicios","Presentador","Contexto","Generar"]

def render_steps(current):
    cols = st.columns(len(STEPS))
    for i,(col,label) in enumerate(zip(cols,STEPS)):
        col_c = "#587579" if i<=current else "#e0ddd8"
        t_c   = "#587579" if i==current else ("#aaa" if i<current else "#ccc")
        w     = "700" if i==current else "400"
        with col:
            st.markdown(f'<div style="height:2px;background:{col_c};border-radius:1px;margin-bottom:5px"></div>'
                        f'<div style="font-size:.6rem;font-weight:{w};color:{t_c};text-align:center;'
                        f'text-transform:uppercase;letter-spacing:.06em">{label}</div>',
                        unsafe_allow_html=True)

# ─── HOME ─────────────────────────────────────────────────────────────────────
def render_home():
    tessaix_b64 = make_ailerons_png("TESSAIX", color=(255,255,255), size=58, width=520, height=84)

    # Load all logos as base64 data URIs
    def _img_b64(path, svg=False):
        try:
            with open(str(path),'rb') as f: data = f.read()
            mime = "image/svg+xml" if svg else "image/png"
            return f"data:{mime};base64,{base64.b64encode(data).decode()}"
        except: return ""

    hc_src  = _img_b64(_find("logo_tessera_human_capital_.svg"), svg=True)
    fi_src  = _img_b64(_find("tessera_finance_logo.svg"), svg=True)
    cd_src  = _img_b64(_find("costdown_logo.png"))
    lt_src  = _img_b64(_find("logo_LT_.svg"), svg=True)

    # HC logo for header — white version via PIL (SVG is navy, need white for dark bg)
    tessera_header = make_ailerons_png("TESSERA", color=(255,255,255), size=26, width=240, height=40)

    # Build logo HTML snippets BEFORE the f-string
    hc_img = f'<img src="{hc_src}" style="height:38px;margin-bottom:12px;filter:brightness(0) invert(1);opacity:.9">' if hc_src else '<div style="font-size:26px;margin-bottom:12px">👥</div>'
    fi_img = f'<img src="{fi_src}" style="height:38px;margin-bottom:12px;filter:brightness(0) invert(1);opacity:.4">' if fi_src else '<div style="font-size:26px;margin-bottom:12px;opacity:.4">🏦</div>'
    cd_img = f'<img src="{cd_src}" style="height:28px;margin-bottom:12px;opacity:.35;filter:brightness(0) invert(1)">' if cd_src else '<div style="font-size:26px;margin-bottom:12px;opacity:.4">💰</div>'
    lt_img = f'<img src="{lt_src}" style="height:24px;margin-bottom:12px;opacity:.35;filter:brightness(0) invert(1)">' if lt_src else '<div style="font-size:26px;margin-bottom:12px;opacity:.4">📊</div>'

    # CSS block
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800;900&display=swap');
    [data-testid="stAppViewContainer"]{background:#202031!important}
    [data-testid="stHeader"]{display:none}
    .stApp{background:#202031!important}
    .block-container{padding-top:0!important}
    .neb{position:fixed;border-radius:50%;filter:blur(90px);pointer-events:none;z-index:0}
    .neb1{width:550px;height:450px;top:-60px;right:-80px;background:rgba(88,117,121,0.4);animation:orb1 5s ease-in-out infinite alternate}
    .neb2{width:450px;height:380px;bottom:-40px;left:-60px;background:rgba(88,117,121,0.28);animation:orb2 6s ease-in-out infinite alternate}
    .neb3{width:320px;height:280px;bottom:100px;right:80px;background:rgba(251,224,160,0.18);animation:orb3 7s ease-in-out infinite alternate}
    .neb4{width:240px;height:210px;top:140px;left:60px;background:rgba(251,224,160,0.12);animation:orb4 4s ease-in-out infinite alternate}
    @keyframes orb1{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-80px,60px) scale(1.2)}}
    @keyframes orb2{0%{transform:translate(0,0) scale(1)}100%{transform:translate(70px,-70px) scale(1.25)}}
    @keyframes orb3{0%{transform:translate(0,0) scale(1)}100%{transform:translate(-60px,50px) scale(0.85)}}
    @keyframes orb4{0%{transform:translate(0,0) scale(1)}100%{transform:translate(50px,60px) scale(1.15)}}
    .home-wrap{position:relative;z-index:2;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:40px 20px}
    .card-grid{display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin:36px 0 24px}
    .card{position:relative;overflow:hidden;background:rgba(255,255,255,0.04);
      border:1.5px solid rgba(255,255,255,0.1);border-radius:14px;padding:28px 24px;
      width:200px;text-align:center;cursor:default;
      transition:border-color .3s, transform .3s;backdrop-filter:blur(4px)}
    .card::before{content:'';position:absolute;inset:0;border-radius:14px;opacity:0;transition:opacity .4s;z-index:0}
    .card:hover{border-color:rgba(88,117,121,0.8);transform:translateY(-3px)}
    .card:hover::before{opacity:1}
    .card-hc::before{background:radial-gradient(ellipse 80% 80% at 50% 50%,rgba(88,117,121,0.35) 0%,transparent 70%)}
    .card-fi::before{background:radial-gradient(ellipse 80% 80% at 50% 50%,rgba(88,117,121,0.2) 0%,transparent 70%)}
    .card-cd::before{background:radial-gradient(ellipse 80% 80% at 50% 50%,rgba(251,224,160,0.2) 0%,transparent 70%)}
    .card-lt::before{background:radial-gradient(ellipse 80% 80% at 50% 50%,rgba(42,61,101,0.4) 0%,transparent 70%)}
    .card-content{position:relative;z-index:1}
    .card-desc{font-size:.68rem;line-height:1.6;font-family:'Raleway',sans-serif}
    .card-title{font-size:.75rem;font-weight:700;letter-spacing:.06em;margin-bottom:6px;font-family:'Raleway',sans-serif}
    .card-badge{display:inline-block;margin-top:10px;font-size:.6rem;font-weight:700;
      letter-spacing:.1em;text-transform:uppercase;
      background:rgba(251,224,160,0.15);color:#FBE0A0;border-radius:3px;padding:2px 8px}
    .bottom-txt{color:rgba(240,232,222,0.3);font-size:.65rem;letter-spacing:.15em;
      font-family:'Raleway',sans-serif;text-align:center;margin-top:8px}
    div.stButton>button{background:rgba(88,117,121,0.2)!important;color:#587579!important;
      border:1.5px solid #587579!important;font-family:'Raleway',sans-serif!important;
      font-weight:700!important;border-radius:8px!important;padding:12px 32px!important;
      font-size:.85rem!important;letter-spacing:.08em!important;transition:all .2s!important}
    div.stButton>button:hover{background:#587579!important;color:white!important}
    </style>
    <div class="neb neb1"></div><div class="neb neb2"></div>
    <div class="neb neb3"></div><div class="neb neb4"></div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="home-wrap">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:32px">
        <img src="data:image/png;base64,{tessera_header}" style="height:24px;opacity:0.75">
        <span style="color:rgba(255,255,255,0.18);font-size:.9rem">|</span>
        <img src="data:image/png;base64,{tessaix_b64}" style="height:26px">
      </div>

      <div style="color:rgba(255,255,255,0.4);font-size:.67rem;letter-spacing:.2em;
                  text-transform:uppercase;font-family:'Raleway',sans-serif;margin-bottom:40px">
        Propuestas comerciales inteligentes
      </div>

      <div class="card-grid">

        <div class="card card-hc">
          <div class="card-content">
            {hc_img}
            <div class="card-desc" style="color:rgba(240,232,222,0.65)">Recruitment · HR Advisory<br>Outsourcing · RPO</div>
          </div>
        </div>

        <div class="card card-fi">
          <div class="card-content">
            {fi_img}
            <div class="card-desc" style="color:rgba(240,232,222,0.3)">CFO · M&amp;A<br>Due Diligence</div>
            <div class="card-badge">🏗️ En obras</div>
          </div>
        </div>

        <div class="card card-cd">
          <div class="card-content">
            {cd_img}
            <div class="card-title" style="color:rgba(240,232,222,0.45);font-size:.75rem;margin-bottom:6px">Cost Down</div>
            <div class="card-desc" style="color:rgba(240,232,222,0.25)">Reducción de costes<br>operativos</div>
            <div class="card-badge">🏗️ En obras</div>
          </div>
        </div>

        <div class="card card-lt">
          <div class="card-content">
            {lt_img}
            <div class="card-title" style="color:rgba(240,232,222,0.45);font-size:.75rem;margin-bottom:6px">LT Impulsa</div>
            <div class="card-desc" style="color:rgba(240,232,222,0.25)">Asesoría fiscal<br>laboral y contable</div>
            <div class="card-badge">🏗️ En obras</div>
          </div>
        </div>

      </div>

      <div class="bottom-txt">Better decisions, together.</div>
    </div>
    """, unsafe_allow_html=True)

    col1,col2,col3 = st.columns([1,2,1])
    with col2:
        if st.button("Entrar a Human Capital →", use_container_width=True):
            st.session_state.screen = "hc"; st.rerun()


# ─── AI CONTENT ───────────────────────────────────────────────────────────────
def generate_content(data):
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    is_en  = data["lang"] == "en"
    svcs   = ", ".join([SERVICES_HC.get(s,s) for s in data["services"]])
    cn, sec = data["client_name"], data["sector"]

    system = """Eres consultor comercial senior de Tessera Human Capital (tesseraservices.com).
Generas contenido para propuestas comerciales en PowerPoint. Devuelves ÚNICAMENTE JSON válido.
REGLAS ABSOLUTAS:
- Sin guiones largos
- Nunca uses "contexto" ni "criterio"
- Español de España natural, directo, tono de negocio cercano
- ESPECÍFICO al sector del cliente — nada de mencionar "publicidad" si el cliente es de otro sector
- Los textos tienen que REEMPLAZAR los existentes completamente, sin mezclar con el original
- Cada texto debe ser completo y autónomo, listo para aparecer en el slide tal cual""" if not is_en else \
    """Senior commercial consultant at Tessera Human Capital. Return ONLY valid JSON.
ABSOLUTE RULES: no em dashes, sector-specific content only, complete standalone texts."""

    prompt = f"""Genera JSON para propuesta comercial de Tessera Human Capital para:
CLIENTE: {cn}
SECTOR: {sec}
PAÍS: {data.get('country','España')}
SERVICIOS: {svcs}
PAIN POINTS: {data.get('pain_points','')}
PERFILES NECESARIOS: {data.get('roles_needed','')}
FEE: {data.get('fee_rate','XX')}%
INFO ADICIONAL: {data.get('extra_info','')}

IMPORTANTE: Todos los textos deben ser específicos para el sector "{sec}". 
NO mencionar publicidad, AdTech, programmatic ni nada relacionado con marketing digital 
a no ser que el cliente sea de ese sector.

Devuelve este JSON con contenido REAL y COMPLETO (no placeholders):
{{
  "why_tessera": {{
    "d1_title": "título diferenciador 1 específico para {sec} (máx 4 palabras)",
    "d1_body": "descripción 2-3 líneas específica para {sec}. Sin mencionar otros sectores.",
    "d2_title": "título diferenciador 2 específico para {sec} (máx 4 palabras)", 
    "d2_body": "descripción 2-3 líneas específica para {sec}",
    "d3_title": "título diferenciador 3 específico para {sec} (máx 4 palabras)",
    "d3_body": "descripción 2-3 líneas específica para {sec}"
  }},
  "services": {{
    {"\"headhunting\": {\"why_col_title1\": \"TÍTULO RAZÓN 1\", \"why_col_body1\": \"descripción específica {sec}\", \"why_col_title2\": \"TÍTULO RAZÓN 2\", \"why_col_body2\": \"descripción específica {sec}\", \"why_col_title3\": \"TÍTULO RAZÓN 3\", \"why_col_body3\": \"descripción específica {sec}\", \"benefits_title1\": \"CERO ESTRÉS\", \"benefits_body1\": \"descripción específica {sec}\", \"benefits_title2\": \"TALENTO A MEDIDA\", \"benefits_body2\": \"descripción específica {sec}\", \"benefits_title3\": \"RELACIÓN A LARGO PLAZO\", \"benefits_body3\": \"descripción específica {sec}\", \"how_title1\": \"NOS CONOCEMOS\", \"how_body1\": \"descripción específica {sec}\", \"how_title2\": \"VAMOS MÁS ALLÁ\", \"how_body2\": \"descripción específica {sec}\", \"how_title3\": \"ELEGIMOS CON PRECISIÓN\", \"how_body3\": \"descripción específica {sec}\", \"how_title4\": \"SOLO LO MEJOR\", \"how_body4\": \"descripción específica {sec}\", \"how_title5\": \"ACOMPAÑAMOS EL PROCESO\", \"how_body5\": \"descripción específica {sec}\"}" if "headhunting" in data.get("services",[]) else ""},
    {"\"outsourcing\": {\"headline\": \"frase impacto para {sec}\", \"body\": \"2 líneas para {sec}\", \"card1_title\": \"Coste variable\", \"card1_body\": \"Pagas solo por horas reales de trabajo\", \"card2_title\": \"Cobertura total\", \"card2_body\": \"Cubrimos bajas y vacaciones sin interrupciones\", \"card3_title\": \"Cero gestión\", \"card3_body\": \"Nóminas y trámites a cargo nuestro\", \"card4_title\": \"Flexibilidad\", \"card4_body\": \"Escalas el equipo cuando lo necesites\"}" if "outsourcing" in data.get("services",[]) else ""},
    {"\"rpo\": {\"headline\": \"titular RPO para {sec}\", \"body\": \"2 líneas RPO para {cn}\", \"p1_title\": \"DEDICACIÓN TOTAL\", \"p1_body\": \"descripción para {sec}\", \"p2_title\": \"EFICIENCIA REAL\", \"p2_body\": \"descripción para {sec}\", \"p3_title\": \"FOCO EN CRECER\", \"p3_body\": \"descripción para {sec}\"}" if "rpo" in data.get("services",[]) else ""},
    {"\"salary\": {\"headline\": \"Inteligencia retributiva para {sec}\", \"body\": \"2 líneas para {cn}\"}" if "salary" in data.get("services",[]) else ""}
  }}
}}"""

    with st.spinner("Generando contenido adaptado al cliente..."):
        msg = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=4000,
            system=system, messages=[{"role":"user","content":prompt}]
        )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n",1)[1].rsplit("```",1)[0].strip()
    # Clean empty keys
    raw = re.sub(r',\s*""\s*:', ',"_empty":', raw)
    raw = re.sub(r':\s*,', ': null,', raw)
    try:
        return json.loads(raw)
    except Exception as e:
        st.error(f"Error parseando JSON: {e}")
        st.code(raw[:800])
        return None

# ─── PPTX BUILDER ─────────────────────────────────────────────────────────────
def build_pptx(data, content):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        work = tmp / "work.pptx"
        shutil.copy(str(PLANTILLA), str(work))

        unpack = tmp / "unpacked"
        with zipfile.ZipFile(work, 'r') as z:
            z.extractall(unpack)

        slides = unpack / "ppt" / "slides"
        media  = unpack / "ppt" / "media"
        r = replace_text_in_xml

        # Logo cliente
        domain = (data.get("client_website","") or f"{data['client_name'].lower().replace(' ','')}.com")
        domain = domain.replace("https://","").replace("http://","").replace("www.","").split("/")[0].strip()
        try:
            import requests as rq
            client_logo = None
            for url in [f"https://logo.clearbit.com/{domain}",
                        f"https://www.google.com/s2/favicons?domain={domain}&sz=256"]:
                try:
                    resp = rq.get(url, timeout=6, headers={"User-Agent":"Mozilla/5.0"})
                    if resp.status_code == 200 and len(resp.content) > 500:
                        client_logo = resp.content; break
                except: continue
            if client_logo:
                img4 = media / "image4.png"
                if img4.exists():
                    img4.write_bytes(client_logo)
            else:
                s1_path = slides / "slide1.xml"
                if s1_path.exists():
                    s1 = s1_path.read_text("utf-8")
                    s1 = re.sub(r'<p:pic>\s*(?:(?!</p:pic>).)*?rId6(?:(?!</p:pic>).)*?</p:pic>','',s1,flags=re.DOTALL)
                    s1_path.write_text(s1,"utf-8")
        except: pass

        why  = content.get("why_tessera",{})
        svcs = content.get("services",{})

        presenter_name  = data.get("presenter_name","Manuel García Pina")
        presenter_phone = data.get("presenter_phone","+34 619 511 155")
        presenter_email = data.get("presenter_email","Manuel.garcia@tesseraservices.com")

        # ── SLIDE 1: portada ──────────────────────────────────────────────
        s = (slides/"slide1.xml").read_text("utf-8")
        s = r(s, "Edward Manrique ", f"{presenter_name} ")
        s = r(s, "+34 695021978", presenter_phone)
        (slides/"slide1.xml").write_text(s,"utf-8")

        # ── SLIDE 3: por qué tessera ──────────────────────────────────────
        if why:
            s = (slides/"slide3.xml").read_text("utf-8")
            if why.get("d1_title"): s = r(s,"Sin CVs al azar", why["d1_title"])
            if why.get("d1_body"):  s = r(s,"No enviamos el primer CV, buscamos a quien encaja de verdad. Candidatos filtrados sobre la mesa en 72 horas.", why["d1_body"])
            if why.get("d2_title"): s = r(s,"Sin pausas", why["d2_title"])
            if why.get("d2_body"):  s = r(s,"Tu servicio no parará.", why["d2_body"])
            if why.get("d3_title"): s = r(s,"Siempre contigo", why["d3_title"])
            if why.get("d3_body"):  s = r(s,"No desaparecemos tras la incorporación: cuidamos a la persona y al cliente durante todo el servicio.", why["d3_body"])
            (slides/"slide3.xml").write_text(s,"utf-8")

        # ── SLIDE 5: headhunting ──────────────────────────────────────────
        if "headhunting" in data.get("services",[]) and svcs.get("headhunting"):
            hh = svcs["headhunting"]
            s  = (slides/"slide5.xml").read_text("utf-8")
            # WHY column
            if hh.get("why_col_title1"): s = r(s,"VELOCIDAD ", hh["why_col_title1"]+" ")
            if hh.get("why_col_body1"):  s = r(s,"Encontramos rápido, pero no a cualquiera. Seleccionamos a los que suman valor real dentro de nuestra base de datos especializada en Adtech.", hh["why_col_body1"])
            if hh.get("why_col_title2"): s = r(s,"ESPECIALIZACIÓN EN PUBLICIDAD & MARKETING", hh["why_col_title2"])
            if hh.get("why_col_body2"):  s = r(s,"Conocemos el ecosistema publicitario y cómo se construyen sus equipos. Evaluamos perfiles en base a su experiencia en Sales, Programmatic, Ad Operations, CSM y Marketing, entendiendo las dinámicas reales del sector y el encaje con tu negocio.", hh["why_col_body2"])
            if hh.get("why_col_title3"): s = r(s,"ÉXITO COMPARTIDO", hh["why_col_title3"])
            if hh.get("why_col_body3"):  s = r(s,"No se trata solo de cubrir vacantes, se trata de construir el equipo que te llevará al siguiente nivel.", hh["why_col_body3"])
            # BENEFITS column
            if hh.get("benefits_body1"): s = r(s,"Nos ocupamos de todo el proceso, tú solo conoces a los mejores.", hh["benefits_body1"])
            if hh.get("benefits_body2"): s = r(s,"Candidatos que no solo encajan, sino que impulsan tu crecimiento.", hh["benefits_body2"])
            if hh.get("benefits_body3"): s = r(s,"Porque para nosotros, tu éxito es también el nuestro.", hh["benefits_body3"])
            # HOW column
            if hh.get("how_body1"): s = r(s,"Escuchamos tus necesidades y objetivos desde el primer día.", hh["how_body1"])
            if hh.get("how_body2"): s = r(s,"Buscamos talento con técnicas avanzadas y un enfoque humano.", hh["how_body2"])
            if hh.get("how_body3"): s = r(s,"Evaluamos habilidades, actitud y encaje cultural.", hh["how_body3"])
            if hh.get("how_body4"): s = r(s,"Te presentamos a los candidatos que realmente suman", hh["how_body4"])
            if hh.get("how_body5"): s = r(s,"Estamos contigo incluso después de la incorporación para un correcto encaje en el equipo.", hh["how_body5"])
            (slides/"slide5.xml").write_text(s,"utf-8")

        # ── SLIDE 6: outsourcing ──────────────────────────────────────────
        if "outsourcing" in data.get("services",[]) and svcs.get("outsourcing"):
            outs = svcs["outsourcing"]
            s    = (slides/"slide6.xml").read_text("utf-8")
            if outs.get("headline"): s = r(s,"Externalización que sí funciona: pagas por trabajo real, sin papeleo.", outs["headline"])
            if outs.get("body"):     s = r(s,"Contratar cuesta más de lo que parece. Con el modelo Time & Material ajustas el equipo a tu actividad real y nosotros nos encargamos de toda la gestión.", outs["body"])
            if outs.get("card1_body"): s = r(s,"Pagas solo por horas reales de trabajo.", outs["card1_body"])
            if outs.get("card2_body"): s = r(s,"Cubrimos bajas y vacaciones.", outs["card2_body"])
            if outs.get("card3_body"): s = r(s,"Nóminas y trámites, a cargo nuestro.", outs["card3_body"])
            if outs.get("card4_body"): s = r(s,"Escalas el equipo cuando quieras.", outs["card4_body"])
            (slides/"slide6.xml").write_text(s,"utf-8")

        # ── SLIDE 7: RPO ──────────────────────────────────────────────────
        if "rpo" in data.get("services",[]) and svcs.get("rpo"):
            rpo = svcs["rpo"]
            s   = (slides/"slide7.xml").read_text("utf-8")
            if rpo.get("headline"): s = r(s,"RPO a tu medida", rpo["headline"])
            if rpo.get("body"):     s = r(s,"Te ofrecemos un equipo de recruiters que se integra en tu compañía como una extensión real de tu equipo, trabajando exclusivamente en tus necesidades durante el tiempo que lo necesites.\nNo solo ejecutamos procesos: entendemos tu negocio, tus retos y tu historia  para atraer el talento que realmente necesitas.", rpo["body"])
            if rpo.get("p1_title"): s = r(s,"DEDICACIÓN TOTAL", rpo["p1_title"])
            if rpo.get("p1_body"):  s = r(s,"Un equipo dedicado en exclusiva a tu compañía, con conocimiento del sector publicitario y alineado c", rpo["p1_body"][:90])
            if rpo.get("p2_title"): s = r(s,"EFICIENCIA REAL", rpo["p2_title"])
            if rpo.get("p2_body"):  s = r(s,"Acceso directo a una red de +12.000 profesionales de publicidad ya identificados y validados. Menos ", rpo["p2_body"][:90])
            if rpo.get("p3_title"): s = r(s,"FOCO EN CRECER", rpo["p3_title"])
            if rpo.get("p3_body"):  s = r(s,"Nosotros buscamos, filtramos y validamos. Tú te concentras en hacer crecer tu equipo y tu negocio.", rpo["p3_body"][:90])
            (slides/"slide7.xml").write_text(s,"utf-8")

        # ── SLIDE 15: fees ────────────────────────────────────────────────
        if data.get("deck_type")=="propuesta" and data.get("fee_rate"):
            s  = (slides/"slide15.xml").read_text("utf-8")
            fr = data.get("fee_rate","XX")
            s  = r(s,"18%",f"{fr}%")
            s  = r(s,"13%",f"{fr}%")
            (slides/"slide15.xml").write_text(s,"utf-8")

        # ── SLIDE 16: cierre ──────────────────────────────────────────────
        s = (slides/"slide16.xml").read_text("utf-8")
        s = r(s,"Edward@tesseraservices.com", presenter_email)
        (slides/"slide16.xml").write_text(s,"utf-8")

        # Pack back to PPTX
        out = tmp / "output.pptx"
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fp in sorted(unpack.rglob('*')):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(unpack))
        return out.read_bytes()

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="TESSAIX",
        page_icon=str(FAVICON) if FAVICON.exists() else "T",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    check_auth()
    inject_css()

    for k,v in [("screen","home"),("step",0),("form",{}),("content",None),("pptx",None)]:
        if k not in st.session_state: st.session_state[k]=v

    if st.session_state.screen == "home":
        render_home(); return

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
            st.session_state.screen="home"; st.rerun()
        return

    render_header("HUMAN CAPITAL")
    step = st.session_state.step
    render_steps(step)

    # STEP 0: Config
    if step == 0:
        st.markdown("### ¿Qué tipo de propuesta?")
        c1,c2 = st.columns(2)
        with c1: lang = st.radio("Idioma", ["Español","English"])
        with c2: deck_type = st.radio("Tipo", ["Presentación (sin fees)","Propuesta completa (con fees)"])
        st.markdown("---")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Inicio"): st.session_state.screen="home"; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                st.session_state.form.update({
                    "lang":"es" if lang=="Español" else "en",
                    "deck_type":"presentacion" if "sin fees" in deck_type else "propuesta"
                })
                st.session_state.step=1; st.rerun()

    # STEP 1: Cliente
    elif step == 1:
        st.markdown("### Datos del cliente")
        c1,c2 = st.columns(2)
        with c1:
            cn  = st.text_input("Nombre de la empresa *", placeholder="Ej: Alimerka, Corpay…")
            sec = st.text_input("Sector / industria *",   placeholder="Ej: Alimentación, Fintech…")
            ctr = st.text_input("País / mercado", value="España")
        with c2:
            sz  = st.text_input("Tamaño de empresa",  placeholder="Ej: 50–200 empleados")
            web = st.text_input("Web del cliente",    placeholder="alimerka.es")
            ccn = st.text_input("Nombre del contacto", placeholder="Nombre y apellido")
            ccr = st.text_input("Cargo del contacto", placeholder="Ej: HR Director")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=0; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                if not cn.strip() or not sec.strip():
                    st.error("Empresa y sector obligatorios.")
                else:
                    st.session_state.form.update({
                        "client_name":cn,"sector":sec,"country":ctr,
                        "size":sz,"client_website":web,
                        "contact_name":ccn,"contact_role":ccr
                    })
                    st.session_state.step=2; st.rerun()

    # STEP 2: Servicios
    elif step == 2:
        st.markdown("### ¿Qué servicios incluimos?")
        selected = st.multiselect("Servicios", list(SERVICES_HC.keys()),
                                   default=["headhunting"], format_func=lambda x: SERVICES_HC[x])
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=1; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                if not selected: st.error("Selecciona al menos uno.")
                else:
                    st.session_state.form["services"]=selected
                    st.session_state.step=3; st.rerun()

    # STEP 3: Presentador ← NUEVO
    elif step == 3:
        st.markdown("### ¿Quién presenta esta propuesta?")
        st.caption("Esta información aparecerá en la portada y en el slide de cierre.")
        c1,c2 = st.columns(2)
        with c1:
            pname  = st.text_input("Nombre completo *", value="Manuel García Pina",
                                    placeholder="Manuel García Pina")
            pemail = st.text_input("Email *", value="Manuel.garcia@tesseraservices.com",
                                    placeholder="email@tesseraservices.com")
        with c2:
            pphone = st.text_input("Teléfono *", value="+34 619 511 155",
                                    placeholder="+34 6XX XXX XXX")
            _      = st.text_input("Cargo", value="Partner",
                                    placeholder="Partner / Head of Recruitment…")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=2; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                if not pname.strip() or not pphone.strip():
                    st.error("Nombre y teléfono obligatorios.")
                else:
                    st.session_state.form.update({
                        "presenter_name":pname,
                        "presenter_phone":pphone,
                        "presenter_email":pemail
                    })
                    st.session_state.step=4; st.rerun()

    # STEP 4: Contexto
    elif step == 4:
        st.markdown("### Personalización")
        pp  = st.text_area("Pain points / retos del cliente", height=80,
                            placeholder="Ej: Crecimiento rápido, rotación alta en tienda, necesitan perfiles con vocación…")
        rn  = st.text_area("Perfiles que necesitan cubrir", height=60,
                            placeholder="Ej: Jefes de sección, reponedores, 1 HR Business Partner…")
        fee = ""
        if st.session_state.form.get("deck_type")=="propuesta":
            fee = st.text_input("Fee propuesto (%)", placeholder="Ej: 16")
        ei  = st.text_area("Info adicional (web, reunión previa, LinkedIn…)", height=90,
                            placeholder="Pega aquí todo lo que sepas del cliente…")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=3; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                st.session_state.form.update({
                    "pain_points":pp,"roles_needed":rn,
                    "fee_rate":fee,"extra_info":ei
                })
                st.session_state.step=5; st.rerun()

    # STEP 5: Generar
    elif step == 5:
        d = st.session_state.form
        st.markdown("### Revisa y genera")
        rows = [
            ("Cliente", d.get("client_name","—")),
            ("Sector", d.get("sector","—")),
            ("Idioma", "Español" if d.get("lang")=="es" else "English"),
            ("Tipo", "Con fees" if d.get("deck_type")=="propuesta" else "Sin fees"),
            ("Servicios", " · ".join([SERVICES_HC.get(s,s) for s in d.get("services",[])])),
            ("Presentado por", d.get("presenter_name","—")),
            ("Teléfono", d.get("presenter_phone","—")),
        ]
        if d.get("fee_rate"): rows.append(("Fee", f"{d['fee_rate']}%"))
        html_r = "".join(f'<div class="sr"><span class="sk">{k}</span><span class="sv">{v}</span></div>'
                          for k,v in rows)
        st.markdown(f'<div class="sum">{html_r}</div>', unsafe_allow_html=True)

        if st.session_state.pptx is None:
            cb,cg = st.columns([1,4])
            with cb:
                if st.button("← Atrás"): st.session_state.step=4; st.rerun()
            with cg:
                if st.button("✦ Generar propuesta"):
                    c = generate_content(d)
                    if c:
                        st.session_state.content = c
                        with st.spinner("Montando el PowerPoint sobre la plantilla de Tessera…"):
                            try:
                                st.session_state.pptx = build_pptx(d, c)
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error generando PPTX: {e}")
                                import traceback
                                st.code(traceback.format_exc())
        else:
            st.markdown('<div class="ok">✅ <strong>Propuesta generada.</strong> Descarga el PowerPoint con la plantilla visual de Tessera.</div>',
                        unsafe_allow_html=True)
            safe = d.get("client_name","Cliente").replace(" ","_")
            st.download_button(
                "⬇ Descargar PowerPoint",
                data=st.session_state.pptx,
                file_name=f"Tessera_{safe}_Propuesta.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            if st.button("+ Nueva propuesta"):
                for k,v in [("step",0),("form",{}),("content",None),("pptx",None)]:
                    st.session_state[k]=v
                st.rerun()

if __name__ == "__main__":
    main()
