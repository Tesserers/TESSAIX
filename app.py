import streamlit as st
import anthropic
import shutil, os, re, io, json, base64
from pathlib import Path
import subprocess, tempfile
from PIL import Image, ImageDraw, ImageFont

HERE      = Path(__file__).parent
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

# ─── AILERONS PNG via PIL ──────────────────────────────────────────────────────
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
    try:
        return base64.b64encode(LOGO_B.read_bytes()).decode()
    except:
        return ""

# ─── AUTH ─────────────────────────────────────────────────────────────────────
def check_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        tessaix_b64 = make_ailerons_png("TESSAIX", color=(255,255,255), size=80, width=600, height=110)
        st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Raleway:wght@300;400;600;700;800&display=swap');
        html,body,[class*="css"]{{font-family:'Raleway',sans-serif!important}}
        [data-testid="stHeader"]{{display:none}}
        [data-testid="stAppViewContainer"]{{background:#202031!important}}
        .stApp{{background:#202031!important;min-height:100vh}}
        .block-container{{padding-top:0!important;position:relative;z-index:10}}
        .login-wrap{{max-width:380px;margin:0 auto;padding-top:18vh;text-align:center;position:relative;z-index:10}}
        .tagline{{color:#587579;font-size:.68rem;letter-spacing:.22em;text-transform:uppercase;
                  margin-top:14px;margin-bottom:52px;font-family:'Raleway',sans-serif;font-weight:600}}
        /* Nebula orbs — posición fija, animadas independientemente, DETRÁS de todo */
        .neb{{position:fixed;border-radius:50%;filter:blur(80px);pointer-events:none;z-index:1}}
        .neb1{{width:600px;height:500px;top:-80px;right:-100px;
               background:rgba(88,117,121,0.45);
               animation:orb1 14s ease-in-out infinite alternate}}
        .neb2{{width:500px;height:400px;bottom:-60px;left:-80px;
               background:rgba(88,117,121,0.30);
               animation:orb2 18s ease-in-out infinite alternate}}
        .neb3{{width:350px;height:300px;bottom:80px;right:60px;
               background:rgba(251,224,160,0.18);
               animation:orb3 22s ease-in-out infinite alternate}}
        .neb4{{width:250px;height:220px;top:120px;left:80px;
               background:rgba(251,224,160,0.12);
               animation:orb4 16s ease-in-out infinite alternate}}
        @keyframes orb1{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(-60px,40px) scale(1.15)}}}}
        @keyframes orb2{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(50px,-50px) scale(1.2)}}}}
        @keyframes orb3{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(-40px,30px) scale(0.9)}}}}
        @keyframes orb4{{0%{{transform:translate(0,0) scale(1)}}100%{{transform:translate(30px,40px) scale(1.1)}}}}
        div.stButton>button{{
          background:transparent!important;color:#587579!important;
          border:1.5px solid #587579!important;font-family:'Raleway',sans-serif!important;
          font-weight:700!important;border-radius:6px!important;width:100%!important;
          padding:14px!important;font-size:.95rem!important;margin-top:10px;
          letter-spacing:.08em!important;position:relative;z-index:10}}
        div.stButton>button:hover{{background:#587579!important;color:white!important}}
        .stTextInput input{{
          background:rgba(42,45,69,0.85)!important;border:1.5px solid #3a3d55!important;
          color:white!important;border-radius:6px!important;
          font-family:'Raleway',sans-serif!important;padding:14px!important;
          position:relative;z-index:10}}
        .stTextInput input::placeholder{{color:#666!important}}
        label{{color:#666!important;font-size:0!important;line-height:0!important}}
        </style>
        <div class="neb neb1"></div>
        <div class="neb neb2"></div>
        <div class="neb neb3"></div>
        <div class="neb neb4"></div>
        <div class="login-wrap">
          <img src="data:image/png;base64,{tessaix_b64}" style="width:320px">
          <div class="tagline">Propuestas comerciales inteligentes</div>
        </div>
        """, unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1,2,1])
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
    /* Home cards */
    .home{min-height:80vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:40px 0}
    .cards{display:flex;gap:20px;justify-content:center;flex-wrap:wrap;margin-top:32px}
    .card{background:transparent;border-radius:12px;padding:28px 36px;width:220px;
           text-align:center;cursor:pointer;transition:all .2s}
    .card-hc{border:1.5px solid #587579}
    .card-hc:hover{background:#58757918}
    .card-fi{border:1.5px solid #3a3850}
    .card-fi:hover{border-color:#555}
    </style>""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
def render_header(section="HUMAN CAPITAL"):
    logo_data = logo_b64()
    logo_html = f'<img src="data:image/png;base64,{logo_data}" style="height:32px">' if logo_data else \
                '<span style="color:white;font-weight:900;letter-spacing:.15em;font-size:.9rem">TESSERA</span>'
    st.markdown(f"""
    <div class="hdr">
      <div style="display:flex;align-items:center;gap:16px">
        {logo_html}
        <span style="color:#3a3d55;font-size:1.1rem">|</span>
        <span style="color:#FBE0A0;font-weight:800;font-size:.85rem;letter-spacing:.2em">TESSAIX</span>
      </div>
      <span class="hb">{section}</span>
    </div>""", unsafe_allow_html=True)

STEPS = ["Config","Cliente","Servicios","Contexto","Generar"]

def render_steps(current):
    cols = st.columns(len(STEPS))
    for i,(col,label) in enumerate(zip(cols,STEPS)):
        col_c = "#587579" if i<=current else "#e0ddd8"
        t_c   = "#587579" if i==current else ("#aaa" if i<current else "#ccc")
        w     = "700" if i==current else "400"
        with col:
            st.markdown(f'<div style="height:2px;background:{col_c};border-radius:1px;margin-bottom:5px"></div>'
                        f'<div style="font-size:.62rem;font-weight:{w};color:{t_c};text-align:center;'
                        f'text-transform:uppercase;letter-spacing:.08em">{label}</div>',
                        unsafe_allow_html=True)

# ─── HOME SCREEN ──────────────────────────────────────────────────────────────
def render_home():
    tessaix_b64 = make_ailerons_png("TESSAIX", color=(32,32,49), size=52, width=480, height=76)
    st.markdown(f"""
    <div style="text-align:center;padding:60px 0 0">
      <div style="font-size:.72rem;letter-spacing:.25em;color:#587579;font-weight:700;
                  text-transform:uppercase;margin-bottom:12px;font-family:'Raleway',sans-serif">Tessera</div>
      <img src="data:image/png;base64,{tessaix_b64}" style="width:240px;margin-bottom:10px">
      <div style="color:#aaa;font-size:.8rem;letter-spacing:.1em;margin-bottom:48px;
                  font-family:'Raleway',sans-serif">Propuestas comerciales inteligentes</div>
    </div>
    <div style="display:flex;gap:20px;justify-content:center;flex-wrap:wrap;padding-bottom:60px">
      <div style="background:transparent;border:1.5px solid #587579;border-radius:12px;
                  padding:28px 36px;width:220px;text-align:center">
        <div style="font-size:28px;margin-bottom:14px">👥</div>
        <div style="font-size:.9rem;font-weight:700;color:#587579;margin-bottom:8px;
                    letter-spacing:.05em;font-family:'Raleway',sans-serif">Human Capital</div>
        <div style="font-size:.72rem;color:#999;line-height:1.6;font-family:'Raleway',sans-serif">
          Recruitment · HR Advisory<br>Outsourcing · RPO</div>
      </div>
      <div style="background:transparent;border:1.5px solid #2a2830;border-radius:12px;
                  padding:28px 36px;width:220px;text-align:center">
        <div style="font-size:28px;margin-bottom:14px;filter:grayscale(.6)">🏗️</div>
        <div style="font-size:.9rem;font-weight:700;color:#555;margin-bottom:8px;
                    letter-spacing:.05em;font-family:'Raleway',sans-serif">Finance</div>
        <div style="font-size:.72rem;color:#888;line-height:1.6;font-family:'Raleway',sans-serif">
          En obras<br><span style="color:#FBE0A060">Próximamente</span></div>
      </div>
    </div>
    <div style="text-align:center;color:#ccc;font-size:.7rem;letter-spacing:.12em;
                font-family:'Raleway',sans-serif;padding-bottom:20px">
      Better decisions, together.
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        if st.button("Entrar a Human Capital →", use_container_width=True):
            st.session_state.screen = "hc"
            st.rerun()

# ─── AI CONTENT ───────────────────────────────────────────────────────────────
def generate_content(data):
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    is_en  = data["lang"] == "en"
    svcs   = ", ".join([SERVICES_HC.get(s,s) for s in data["services"]])

    system = ("Eres consultor comercial senior de Tessera Human Capital. "
              "Devuelves ÚNICAMENTE JSON válido. Sin guiones largos, sin 'contexto' ni 'criterio', "
              "español de España natural, específico al sector del cliente.") if not is_en else \
             ("Senior commercial consultant at Tessera Human Capital. "
              "Return ONLY valid JSON. No em dashes, no 'context'/'criteria', natural English.")

    svc_json = {}
    for svc in data["services"]:
        cn, sec = data["client_name"], data["sector"]
        if svc == "headhunting":
            svc_json[svc] = {
                "why_col": f"3 diferenciadores específicos para {sec}",
                "benefits_col": f"3 beneficios para {cn} en {sec}",
                "how_col": "5 pasos adaptados al sector"
            }
        elif svc == "outsourcing":
            svc_json[svc] = {
                "headline": f"frase de impacto para {sec}",
                "body": f"2 líneas sobre outsourcing para {cn}",
                "benefit_cards": [{"title":"Coste variable","body":"adaptado"},
                                   {"title":"Cobertura total","body":"adaptado"},
                                   {"title":"Cero gestión","body":"adaptado"},
                                   {"title":"Flexibilidad","body":"adaptado"}],
                "differentials": ["diferencial 1","diferencial 2","diferencial 3"]
            }
        elif svc == "rpo":
            svc_json[svc] = {
                "headline": f"titular RPO para {sec}", "body": f"2 líneas RPO para {cn}",
                "pillar1_title":"DEDICACIÓN TOTAL","pillar1_body":f"adaptado a {sec}",
                "pillar2_title":"EFICIENCIA REAL","pillar2_body":"con dato concreto",
                "pillar3_title":"FOCO EN CRECER","pillar3_body":"resultados específicos"
            }
        elif svc == "salary":
            svc_json[svc] = {
                "headline": f"Inteligencia retributiva para {sec}",
                "body": f"2 líneas consultoría salarial para {cn}",
                "services": [{"title":"Benchmarking salarial","body":"adaptado"},
                              {"title":"Transparencia retributiva","body":"adaptado"},
                              {"title":"Estudio de mercado","body":"adaptado"},
                              {"title":"Diseño de bandas","body":"adaptado"},
                              {"title":"Asesoría People","body":"adaptado"}]
            }

    template = {
        "cover": {
            "headline": "Queremos ser tu partner de talento",
            "tagline": f"frase específica para {data['client_name']} en {data['sector']} (1 línea directa)",
            "contact": "Manuel.garcia@tesseraservices.com"
        },
        "why_tessera": {
            "d1_title": f"diferenciador 1 para {data['sector']}", "d1_body": "2 líneas",
            "d2_title": f"diferenciador 2 para {data['sector']}", "d2_body": "2 líneas",
            "d3_title": f"diferenciador 3 para {data['sector']}", "d3_body": "2 líneas"
        },
        "first_week": {
            "step1_body": f"descripción 24h para {data['sector']}",
            "step2_body": "descripción 72h",
            "step3_body": "descripción semana 1"
        },
        "services": svc_json,
        "fees": {
            "fee_rate": data.get("fee_rate","XX"),
            "success_body": "descripción modelo éxito",
            "partner_body": "descripción modelo partner"
        },
        "contact": {"email":"Manuel.garcia@tesseraservices.com","phone":"+34 619 511 155","linkedin":"@tesseraservices"}
    }

    prompt = (f"Genera JSON de propuesta comercial para:\n"
              f"CLIENTE: {data['client_name']}\nSECTOR: {data['sector']}\n"
              f"PAÍS: {data.get('country','España')}\nTAMAÑO: {data.get('size','')}\n"
              f"CONTACTO: {data.get('contact_name','')} ({data.get('contact_role','')})\n"
              f"SERVICIOS: {svcs}\nPAIN POINTS: {data.get('pain_points','')}\n"
              f"PERFILES: {data.get('roles_needed','')}\nFEE: {data.get('fee_rate','XX')}%\n"
              f"INFO: {data.get('extra_info','')}\n\n"
              f"Rellena TODOS los campos con contenido real y específico:\n"
              f"{json.dumps(template, ensure_ascii=False, indent=2)}")

    with st.spinner("Generando contenido adaptado al cliente..."):
        msg = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=4000,
            system=system, messages=[{"role":"user","content":prompt}]
        )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n",1)[1].rsplit("```",1)[0].strip()
    try:
        return json.loads(raw)
    except Exception as e:
        st.error(f"Error parseando JSON: {e}"); st.code(raw[:600]); return None

# ─── PPTX BUILDER ─────────────────────────────────────────────────────────────
def replace_text(xml, old, new):
    esc_old = old.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    esc_new = new.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    if esc_old in xml: return xml.replace(esc_old, esc_new)
    if old in xml: return xml.replace(old, esc_new)
    return xml

def build_pptx(data, content):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        work = tmp / "work.pptx"
        shutil.copy(PLANTILLA, work)
        scripts = Path("/mnt/skills/public/pptx/scripts")
        unpack  = tmp / "unpacked"
        subprocess.run(["python3", str(scripts/"office/unpack.py"), str(work), str(unpack)], capture_output=True)
        slides = unpack / "ppt" / "slides"
        media  = unpack / "ppt" / "media"

        # Logo cliente: reemplaza image4.png (Tandem) con logo buscado por dominio
        domain = (data.get("client_website","") or f"{data['client_name'].lower().replace(' ','')}.com")
        domain = domain.replace("https://","").replace("http://","").replace("www.","").split("/")[0].strip()
        client_logo = None
        try:
            import requests as rq
            for url in [f"https://logo.clearbit.com/{domain}",
                        f"https://www.google.com/s2/favicons?domain={domain}&sz=128"]:
                try:
                    r = rq.get(url, timeout=6, headers={"User-Agent":"Mozilla/5.0"})
                    if r.status_code == 200 and len(r.content) > 500:
                        client_logo = r.content; break
                except: continue
        except: pass

        if client_logo:
            (media/"image4.png").write_bytes(client_logo)
        else:
            s1 = (slides/"slide1.xml").read_text("utf-8")
            s1 = re.sub(r'<p:pic>\s*(?:(?!</p:pic>).)*?rId6(?:(?!</p:pic>).)*?</p:pic>','',s1,flags=re.DOTALL)
            (slides/"slide1.xml").write_text(s1,"utf-8")

        cover = content.get("cover",{})
        why   = content.get("why_tessera",{})
        fw    = content.get("first_week",{})
        fees  = content.get("fees",{})
        svcs  = content.get("services",{})

        # SLIDE 1
        s = (slides/"slide1.xml").read_text("utf-8")
        s = replace_text(s,"Edward Manrique ","Manuel García Pina ")
        s = replace_text(s,"+34 695021978","+34 619 511 155")
        if data["lang"]=="en":
            s = replace_text(s,"Queremos ser tu partner de talento",
                             cover.get("headline","From now on, you will enjoy your selection processes"))
        (slides/"slide1.xml").write_text(s,"utf-8")

        # SLIDE 3
        s = (slides/"slide3.xml").read_text("utf-8")
        if why.get("d1_title"): s = replace_text(s,"Sin CVs al azar",why["d1_title"])
        if why.get("d1_body"):  s = replace_text(s,"No enviamos el primer CV, buscamos a quien encaja de verdad. Candidatos filtrados sobre la mesa en 72 horas.",why["d1_body"])
        if why.get("d2_title"): s = replace_text(s,"Sin pausas",why["d2_title"])
        if why.get("d2_body"):  s = replace_text(s,"Tu servicio no parará.\nNos aseguramos de cubrir tus vacantes incluso si el responsable directo no se encuentra disponible.",why["d2_body"])
        if why.get("d3_title"): s = replace_text(s,"Siempre contigo",why["d3_title"])
        if why.get("d3_body"):  s = replace_text(s,"No desaparecemos tras la incorporación: cuidamos a la persona y al cliente durante todo el servicio.",why["d3_body"])
        (slides/"slide3.xml").write_text(s,"utf-8")

        # SLIDE 4
        s = (slides/"slide4.xml").read_text("utf-8")
        if fw.get("step1_body"): s = replace_text(s,"Nos integramos con tu equipo para entender a fondo el rol, las necesidades y tu forma de trabajar. Lanzamos la búsqueda desde el primer día, publicando en los canales necesarios sin coste adicional.",fw["step1_body"])
        (slides/"slide4.xml").write_text(s,"utf-8")

        # SLIDE 5 headhunting
        if "headhunting" in data.get("services",[]):
            hh = svcs.get("headhunting",{})
            s  = (slides/"slide5.xml").read_text("utf-8")
            if hh.get("why_col"):     s = replace_text(s,"VELOCIDAD \nEncontramos rápido, pero no a cualquiera. Seleccionamos a los que suman valor real dentro de nuestra base de datos especializada en ",hh["why_col"][:400])
            if hh.get("benefits_col"):s = replace_text(s,"CERO ESTRÉSNos ocupamos de todo el proceso, tú solo conoces a los mejores.TALENTO A MEDIDACandidatos",hh["benefits_col"][:300])
            if hh.get("how_col"):     s = replace_text(s,"NOS CONOCEMOSEscuchamos tus necesidades y objetivos desde el primer día.VAMOS MÁS ALLÁBuscamos talen",hh["how_col"][:300])
            (slides/"slide5.xml").write_text(s,"utf-8")

        # SLIDE 6 outsourcing
        if "outsourcing" in data.get("services",[]):
            outs = svcs.get("outsourcing",{})
            s    = (slides/"slide6.xml").read_text("utf-8")
            if outs.get("headline"): s = replace_text(s,"Externalización que sí funciona: pagas por trabajo real, sin papeleo.",outs["headline"])
            if outs.get("body"):     s = replace_text(s,"Contratar cuesta más de lo que parece. Con el modelo Time & Material ajustas el equipo a tu actividad real y nosotros nos encargamos de toda la gestión.",outs["body"])
            cards = outs.get("benefit_cards",[])
            for i,(ot,ob) in enumerate([("Coste variable","Pagas solo por horas reales de trabajo."),
                                         ("Cobertura total","Cubrimos bajas y vacaciones."),
                                         ("Cero gestión","Nóminas y trámites, a cargo nuestro."),
                                         ("Flexibilidad","Escalas el equipo cuando quieras.")]):
                if i<len(cards):
                    s = replace_text(s,ot,cards[i].get("title",ot))
                    s = replace_text(s,ob,cards[i].get("body",ob))
            diffs = outs.get("differentials",[])
            for i,od in enumerate(["Headhunters especializados","Cuidamos a las personas","Partner de principio a fin"]):
                if i<len(diffs): s = replace_text(s,od,diffs[i])
            (slides/"slide6.xml").write_text(s,"utf-8")

        # SLIDE 7 RPO
        if "rpo" in data.get("services",[]):
            rpo = svcs.get("rpo",{})
            s   = (slides/"slide7.xml").read_text("utf-8")
            if rpo.get("headline"): s = replace_text(s,"RPO a tu medida",rpo["headline"])
            if rpo.get("body"):     s = replace_text(s,"Te ofrecemos un equipo de recruiters que se integra en tu compañía como una extensión real de tu equipo, trabajando exclusivamente en tus necesidades durante el tiempo que lo necesites.\nNo solo ejecutamos procesos: entendemos tu negocio, tus retos y tu historia  para atraer el talento que realmente necesitas.",rpo["body"])
            for key,ot,ob in [("pillar1","DEDICACIÓN TOTAL","Un equipo dedicado en exclusiva a tu compañía, con conocimiento del sector publicitario y alineado c"),
                               ("pillar2","EFICIENCIA REAL","Acceso directo a una red de +12.000 profesionales de publicidad ya identificados y validados. Menos "),
                               ("pillar3","FOCO EN CRECER","Nosotros buscamos, filtramos y validamos. Tú te concentras en hacer crecer tu equipo y tu negocio.")]:
                if rpo.get(f"{key}_title"): s = replace_text(s,ot,rpo[f"{key}_title"])
                if rpo.get(f"{key}_body"):  s = replace_text(s,ob,rpo[f"{key}_body"][:len(ob)])
            (slides/"slide7.xml").write_text(s,"utf-8")

        # SLIDE 15 fees
        if data.get("deck_type")=="propuesta":
            s  = (slides/"slide15.xml").read_text("utf-8")
            fr = fees.get("fee_rate",data.get("fee_rate","XX"))
            s  = replace_text(s,"18%",f"{fr}%")
            s  = replace_text(s,"13%",f"{fr}%")
            (slides/"slide15.xml").write_text(s,"utf-8")

        # SLIDE 16 cierre
        s = (slides/"slide16.xml").read_text("utf-8")
        s = replace_text(s,"Edward@tesseraservices.com",content.get("contact",{}).get("email","Manuel.garcia@tesseraservices.com"))
        (slides/"slide16.xml").write_text(s,"utf-8")

        subprocess.run(["python3",str(scripts/"clean.py"),str(unpack)],capture_output=True)
        out = tmp/"output.pptx"
        subprocess.run(["python3",str(scripts/"office/pack.py"),str(unpack),str(out),"--original",str(work)],capture_output=True)
        return out.read_bytes()

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="TESSAIX", page_icon=str(FAVICON) if FAVICON.exists() else "T",
                       layout="centered", initial_sidebar_state="collapsed")
    check_auth()
    inject_css()

    for k,v in [("screen","home"),("step",0),("form",{}),("content",None),("pptx",None)]:
        if k not in st.session_state: st.session_state[k]=v

    # HOME
    if st.session_state.screen == "home":
        render_home()
        return

    # FINANCE
    if st.session_state.screen == "finance":
        render_header("FINANCE")
        st.markdown("""
        <div style="text-align:center;padding:80px 0">
          <div style="font-size:64px;margin-bottom:20px">🏗️</div>
          <div style="font-size:1.2rem;font-weight:700;color:#FBE0A0;margin-bottom:12px;font-family:'Raleway',sans-serif">En obras</div>
          <div style="color:#888;font-size:.9rem;line-height:1.8;font-family:'Raleway',sans-serif">
            Tessera Finance estará disponible muy pronto.<br>Estamos construyendo algo especial.
          </div>
        </div>""", unsafe_allow_html=True)
        if st.button("← Volver"):
            st.session_state.screen="home"; st.rerun()
        return

    # HC FLOW
    render_header("HUMAN CAPITAL")
    step = st.session_state.step
    render_steps(step)

    if step == 0:
        st.markdown("### ¿Qué tipo de propuesta?")
        c1,c2 = st.columns(2)
        with c1: lang      = st.radio("Idioma",["Español","English"])
        with c2: deck_type = st.radio("Tipo",["Presentación (sin fees)","Propuesta completa (con fees)"])
        st.markdown("---")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Inicio"): st.session_state.screen="home"; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                st.session_state.form.update({"lang":"es" if lang=="Español" else "en",
                    "deck_type":"presentacion" if "sin fees" in deck_type else "propuesta"})
                st.session_state.step=1; st.rerun()

    elif step == 1:
        st.markdown("### Datos del cliente")
        c1,c2 = st.columns(2)
        with c1:
            cn  = st.text_input("Nombre de la empresa *", placeholder="Ej: Moneycorp, Corpay…")
            sec = st.text_input("Sector / industria *",   placeholder="Ej: Fintech, Banca privada…")
            ctr = st.text_input("País / mercado", value="España")
        with c2:
            sz  = st.text_input("Tamaño de empresa",   placeholder="Ej: 50–200 empleados")
            web = st.text_input("Web del cliente",     placeholder="moneycorp.com")
            ccn = st.text_input("Nombre del contacto", placeholder="Nombre y apellido")
            ccr = st.text_input("Cargo del contacto",  placeholder="Ej: HR Director, CFO…")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=0; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                if not cn.strip() or not sec.strip(): st.error("Empresa y sector obligatorios.")
                else:
                    st.session_state.form.update({"client_name":cn,"sector":sec,"country":ctr,
                        "size":sz,"client_website":web,"contact_name":ccn,"contact_role":ccr})
                    st.session_state.step=2; st.rerun()

    elif step == 2:
        st.markdown("### ¿Qué servicios incluimos?")
        selected = st.multiselect("Servicios",list(SERVICES_HC.keys()),
                                   default=["headhunting"],format_func=lambda x:SERVICES_HC[x])
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=1; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                if not selected: st.error("Selecciona al menos uno.")
                else: st.session_state.form["services"]=selected; st.session_state.step=3; st.rerun()

    elif step == 3:
        st.markdown("### Personalización")
        pp  = st.text_area("Pain points / retos del cliente",height=80,
                            placeholder="Ej: Crecimiento rápido sin proceso de selección consolidado…")
        rn  = st.text_area("Perfiles que necesitan cubrir",height=60,
                            placeholder="Ej: 3 Account Managers, 1 CFO…")
        fee = ""
        if st.session_state.form.get("deck_type")=="propuesta":
            fee = st.text_input("Fee propuesto (%)",placeholder="Ej: 16")
        ei  = st.text_area("Info adicional (web, reunión previa, LinkedIn…)",height=90,
                            placeholder="Pega aquí todo lo que sepas del cliente…")
        cb,cn_ = st.columns([1,4])
        with cb:
            if st.button("← Atrás"): st.session_state.step=2; st.rerun()
        with cn_:
            if st.button("Siguiente →"):
                st.session_state.form.update({"pain_points":pp,"roles_needed":rn,"fee_rate":fee,"extra_info":ei})
                st.session_state.step=4; st.rerun()

    elif step == 4:
        d = st.session_state.form
        st.markdown("### Revisa y genera")
        rows = [("Cliente",d.get("client_name","—")),("Sector",d.get("sector","—")),
                ("Idioma","Español" if d.get("lang")=="es" else "English"),
                ("Tipo","Con fees" if d.get("deck_type")=="propuesta" else "Sin fees"),
                ("Servicios"," · ".join([SERVICES_HC.get(s,s) for s in d.get("services",[])]))]
        if d.get("fee_rate"): rows.append(("Fee",f"{d['fee_rate']}%"))
        html_r = "".join(f'<div class="sr"><span class="sk">{k}</span><span class="sv">{v}</span></div>' for k,v in rows)
        st.markdown(f'<div class="sum">{html_r}</div>',unsafe_allow_html=True)

        if st.session_state.pptx is None:
            cb,cg = st.columns([1,4])
            with cb:
                if st.button("← Atrás"): st.session_state.step=3; st.rerun()
            with cg:
                if st.button("✦ Generar propuesta"):
                    c = generate_content(d)
                    if c:
                        st.session_state.content=c
                        with st.spinner("Montando el PowerPoint…"):
                            try:
                                st.session_state.pptx=build_pptx(d,c); st.rerun()
                            except Exception as e:
                                st.error(f"Error: {e}")
        else:
            st.markdown('<div class="ok">✅ <strong>Propuesta generada.</strong> Descarga el PowerPoint con la plantilla visual de Tessera completa.</div>',unsafe_allow_html=True)
            safe = d.get("client_name","Cliente").replace(" ","_")
            st.download_button("⬇ Descargar PowerPoint",data=st.session_state.pptx,
                               file_name=f"Tessera_{safe}_Propuesta.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)
            if st.button("+ Nueva propuesta"):
                for k,v in [("step",0),("form",{}),("content",None),("pptx",None)]:
                    st.session_state[k]=v
                st.rerun()

if __name__ == "__main__":
    main()
