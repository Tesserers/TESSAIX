"""
Construye `plantilla_hc_finance.pptx` (Human Capital + Finance) desde cero
con el sistema de diseño del rebrand 2026 (ver brand_kit.py). Mismo criterio
que build_hc.py: misma información y estructura que la plantilla anterior,
solo cambia la piel visual.

Uso: python scripts/build_templates/build_hc_finance.py
"""
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

import brand_kit as bk

ROOT = Path(__file__).resolve().parent.parent.parent
ASSETS = Path(__file__).resolve().parent / "assets"
OUT = ROOT / "plantilla_hc_finance.pptx"

SHAPE_LOG: dict[str, dict] = {}


def log(slide_key, label, shape_or_value):
    val = shape_or_value.shape_id if hasattr(shape_or_value, "shape_id") else shape_or_value
    SHAPE_LOG.setdefault(slide_key, {})[label] = val


# ───────────────────────── 1. Portada ─────────────────────────────────
def slide_portada(prs):
    s = bk.new_slide(prs)
    bk.set_gradient_bg(s, bk.TEAL, bk.NAVY, angle=45)
    bk.logo_picture(s, bk.MARGIN, 60, 22, dark_bg=True)
    bk.add_text(s, bk.MARGIN, 150, 500, 18, "PROPUESTA COMERCIAL", typeface=bk.RALEWAY_MEDIUM,
                size=11, color=bk.TEAL_LIGHT, uppercase=True, spacing_pt=1.5)
    bk.add_text(s, bk.MARGIN, 180, 760, 90, "Queremos ser tu partner.",
                typeface=bk.RALEWAY_LIGHT, size=44, color=bk.WHITE)
    bk.add_text(s, bk.MARGIN, 260, 700, 70,
                "Finance & Human Capital. Soluciones financieras y de talento para fondos "
                "de inversión y sus compañías participadas.",
                typeface=bk.RALEWAY, size=14, color=bk.WHITE, line_spacing=1.25)
    bk.add_text(s, bk.MARGIN, 478, 400, 18, "tesseraservices.com",
                typeface=bk.RALEWAY, size=9, color=bk.TEAL_LIGHT)
    bk.add_number(s, bk.SLIDE_W - bk.MARGIN - 80, 478, 80, 18, "2026",
                  size=11, color=bk.TEAL_LIGHT, weight="medium", align=PP_ALIGN.RIGHT)
    return s


# ─────────── 2 / 3. Tessera en una página (todas sedes / Madrid) ──────
def _pagina_header(s, page_no):
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "TESSERA EN UNA PÁGINA")
    bk.add_text(s, bk.MARGIN, 70, 700, 50, "Tessera en una página",
                typeface=bk.RALEWAY_LIGHT, size=34, color=bk.NAVY)
    bk.add_text(s, bk.MARGIN, 116, 300, 16, "Mejores decisiones, juntos.",
                typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.TEAL)
    bk.add_text(s, bk.MARGIN, 138, 760, 36,
                "Consultoría estratégica en Finanzas y Talento para compañías, inversores y "
                "sus participadas. Un equipo senior que se integra en el proyecto y responde "
                "por el resultado.", typeface=bk.RALEWAY, size=10.5, color=bk.NAVY, line_spacing=1.2)
    stats = [("2", "áreas: Finance y Human Capital"), ("+21", "consultores y equipo de CFOs"),
             ("3", "oficinas en España")]
    stat_w = 760 / 3
    for i, (num, label) in enumerate(stats):
        x = bk.MARGIN + i * stat_w
        bk.add_number(s, x, 184, 80, 34, num, size=26, color=bk.NAVY, weight="extralight")
        bk.add_text(s, x, 216, stat_w - 20, 28, label, typeface=bk.RALEWAY, size=9.5, color=bk.NAVY, line_spacing=1.1)
    bk.footer_chrome(s, page_no)


def _location_cards(s, cities, y=260):
    card_w, gap = 266.0, 21.0
    for i, (city, sub) in enumerate(cities):
        x = bk.MARGIN + i * (card_w + gap)
        bk.flat_rect(s, x, y, card_w, 100, bk.BRUMA)
        bk.loop_accent_line(s, x + 18, y + 16, 44)
        bk.add_text(s, x + 18, y + 30, card_w - 36, 24, city, typeface=bk.RALEWAY_SEMIBOLD, size=15, color=bk.NAVY)
        bk.add_text(s, x + 18, y + 58, card_w - 36, 20, sub, typeface=bk.RALEWAY, size=10, color=bk.TEAL)


def _diff_strip(s, y=380):
    tiles = [("BOUTIQUE", "Equipo senior en cada proyecto."),
             ("TRATO DIRECTO", "Un interlocutor, de tú a tú."),
             ("VISIÓN DE NEGOCIO", "Al servicio de la tesis de inversión.")]
    tile_w = (760 - 2 * 16) / 3
    for i, (label, text) in enumerate(tiles):
        x = bk.MARGIN + i * (tile_w + 16)
        bk.flat_rect(s, x, y, tile_w, 70, bk.WHITE if y > 300 else bk.BRUMA)
        bk.add_text(s, x + 16, y + 14, tile_w - 32, 16, label, typeface=bk.RALEWAY_MEDIUM,
                    size=9, color=bk.TEAL, uppercase=True, spacing_pt=0.8)
        bk.add_text(s, x + 16, y + 34, tile_w - 32, 28, text, typeface=bk.RALEWAY, size=10, color=bk.NAVY)


def slide_pagina_todas(prs):
    s = bk.new_slide(prs)
    _pagina_header(s, 2)
    _location_cards(s, [("MADRID", "Sede central"), ("BILBAO", "País Vasco - España"), ("OVIEDO", "Asturias - España")])
    _diff_strip(s, y=378)
    return s


def slide_pagina_madrid(prs):
    s = bk.new_slide(prs)
    _pagina_header(s, 2)
    _location_cards(s, [("MADRID", "Sede central")])
    _diff_strip(s, y=378)
    return s


# ─────────── 4 / 5. Equipo (sin / con Eduardo) ─────────────────────────
TEAM_BASE = [
    ("team_hcf_manuel_pina.png", "MANUEL PINA", "PARTNER - FINANCE",
     "Más de 10 años en Big Four en auditoría, due diligence y M&A, y posteriormente CFO en "
     "grandes pymes españolas. Lidera Tessera Finance, coordinando el equipo de CFOs, "
     "transacciones y asesoramiento financiero estratégico."),
    ("team_hcf_monica_mayoral.png", "MÓNICA MAYORAL", "PARTNER - FINANCE",
     "Más de 20 años como Directora Financiera en empresa multinacional, muy involucrada en "
     "el día a día de la compañía con visión y estrategia global: reducción de riesgos y "
     "planes de crecimiento."),
    ("team_hcf_edward_manrique.png", "EDWARD MANRIQUE", "HEAD OF RECRUITMENT & TALENT OPERATIONS",
     "Más de 8 años de experiencia en identificación y evaluación de talento. En Seedtag "
     "lideró la contratación de más de 200 profesionales en 16 países, y ha trabajado con "
     "entidades financieras como Santander, BNP Paribas y BBVA."),
]
TEAM_EDUARDO = ("team_hcf_eduardo_serrano.png", "EDUARDO SERRANO", "PARTNER - FINANCE",
    "17 años en el mundo del emprendimiento como CFO de startups internacionales y de fondos "
    "líderes de Venture Capital y Venture Builders. Especializado en aportar valor desde el "
    "área financiera, la tesorería y la consecución de financiación y capital.")


def _equipo_slide(prs, team):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "NUESTRO EQUIPO")
    bk.add_text(s, bk.MARGIN, 70, 500, 50, "El Equipo", typeface=bk.RALEWAY_LIGHT, size=36, color=bk.NAVY)
    bk.add_text(s, bk.MARGIN, 122, 700, 18, "Equipo directivo - +21 consultores - Madrid, Bilbao y Oviedo",
                typeface=bk.RALEWAY, size=10.5, color=bk.TEAL)

    n = len(team)
    card_w, gap = (760 - (n - 1) * 20) / n, 20.0
    for i, (photo, name, role, bio) in enumerate(team):
        x = bk.MARGIN + i * (card_w + gap)
        bk.cover_picture(s, ASSETS / photo, x, 155, card_w, 135)
        bk.loop_accent_line(s, x, 300, 50)
        bk.add_text(s, x, 312, card_w, 18, name, typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.NAVY)
        bk.add_text(s, x, 330, card_w, 16, role, typeface=bk.RALEWAY_MEDIUM, size=8,
                    color=bk.TEAL, uppercase=True, spacing_pt=0.6)
        bk.add_paragraphs(s, x, 350, card_w, 150, [bio], size=8.3, color=bk.NAVY, line_spacing=1.15)
    bk.footer_chrome(s, 3)
    return s


def slide_equipo_sin_eduardo(prs):
    return _equipo_slide(prs, TEAM_BASE)


def slide_equipo_con_eduardo(prs):
    team = [TEAM_BASE[0], TEAM_EDUARDO, TEAM_BASE[1], TEAM_BASE[2]]
    return _equipo_slide(prs, team)


# ───────────────── 6. Para el fondo (ciclo de inversión) ───────────────
def slide_para_el_fondo(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.BRUMA)
    bk.header_chrome(s, "PARA EL FONDO Y SUS PARTICIPADAS")
    bk.add_text(s, bk.MARGIN, 70, 760, 50, "Un partner para el fondo",
                typeface=bk.RALEWAY_LIGHT, size=32, color=bk.NAVY)
    bk.add_text(s, bk.MARGIN, 118, 760, 24,
                "Finance y Human Capital en un mismo equipo, en cada fase del ciclo de inversión.",
                typeface=bk.RALEWAY, size=11, color=bk.NAVY)

    stages = [
        ("01", "PRE-INVERSIÓN", "Antes de firmar",
         ["Due diligence financiera del target", "Valoración independiente",
          "Revisión del business plan", "Estructura de la operación"]),
        ("02", "ACOMPAÑAMIENTO", "Durante la inversión",
         ["CFO part-time y reporting al fondo", "Financiación y ahorro de costes",
          "Headhunting de equipos directivos", "Outsourcing y consultoría salarial"]),
        ("03", "SALIDA", "Al preparar la venta",
         ["Compañía preparada para la venta", "Vendor due diligence y valoración",
          "Proceso de M&A hasta el cierre", "Búsqueda del comprador o socio"]),
    ]
    card_w, gap = 266.0, 21.0
    for i, (num, title, sub, bullets) in enumerate(stages):
        x = bk.MARGIN + i * (card_w + gap)
        bk.flat_rect(s, x, 160, card_w, 230, bk.WHITE)
        bk.add_number(s, x + 18, 176, 50, 26, num, size=16, color=bk.TEAL, weight="medium")
        bk.add_text(s, x + 18, 204, card_w - 36, 20, title, typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.NAVY)
        bk.add_text(s, x + 18, 224, card_w - 36, 16, sub, typeface=bk.RALEWAY, size=9.5, color=bk.TEAL)
        bk.add_paragraphs(s, x + 18, 246, card_w - 36, 140, [f"-  {b}" for b in bullets],
                          size=9, color=bk.NAVY, line_spacing=1.2, space_after=5)

    bk.flat_rect(s, bk.MARGIN, 412, 760, 86, bk.NAVY)
    bk.add_text(s, bk.MARGIN + 20, 432, 520, 50,
                "Un único interlocutor para todo el ciclo de inversión, con visibilidad "
                "financiera homogénea entre participadas.",
                typeface=bk.RALEWAY, size=11, color=bk.WHITE, line_spacing=1.25)
    bk.add_number(s, bk.MARGIN + 600, 426, 90, 36, "360º", size=24, color=bk.WHITE, weight="extralight")
    bk.add_text(s, bk.MARGIN + 600, 462, 150, 24, "Cobertura del ciclo\nde inversión",
                typeface=bk.RALEWAY, size=8.5, color=bk.TEAL_LIGHT, line_spacing=1.05)
    bk.footer_chrome(s, 4)
    return s


# ───────────────── 7. Servicios Finance ────────────────────────────────
def slide_servicios_finance(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "SERVICIOS")
    bk.add_text(s, bk.MARGIN, 70, 700, 50, "Finance", typeface=bk.RALEWAY_LIGHT, size=36, color=bk.NAVY)
    bk.add_text(s, bk.MARGIN, 118, 760, 20,
                "Todo el ciclo financiero de una participada, desde la due diligence hasta la salida.",
                typeface=bk.RALEWAY, size=11, color=bk.NAVY)

    cards = [
        ("Due Diligence Financiera", "Calidad de ingresos, EBITDA ajustado, deuda neta, circulante y red flags antes de invertir."),
        ("Valoraciones", "DCF y múltiplos de mercado, con informe defendible ante compradores, inversores y auditores."),
        ("Business Plan", "Estrategia y proyecciones financieras completas, con escenarios y sensibilidad."),
        ("M&A", "Compra, venta y búsqueda de socios: del primer contacto al cierre, con red internacional."),
        ("CFO part-time", "Dirección financiera senior para la participada y reporting homogéneo para el fondo."),
        ("Financiación", "Capital, deuda bancaria y privada, circulante y financiación pública (93% de éxito en ENISA)."),
        ("Cost Down", "Ahorro de costes a éxito: si no ahorramos, no cobramos. Media del 30% en los costes analizados."),
        ("Otros servicios", "Restructuring, tesorería, recobros, interim team, integraciones contables y soporte a auditoría."),
    ]
    card_w, gap = 184.0, 12.0
    for i, (title, body) in enumerate(cards):
        col = i % 4
        row = i // 4
        x = bk.MARGIN + col * (card_w + gap)
        y = 156 + row * 150
        bk.flat_rect(s, x, y, card_w, 136, bk.BRUMA)
        bk.loop_accent_line(s, x + 14, y + 14, 34)
        bk.add_text(s, x + 14, y + 28, card_w - 28, 32, title, typeface=bk.RALEWAY_SEMIBOLD, size=10, color=bk.NAVY, line_spacing=1.05)
        bk.add_text(s, x + 14, y + 62, card_w - 28, 68, body, typeface=bk.RALEWAY, size=7.8, color=bk.NAVY, line_spacing=1.15)

    bk.add_text(s, bk.MARGIN, 466, 760, 20,
                "Un solo equipo para el fondo y para cada compañía: menos interlocutores, más control.",
                typeface=bk.RALEWAY_SEMIBOLD, size=10.5, color=bk.TEAL)
    bk.footer_chrome(s, 5)
    return s


# ───────────────── 8. Servicios Human Capital (resumen) ────────────────
def slide_servicios_hc_resumen(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "SERVICIOS")
    bk.add_text(s, bk.MARGIN, 70, 700, 50, "Human Capital", typeface=bk.RALEWAY_LIGHT, size=36, color=bk.NAVY)
    bk.add_text(s, bk.MARGIN, 118, 760, 20,
                "Expertos en el flujo de talento: procesos claros y compromisos que se cumplen.",
                typeface=bk.RALEWAY, size=11, color=bk.NAVY)

    cards = [
        ("Headhunting", "Equipo senior en cada búsqueda y metodología propia apoyada en nuestra IA de "
         "selección (TESSA). Primera shortlist en 72 horas y acompañamiento tras la incorporación."),
        ("Outsourcing", "Equipos que se ajustan a la actividad real: pagas por horas trabajadas, "
         "cubrimos bajas y vacaciones y nos ocupamos de nóminas y trámites."),
        ("Consultoría salarial", "Benchmarking, diseño de bandas y transparencia retributiva alineada "
         "con la Directiva EU 2023/970. Apoyo en ofertas, revisiones y retención."),
        ("Formación", "Programas a medida con metodología ágil: gestión del talento, employee "
         "experience, procesos de RRHH e IA aplicada a HR. Se mide en resultados."),
    ]
    card_w, gap = 266.0, 21.0
    for i, (title, body) in enumerate(cards):
        col = i % 2
        row = i // 2
        x = bk.MARGIN + col * (card_w + gap)
        y = 156 + row * 130
        bk.flat_rect(s, x, y, 2 * card_w + gap if False else card_w, 112, bk.BRUMA)
        bk.loop_accent_line(s, x + 18, y + 16, 44)
        bk.add_text(s, x + 18, y + 30, card_w - 36, 18, title, typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.NAVY)
        bk.add_text(s, x + 18, y + 50, card_w - 36, 56, body, typeface=bk.RALEWAY, size=8.5, color=bk.NAVY, line_spacing=1.15)

    bk.flat_rect(s, bk.MARGIN + 2 * (card_w + gap), 156, card_w, 242, bk.NAVY)
    bk.add_text(s, bk.MARGIN + 2 * (card_w + gap) + 18, 174, card_w - 36, 70,
                "Para un fondo: equipos directivos para las participadas, cobertura de vacantes "
                "críticas sin pausas y estructuras salariales que retienen al talento clave.",
                typeface=bk.RALEWAY, size=9.5, color=bk.WHITE, line_spacing=1.25)
    bk.add_number(s, bk.MARGIN + 2 * (card_w + gap) + 18, 260, 90, 30, "72h", size=20, color=bk.WHITE, weight="extralight")
    bk.add_text(s, bk.MARGIN + 2 * (card_w + gap) + 18, 292, card_w - 36, 24, "Primera shortlist sobre tu mesa",
                typeface=bk.RALEWAY, size=8, color=bk.TEAL_LIGHT, line_spacing=1.1)
    bk.add_number(s, bk.MARGIN + 2 * (card_w + gap) + 18, 330, 90, 30, "+200", size=20, color=bk.WHITE, weight="extralight")
    bk.add_text(s, bk.MARGIN + 2 * (card_w + gap) + 18, 362, card_w - 36, 24, "Contrataciones lideradas en 16 países",
                typeface=bk.RALEWAY, size=8, color=bk.TEAL_LIGHT, line_spacing=1.1)
    bk.footer_chrome(s, 6)
    return s


# ───────────────── 9. Por qué Tessera (TAMBIÉN donante de Índice/Contexto) ─
PQT_EYEBROW = "POR QUÉ TESSERA"
PQT_TITLE = "Por qué Tessera"
PQT_SUBTITLE = "Control, rigor financiero y talento al servicio de la tesis de inversión."

WHY_CARDS = [
    ("Control", "Visibilidad real del fondo sobre cada participada: reporting homogéneo, cuadro de mando y seguimiento de caja."),
    ("Rigor financiero", "Metodología de Big Four y de CFOs en activo: información validada y decisiones basadas en datos."),
    ("Ejecución", "No nos quedamos en el diagnóstico: implantamos, acompañamos al equipo directivo y medimos el resultado."),
]
STRIP_TILES = [
    ("UN SOLO EQUIPO", "Finance y talento, un interlocutor."),
    ("SOCIOS DENTRO", "Los socios ejecutan el proyecto."),
    ("INDEPENDENCIA", "Decimos lo que vemos."),
]


def slide_por_que_tessera(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.BRUMA)
    bk.header_chrome(s, PQT_EYEBROW)
    title_shp = bk.add_text(s, bk.MARGIN, 70, 760, 60, PQT_TITLE, typeface=bk.RALEWAY_LIGHT, size=40, color=bk.NAVY)
    body_shp = bk.add_text(s, bk.MARGIN, 128, 760, 24, PQT_SUBTITLE, typeface=bk.RALEWAY, size=12, color=bk.NAVY)
    log("pqt", "title", title_shp)
    log("pqt", "body", body_shp)

    card_w, gap = 266.0, 21.0
    prune = []
    for i, (title, body) in enumerate(WHY_CARDS):
        x = bk.MARGIN + i * (card_w + gap)
        bg = bk.flat_rect(s, x, 172, card_w, 190, bk.WHITE)
        acc = bk.loop_accent_line(s, x + 20, 194, 50)
        box = bk.add_textbox(s, x + 20, 210, card_w - 40, 140)
        tf = box.text_frame
        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = title
        bk.style_run(r1, typeface=bk.RALEWAY_SEMIBOLD, size=14, color=bk.NAVY)
        p1.space_after = Pt(6)
        p2 = tf.add_paragraph()
        p2.line_spacing = 1.2
        r2 = p2.add_run()
        r2.text = body
        bk.style_run(r2, typeface=bk.RALEWAY, size=10.5, color=bk.NAVY)
        prune += [bg.shape_id, acc.shape_id, box.shape_id]

    strip_y = 392
    tile_w = (760 - 2 * 16) / 3
    for i, (label, text) in enumerate(STRIP_TILES):
        x = bk.MARGIN + i * (tile_w + 16)
        bg = bk.flat_rect(s, x, strip_y, tile_w, 70, bk.WHITE)
        lbl = bk.add_text(s, x + 16, strip_y + 14, tile_w - 32, 16, label, typeface=bk.RALEWAY_MEDIUM,
                          size=9, color=bk.TEAL, uppercase=True, spacing_pt=0.8)
        txt = bk.add_text(s, x + 16, strip_y + 34, tile_w - 32, 28, text, typeface=bk.RALEWAY, size=10, color=bk.NAVY)
        prune += [bg.shape_id, lbl.shape_id, txt.shape_id]

    log("pqt", "prune_ids", prune)
    bk.footer_chrome(s, 7)
    return s


# ───────────────── 10 / 11. Términos y condiciones (idéntico a HC) ──────
WARRANTY_OLD_SENTENCE = (
    "Si el candidato o el cliente rescinde el contrato de trabajo dentro de los 3 meses "
    "siguientes a la fecha de contratación, Tessera Human Capital se compromete a "
    "proporcionar otro candidato y/o a activar otro proceso de selección, a elección del "
    "cliente, con el fin de sustituir al candidato sin coste adicional."
)
DATE_PLACEHOLDER = "{dd/mm/aaaa}"


def _terms_header(s, page_no):
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "TÉRMINOS Y CONDICIONES")
    bk.add_text(s, bk.MARGIN, 70, 700, 40, "Términos y condiciones", typeface=bk.RALEWAY_LIGHT, size=28, color=bk.NAVY)
    bk.footer_chrome(s, page_no)


def _render_legal_column(s, x, items):
    col_w = 360.0
    box = bk.add_textbox(s, x, 140, col_w, 350)
    tf = box.text_frame
    for i, text in enumerate(items):
        is_heading = text.startswith(("1.", "2.", "3.", ".4", ".5", "6.", "7.", "8."))
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.08
        p.space_after = Pt(4)
        run = p.add_run()
        run.text = text
        bk.style_run(run, typeface=bk.RALEWAY_SEMIBOLD if is_heading else bk.RALEWAY,
                    size=9 if is_heading else 7.3, color=bk.TEAL if is_heading else bk.NAVY)
    return box


def slide_terminos_1(prs):
    s = bk.new_slide(prs)
    _terms_header(s, 10)
    bk.add_text(s, bk.MARGIN, 112, 700, 16, "CONDICIONES SUBYACENTES DE LA PROPUESTA",
                typeface=bk.RALEWAY_MEDIUM, size=9.5, color=bk.TEAL, uppercase=True, spacing_pt=0.8)

    left = [
        "Las presentes condiciones generales se aplican a cualquier servicio profesional "
        "prestado entre las partes y son adicionales a cualquier acuerdo específico "
        "suscrito, el cual prevalecerá sobre estas en caso de discrepancia. En este "
        "documento: (i) «Tessera» significa Tessera Human Capital, S.L.; (ii) «Cliente» es "
        "el destinatario de la Propuesta y (iii) «Servicios» son los servicios "
        "profesionales cubiertos por la Propuesta.",
        "1. CONFIDENCIALIDAD",
        "Toda la información relativa a la empresa cliente, sus miembros y los candidatos "
        "presentados o contactados será tratada por ambas partes con la máxima "
        "confidencialidad, manteniendo la obligación de secreto profesional sobre toda "
        "dicha información y, en particular, sobre los datos personales de los candidatos, "
        "incluso una vez finalizado este acuerdo, todo ello de conformidad con la "
        "legislación española.",
        "Ambas partes se comprometen mutuamente a mantener la confidencialidad y a no "
        "revelar a terceros ninguna información confidencial facilitada por la otra parte, "
        "durante la vigencia de este acuerdo y con posterioridad al mismo. Este compromiso "
        "no impedirá a Tessera Human Capital hacer referencia a los servicios prestados al "
        "Cliente y al propio Cliente como parte de su cartera de clientes, con fines "
        "publicitarios y de marketing.",
        "2. CANDIDATURAS",
        "2.1 Todos los candidatos identificados y presentados por Tessera Human Capital al "
        "cliente quedarán reservados exclusivamente para la presente selección. Si alguno "
        "de dichos candidatos fuera contratado por el cliente en un puesto correspondiente "
        "al encargo, o en otro puesto, o en otra empresa del grupo, dentro de un plazo de "
        "12 meses tras la presentación del candidato, el cliente deberá abonar a Tessera "
        "Human Capital el importe previsto para la contratación inicial, o la totalidad de "
        "los honorarios calculados sobre la retribución bruta anual, incluidas las "
        "retribuciones en especie.",
        "2.2 Esta cláusula se aplica en los mismos términos si el cliente presenta a alguno "
        "de estos candidatos a otra persona física o jurídica y si el candidato se "
        "compromete con dicha persona.",
    ]
    right = [
        "3. GARANTÍA",
        WARRANTY_OLD_SENTENCE,
        "Esta garantía podrá aplicarse salvo en los siguientes casos:",
        "3.1 Si la rescisión del contrato de trabajo se debe a un exceso de plantilla o a "
        "un cambio geográfico inicialmente no previsto.",
        "3.2 Si las condiciones laborales del candidato contratado difieren "
        "significativamente de las inicialmente definidas en el contrato de trabajo.",
        "3.3 Si la rescisión del contrato de trabajo se debe a una falta grave por parte de "
        "la empresa contratante.",
        "3.4 Si no se han abonado todas las facturas adeudadas por el cliente dentro de los "
        "plazos establecidos y conforme a las condiciones generales.",
        "Si el candidato abandona la empresa transcurridos 6 meses desde la fecha de "
        "contratación (periodo de prueba), Tessera Human Capital no se hace responsable y, "
        "por tanto, no tiene obligación de cubrir el puesto.",
    ]
    log("terminos1", "left", _render_legal_column(s, bk.MARGIN, left))
    log("terminos1", "right", _render_legal_column(s, bk.MARGIN + 400, right))
    return s


def slide_terminos_2(prs):
    s = bk.new_slide(prs)
    _terms_header(s, 11)

    left = [
        ".4. RESPONSABILIDAD",
        "Las siguientes aclaraciones legales no excluyen nuestras condiciones de garantía, "
        "que permanecen válidas conforme a los términos definidos en el apartado 3.",
        "4.1 Tessera Human Capital se compromete a prestar un servicio de calidad y "
        "garantiza, como mínimo, la devolución de las cantidades abonadas en aquellos casos "
        "en los que hubiera advertido al cliente de un incumplimiento grave y directo de "
        "sus obligaciones.",
        "4.2 Sin perjuicio de lo anterior, salvo en caso de dolo, Tessera Human Capital no "
        "asume responsabilidad alguna, ni tampoco su personal, por cualquier negligencia o "
        "incapacidad, mala conducta o práctica inadecuada o cualquier otra reclamación "
        "atribuida a los candidatos presentados.",
        ".5. PROTECCIÓN DE DATOS",
        "Ambas partes reconocen que cumplen con la normativa de protección de datos "
        "personales y han sido informadas de las finalidades del tratamiento de los datos "
        "facilitados en esta oferta comercial, basadas en el interés legítimo de ambas "
        "partes. Cualquier cuestión relativa al tratamiento de datos podrá dirigirse a "
        "Tessera Human Capital.",
        "La cesión de los CV se realiza exclusivamente en el marco de la selección "
        "contratada; una vez finalizada, el cliente se compromete a devolverlos o "
        "destruirlos, a la confidencialidad de los datos personales facilitados, a su no "
        "utilización fuera de la relación contractual y al cumplimiento del RGPD, siendo el "
        "único responsable de cualquier incumplimiento de dicha normativa.",
        "Tessera Human Capital, por su parte, garantiza que cuenta con el consentimiento de "
        "los candidatos para la cesión de sus datos a sus clientes durante los procesos de "
        "selección de personal.",
    ]
    right_top = [
        "6. LEY APLICABLE Y JURISDICCIÓN",
        "Este acuerdo se regirá e interpretará de conformidad con la legislación española. "
        "Para la resolución de cualquier controversia derivada de la interpretación o "
        "ejecución de este acuerdo, las partes acuerdan someterse a la jurisdicción de los "
        "Juzgados y Tribunales de Madrid, renunciando expresamente a cualquier otro fuero "
        "que pudiera corresponderles.",
        "7. VIGENCIA DE ESTE ACUERDO",
        "Este acuerdo entrará en vigor desde la fecha de su firma, con fecha de "
        "finalización en mayo de 2026.",
        "8. FIRMA DEL ACUERDO DE COLABORACIÓN",
        f"En Madrid, a {DATE_PLACEHOLDER}",
    ]
    log("terminos2", "left", _render_legal_column(s, bk.MARGIN, left))

    rx = bk.MARGIN + 400
    box_r = bk.add_textbox(s, rx, 140, 360, 170)
    tf = box_r.text_frame
    for i, text in enumerate(right_top):
        is_heading = text.startswith(("6.", "7.", "8."))
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.08
        p.space_after = Pt(4)
        run = p.add_run()
        run.text = text
        bk.style_run(run, typeface=bk.RALEWAY_SEMIBOLD if is_heading else bk.RALEWAY,
                    size=9 if is_heading else 7.3, color=bk.TEAL if is_heading else bk.NAVY)
    log("terminos2", "right", box_r)

    sig_y = 330
    bk.thin_rule(s, rx, sig_y, 360, color=bk.BRUMA, height_pt=1.0)
    bk.add_text(s, rx, sig_y + 12, 200, 16, "FIRMA DE TESSERA", typeface=bk.RALEWAY_MEDIUM,
                size=9.5, color=bk.TEAL, uppercase=True, spacing_pt=0.8)
    firma_shp = bk.add_paragraphs(s, rx, sig_y + 34, 360, 70,
                      ["Nombre: Manuel García Pina", f"Fecha: {DATE_PLACEHOLDER}", "Firma:"],
                      size=9.5, color=bk.NAVY, line_spacing=1.3, space_after=4)
    log("terminos2", "firma_tessera", firma_shp)
    bk.add_text(s, rx, sig_y + 120, 200, 16, "FIRMA DEL CLIENTE", typeface=bk.RALEWAY_MEDIUM,
                size=9.5, color=bk.TEAL, uppercase=True, spacing_pt=0.8)
    bk.thin_rule(s, rx, sig_y + 170, 360, color=bk.BRUMA, height_pt=1.0)
    return s


# ───────────────── 12. Honorarios (idéntico a HC) ───────────────────────
FEE_OLD = "13%"


def slide_honorarios(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "HONORARIOS")
    bk.add_text(s, bk.MARGIN, 70, 760, 40, "El valor que aportamos a tu proyecto",
                typeface=bk.RALEWAY_LIGHT, size=24, color=bk.NAVY)

    cards = [
        ("01", "Experiencia Directa",
         "Entendemos tu entorno, lo que nos permite analizar con precisión las dinámicas "
         "de negocio y los aspectos críticos que impactan en tu crecimiento."),
        ("02", "Enfoque Práctico",
         "Orientados a resultados, facilitando decisiones claras y ejecutables que "
         "optimizan procesos."),
        ("03", "Impacto Real",
         "Soluciones que impulsan eficiencias y mejoras sostenibles en el negocio de "
         "nuestros clientes."),
    ]
    card_w, gap = 246.0, 16.0
    for i, (num, title, body) in enumerate(cards):
        x = bk.MARGIN + i * (card_w + gap)
        bk.flat_rect(s, x, 130, card_w, 150, bk.BRUMA)
        bk.add_number(s, x + 18, 142, 60, 30, num, size=20, color=bk.TEAL, weight="medium")
        bk.add_text(s, x + 18, 176, card_w - 36, 20, title, typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.NAVY)
        bk.add_text(s, x + 18, 198, card_w - 36, 72, body, typeface=bk.RALEWAY, size=9, color=bk.NAVY, line_spacing=1.2)

    fee_y = 310
    bk.flat_rect(s, bk.MARGIN, fee_y, 760, 120, bk.NAVY)
    bk.add_text(s, bk.MARGIN + 24, fee_y + 20, 300, 16, "HONORARIOS", typeface=bk.RALEWAY_MEDIUM,
                size=9.5, color=bk.TEAL_LIGHT, uppercase=True, spacing_pt=0.8)
    bk.add_text(s, bk.MARGIN + 24, fee_y + 42, 400, 24, "Modalidad a éxito en Headhunting",
                typeface=bk.RALEWAY_SEMIBOLD, size=13, color=bk.WHITE)
    bk.add_text(s, bk.MARGIN + 24, fee_y + 72, 300, 18, "Salario Bruto Anual",
                typeface=bk.RALEWAY, size=10, color=bk.TEAL_LIGHT)
    fee_shape = bk.add_number(s, bk.MARGIN + 640, fee_y + 30, 90, 50, FEE_OLD, size=34,
                              color=bk.WHITE, weight="extralight", align=PP_ALIGN.RIGHT)
    log("honorarios", "fee", fee_shape)
    bk.footer_chrome(s, 12)
    return s


# ───────────────── 13. Nuestros Partners (logo wall) ────────────────────
def slide_partners(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "NUESTROS PARTNERS")
    bk.add_text(s, bk.MARGIN, 70, 700, 50, "Nuestros Partners", typeface=bk.RALEWAY_LIGHT, size=34, color=bk.NAVY)

    from PIL import Image
    logos = sorted((ASSETS / "logos").glob("logo_*.png"))
    cols, rows = 6, 3
    cell_w = (bk.SLIDE_W - 2 * bk.MARGIN) / cols
    cell_h = 90.0
    start_y = 150.0
    for i, logo_path in enumerate(logos[: cols * rows]):
        col, row = i % cols, i // cols
        x = bk.MARGIN + col * cell_w
        y = start_y + row * cell_h
        with Image.open(logo_path) as im:
            iw, ih = im.size
        max_w, max_h = cell_w - 24, 44.0
        scale = min(max_w / iw, max_h / ih)
        w, h = iw * scale, ih * scale
        bk.photo_picture(s, logo_path, x + (cell_w - w) / 2, y + (cell_h - h) / 2, w, h)
    bk.footer_chrome(s, 13)
    return s


# ───────────────── 14. Contacto ─────────────────────────────────────────
def slide_contacto(prs):
    s = bk.new_slide(prs)
    bk.set_solid_bg(s, bk.WHITE)
    bk.header_chrome(s, "CONTACTO")
    bk.add_text(s, bk.MARGIN, 70, 700, 50, "Contacta con nosotros", typeface=bk.RALEWAY_LIGHT, size=34, color=bk.NAVY)

    people = [
        ("MANUEL GARCÍA PINA", "Partner - Finance", "+34 669 01 49 58", "manuel.garcia@tesseraservices.com"),
        ("EDUARDO SERRANO CARRETERO", "Partner - Finance", "+34 619 51 11 55", "eduardo.serrano@tesseraservices.com"),
        ("MÓNICA MAYORAL", "Partner - Finance", "+34 696 39 38 49", "monica.mayoral@tesseraservices.com"),
        ("EDWARD MANRIQUE", "Head of Recruitment & Talent Operations", "+34 695 02 19 78", "edward@tesseraservices.com"),
    ]
    card_w, gap = 266.0, 21.0
    for i, (name, role, phone, email) in enumerate(people):
        col = i % 2
        row = i // 2
        x = bk.MARGIN + col * (card_w + gap)
        y = 160 + row * 150
        bk.loop_accent_line(s, x, y, 50)
        bk.add_text(s, x, y + 14, card_w, 20, name, typeface=bk.RALEWAY_SEMIBOLD, size=12, color=bk.NAVY)
        bk.add_text(s, x, y + 34, card_w, 28, role, typeface=bk.RALEWAY, size=9, color=bk.TEAL, line_spacing=1.1)
        bk.add_number(s, x, y + 58, card_w, 16, phone, size=10, color=bk.NAVY, weight="medium")
        bk.add_text(s, x, y + 76, card_w, 16, email, typeface=bk.RALEWAY, size=9, color=bk.NAVY)
    bk.footer_chrome(s, 14)
    return s


# ───────────────── 15. Cierre ───────────────────────────────────────────
def slide_cierre(prs):
    s = bk.new_slide(prs)
    bk.set_gradient_bg(s, bk.TEAL, bk.NAVY, angle=45)
    bk.logo_picture(s, bk.MARGIN, 60, 22, dark_bg=True)
    bk.add_text(s, bk.MARGIN, 190, 700, 60, "Better decisions,\ntogether.",
                typeface=bk.RALEWAY_LIGHT, size=34, color=bk.WHITE, line_spacing=1.1)
    bk.add_text(s, bk.MARGIN, 310, 700, 60,
                "Gracias por tu tiempo. Estaremos encantados de concretar cómo podemos "
                "acompañar al fondo y a sus participadas.",
                typeface=bk.RALEWAY, size=13, color=bk.WHITE, line_spacing=1.3)
    bk.add_text(s, bk.MARGIN, 478, 500, 18, "tesseraservices.com - LinkedIn @tesseraservices",
                typeface=bk.RALEWAY, size=9, color=bk.TEAL_LIGHT)
    return s


def build():
    prs = Presentation()
    prs.slide_width = Pt(bk.SLIDE_W)
    prs.slide_height = Pt(bk.SLIDE_H)

    slide_portada(prs)
    slide_pagina_todas(prs)
    slide_pagina_madrid(prs)
    slide_equipo_sin_eduardo(prs)
    slide_equipo_con_eduardo(prs)
    slide_para_el_fondo(prs)
    slide_servicios_finance(prs)
    slide_servicios_hc_resumen(prs)
    slide_por_que_tessera(prs)
    slide_terminos_1(prs)
    slide_terminos_2(prs)
    slide_honorarios(prs)
    slide_partners(prs)
    slide_contacto(prs)
    slide_cierre(prs)

    prs.save(str(OUT))
    print(f"Guardado {OUT} con {len(prs.slides)} diapositivas")
    for slide_key, shapes in SHAPE_LOG.items():
        print(f"  [{slide_key}]")
        for label, val in shapes.items():
            print(f"    {label}: {val}")


if __name__ == "__main__":
    build()
