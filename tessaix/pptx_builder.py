"""Ensamblado del .pptx a partir de las plantillas de Tessera.

Las dos plantillas maestras (ver config.py: `PLANTILLA_HC` y
`PLANTILLA_HC_FINANCE`) son, literalmente, los archivos reales de la
identidad 2026 que dio el equipo (ver scripts/build_templates/rebuild_v3.py)
— no una reconstrucción a mano. `PLANTILLA_HC` sale de "nueva versi con fee
parter" (equipo + fee "Modalidad Partner"); `PLANTILLA_HC_FINANCE` de
"Catalogo_Servicios_Tessera_Ejecutivo". Esto importa porque el texto de
cada diapositiva de servicio YA ES el texto final de marca — los
marcadores "antiguos" que este módulo busca y sustituye son, literalmente,
el texto tal cual aparece en esos archivos, no una paráfrasis.

Cada plantilla es un SUPERCONJUNTO: trae ya dentro los bloques opcionales
(equipo con/sin Eduardo Serrano, términos y condiciones + honorarios)
como diapositivas completas, y este módulo decide en tiempo de generación
cuáles conservar y cuáles quitar, en vez de tener que maquetar nada desde
cero. La única alternativa que NO viene de fábrica en ningún archivo real
es "solo Madrid" para plantilla_hc.pptx — se clonó la diapositiva real de
sedes y se le quitaron Bilbao/Oviedo (ver rebuild_v3.py); Human Capital +
Finance no tiene esa alternativa (decisión del equipo).

Las diapositivas de "Índice" y "Contexto" son la excepción: no vienen en
la plantilla, así que se CLONAN a partir de una diapositiva ya existente
que tiene el "cascarón" correcto (logo, línea, texto de sección, título,
subtítulo, número de página) — se le quitan las tarjetas propias de esa
diapositiva y se reutiliza el subtítulo como cuerpo de texto, dejando que
el motor de ajuste (text_fit) lo redimensione. Así no hay que maquetar
nada nuevo a mano ni arriesgarse a que no case con el estilo del resto.
"""
from __future__ import annotations

import datetime
import io
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pptx import Presentation

from .ai_content import generate_context_paragraph
from .config import BUSINESS_LINE_HC_FINANCE, PLANTILLA_HC, PLANTILLA_HC_FINANCE, SERVICES_HC
from .logo_client import compose_logo_for_placeholder
from .team import LOCATIONS_MADRID_ONLY
from .xml_utils import (
    center_vertical_anchor,
    enable_shrink_autofit,
    find_pic_embed_rid,
    get_picture_frame_emu,
    get_shape_ext_emu,
    remove_content_type_override,
    remove_relationship,
    remove_shape_by_id,
    remove_sld_id,
    replace_text_and_fit,
    replace_text_in_xml,
    resolve_rel_target,
    resolve_rid_for_target,
    shift_shape_y,
)

# ─── SLIDES DE CADA PLANTILLA (numeración fija, 1-indexada) ───────────────
# plantilla_hc.pptx (18 diapositivas antes de añadir Índice/Contexto):
#  1 portada · 2 equipo SIN Eduardo Serrano (con fee) / 3 equipo CON él
#  (sin fee) — alternativa según deck_type, no hay nada que elegir a mano
#  · 4 sedes (3 ciudades) · 4b sedes (solo Madrid, clonada — no viene de
#  fábrica) · 5 "Así vamos a trabajar juntos" (por qué Tessera, igual en
#  los dos modelos de fee) · 6 "No somos un proveedor de CVs" (comparativa,
#  SIEMPRE se conserva) · 7 headhunting · 8 outsourcing · 9 salarial
#  · 10 formación · 11 Tessera Services (estático) · 12 Próximos pasos
#  (solo con fee) · 13-14 términos y condiciones (solo con fee) · 15
#  honorarios — Modalidad Partner (solo con fee) · 16 confían en nosotros
#  · 17 cierre
HC_SLIDES_DYNAMIC_TEXT = ["slide5.xml", "slide7.xml", "slide8.xml", "slide9.xml", "slide10.xml"]
HC_LOCATION_SLIDES = {"all": "slide4.xml", "madrid": "slide4b.xml"}
HC_TEAM_SLIDES = {"sin_eduardo": "slide2.xml", "con_eduardo": "slide3.xml"}
HC_NEXT_STEPS_SLIDE = "slide12.xml"
HC_TERMS_SLIDES = ["slide13.xml", "slide14.xml"]
HC_FEE_SLIDE = "slide15.xml"
HC_SERVICE_SLIDES = {"headhunting": "slide7.xml", "outsourcing": "slide8.xml",
                      "salary": "slide9.xml", "formacion": "slide10.xml"}

# plantilla_hc_finance.pptx (11 diapositivas antes de añadir Índice/Contexto):
#  1 portada · 2 "Tessera en una página" (una sola versión, sin alternativa
#  de sedes — decisión del equipo) · 3/4 equipo (sin/con Eduardo Serrano,
#  se conserva una) · 5 para el fondo · 6 servicios Finance
#  · 7 servicios Human Capital (resumen) · 8 por qué Tessera
#  · 9 confían en nosotros · 10 contacto · 11 cierre
# Human Capital + Finance no tiene ningún flujo de fee/honorarios ni de
# términos y condiciones (el archivo real nunca los trajo).
HCF_TEAM_SLIDES = {"no_eduardo": "slide3.xml", "con_eduardo": "slide4.xml"}

# Texto exacto (tal y como está en la plantilla real) del bloque de
# garantía y de la firma — solo en plantilla_hc.pptx, la única con flujo
# de fee.
_WARRANTY_OLD = "dentro de los 3 meses siguientes a la fecha de contratación"
# La plantilla trae una fecha de ejemplo real (no un token) en la firma y
# en "Firma del acuerdo de colaboración" — se sustituye por la de hoy.
_DATE_OLD = "08/10/2026"
# "8. VIGENCIA DE ESTE ACUERDO": un año desde la fecha de la propuesta,
# en formato largo español ("el 8 de octubre de 2027").
_VIGENCIA_OLD = "hasta el 8 de octubre de 2027"
_MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# "Modalidad Partner" (diapositiva de Honorarios): fee_rate% sobre un SBA
# de referencia, en dos plazos (1/3 a la firma, 2/3 ajustable a la
# incorporación). Los textos de abajo son el ejemplo de fábrica (12% sobre
# 36.000 €) tal cual aparece en la plantilla — varias formas distintas,
# cada una con su propio marcador exacto (ver _modalidad_partner_subs).
_FEE_RATE_OLD = "12%"
_FEE_ANTICIPO_OLD = "1/3     1.440 €"
_FEE_SEGUNDO_OLD = "2/3     2.880 € ajustable"
_FEE_CONDICIONES_OLD = "Si el SBA real es mayor o menor de 36.000 €, se regulariza en los 2/3."
# "36.000" y "4.320" aparecen en una frase ("...de 36.000 € (4.320 €)")
# donde el símbolo € quedó mal codificado en el archivo real (un carácter
# de reemplazo en vez de €) — como texto, no coincide con un "€" normal.
# En vez de intentar adivinar esa codificación exacta, se sustituyen los
# dos números sueltos directamente por su <a:t> (ver _replace_isolated_numbers).
_SBA_REF_NUMBER_OLD = "36.000"
_SBA_FEE_TOTAL_NUMBER_OLD = "4.320"

# ─── DONANTES PARA CLONAR "ÍNDICE" Y "CONTEXTO" ───────────────────────────
# Cada plantilla tiene ya una diapositiva con el "cascarón" que se necesita
# (logo, texto de sección, título grande, subtítulo, número de página): en
# plantilla_hc.pptx es "Así vamos a trabajar juntos" (slide5), en
# plantilla_hc_finance.pptx "Por qué Tessera" (slide8). Se clona dos veces,
# se le quitan las tarjetas propias de esa diapositiva (queda solo el
# título y el subtítulo, reutilizado como cuerpo) y se reescribe el
# contenido.
_HC_DONOR = {
    "slide": "slide5.xml",
    "prune_ids": [7, 9, 10, 11, 13, 14, 15, 17, 18, 19, 20, 21, 22, 24, 25, 27, 28],
    "eyebrow_old": "NUESTRA FUERZA",
    "title_old": "Así vamos a trabajar juntos",
    "body_old": "No somos la consultora que envía CVs y desaparece: trabajamos como parte de tu equipo.",
}
_HCF_DONOR = {
    "slide": "slide8.xml",
    "prune_ids": [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28],
    "eyebrow_old": "NOSOTROS",
    "title_old": "Por qué Tessera",
    "body_old": "Control, rigor financiero y talento al servicio de la tesis de inversión.",
}

_INDEX_SLIDE_FILE = "slide90.xml"
_CONTEXT_SLIDE_FILE = "slide91.xml"

# El cuerpo donante es un subtítulo de una sola línea (pensado para una
# frase corta), así que de fábrica su caja es minúscula. Tanto el Índice
# (8-10 líneas) como el Contexto (un párrafo de varias frases) lo
# forzarían a encogerse hasta quedar diminuto; a los dos les damos
# explícitamente más tamaño de partida, un suelo más alto y mucho más alto
# de caja disponible (hay hueco de sobra entre el título y el número de
# página, una vez quitadas las tarjetas propias de la diapositiva donante).
_CLONE_BODY_BASE_SCALE = 1.3
_CLONE_BODY_MIN_FONT_PT = 14.0
_CLONE_BODY_PRESET_HEIGHT_PT = 330.0

# A diferencia de la plantilla anterior, aquí el título y el cuerpo de
# cada tarjeta son DOS formas independientes apiladas (no un único cuadro
# con dos párrafos) — así es como vienen en el archivo real.
#
# En Headhunting el hueco entre una forma y la siguiente es mínimo (están
# pensadas para un título de una palabra y una frase corta): el tope de
# crecimiento es el propio alto de fábrica, así que nunca crecen —
# cualquier texto más largo se encoge en vez de arriesgarse a pisar la
# forma de abajo.
_HH_TITLE_CEILING_PT = 18.0
_HH_BODY_CEILING_PT = 40.0

# En "Así vamos a trabajar juntos" el título SÍ tiene hueco de sobra de
# fábrica (pensado para 1-2 líneas), así que se le deja crecer un poco más
# — pero si aun así crece, hay que desplazar el cuerpo hacia abajo lo
# mismo que haya crecido el título, o se solapan (ver _replace_title_push).
_WHY_TITLE_CEILING_PT = 50.0
_WHY_BODY_CEILING_PT = 85.0


def _read(path: Path) -> str:
    return path.read_text("utf-8")


def _write(path: Path, content: str) -> None:
    path.write_text(content, "utf-8")


# ─── UTILIDADES DE PAQUETE (borrar / duplicar diapositivas) ───────────────
def _remove_slides(unpack: Path, slide_files: list[str]) -> None:
    """
    Quita por completo una o varias diapositivas del paquete .pptx: la
    referencia en <p:sldIdLst> de presentation.xml, la relación en
    presentation.xml.rels, el Override en [Content_Types].xml y el propio
    archivo de la diapositiva (con sus _rels). Es la forma correcta de
    borrar una diapositiva a nivel de paquete OOXML.
    """
    pres_path = unpack / "ppt" / "presentation.xml"
    pres_rels_path = unpack / "ppt" / "_rels" / "presentation.xml.rels"
    ct_path = unpack / "[Content_Types].xml"

    presentation_xml = _read(pres_path)
    pres_rels_xml = _read(pres_rels_path)
    content_types_xml = _read(ct_path)

    for slide_file in slide_files:
        target = f"slides/{slide_file}"
        rid = resolve_rid_for_target(pres_rels_xml, target)
        if rid:
            presentation_xml = remove_sld_id(presentation_xml, rid)
            pres_rels_xml = remove_relationship(pres_rels_xml, rid)
        content_types_xml = remove_content_type_override(content_types_xml, f"/ppt/slides/{slide_file}")

        slide_path = unpack / "ppt" / "slides" / slide_file
        slide_rels_path = unpack / "ppt" / "slides" / "_rels" / f"{slide_file}.rels"
        slide_path.unlink(missing_ok=True)
        slide_rels_path.unlink(missing_ok=True)

    _write(pres_path, presentation_xml)
    _write(pres_rels_path, pres_rels_xml)
    _write(ct_path, content_types_xml)


def _duplicate_slide(unpack: Path, source_slide_file: str, new_slide_file: str, insert_after_file: str) -> None:
    """
    Clona una diapositiva existente (XML + relaciones) y la inserta en el
    orden visual justo después de `insert_after_file`. No duplica media
    nueva: comparte las relaciones de la original (layout, imágenes
    decorativas) — con eso basta para clonar un "cascarón" ya con el
    estilo correcto, listo para que quien llame sustituya su contenido.

    Las relaciones a un diseño o a una imagen pueden compartirse sin
    problema entre varias diapositivas. La relación a una página de NOTAS
    del orador, no: cada diapositiva de notas pertenece a UNA sola
    diapositiva, así que si se copia tal cual, la diapositiva original y
    el clon quedan ambas "dueñas" de la misma página de notas — un paquete
    así de ambiguo falla la validación estricta de PowerPoint (ya no solo
    pide reparar: el archivo queda inestable y se cierra solo al poco de
    abrirlo). Por eso esa relación en concreto se quita al clonar.
    """
    slides_dir = unpack / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    shutil.copy(slides_dir / source_slide_file, slides_dir / new_slide_file)
    src_rels = rels_dir / f"{source_slide_file}.rels"
    if src_rels.exists():
        rels_xml = _read(src_rels)
        rels_xml = re.sub(
            r'<Relationship\b[^>]*\bType="[^"]*/relationships/notesSlide"[^>]*/>', '', rels_xml,
        )
        _write(rels_dir / f"{new_slide_file}.rels", rels_xml)

    ct_path = unpack / "[Content_Types].xml"
    ct = _read(ct_path)
    override = (
        f'<Override PartName="/ppt/slides/{new_slide_file}" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    )
    _write(ct_path, ct.replace("</Types>", override + "</Types>"))

    pres_rels_path = unpack / "ppt" / "_rels" / "presentation.xml.rels"
    pres_rels = _read(pres_rels_path)
    existing_rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels)]
    new_rid = f"rId{max(existing_rids) + 1}"
    relationship = (
        f'<Relationship Id="{new_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/{new_slide_file}"/>'
    )
    pres_rels = pres_rels.replace("</Relationships>", relationship + "</Relationships>")
    _write(pres_rels_path, pres_rels)

    pres_path = unpack / "ppt" / "presentation.xml"
    pres = _read(pres_path)
    existing_ids = [int(m) for m in re.findall(r'<p:sldId\s+id="(\d+)"', pres)]
    new_id = max(existing_ids) + 1
    new_sldid = f'<p:sldId id="{new_id}" r:id="{new_rid}"/>'
    insert_after_rid = resolve_rid_for_target(pres_rels, f"slides/{insert_after_file}")
    anchor = re.search(r'<p:sldId\b[^>]*\br:id="%s"[^>]*/>' % re.escape(insert_after_rid), pres)
    pres = pres[:anchor.end()] + new_sldid + pres[anchor.end():]
    _write(pres_path, pres)


def _renumber_visible_pages(pptx_bytes: bytes) -> bytes:
    """
    Tras quitar o añadir diapositivas, el número de página visible en la
    esquina inferior derecha deja de ser correlativo. Se recorren las
    diapositivas ya ensambladas y se renumeran de forma secuencial las que
    tengan una forma con esa pinta (texto de 1-3 dígitos, en la esquina
    inferior derecha) — la portada y el cierre normalmente no llevan
    número, así que se saltan solas.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide_w, slide_h = prs.slide_width, prs.slide_height
    visible = 0
    for slide in prs.slides:
        for shp in slide.shapes:
            if shp.left is None or not shp.has_text_frame:
                continue
            txt = shp.text_frame.text.strip()
            if not re.fullmatch(r"0?\d{1,2}", txt) or txt in ("2025", "2026"):
                continue
            if shp.left < slide_w * 0.6 or shp.top < slide_h * 0.8:
                continue
            visible += 1
            new_txt = f"{visible:02d}"
            tf = shp.text_frame
            first = True
            for p in tf.paragraphs:
                for r in p.runs:
                    if first:
                        r.text = new_txt
                        first = False
                    else:
                        r.text = ""
            break
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── ÍNDICE Y CONTEXTO (compartido) ────────────────────────────────────
def _build_info_slide(
    unpack: Path, donor: dict, new_slide_file: str, insert_after_file: str,
    eyebrow: str, title: str, body: str,
    body_base_scale: float = 1.0, body_min_font_pt: float | None = None,
    body_preset_height_pt: float | None = None,
) -> None:
    _duplicate_slide(unpack, donor["slide"], new_slide_file, insert_after_file)
    path = unpack / "ppt" / "slides" / new_slide_file
    s = _read(path)
    for shape_id in donor["prune_ids"]:
        s = remove_shape_by_id(s, shape_id)
    s = replace_text_in_xml(s, donor["eyebrow_old"], eyebrow)
    s = replace_text_and_fit(s, donor["title_old"], title, role="title")
    # Alineado a la izquierda (igual que el resto de títulos/cuerpos de la
    # plantilla) — solo se centra verticalmente el cuerpo dentro de su
    # caja, que es mucho más alta que el texto corto del donante.
    s = center_vertical_anchor(s, donor["body_old"])
    s = replace_text_and_fit(
        s, donor["body_old"], body,
        base_scale=body_base_scale, min_font_pt=body_min_font_pt, preset_height_pt=body_preset_height_pt,
    )
    s = enable_shrink_autofit(s)
    _write(path, s)


def _index_items(data: dict) -> list[str]:
    is_hc_finance = data.get("business_line") == BUSINESS_LINE_HC_FINANCE
    include_fee = data.get("deck_type") == "propuesta"
    if is_hc_finance:
        # Human Capital + Finance ya no tiene flujo de fee: su índice nunca
        # incluye "Honorarios".
        items = [
            "Contexto", "Tessera en una página", "Nuestro equipo",
            "Para el fondo", "Servicios Finance", "Servicios Human Capital",
            "Por qué Tessera", "Confían en nosotros", "Contacto",
        ]
    else:
        items = ["Contexto", "Nuestro equipo", "Dónde estamos", "Por qué Tessera", "Cómo trabajamos"]
        items += [SERVICES_HC.get(s, s) for s in data.get("services", [])]
        if include_fee:
            items += ["Próximos pasos", "Honorarios"]
        items.append("Confían en nosotros")
    return items


def _build_index_and_context(data: dict, unpack: Path, donor: dict) -> None:
    items = _index_items(data)
    index_body = "\n".join(f"{i:02d}.  {name}" for i, name in enumerate(items, 1))
    _build_info_slide(
        unpack, donor, _INDEX_SLIDE_FILE, "slide1.xml",
        eyebrow="NOSOTROS", title="Índice", body=index_body,
        body_base_scale=_CLONE_BODY_BASE_SCALE, body_min_font_pt=_CLONE_BODY_MIN_FONT_PT,
        body_preset_height_pt=_CLONE_BODY_PRESET_HEIGHT_PT,
    )

    try:
        context_body = generate_context_paragraph(data)
    except Exception:
        context_body = None
    context_body = context_body or (
        f"Tessera trabaja junto a {data.get('client_name', 'tu empresa')} para responder "
        f"a sus necesidades actuales con un equipo senior y un enfoque directo."
    )
    _build_info_slide(
        unpack, donor, _CONTEXT_SLIDE_FILE, _INDEX_SLIDE_FILE,
        eyebrow="NOSOTROS", title="Contexto", body=context_body,
        body_base_scale=_CLONE_BODY_BASE_SCALE, body_min_font_pt=_CLONE_BODY_MIN_FONT_PT,
        body_preset_height_pt=_CLONE_BODY_PRESET_HEIGHT_PT,
    )


# ─── TÉRMINOS Y CONDICIONES + HONORARIOS (solo Human Capital) ─────────────
def _fmt_eur(value: float) -> str:
    """1440.0 -> '1.440 €' (separador de miles con punto, sin decimales —
    así es como aparece ya escrito en la plantilla)."""
    return f"{value:,.0f}".replace(",", ".") + " €"


def _parse_number(raw: str | None) -> float | None:
    """Admite tanto '36000' como '36.000' o '36.000,50' (formato español)."""
    if not raw:
        return None
    cleaned = str(raw).strip().replace(".", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        try:
            return float(raw)
        except ValueError:
            return None


def _spanish_long_date(d: datetime.date) -> str:
    """08/10/2026 -> '8 de octubre de 2026' (sin depender del locale del sistema)."""
    return f"{d.day} de {_MESES_ES[d.month - 1]} de {d.year}"


def _replace_isolated_numbers(slide_xml: str, pairs: list[tuple[str, str]]) -> str:
    """Sustituye un número que está solo, en su propio <a:t>...</a:t>, sin
    tocar el resto del párrafo — hace falta cuando el texto que lo rodea no
    se puede buscar literal de forma fiable (ver _SBA_REF_NUMBER_OLD)."""
    for old, new in pairs:
        slide_xml = re.sub(r'(<a:t>)%s(</a:t>)' % re.escape(old), rf'\g<1>{new}\g<2>', slide_xml)
    return slide_xml


def _modalidad_partner_subs(fee_rate: float, reference_salary: float) -> list[tuple[str, str]]:
    """"Modalidad Partner": fee_rate% de un SBA de referencia, pagado en dos
    plazos — 1/3 a la firma (anticipo fijo) y 2/3 a la incorporación,
    ajustable al SBA real del candidato. La plantilla real reparte estos
    números en varias formas distintas (el "12%" grande, cada plazo, la
    nota de condiciones) en vez de un único bloque de texto — cada
    marcador de abajo es el texto exacto de una de esas formas."""
    fee_total = reference_salary * fee_rate / 100
    anticipo = fee_total / 3
    segundo_pago = fee_total - anticipo
    rate_label = f"{fee_rate:g}"
    salary_fmt = _fmt_eur(reference_salary)
    return [
        (_FEE_RATE_OLD, f"{rate_label}%"),
        (_FEE_ANTICIPO_OLD, f"1/3     {_fmt_eur(anticipo)}"),
        (_FEE_SEGUNDO_OLD, f"2/3     {_fmt_eur(segundo_pago)} ajustable"),
        (_FEE_CONDICIONES_OLD, f"Si el SBA real es mayor o menor de {salary_fmt}, se regulariza en los 2/3."),
    ]


def _apply_terms_and_fee(
    unpack: Path, terms_slides: list[str], fee_slide: str, next_steps_slide: str,
    include_fee: bool, warranty_months: int, fee_rate: str | None, reference_salary: str | None,
) -> None:
    slides = unpack / "ppt" / "slides"

    if not include_fee:
        _remove_slides(unpack, terms_slides + [fee_slide, next_steps_slide])
        return

    warranty_new = _WARRANTY_OLD.replace("3 meses", f"{warranty_months} meses")
    today = datetime.date.today()
    today_fmt = today.strftime("%d/%m/%Y")
    vigencia_date = today.replace(year=today.year + 1) if not (today.month == 2 and today.day == 29) \
        else today.replace(year=today.year + 1, day=28)

    for slide_file in terms_slides:
        path = slides / slide_file
        if not path.exists():
            continue
        s = _read(path)
        s = replace_text_in_xml(s, _WARRANTY_OLD, warranty_new)
        s = replace_text_in_xml(s, _DATE_OLD, today_fmt)
        s = replace_text_in_xml(s, _VIGENCIA_OLD, f"hasta el {_spanish_long_date(vigencia_date)}")
        _write(path, s)

    if fee_rate:
        rate_f = _parse_number(fee_rate)
        salary_f = _parse_number(reference_salary)
        if rate_f is not None and salary_f is not None:
            path = slides / fee_slide
            if path.exists():
                s = _read(path)
                for old, new in _modalidad_partner_subs(rate_f, salary_f):
                    s = replace_text_in_xml(s, old, new)
                s = _replace_isolated_numbers(s, [
                    (_SBA_REF_NUMBER_OLD, _fmt_eur(salary_f).replace(" €", "")),
                    (_SBA_FEE_TOTAL_NUMBER_OLD, _fmt_eur(rate_f * salary_f / 100).replace(" €", "")),
                ])
                _write(path, s)


def _replace_title_push_body(
    slide_xml: str, title_old: str, title_new: str, title_id: int, body_id: int, grow_ceiling_pt: float,
) -> str:
    """Sustituye el título (puede crecer hasta `grow_ceiling_pt`, ver
    replace_text_and_fit) y, si de verdad ha crecido, desplaza el cuerpo
    (una forma independiente, no un párrafo dentro del mismo cuadro) hacia
    abajo exactamente lo que haya crecido — si no, un título de dos líneas
    se come la primera línea del cuerpo."""
    before = get_shape_ext_emu(slide_xml, title_id)
    slide_xml = replace_text_and_fit(slide_xml, title_old, title_new, role="title", grow_ceiling_pt=grow_ceiling_pt)
    after = get_shape_ext_emu(slide_xml, title_id)
    if before and after and after[1] > before[1]:
        slide_xml = shift_shape_y(slide_xml, body_id, after[1] - before[1])
    return slide_xml


# ─── HUMAN CAPITAL (solo) ──────────────────────────────────────────────
def _build_hc(data: dict, content: dict, unpack: Path) -> None:
    slides = unpack / "ppt" / "slides"
    r = replace_text_and_fit

    _build_index_and_context(data, unpack, _HC_DONOR)

    show_all_locations = data.get("locations", "all") != LOCATIONS_MADRID_ONLY
    drop_location = HC_LOCATION_SLIDES["madrid" if show_all_locations else "all"]

    # Equipo: ya no se pregunta en el wizard — lo decide el tipo de
    # propuesta. Sin fee (presentación) va con Eduardo Serrano; con fee
    # (propuesta) va sin él, de momento (ver conversación con el equipo).
    include_fee = data.get("deck_type") == "propuesta"
    drop_team = HC_TEAM_SLIDES["con_eduardo" if include_fee else "sin_eduardo"]

    _remove_slides(unpack, [drop_location, drop_team])

    why = content.get("why_tessera", {})
    svcs = content.get("services", {})
    services = data.get("services", [])

    # ── Así vamos a trabajar juntos (por qué Tessera) ────────────────
    # Título y cuerpo de cada uno de los 3 pasos son DOS formas
    # independientes apiladas (no un único cuadro) — así vienen en el
    # archivo real. El título tiene hueco de fábrica para crecer un poco;
    # si crece, el cuerpo se desplaza hacia abajo lo mismo (ver
    # _replace_title_push_body) para no solaparse.
    if why:
        s = _read(slides / "slide5.xml")
        if why.get("d1_title"): s = _replace_title_push_body(s, "Inmersión", why["d1_title"], 10, 11, _WHY_TITLE_CEILING_PT)
        if why.get("d1_body"):  s = r(s, "Arrancamos conociendo tu equipo, tu cultura y tus retos. Con eso, primer shortlist filtrado en 72 horas.", why["d1_body"], grow_ceiling_pt=_WHY_BODY_CEILING_PT)
        if why.get("d2_title"): s = _replace_title_push_body(s, "En cada paso", why["d2_title"], 14, 15, _WHY_TITLE_CEILING_PT)
        if why.get("d2_body"):  s = r(s, "Seguimiento semanal contigo. Un interlocutor fijo y un equipo de respaldo: tus vacantes nunca se quedan paradas.", why["d2_body"], grow_ceiling_pt=_WHY_BODY_CEILING_PT)
        if why.get("d3_title"): s = _replace_title_push_body(s, "Siempre presentes", why["d3_title"], 18, 19, _WHY_TITLE_CEILING_PT)
        if why.get("d3_body"):  s = r(s, "Tras la incorporación seguimos ahí: acompañamos a la persona y a tu equipo.", why["d3_body"], grow_ceiling_pt=_WHY_BODY_CEILING_PT)
        _write(slides / "slide5.xml", s)

    # ── Headhunting ──────────────────────────────────────────────────
    # 3 columnas, cada una con 3 pares título+cuerpo apilados — aquí el
    # hueco de fábrica es mínimo (pensado para una palabra y una frase
    # corta), así que el tope de crecimiento es el propio alto de fábrica:
    # nunca crecen, cualquier texto más largo se encoge en su sitio.
    ht, hb = _HH_TITLE_CEILING_PT, _HH_BODY_CEILING_PT
    if "headhunting" in services and svcs.get("headhunting"):
        hh = svcs["headhunting"]
        s = _read(slides / "slide7.xml")
        if hh.get("why_col_title1"): s = r(s, "VELOCIDAD", hh["why_col_title1"], role="title", grow_ceiling_pt=ht)
        if hh.get("why_col_body1"):  s = r(s, "Disponemos de nuestra propia metodología, la cual nos permite presentar candidatos a tiempo.", hh["why_col_body1"], grow_ceiling_pt=hb)
        if hh.get("why_col_title2"): s = r(s, "ESPECIALIZACIÓN", hh["why_col_title2"], role="title", grow_ceiling_pt=ht)
        if hh.get("why_col_body2"):  s = r(s, "Conocemos tu sector y cómo se construyen sus equipos.", hh["why_col_body2"], grow_ceiling_pt=hb)
        if hh.get("why_col_title3"): s = r(s, "ÉXITO COMPARTIDO", hh["why_col_title3"], role="title", grow_ceiling_pt=ht)
        if hh.get("why_col_body3"):  s = r(s, "El equipo que lleva tu negocio al siguiente nivel.", hh["why_col_body3"], grow_ceiling_pt=hb)
        if hh.get("benefits_body1"): s = r(s, "Nos ocupamos de todo. Tú solo conoces a los mejores.", hh["benefits_body1"], grow_ceiling_pt=hb)
        if hh.get("benefits_body2"): s = r(s, "Candidatos que no solo encajan: impulsan tu crecimiento.", hh["benefits_body2"], grow_ceiling_pt=hb)
        if hh.get("benefits_body3"): s = r(s, "Tu éxito es también el nuestro.", hh["benefits_body3"], grow_ceiling_pt=hb)
        if hh.get("how_body3"): s = r(s, "Habilidades, actitud y encaje cultural.", hh["how_body3"], grow_ceiling_pt=hb)
        if hh.get("how_body4"): s = r(s, "Únicamente candidatos que realmente suman.", hh["how_body4"], grow_ceiling_pt=hb)
        if hh.get("how_body5"): s = r(s, "Contigo también después de la incorporación.", hh["how_body5"], grow_ceiling_pt=hb)
        _write(slides / "slide7.xml", s)

    # ── Outsourcing ──────────────────────────────────────────────────
    if "outsourcing" in services and svcs.get("outsourcing"):
        outs = svcs["outsourcing"]
        s = _read(slides / "slide8.xml")
        if outs.get("headline"): s = r(s, "Externalización que sí funciona: pagas por trabajo real, sin papeleos.", outs["headline"])
        if outs.get("body"):     s = r(s, "Contratar cuesta más de lo que parece. Con el modelo Time & Material ajustas el equipo a tu actividad real y nosotros nos encargamos de toda la gestión.", outs["body"])
        if outs.get("card1_body"): s = r(s, "Pagas solo por horas reales de trabajo.", outs["card1_body"])
        if outs.get("card2_body"): s = r(s, "Cubrimos bajas y vacaciones.", outs["card2_body"])
        if outs.get("card3_body"): s = r(s, "Nóminas y trámites, a cargo nuestro.", outs["card3_body"])
        if outs.get("card4_body"): s = r(s, "Escalas el equipo cuando quieras.", outs["card4_body"])
        _write(slides / "slide8.xml", s)

    # ── Consultoría Salarial ─────────────────────────────────────────
    if "salary" in services and svcs.get("salary"):
        sal = svcs["salary"]
        s = _read(slides / "slide9.xml")
        if sal.get("headline"): s = r(s, "Inteligencia retributiva para decisiones que importan.", sal["headline"])
        if sal.get("body"):     s = r(s, "Acompañamos a organizaciones en el diseño de estructuras salariales coherentes, alineadas con el mercado y preparadas para el nuevo marco regulatorio europeo.", sal["body"])
        _write(slides / "slide9.xml", s)

    # ── Formación ────────────────────────────────────────────────────
    if "formacion" in services and svcs.get("formacion"):
        form = svcs["formacion"]
        s = _read(slides / "slide10.xml")
        if form.get("body"): s = r(s, "Potenciamos el talento interno: consultoría en formación orientada a procesos y resultados ágiles.", form["body"])
        _write(slides / "slide10.xml", s)

    # ── Diapositivas de servicio no seleccionadas: se quitan ─────────
    to_remove = [fn for key, fn in HC_SERVICE_SLIDES.items() if key not in services]
    if to_remove:
        _remove_slides(unpack, to_remove)

    # ── Autofit dinámico en las diapositivas con texto variable ──────
    for name in HC_SLIDES_DYNAMIC_TEXT:
        path = slides / name
        if path.exists():
            _write(path, enable_shrink_autofit(_read(path)))

    _apply_terms_and_fee(
        unpack, HC_TERMS_SLIDES, HC_FEE_SLIDE, HC_NEXT_STEPS_SLIDE,
        include_fee=include_fee,
        warranty_months=data.get("warranty_months", 3),
        fee_rate=data.get("fee_rate"),
        reference_salary=data.get("reference_salary"),
    )


# ─── HUMAN CAPITAL + FINANCE ────────────────────────────────────────────
def _build_hc_finance(data: dict, content: dict, unpack: Path) -> None:
    # Human Capital + Finance no tiene ningún flujo de fee/honorarios ni de
    # términos y condiciones (el archivo real nunca los trajo) y tampoco
    # alternativa de sedes (decisión del equipo) — el contenido es fijo,
    # solo cambia el equipo.
    _build_index_and_context(data, unpack, _HCF_DONOR)

    include_eduardo = bool(data.get("include_eduardo"))
    drop_team = HCF_TEAM_SLIDES["no_eduardo" if include_eduardo else "con_eduardo"]

    _remove_slides(unpack, [drop_team])


# ─── LOGO DEL CLIENTE (portada + cierre) ──────────────────────────────────
# Qué diapositiva/forma sustituir por línea de negocio — el hueco (un
# <p:pic> de relleno transparente) se añadió a las plantillas con
# scripts/build_templates/rebuild_v3.py; estos ids son los que asignó ese
# script en cada plantilla.
_LOGO_SLOTS: dict[str, list[tuple[str, int]]] = {
    "hc": [("slide1.xml", 8), ("slide17.xml", 10)],
    "hc_finance": [("slide1.xml", 10), ("slide11.xml", 9)],
}


def _replace_client_logo(unpack: Path, slide_file: str, shape_id: int, logo_bytes: bytes) -> None:
    """
    Sustituye la imagen de la forma `shape_id` (un <p:pic>, el hueco de
    logo del cliente) por `logo_bytes`, compuesto sin deformar y sin caja
    de color alrededor (ver `compose_logo_for_placeholder`). Solo toca el
    archivo de media al que ya apunta esa forma — el tamaño/posición del
    hueco en la diapositiva no cambia.
    """
    slide_path = unpack / "ppt" / "slides" / slide_file
    if not slide_path.exists():
        return
    slide_xml = _read(slide_path)
    rid = find_pic_embed_rid(slide_xml, shape_id)
    if not rid:
        return
    frame = get_picture_frame_emu(slide_xml, rid)
    if not frame:
        return
    rels_path = unpack / "ppt" / "slides" / "_rels" / f"{slide_file}.rels"
    target = resolve_rel_target(_read(rels_path), rid)
    if not target:
        return
    media_path = (slide_path.parent / target).resolve()
    if not media_path.exists():
        return
    composed = compose_logo_for_placeholder(logo_bytes, frame[0], frame[1])
    media_path.write_bytes(composed)


# ─── PUNTO DE ENTRADA ────────────────────────────────────────────────────
def build_pptx(data: dict, content: dict, logo_bytes: bytes | None = None) -> bytes:
    business_line = data.get("business_line", "hc")
    plantilla = PLANTILLA_HC_FINANCE if business_line == BUSINESS_LINE_HC_FINANCE else PLANTILLA_HC

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        work = tmp / "work.pptx"
        shutil.copy(str(plantilla), str(work))

        unpack = tmp / "unpacked"
        with zipfile.ZipFile(work, 'r') as z:
            z.extractall(unpack)

        if business_line == BUSINESS_LINE_HC_FINANCE:
            _build_hc_finance(data, content, unpack)
        else:
            _build_hc(data, content, unpack)

        if logo_bytes:
            for slide_file, shape_id in _LOGO_SLOTS.get(business_line, []):
                _replace_client_logo(unpack, slide_file, shape_id, logo_bytes)

        out = tmp / "output.pptx"
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fp in sorted(unpack.rglob('*')):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(unpack))

        return _renumber_visible_pages(out.read_bytes())
