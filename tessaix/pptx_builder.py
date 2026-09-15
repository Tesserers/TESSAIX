"""Ensamblado del .pptx a partir de las plantillas de Tessera.

Hay dos plantillas maestras, una por línea de negocio (ver config.py):
`PLANTILLA_HC` (Human Capital solo) y `PLANTILLA_HC_FINANCE` (Human Capital
+ Finance). Cada una es un SUPERCONJUNTO: trae ya dentro los bloques
opcionales (equipo con/sin Eduardo Serrano, sedes Madrid-sola o las 3,
términos y condiciones + honorarios) como diapositivas o formas completas,
y este módulo decide en tiempo de generación cuáles conservar y cuáles
quitar, en vez de tener que maquetar nada desde cero.
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

from .config import BUSINESS_LINE_HC_FINANCE, PLANTILLA_HC, PLANTILLA_HC_FINANCE
from .team import LOCATIONS_MADRID_ONLY
from .xml_utils import (
    enable_shrink_autofit,
    remove_content_type_override,
    remove_relationship,
    remove_shape_by_id,
    remove_sld_id,
    replace_text_and_fit,
    replace_text_in_xml,
    resolve_rid_for_target,
)

# ─── SLIDES DE CADA PLANTILLA (numeración fija, 1-indexada) ───────────────
# plantilla_hc.pptx (14 diapositivas):
#  1 portada · 2 equipo (fijo) · 3 sedes (3 ciudades)
#  4 por qué Tessera · 5 headhunting · 6 outsourcing · 7 salarial · 8 formación
#  9 Tessera Services (estático) · 10-11 términos y condiciones · 12 honorarios
#  13 confían en nosotros · 14 cierre
HC_SLIDES_DYNAMIC_TEXT = ["slide4.xml", "slide5.xml", "slide6.xml", "slide7.xml", "slide8.xml"]
HC_LOCATION_SLIDE = "slide3.xml"
HC_TERMS_SLIDES = ["slide10.xml", "slide11.xml"]
HC_FEE_SLIDE = "slide12.xml"

# plantilla_hc_finance.pptx (15 diapositivas):
#  1 portada · 2/3 "Tessera en una página" (3 sedes / solo Madrid, se
#  conserva una y se borra la otra) · 4/5 equipo (sin/con Eduardo Serrano,
#  igual: se conserva una) · 6 para el fondo · 7 servicios Finance
#  8 servicios Human Capital (resumen) · 9 por qué Tessera
#  10-11 términos y condiciones · 12 honorarios · 13 confían en nosotros
#  14 contacto · 15 cierre
HCF_LOCATION_SLIDES = {"all": "slide2.xml", "madrid": "slide3.xml"}
HCF_TEAM_SLIDES = {"no_eduardo": "slide4.xml", "con_eduardo": "slide5.xml"}
HCF_TERMS_SLIDES = ["slide10.xml", "slide11.xml"]
HCF_FEE_SLIDE = "slide12.xml"

# Dentro de la diapositiva de sedes de plantilla_hc.pptx, estos son los
# shapes de las tarjetas de Bilbao y Oviedo — se quitan para dejar solo
# Madrid (no existe, en esta plantilla, una diapositiva alternativa ya
# hecha como en la de HC+Finance).
HC_LOCATION_BILBAO_SHAPES = [41, 36, 26, 27]
HC_LOCATION_OVIEDO_SHAPES = [42, 37, 29, 30]

# Texto de "3 ciudades" que hay que ajustar cuando solo se muestra Madrid.
HC_LOCATION_CITY_COUNT_OLD = "3 CIUDADES"
HC_LOCATION_CITY_COUNT_NEW = "1 CIUDAD"
HC_LOCATION_CITIES_OLD = "Madrid · Bilbao · Oviedo"
HC_LOCATION_CITIES_NEW = "Madrid"

# Texto exacto (tal y como está en la plantilla) del bloque de garantía y
# de la firma, común a las dos plantillas — ver team.py para las opciones.
_WARRANTY_OLD = "dentro de los 3 meses siguientes a la fecha de contratación"
_DATE_PLACEHOLDER = "{dd/mm/aaaa}"
_FEE_OLD = "13%"


def _read(path: Path) -> str:
    return path.read_text("utf-8")


def _write(path: Path, content: str) -> None:
    path.write_text(content, "utf-8")


# ─── UTILIDADES DE PAQUETE (borrar diapositivas) ──────────────────────────
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


def _renumber_visible_pages(pptx_bytes: bytes) -> bytes:
    """
    Tras quitar diapositivas (sedes, equipo, términos y condiciones...) el
    número de página visible en la esquina inferior derecha deja de ser
    correlativo. Se recorren las diapositivas ya ensambladas y se
    renumeran de forma secuencial las que tengan una forma con esa pinta
    (texto de 1-3 dígitos, en la esquina inferior derecha) — la portada y
    el cierre normalmente no llevan número, así que se saltan solas.
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


# ─── TÉRMINOS Y CONDICIONES + HONORARIOS (compartido) ─────────────────────
def _apply_terms_and_fee(
    unpack: Path, terms_slides: list[str], fee_slide: str,
    include_fee: bool, warranty_months: int, fee_rate: str | None,
) -> None:
    slides = unpack / "ppt" / "slides"

    if not include_fee:
        _remove_slides(unpack, terms_slides + [fee_slide])
        return

    warranty_new = _WARRANTY_OLD.replace("3 meses", f"{warranty_months} meses")
    today = datetime.date.today().strftime("%d/%m/%Y")

    for slide_file in terms_slides:
        path = slides / slide_file
        if not path.exists():
            continue
        s = _read(path)
        s = replace_text_in_xml(s, _WARRANTY_OLD, warranty_new)
        s = replace_text_in_xml(s, _DATE_PLACEHOLDER, today)
        _write(path, s)

    if fee_rate:
        path = slides / fee_slide
        if path.exists():
            _write(path, replace_text_in_xml(_read(path), _FEE_OLD, f"{fee_rate}%"))


# ─── HUMAN CAPITAL (solo) ──────────────────────────────────────────────
def _apply_hc_location(unpack: Path, show_all: bool) -> None:
    if show_all:
        return
    path = unpack / "ppt" / "slides" / HC_LOCATION_SLIDE
    if not path.exists():
        return
    s = _read(path)
    for shape_id in HC_LOCATION_BILBAO_SHAPES + HC_LOCATION_OVIEDO_SHAPES:
        s = remove_shape_by_id(s, shape_id)
    s = replace_text_in_xml(s, HC_LOCATION_CITY_COUNT_OLD, HC_LOCATION_CITY_COUNT_NEW)
    s = replace_text_in_xml(s, HC_LOCATION_CITIES_OLD, HC_LOCATION_CITIES_NEW)
    _write(path, s)


def _build_hc(data: dict, content: dict, unpack: Path) -> None:
    slides = unpack / "ppt" / "slides"
    r = replace_text_and_fit

    _apply_hc_location(unpack, data.get("locations", "all") != LOCATIONS_MADRID_ONLY)

    why = content.get("why_tessera", {})
    svcs = content.get("services", {})
    services = data.get("services", [])

    # ── Por qué Tessera ─────────────────────────────────────────────
    if why:
        s = _read(slides / "slide4.xml")
        if why.get("d1_title"): s = r(s, "Sin CVs al azar", why["d1_title"])
        if why.get("d1_body"):  s = r(s, "No enviamos el primer CV, buscamos a quien encaja de verdad. Candidatos filtrados sobre la mesa en 72 horas.", why["d1_body"])
        if why.get("d2_title"): s = r(s, "Sin pausas", why["d2_title"])
        if why.get("d2_body"):  s = r(s, "Tu servicio no parará.", why["d2_body"])
        if why.get("d3_title"): s = r(s, "Siempre contigo", why["d3_title"])
        if why.get("d3_body"):  s = r(s, "No desaparecemos tras la incorporación: cuidamos a la persona y al cliente.", why["d3_body"])
        _write(slides / "slide4.xml", s)

    # ── Headhunting ──────────────────────────────────────────────────
    if "headhunting" in services and svcs.get("headhunting"):
        hh = svcs["headhunting"]
        s = _read(slides / "slide5.xml")
        if hh.get("why_col_title1"): s = r(s, "VELOCIDAD", hh["why_col_title1"])
        if hh.get("why_col_body1"):  s = r(s, "Disponemos de nuestra propia metodología, la cual nos permite presentar candidatos a tiempo.", hh["why_col_body1"])
        if hh.get("why_col_title2"): s = r(s, "ESPECIALIZACIÓN", hh["why_col_title2"])
        if hh.get("why_col_body2"):  s = r(s, "Conocemos tu sector y cómo se construyen sus equipos.", hh["why_col_body2"])
        if hh.get("why_col_title3"): s = r(s, "ÉXITO COMPARTIDO", hh["why_col_title3"])
        if hh.get("why_col_body3"):  s = r(s, "El equipo que lleva tu negocio al siguiente nivel.", hh["why_col_body3"])
        if hh.get("benefits_body1"): s = r(s, "Nos ocupamos de todo. Tú solo conoces a los mejores.", hh["benefits_body1"])
        if hh.get("benefits_body2"): s = r(s, "Candidatos que no solo encajan: impulsan tu crecimiento.", hh["benefits_body2"])
        if hh.get("benefits_body3"): s = r(s, "Tu éxito es también el nuestro.", hh["benefits_body3"])
        if hh.get("how_body3"): s = r(s, "Habilidades, actitud y encaje cultural.", hh["how_body3"])
        if hh.get("how_body4"): s = r(s, "Únicamente candidatos que realmente suman.", hh["how_body4"])
        if hh.get("how_body5"): s = r(s, "Contigo también después de la incorporación.", hh["how_body5"])
        _write(slides / "slide5.xml", s)

    # ── Outsourcing ──────────────────────────────────────────────────
    if "outsourcing" in services and svcs.get("outsourcing"):
        outs = svcs["outsourcing"]
        s = _read(slides / "slide6.xml")
        if outs.get("headline"): s = r(s, "Externalización que sí funciona: pagas por trabajo real, sin papeleos.", outs["headline"])
        if outs.get("body"):     s = r(s, "Contratar cuesta más de lo que parece. Con el modelo Time & Material ajustas el equipo a tu actividad real y nosotros nos encargamos de toda la gestión.", outs["body"])
        if outs.get("card1_body"): s = r(s, "Pagas solo por horas reales de trabajo.", outs["card1_body"])
        if outs.get("card2_body"): s = r(s, "Cubrimos bajas y vacaciones.", outs["card2_body"])
        if outs.get("card3_body"): s = r(s, "Nóminas y trámites, a cargo nuestro.", outs["card3_body"])
        if outs.get("card4_body"): s = r(s, "Escalas el equipo cuando quieras.", outs["card4_body"])
        _write(slides / "slide6.xml", s)

    # ── Consultoría Salarial ─────────────────────────────────────────
    if "salary" in services and svcs.get("salary"):
        sal = svcs["salary"]
        s = _read(slides / "slide7.xml")
        if sal.get("headline"): s = r(s, "Inteligencia retributiva para decisiones que importan.", sal["headline"])
        if sal.get("body"):     s = r(s, "Acompañamos a organizaciones en el diseño de estructuras salariales coherentes, alineadas con el mercado y preparadas para el nuevo marco regulatorio europeo.", sal["body"])
        _write(slides / "slide7.xml", s)

    # ── Formación ────────────────────────────────────────────────────
    if "formacion" in services and svcs.get("formacion"):
        form = svcs["formacion"]
        s = _read(slides / "slide8.xml")
        if form.get("body"): s = r(s, "Potenciamos el talento interno: consultoría en formación orientada a procesos y resultados ágiles.", form["body"])
        _write(slides / "slide8.xml", s)

    # ── Diapositivas de servicio no seleccionadas: se quitan ─────────
    all_service_slides = {"headhunting": "slide5.xml", "outsourcing": "slide6.xml",
                           "salary": "slide7.xml", "formacion": "slide8.xml"}
    to_remove = [fn for key, fn in all_service_slides.items() if key not in services]
    if to_remove:
        _remove_slides(unpack, to_remove)

    # ── Autofit dinámico en las diapositivas con texto variable ──────
    for name in HC_SLIDES_DYNAMIC_TEXT:
        path = slides / name
        if path.exists():
            _write(path, enable_shrink_autofit(_read(path)))

    _apply_terms_and_fee(
        unpack, HC_TERMS_SLIDES, HC_FEE_SLIDE,
        include_fee=data.get("deck_type") == "propuesta",
        warranty_months=data.get("warranty_months", 3),
        fee_rate=data.get("fee_rate"),
    )


# ─── HUMAN CAPITAL + FINANCE ────────────────────────────────────────────
def _build_hc_finance(data: dict, content: dict, unpack: Path) -> None:
    show_all_locations = data.get("locations", "all") == "all"
    keep_location = HCF_LOCATION_SLIDES["all" if show_all_locations else "madrid"]
    drop_location = HCF_LOCATION_SLIDES["madrid" if show_all_locations else "all"]

    include_eduardo = bool(data.get("include_eduardo"))
    keep_team = HCF_TEAM_SLIDES["con_eduardo" if include_eduardo else "no_eduardo"]
    drop_team = HCF_TEAM_SLIDES["no_eduardo" if include_eduardo else "con_eduardo"]

    _remove_slides(unpack, [drop_location, drop_team])

    _apply_terms_and_fee(
        unpack, HCF_TERMS_SLIDES, HCF_FEE_SLIDE,
        include_fee=data.get("deck_type") == "propuesta",
        warranty_months=data.get("warranty_months", 3),
        fee_rate=data.get("fee_rate"),
    )


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

        out = tmp / "output.pptx"
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fp in sorted(unpack.rglob('*')):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(unpack))

        return _renumber_visible_pages(out.read_bytes())
