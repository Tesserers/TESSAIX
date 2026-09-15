"""Ensamblado final del .pptx a partir de la plantilla de Tessera."""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from .config import PLANTILLA
from .logo_client import compose_logo_for_placeholder
from .xml_utils import (
    clamp_picture_to_slide,
    enable_shrink_autofit,
    get_picture_frame_emu,
    get_shape_ext_emu,
    get_slide_size_emu,
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

# Diapositivas cuyo contenido (título + cuerpo) viene de la IA y por tanto
# puede ser más largo o más corto de lo previsto en la plantilla. En todas
# ellas se activa el "shrink to fit" dinámico para que un texto largo nunca
# se superponga con lo que hay debajo — ver xml_utils.enable_shrink_autofit.
_DYNAMIC_TEXT_SLIDES = ["slide3.xml", "slide5.xml", "slide6.xml", "slide7.xml"]

# rId del <a:blip> que referencia el placeholder del logo del cliente en la
# portada (slide1.xml), tal y como está definido en la plantilla actual.
_CLIENT_LOGO_RID = "rId6"

# Margen mínimo respecto al borde de la diapositiva al recolocar el hueco
# del logo si se sale del lienzo (ver _apply_client_logo).
_SLIDE_EDGE_MARGIN_EMU = 91440  # 0.1"

# Bloque de "oferta cruzada" de Finance/M&A (Soporte financiero, Fiscalidad,
# Auditoría, divisor "Tessera Group" y "Tessera Services") que viene
# incrustado en la plantilla de Human Capital. Es opcional por cliente: se
# incluye sólo si el consultor marca la casilla correspondiente en el paso
# de Servicios (ver ui_wizard.step_servicios / data["include_finance_crosssell"]).
_FINANCE_CROSSSELL_SLIDES = ["slide8.xml", "slide9.xml", "slide10.xml", "slide11.xml", "slide12.xml"]

# Al quitar ese bloque, las diapositivas que quedan detrás (Referencias,
# Honorarios, Fees) dejan de ser correlativas (016/017/018 tras el 07 de
# RPO). Se renumeran a mano para que la numeración visible siga una
# secuencia limpia. Si el bloque SÍ se incluye, no se toca nada.
_PAGE_RENUMBER_WHEN_FINANCE_EXCLUDED = {
    "slide13.xml": ("016", "08"),
    "slide14.xml": ("017", "09"),
    "slide15.xml": ("018", "10"),
}

# Dentro de slide15.xml (Honorarios) hay, además, un bloque de precios de
# Finance (CFO Part time / Plan Negocio) que solo tiene sentido si se está
# ofreciendo esa oferta cruzada — igual que las diapositivas 8-12. Son estos
# shapes (ver CFOPrice/PlanPrice/Footnote en la plantilla):
#   153 CFOPrice, 154 PlanPrice, 155 "*Revisables a 6 meses..."
#   39  "Condicionado a la colaboración en Headhunting/Outsourcing*"
# El título "HONORARIOS" y su línea (151/152) se dejan porque también
# encabezan las tablas de Headhunting/RRHH que sí se mantienen siempre.
_FINANCE_FEE_SHAPE_IDS_ON_SLIDE15 = [153, 154, 155, 39]


def _read(path: Path) -> str:
    return path.read_text("utf-8")


def _write(path: Path, content: str) -> None:
    path.write_text(content, "utf-8")


def get_client_logo_frame_emu() -> tuple[int, int] | None:
    """
    Lee directamente de la plantilla el tamaño (cx, cy) en EMU del hueco
    donde se inserta el logo del cliente. Se usa para renderizar en la
    pantalla de revisión una vista previa EXACTA de cómo va a quedar el
    logo antes de generar el .pptx (misma composición que _apply_client_logo).
    """
    with zipfile.ZipFile(PLANTILLA, 'r') as z:
        s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
    return get_picture_frame_emu(s1, _CLIENT_LOGO_RID)


def _apply_client_logo(unpack: Path, logo_bytes: bytes | None) -> None:
    """
    Inserta el logo del cliente en el hueco de la portada sin deformarlo, o
    si no hay logo disponible, quita el placeholder para no dejar una imagen
    incorrecta en el PowerPoint final.
    """
    slides = unpack / "ppt" / "slides"
    rels_dir = slides / "_rels"
    s1_path = slides / "slide1.xml"
    rels_path = rels_dir / "slide1.xml.rels"
    if not s1_path.exists():
        return

    s1 = _read(s1_path)

    if not logo_bytes:
        # Sin logo: se retira el <p:pic> del placeholder en vez de dejar la
        # imagen de muestra de la plantilla.
        s1 = re.sub(
            rf'<p:pic>\s*(?:(?!</p:pic>).)*?r:embed="{_CLIENT_LOGO_RID}"(?:(?!</p:pic>).)*?</p:pic>',
            '', s1, flags=re.DOTALL,
        )
        _write(s1_path, s1)
        return

    frame = get_picture_frame_emu(s1, _CLIENT_LOGO_RID)
    if frame is None:
        return  # la plantilla cambió de forma inesperada; no arriesgamos el resto del build
    target_w, target_h = frame

    composed = compose_logo_for_placeholder(logo_bytes, target_w, target_h)

    media_file = None
    if rels_path.exists():
        target = resolve_rel_target(_read(rels_path), _CLIENT_LOGO_RID)
        if target:
            media_file = (slides / target).resolve()
    if media_file is None or not media_file.exists():
        media_file = unpack / "ppt" / "media" / "image4.png"  # fallback conocido de esta plantilla

    if media_file.exists():
        media_file.write_bytes(composed)

    # La plantilla posiciona este hueco parcialmente FUERA de los límites
    # de la diapositiva (se comprobó: se sale ~0.57" por la derecha y
    # ~0.18" por abajo) — PowerPoint recorta silenciosamente lo que cae
    # fuera del lienzo, así que el logo se veía cortado. Se recoloca sin
    # tocar su tamaño para que quede completamente visible.
    pres_path = unpack / "ppt" / "presentation.xml"
    if pres_path.exists():
        slide_size = get_slide_size_emu(_read(pres_path))
        if slide_size:
            slide_w, slide_h = slide_size
            s1 = clamp_picture_to_slide(s1, _CLIENT_LOGO_RID, slide_w, slide_h, _SLIDE_EDGE_MARGIN_EMU)
            _write(s1_path, s1)


def _remove_slides(unpack: Path, slide_files: list[str]) -> None:
    """
    Quita por completo una o varias diapositivas del paquete .pptx: la
    referencia en <p:sldIdLst> de presentation.xml, la relación en
    presentation.xml.rels, el Override en [Content_Types].xml y el propio
    archivo de la diapositiva (con sus _rels). Es la forma correcta de
    borrar una diapositiva a nivel de paquete OOXML (equivalente a lo que
    hace PowerPoint al eliminarla desde la UI).
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


def _apply_finance_crosssell(unpack: Path, include: bool) -> None:
    """
    El bloque de Finance/M&A es una oferta cruzada opcional por cliente: si
    no se marca, se quita del paquete y se renumeran las diapositivas que
    quedan detrás para que la numeración visible siga siendo correlativa.
    """
    if include:
        return

    _remove_slides(unpack, _FINANCE_CROSSSELL_SLIDES)

    slides = unpack / "ppt" / "slides"
    for slide_file, (old_num, new_num) in _PAGE_RENUMBER_WHEN_FINANCE_EXCLUDED.items():
        path = slides / slide_file
        if path.exists():
            _write(path, replace_text_in_xml(_read(path), old_num, new_num))

    # El bloque de precios de Finance (CFO Part time / Plan Negocio) dentro
    # de la propia slide de Honorarios tampoco tiene sentido sin la oferta
    # cruzada — se quita igual que las diapositivas 8-12.
    slide15_path = slides / "slide15.xml"
    if slide15_path.exists():
        s15 = _read(slide15_path)
        for shape_id in _FINANCE_FEE_SHAPE_IDS_ON_SLIDE15:
            s15 = remove_shape_by_id(s15, shape_id)
        _write(slide15_path, s15)


def _replace_title_and_push(s: str, old: str, new: str, title_id: int, body_id: int) -> str:
    """
    Sustituye un título que NUNCA se encoge (role="title"): si el texto
    nuevo necesita más espacio del que tiene su caja original, esta crece
    hacia abajo en vez de reducir la letra. Como el cuerpo está pegado justo
    debajo, si el título creció se desplaza el cuerpo hacia abajo esa misma
    distancia para que no se superpongan.
    """
    before = get_shape_ext_emu(s, title_id)
    s = replace_text_and_fit(s, old, new, role="title")
    after = get_shape_ext_emu(s, title_id)
    if before and after and after[1] > before[1]:
        s = shift_shape_y(s, body_id, after[1] - before[1])
    return s


def build_pptx(data: dict, content: dict, logo_bytes: bytes | None = None) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        work = tmp / "work.pptx"
        shutil.copy(str(PLANTILLA), str(work))

        unpack = tmp / "unpacked"
        with zipfile.ZipFile(work, 'r') as z:
            z.extractall(unpack)

        slides = unpack / "ppt" / "slides"
        r = replace_text_and_fit

        _apply_client_logo(unpack, logo_bytes)
        _apply_finance_crosssell(unpack, data.get("include_finance_crosssell", False))

        why = content.get("why_tessera", {})
        svcs = content.get("services", {})

        presenter_name = data.get("presenter_name", "Manuel García Pina")
        presenter_phone = data.get("presenter_phone", "+34 619 511 155")
        presenter_email = data.get("presenter_email", "Manuel.garcia@tesseraservices.com")

        # ── SLIDE 1: portada ──────────────────────────────────────────
        s = _read(slides / "slide1.xml")
        s = r(s, "Edward Manrique ", f"{presenter_name} ")
        s = r(s, "+34 695021978", presenter_phone)
        _write(slides / "slide1.xml", s)

        # ── SLIDE 3: por qué tessera ────────────────────────────────
        # Título y cuerpo son cajas SEPARADAS, una justo debajo de la otra
        # (id título → id cuerpo): el título nunca se encoge — si necesita
        # más sitio, crece y empuja el cuerpo hacia abajo en vez de
        # superponerse. El cuerpo puede encogerse un poco, pero nunca por
        # debajo de 10pt (text_fit.MIN_FONT_SIZE_PT); si ni así cabe, crece
        # su propia caja (no hay nada pegado justo debajo que lo impida).
        if why:
            s = _read(slides / "slide3.xml")
            if why.get("d1_title"): s = _replace_title_and_push(s, "Sin CVs al azar", why["d1_title"], 286, 287)
            if why.get("d1_body"):  s = r(s, "No enviamos el primer CV, buscamos a quien encaja de verdad. Candidatos filtrados sobre la mesa en 72 horas.", why["d1_body"])
            if why.get("d2_title"): s = _replace_title_and_push(s, "Sin pausas", why["d2_title"], 289, 290)
            if why.get("d2_body"):  s = r(s, "Tu servicio no parará.", why["d2_body"])
            if why.get("d3_title"): s = _replace_title_and_push(s, "Siempre contigo", why["d3_title"], 292, 293)
            if why.get("d3_body"):  s = r(s, "No desaparecemos tras la incorporación: cuidamos a la persona y al cliente durante todo el servicio.", why["d3_body"])
            _write(slides / "slide3.xml", s)

        # ── SLIDE 5: headhunting ──────────────────────────────────────
        if "headhunting" in data.get("services", []) and svcs.get("headhunting"):
            hh = svcs["headhunting"]
            s = _read(slides / "slide5.xml")
            # WHY column
            if hh.get("why_col_title1"): s = r(s, "VELOCIDAD ", hh["why_col_title1"] + " ")
            if hh.get("why_col_body1"):  s = r(s, "Encontramos rápido, pero no a cualquiera. Seleccionamos a los que suman valor real dentro de nuestra base de datos especializada en Adtech.", hh["why_col_body1"])
            if hh.get("why_col_title2"): s = r(s, "ESPECIALIZACIÓN EN PUBLICIDAD & MARKETING", hh["why_col_title2"])
            if hh.get("why_col_body2"):  s = r(s, "Conocemos el ecosistema publicitario y cómo se construyen sus equipos. Evaluamos perfiles en base a su experiencia en Sales, Programmatic, Ad Operations, CSM y Marketing, entendiendo las dinámicas reales del sector y el encaje con tu negocio.", hh["why_col_body2"])
            if hh.get("why_col_title3"): s = r(s, "ÉXITO COMPARTIDO", hh["why_col_title3"])
            if hh.get("why_col_body3"):  s = r(s, "No se trata solo de cubrir vacantes, se trata de construir el equipo que te llevará al siguiente nivel.", hh["why_col_body3"])
            # BENEFITS column
            if hh.get("benefits_body1"): s = r(s, "Nos ocupamos de todo el proceso, tú solo conoces a los mejores.", hh["benefits_body1"])
            if hh.get("benefits_body2"): s = r(s, "Candidatos que no solo encajan, sino que impulsan tu crecimiento.", hh["benefits_body2"])
            if hh.get("benefits_body3"): s = r(s, "Porque para nosotros, tu éxito es también el nuestro.", hh["benefits_body3"])
            # HOW column
            if hh.get("how_body1"): s = r(s, "Escuchamos tus necesidades y objetivos desde el primer día.", hh["how_body1"])
            if hh.get("how_body2"): s = r(s, "Buscamos talento con técnicas avanzadas y un enfoque humano.", hh["how_body2"])
            if hh.get("how_body3"): s = r(s, "Evaluamos habilidades, actitud y encaje cultural.", hh["how_body3"])
            if hh.get("how_body4"): s = r(s, "Te presentamos a los candidatos que realmente suman", hh["how_body4"])
            if hh.get("how_body5"): s = r(s, "Estamos contigo incluso después de la incorporación para un correcto encaje en el equipo.", hh["how_body5"])
            _write(slides / "slide5.xml", s)

        # ── SLIDE 6: outsourcing ────────────────────────────────────
        if "outsourcing" in data.get("services", []) and svcs.get("outsourcing"):
            outs = svcs["outsourcing"]
            s = _read(slides / "slide6.xml")
            if outs.get("headline"): s = _replace_title_and_push(s, "Externalización que sí funciona: pagas por trabajo real, sin papeleo.", outs["headline"], 17, 3)
            if outs.get("body"):     s = r(s, "Contratar cuesta más de lo que parece. Con el modelo Time & Material ajustas el equipo a tu actividad real y nosotros nos encargamos de toda la gestión.", outs["body"])
            if outs.get("card1_body"): s = r(s, "Pagas solo por horas reales de trabajo.", outs["card1_body"])
            if outs.get("card2_body"): s = r(s, "Cubrimos bajas y vacaciones.", outs["card2_body"])
            if outs.get("card3_body"): s = r(s, "Nóminas y trámites, a cargo nuestro.", outs["card3_body"])
            if outs.get("card4_body"): s = r(s, "Escalas el equipo cuando quieras.", outs["card4_body"])
            _write(slides / "slide6.xml", s)

        # ── SLIDE 7: RPO ──────────────────────────────────────────────
        # Mismo patrón que la slide 3: título/cuerpo en cajas separadas y
        # pegadas — el título crece y empuja el cuerpo si hace falta.
        if "rpo" in data.get("services", []) and svcs.get("rpo"):
            rpo = svcs["rpo"]
            s = _read(slides / "slide7.xml")
            if rpo.get("headline"): s = _replace_title_and_push(s, "RPO a tu medida", rpo["headline"], 48, 49)
            if rpo.get("body"):     s = r(s, "Te ofrecemos un equipo de recruiters que se integra en tu compañía como una extensión real de tu equipo, trabajando exclusivamente en tus necesidades durante el tiempo que lo necesites.\nNo solo ejecutamos procesos: entendemos tu negocio, tus retos y tu historia  para atraer el talento que realmente necesitas.", rpo["body"])
            if rpo.get("p1_title"): s = _replace_title_and_push(s, "DEDICACIÓN TOTAL", rpo["p1_title"], 5, 9)
            if rpo.get("p1_body"):  s = r(s, "Un equipo dedicado en exclusiva a tu compañía, con conocimiento del sector publicitario y alineado c", rpo["p1_body"][:90])
            if rpo.get("p2_title"): s = _replace_title_and_push(s, "EFICIENCIA REAL", rpo["p2_title"], 12, 13)
            if rpo.get("p2_body"):  s = r(s, "Acceso directo a una red de +12.000 profesionales de publicidad ya identificados y validados. Menos ", rpo["p2_body"][:90])
            if rpo.get("p3_title"): s = _replace_title_and_push(s, "FOCO EN CRECER", rpo["p3_title"], 15, 16)
            if rpo.get("p3_body"):  s = r(s, "Nosotros buscamos, filtramos y validamos. Tú te concentras en hacer crecer tu equipo y tu negocio.", rpo["p3_body"][:90])
            _write(slides / "slide7.xml", s)

        # ── SLIDE 15: fees ──────────────────────────────────────────
        if data.get("deck_type") == "propuesta" and data.get("fee_rate"):
            s = _read(slides / "slide15.xml")
            fr = data.get("fee_rate", "XX")
            s = r(s, "18%", f"{fr}%")
            s = r(s, "13%", f"{fr}%")
            _write(slides / "slide15.xml", s)

        # ── SLIDE 16: cierre ──────────────────────────────────────────
        s = _read(slides / "slide16.xml")
        s = r(s, "Edward@tesseraservices.com", presenter_email)
        _write(slides / "slide16.xml", s)

        # ── Autofit dinámico: evita que texto largo se superponga ────
        for name in _DYNAMIC_TEXT_SLIDES:
            path = slides / name
            if path.exists():
                _write(path, enable_shrink_autofit(_read(path)))

        # Pack back to PPTX
        out = tmp / "output.pptx"
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fp in sorted(unpack.rglob('*')):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(unpack))
        return out.read_bytes()
