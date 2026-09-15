"""Manipulación del XML interno de un .pptx (ppt/slides/slideN.xml).

Todo lo que toca el marcado en crudo del OOXML vive aquí, aislado del resto
de la app, porque es la parte más delicada: un regex demasiado permisivo
puede dejar el XML mal formado y que PowerPoint se niegue a abrir el archivo.
"""
from __future__ import annotations

import re

from . import text_fit

# ─── REEMPLAZO DE TEXTO — CONSOLIDA RUNS ──────────────────────────────────
# Sólo debe casar la etiqueta real <a:t ...>, nunca otras que empiecen
# igual (p.ej. <a:tabLst/>, <a:theme>). El lookahead exige que justo
# después de "a:t" venga espacio, "/" o ">" — nunca otra letra.
_A_T_OPEN = r'<a:t(?=[\s/>])[^>]*>'


def replace_text_in_xml(xml: str, old: str, new: str) -> str:
    """
    Reemplaza texto que puede estar repartido en varios runs <a:t> dentro de
    un mismo párrafo. Estrategia: por cada <a:p>, se consolida todo el texto,
    se comprueba si el texto viejo está presente y, si lo está, se coloca el
    texto completo nuevo en el primer run y se vacían los demás.
    """
    new_esc = new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    old_esc = old.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def process_para(m):
        para = m.group(0)
        runs_text = re.findall(_A_T_OPEN + r'(.*?)</a:t>', para, re.DOTALL)
        full_text = ''.join(runs_text)

        if old_esc in full_text:
            new_full = full_text.replace(old_esc, new_esc)
        elif old in full_text:
            new_full = full_text.replace(old, new_esc)
        else:
            return para  # sin coincidencia, no se toca

        count = [0]

        def repl(m2):
            count[0] += 1
            if count[0] == 1:
                return f'{m2.group(1)}{new_full}{m2.group(3)}'
            return f'{m2.group(1)}{m2.group(3)}'

        return re.sub(f'({_A_T_OPEN})([^<]*)(</a:t>)', repl, para)

    return re.sub(r'<a:p\b[^>]*>.*?</a:p>', process_para, xml, flags=re.DOTALL)


# ─── AUTOFIT DINÁMICO — EVITA SUPERPOSICIONES ─────────────────────────────
# Las cajas de texto de la plantilla vienen con altura fija (o con
# "spAutoFit", que hace crecer la CAJA sin mover lo que hay debajo). Cuando
# el contenido generado por IA es más largo de lo previsto, el texto se sale
# de su caja y se superpone con el siguiente elemento.
#
# La solución nativa de PowerPoint para esto es <a:normAutofit/>: en vez de
# agrandar la caja o desbordar el texto, PowerPoint reduce dinámicamente el
# tamaño de fuente para que quepa dentro del tamaño fijo de la caja. Por eso
# sustituimos cualquier "spAutoFit"/"noAutofit"/autofit-ausente por
# "normAutofit" en las diapositivas con contenido variable.
_BODY_PR_RE = re.compile(r'<a:bodyPr\b[^>]*?(?:/>|>.*?</a:bodyPr>)', re.DOTALL)
_AUTOFIT_CHILD_RE = re.compile(r'<a:(spAutoFit|noAutofit)\s*/>')


def _shrink_body_pr(body_pr_xml: str) -> str:
    if "normAutofit" in body_pr_xml:
        return body_pr_xml
    if _AUTOFIT_CHILD_RE.search(body_pr_xml):
        return _AUTOFIT_CHILD_RE.sub("<a:normAutofit/>", body_pr_xml)
    if body_pr_xml.endswith("/>"):
        # <a:bodyPr .../>  →  <a:bodyPr ...><a:normAutofit/></a:bodyPr>
        return body_pr_xml[:-2] + "><a:normAutofit/></a:bodyPr>"
    # <a:bodyPr ...>...</a:bodyPr> sin autofit declarado: se inserta tras
    # la etiqueta de apertura (el resto del contenido puede tener sus
    # propios hijos como <a:normAutofit> ya cubierto arriba, o cosas como
    # <a:spAutoFit> ya cubiertas por la rama anterior).
    return re.sub(r'(<a:bodyPr[^>]*>)', r'\1<a:normAutofit/>', body_pr_xml, count=1)


def enable_shrink_autofit(slide_xml: str) -> str:
    """
    Fuerza "reducir texto al desbordar" en TODAS las cajas de texto de una
    diapositiva. Es un cambio seguro: si el texto ya cabe, no pasa nada
    (fontScale se queda al 100%); si no cabe, PowerPoint lo encoge en vez de
    superponerlo con lo que hay debajo.
    """
    return _BODY_PR_RE.sub(lambda m: _shrink_body_pr(m.group(0)), slide_xml)


# ─── REEMPLAZO CON AJUSTE DE TAMAÑO — evita solapes de verdad ─────────────
# `enable_shrink_autofit` (arriba) marca las cajas para que PowerPoint las
# encoja SOLO. El problema: esa señal (<a:normAutofit/>) únicamente surte
# efecto si el programa que abre el archivo la recalcula, y no todos lo
# hacen de forma fiable al abrir (por eso seguían viéndose solapes en la
# práctica). Esta variante calcula el tamaño correcto EN PYTHON y lo deja
# ya escrito en el `sz` de cada run — funciona en cualquier visor, sin
# depender de que nadie recalcule nada. Ver tessaix.text_fit.
#
# Limitación conocida: asume que la forma con el texto no está dentro de un
# <p:grpSp> (grupo), cuyo sistema de coordenadas hijo puede tener una
# escala distinta a la de la diapositiva. Ninguno de los campos que hoy
# rellena la IA vive dentro de un grupo, así que no aplica — si en el
# futuro se añadiera uno, esta función seguiría reemplazando el texto bien,
# solo que sin el ajuste de tamaño (se limitaría a no encoger nada).
_SHAPE_RE = re.compile(r'<p:sp>(?:(?!</p:sp>).)*?</p:sp>', re.DOTALL)
_SHAPE_EXT_RE = re.compile(r'<p:spPr>.*?<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*/>', re.DOTALL)
_EXT_RE = re.compile(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*/>')
_BODY_PR_OPEN_RE = re.compile(r'<a:bodyPr\b([^>]*)>')
_PARAGRAPH_RE = re.compile(r'<a:p\b[^>]*>.*?</a:p>', re.DOTALL)
_RUN_SZ_RE = re.compile(r'<a:rPr\b[^>]*?\bsz="(\d+)"')
_LNSPC_PCT_RE = re.compile(r'<a:lnSpc>\s*<a:spcPct\s+val="(\d+)"\s*/>\s*</a:lnSpc>')
_SPCBEF_PTS_RE = re.compile(r'<a:spcBef>\s*<a:spcPts\s+val="(\d+)"\s*/>\s*</a:spcBef>')
_ANY_SZ_ATTR_RE = re.compile(r'\bsz="(\d+)"')


def _inset_attr(attrs: str, name: str, default: int) -> int:
    m = re.search(rf'\b{name}="(\d+)"', attrs)
    return int(m.group(1)) if m else default


def _shape_insets_emu(shape_xml: str) -> tuple[int, int, int, int]:
    """Márgenes internos (l, t, r, b) en EMU de <a:bodyPr>, con los valores
    por defecto de OOXML cuando no se especifican."""
    body_pr = _BODY_PR_OPEN_RE.search(shape_xml)
    attrs = body_pr.group(1) if body_pr else ""
    l_ins = _inset_attr(attrs, "lIns", text_fit.DEFAULT_INSET_LR_EMU)
    r_ins = _inset_attr(attrs, "rIns", text_fit.DEFAULT_INSET_LR_EMU)
    t_ins = _inset_attr(attrs, "tIns", text_fit.DEFAULT_INSET_TB_EMU)
    b_ins = _inset_attr(attrs, "bIns", text_fit.DEFAULT_INSET_TB_EMU)
    return l_ins, t_ins, r_ins, b_ins


def _shape_usable_area_pt(shape_xml: str) -> tuple[float, float] | None:
    """Alto/ancho útil de una forma en puntos: su <a:ext> menos los
    márgenes internos (<a:bodyPr lIns/tIns/rIns/bIns>)."""
    ext = _SHAPE_EXT_RE.search(shape_xml)
    if not ext:
        return None
    box_w_pt = text_fit.emu_to_pt(int(ext.group(1)))
    box_h_pt = text_fit.emu_to_pt(int(ext.group(2)))

    l_ins, t_ins, r_ins, b_ins = _shape_insets_emu(shape_xml)
    usable_w = max(1.0, box_w_pt - text_fit.emu_to_pt(l_ins + r_ins))
    usable_h = max(1.0, box_h_pt - text_fit.emu_to_pt(t_ins + b_ins))
    return usable_w, usable_h


def _paragraph_fit_info(para_xml: str) -> dict | None:
    runs_text = re.findall(_A_T_OPEN + r'(.*?)</a:t>', para_xml, re.DOTALL)
    text = ''.join(runs_text).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")

    sz_match = _RUN_SZ_RE.search(para_xml)
    if not sz_match:
        return None  # sin tamaño de fuente detectable: no arriesgamos el cálculo
    size_pt = int(sz_match.group(1)) / 100.0

    lnspc_match = _LNSPC_PCT_RE.search(para_xml)
    line_spacing_pct = int(lnspc_match.group(1)) / 1000.0 if lnspc_match else 100.0

    spc_match = _SPCBEF_PTS_RE.search(para_xml)
    space_before_pt = int(spc_match.group(1)) / 100.0 if spc_match else 0.0

    return {
        "text": text, "size_pt": size_pt,
        "line_spacing_pct": line_spacing_pct, "space_before_pt": space_before_pt,
    }


def _grow_shape_height(shape_xml: str, min_cy_emu: int) -> str:
    """Aumenta (nunca reduce) el <a:ext cy="..."> de esta forma para que
    tenga al menos `min_cy_emu` de alto. Mantiene cx intacto."""
    m = _EXT_RE.search(shape_xml)
    if not m:
        return shape_xml
    cx, cy = m.group(1), int(m.group(2))
    if cy >= min_cy_emu:
        return shape_xml
    return shape_xml.replace(m.group(0), f'<a:ext cx="{cx}" cy="{min_cy_emu}"/>', 1)


# Suelo "normal" de un título: casi nunca se toca (mantiene el espíritu de
# "los subtítulos siguen igual"), pero deja un pequeño margen antes de
# recurrir a crecer la caja, que para columnas estrechas puede necesitar
# bastante espacio incluso para un título de longitud normal.
_TITLE_MIN_SCALE = 0.90


def _apply_sz_scale(shape_xml: str, scale: float) -> str:
    return _ANY_SZ_ATTR_RE.sub(
        lambda sm: f'sz="{max(100, round(int(sm.group(1)) * scale))}"', shape_xml,
    )


def replace_text_and_fit(slide_xml: str, old: str, new: str, role: str = "body") -> str:
    """
    Igual que `replace_text_in_xml`, pero además comprueba si el texto
    resultante cabe en la caja real de la forma que lo contiene (todos sus
    párrafos, no solo el que se acaba de sustituir), en tres niveles:

    1. Si cabe encogiendo un poco (sin bajar del suelo de su rol), se
       reduce el tamaño de fuente de TODOS los runs de esa forma — el
       cálculo queda horneado en el archivo, así que funciona igual en
       PowerPoint, LibreOffice, Google Slides, etc.
    2. Si ni así cabe, no se sigue encogiendo (quedaría ilegible): se hace
       crecer la caja hasta el alto que haga falta, limitado a
       `text_fit.MAX_GROW_MULTIPLIER` veces su alto original. Quien llama
       es responsable de comprobar si la forma creció (con
       `get_shape_ext_emu` antes/después) y desplazar lo que tenga debajo
       — ver `shift_shape_y` y `pptx_builder._replace_title_and_push`.
    3. Si ni haciendo crecer la caja hasta ese tope cabe (texto realmente
       desmedido), se prioriza no salirse de la diapositiva por encima de
       mantener el tamaño: se encoge más allá del suelo normal, hasta
       `text_fit.ABSOLUTE_MIN_SCALE`, dentro de la caja ya crecida al tope.

    `role="title"` usa un suelo de encogido muy leve (rara vez se nota) en
    el paso 1; `role="body"` (por defecto) usa el suelo absoluto de
    `text_fit.MIN_FONT_SIZE_PT`. El paso 3 (último recurso) es igual para
    ambos roles.
    """
    min_scale = _TITLE_MIN_SCALE if role == "title" else None

    def process_shape(m):
        shape_xml = m.group(0)
        # OJO: no vale hacer aquí un atajo tipo "if old not in shape_xml:
        # skip" como optimización — el texto viejo puede estar repartido en
        # varios <a:r> (p.ej. por una palabra en negrita a mitad de frase),
        # así que no aparece como substring contiguo aunque SÍ vaya a
        # coincidir en `replace_text_in_xml` (que consolida por párrafo).
        # Ya nos pasó una vez: se saltaba en silencio justo los textos que
        # esta función existe para arreglar.
        replaced = replace_text_in_xml(shape_xml, old, new)
        if replaced == shape_xml:
            return shape_xml  # este shape no contenía el texto a sustituir

        usable_area = _shape_usable_area_pt(replaced)
        if usable_area is None:
            return replaced
        usable_w_pt, usable_h_pt = usable_area

        paragraphs = [
            info for para in _PARAGRAPH_RE.findall(replaced)
            if (info := _paragraph_fit_info(para)) is not None
        ]
        if not paragraphs:
            return replaced

        result = text_fit.fit_scale(paragraphs, usable_w_pt, usable_h_pt, min_scale=min_scale)
        if result.scale >= 0.999 and result.fits:
            return replaced

        if result.scale < 0.999:
            replaced = _apply_sz_scale(replaced, result.scale)

        if result.fits:
            return replaced

        # Ni encogiendo hasta el suelo del rol cabe: se hace crecer la
        # caja. El alto de contenido no incluye los márgenes internos —hay
        # que devolvérselos para obtener el alto TOTAL de la forma.
        _, t_ins, _, b_ins = _shape_insets_emu(replaced)
        insets_pt = text_fit.emu_to_pt(t_ins + b_ins)
        original_box_h_pt = usable_h_pt + insets_pt
        max_box_h_pt = original_box_h_pt * text_fit.MAX_GROW_MULTIPLIER
        needed_box_h_pt = result.required_height_pt * text_fit.GROW_BUFFER + insets_pt

        if needed_box_h_pt <= max_box_h_pt:
            replaced = _grow_shape_height(replaced, text_fit.pt_to_emu(needed_box_h_pt))
            return replaced

        # El crecimiento necesario sería desmedido (texto muy por encima de
        # los límites de longitud habituales): se prioriza no salirse de la
        # diapositiva. Se crece hasta el tope y se encoge lo que haga falta
        # — sin bajar de ABSOLUTE_MIN_SCALE — para que quepa ahí dentro.
        capped_usable_h_pt = max_box_h_pt - insets_pt
        fallback = text_fit.fit_scale(
            paragraphs, usable_w_pt, capped_usable_h_pt, min_scale=text_fit.ABSOLUTE_MIN_SCALE,
        )
        replaced = _apply_sz_scale(replaced, fallback.scale / result.scale)
        replaced = _grow_shape_height(replaced, text_fit.pt_to_emu(max_box_h_pt))
        return replaced

    return _SHAPE_RE.sub(process_shape, slide_xml)


def get_shape_ext_emu(slide_xml: str, shape_id: int) -> tuple[int, int] | None:
    """Tamaño (cx, cy) en EMU de la forma cuyo <p:cNvPr id="shape_id"> coincide."""
    pattern = re.compile(
        r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:sp>).)*?</p:sp>' % shape_id,
        re.DOTALL,
    )
    m = pattern.search(slide_xml)
    if not m:
        return None
    ext = _EXT_RE.search(m.group(0))
    return (int(ext.group(1)), int(ext.group(2))) if ext else None


def shift_shape_y(slide_xml: str, shape_id: int, delta_emu: int) -> str:
    """Desplaza verticalmente (delta_emu puede ser negativo) la forma cuyo
    <p:cNvPr id="shape_id"> coincide, sin tocar su tamaño."""
    if delta_emu == 0:
        return slide_xml
    pattern = re.compile(
        r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:sp>).)*?</p:sp>' % shape_id,
        re.DOTALL,
    )
    m = pattern.search(slide_xml)
    if not m:
        return slide_xml
    shape_xml = m.group(0)
    off = re.search(r'<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*/>', shape_xml)
    if not off:
        return slide_xml
    x, y = int(off.group(1)), int(off.group(2))
    new_shape_xml = shape_xml.replace(off.group(0), f'<a:off x="{x}" y="{y + delta_emu}"/>', 1)
    return slide_xml[:m.start()] + new_shape_xml + slide_xml[m.end():]


# ─── LECTURA DE RELACIONES / MARCOS DE IMAGEN ─────────────────────────────
def resolve_rel_target(rels_xml: str, rid: str) -> str | None:
    """Devuelve el Target (ruta relativa) de una relación r:id concreta."""
    m = re.search(
        r'<Relationship\b[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"' % re.escape(rid),
        rels_xml,
    ) or re.search(
        r'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="%s"' % re.escape(rid),
        rels_xml,
    )
    return m.group(1) if m else None


def resolve_rid_for_target(rels_xml: str, target: str) -> str | None:
    """Inversa de resolve_rel_target: dado un Target, devuelve su r:id."""
    m = re.search(
        r'<Relationship\b[^>]*\bTarget="%s"[^>]*\bId="(rId\d+)"' % re.escape(target),
        rels_xml,
    ) or re.search(
        r'<Relationship\b[^>]*\bId="(rId\d+)"[^>]*\bTarget="%s"' % re.escape(target),
        rels_xml,
    )
    return m.group(1) if m else None


def remove_relationship(rels_xml: str, rid: str) -> str:
    """Quita la <Relationship Id="rid" .../> de un .rels."""
    return re.sub(
        r'<Relationship\b[^>]*\bId="%s"[^>]*/>' % re.escape(rid), '', rels_xml,
    )


def remove_sld_id(presentation_xml: str, rid: str) -> str:
    """Quita la <p:sldId .../> que referencia `rid` de <p:sldIdLst>."""
    return re.sub(
        r'<p:sldId\b[^>]*\br:id="%s"[^>]*/>' % re.escape(rid), '', presentation_xml,
    )


def remove_content_type_override(content_types_xml: str, part_name: str) -> str:
    """Quita el <Override PartName="part_name" .../> de [Content_Types].xml."""
    return re.sub(
        r'<Override\s+PartName="%s"[^>]*/>' % re.escape(part_name), '', content_types_xml,
    )


def remove_shape_by_id(slide_xml: str, shape_id: int) -> str:
    """
    Quita por completo el <p:sp>...</p:sp> cuyo <p:cNvPr id="shape_id" .../>
    coincide. Un <p:sp> normal nunca anida otro <p:sp> dentro (solo lo hacen
    los <p:grpSp>), así que basta con exigir que no aparezca un </p:sp> de
    cierre antes de encontrar el id buscado, igual que ya se hace para
    aislar el <p:pic> del logo del cliente.
    """
    pattern = re.compile(
        r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:sp>).)*?</p:sp>' % shape_id,
        re.DOTALL,
    )
    return pattern.sub('', slide_xml, count=1)


def get_picture_frame_emu(slide_xml: str, rid: str) -> tuple[int, int] | None:
    """
    Busca el <p:pic> que referencia `rid` (vía <a:blip r:embed="rid"/>) y
    devuelve el tamaño (cx, cy) en EMU de su <a:xfrm><a:ext>, es decir, el
    tamaño real del hueco donde se va a insertar la imagen en la diapositiva.
    """
    pic_re = re.compile(
        r'<p:pic>(?:(?!</p:pic>).)*?</p:pic>', re.DOTALL
    )
    for m in pic_re.finditer(slide_xml):
        block = m.group(0)
        if f'r:embed="{rid}"' not in block:
            continue
        ext = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*/>', block)
        if ext:
            return int(ext.group(1)), int(ext.group(2))
    return None


def get_slide_size_emu(presentation_xml: str) -> tuple[int, int] | None:
    """Tamaño (cx, cy) en EMU de la diapositiva, desde <p:sldSz> en presentation.xml."""
    m = re.search(r'<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"', presentation_xml)
    return (int(m.group(1)), int(m.group(2))) if m else None


def clamp_picture_to_slide(
    slide_xml: str, rid: str, slide_w_emu: int, slide_h_emu: int, margin_emu: int = 0,
) -> str:
    """
    Si el <p:pic> que referencia `rid` se sale de los límites de la
    diapositiva (total o parcialmente — puede pasar si la plantilla la
    posicionó pensando en otro tamaño de lienzo), lo desplaza hacia dentro
    lo justo para que quede completamente visible, sin tocar su tamaño ni
    su relación de aspecto.
    """
    pic_re = re.compile(r'<p:pic>(?:(?!</p:pic>).)*?</p:pic>', re.DOTALL)

    def process(m):
        block = m.group(0)
        if f'r:embed="{rid}"' not in block:
            return block
        off = re.search(r'<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*/>', block)
        ext = re.search(r'<a:ext\s+cx="(\d+)"\s+cy="(\d+)"\s*/>', block)
        if not off or not ext:
            return block
        x, y = int(off.group(1)), int(off.group(2))
        cx, cy = int(ext.group(1)), int(ext.group(2))

        new_x = min(x, slide_w_emu - margin_emu - cx)
        new_y = min(y, slide_h_emu - margin_emu - cy)
        new_x = max(new_x, margin_emu)
        new_y = max(new_y, margin_emu)

        if new_x == x and new_y == y:
            return block
        return block.replace(off.group(0), f'<a:off x="{new_x}" y="{new_y}"/>', 1)

    return pic_re.sub(process, slide_xml)
