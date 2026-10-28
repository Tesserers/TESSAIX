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
_A_R_RE = re.compile(r'<a:r\b[^>]*>(?:(?!</a:r>).)*?</a:r>', re.DOTALL)
_RPR_RE = re.compile(r'<a:rPr\b[^>]*(?:/>|>(?:(?!</a:rPr>).)*?</a:rPr>)', re.DOTALL)
# Un número "suelto" dentro de una frase: dígitos con separadores de miles
# o decimales incluidos (p.ej. "72", "36.000", "4,5") — nunca una letra ni
# el símbolo que vaya detrás (€, %, h...), que se quedan en la tipografía
# del texto que los rodea.
_NUMBER_RE = re.compile(r'(\d[\d.,]*\d|\d)')


def _manrope_rpr(rpr_xml: str) -> str:
    """
    Copia `rpr_xml` (el <a:rPr> de un run existente) forzando su tipografía
    a Manrope, manteniendo todo lo demás (tamaño, color, negrita…) — regla
    de marca: cualquier NÚMERO va siempre en Manrope, aunque esté dentro de
    una frase en Raleway. Si el run no tenía <a:rPr> o no tenía <a:latin>,
    se añade uno nuevo.
    """
    if not rpr_xml:
        return '<a:rPr><a:latin typeface="Manrope"/></a:rPr>'
    if '<a:latin' in rpr_xml:
        return re.sub(r'<a:latin\s+typeface="[^"]*"', '<a:latin typeface="Manrope"', rpr_xml, count=1)
    if rpr_xml.endswith('/>'):
        return rpr_xml[:-2] + '><a:latin typeface="Manrope"/></a:rPr>'
    return re.sub(r'(<a:rPr\b[^>]*>)', r'\1<a:latin typeface="Manrope"/>', rpr_xml, count=1)


def _runs_with_manrope_numbers(text: str, rpr_xml: str) -> str:
    """
    Reparte `text` en los <a:r> que hagan falta: el texto normal con
    `rpr_xml` tal cual (la tipografía que ya tenía ese run), y cualquier
    número suelto con la misma `rpr_xml` pero en Manrope. Si no hay ningún
    número, es un único run — igual que antes de esta regla de marca.
    """
    chunks = _NUMBER_RE.split(text)
    if len(chunks) == 1:
        return f'<a:r>{rpr_xml}<a:t xml:space="preserve">{text}</a:t></a:r>'
    manrope_rpr = _manrope_rpr(rpr_xml)
    runs = []
    for i, chunk in enumerate(chunks):
        if not chunk:
            continue
        rpr = manrope_rpr if i % 2 == 1 else rpr_xml
        runs.append(f'<a:r>{rpr}<a:t xml:space="preserve">{chunk}</a:t></a:r>')
    return ''.join(runs)


def replace_text_in_xml(xml: str, old: str, new: str) -> str:
    """
    Reemplaza texto que puede estar repartido en varios runs <a:t> dentro de
    un mismo párrafo. Estrategia: por cada <a:p>, se consolida todo el texto,
    se comprueba si el texto viejo está presente y, si lo está, se reparte el
    texto completo nuevo en uno o más runs a partir del primero (con su
    mismo <a:rPr> — tamaño, color, tipografía — salvo en los números, que
    siempre van en Manrope) y se vacían los demás.
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

        first_run_m = _A_R_RE.search(para)
        rpr_m = _RPR_RE.search(first_run_m.group(0)) if first_run_m else None
        rpr_xml = rpr_m.group(0) if rpr_m else ''
        new_runs_xml = _runs_with_manrope_numbers(new_full, rpr_xml)

        count = [0]

        def repl(rm):
            count[0] += 1
            return new_runs_xml if count[0] == 1 else ''

        return _A_R_RE.sub(repl, para)

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


def replace_text_and_fit(
    slide_xml: str, old: str, new: str, role: str = "body",
    base_scale: float = 1.0, min_font_pt: float | None = None,
    preset_height_pt: float | None = None, grow_ceiling_pt: float | None = None,
    max_usable_height_pt: float | None = None,
) -> str:
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
       `text_fit.MAX_GROW_MULTIPLIER` veces su alto original (o a
       `max_height_pt`, si se indica). Quien llama es responsable de
       comprobar si la forma creció (con `get_shape_ext_emu` antes/después)
       y desplazar lo que tenga debajo — ver `shift_shape_y` y
       `pptx_builder._replace_title_and_push`.
    3. Si ni haciendo crecer la caja hasta ese tope cabe (texto realmente
       desmedido), se prioriza no salirse de la diapositiva por encima de
       mantener el tamaño: se encoge más allá del suelo normal, hasta
       `text_fit.ABSOLUTE_MIN_SCALE`, dentro de la caja ya crecida al tope.

    `role="title"` usa un suelo de encogido muy leve (rara vez se nota) en
    el paso 1; `role="body"` (por defecto) usa el suelo absoluto de
    `text_fit.MIN_FONT_SIZE_PT` (o `min_font_pt`, si se indica). El paso 3
    (último recurso) es igual para ambos roles.

    `base_scale` escala el tamaño de fuente ORIGINAL de la plantilla antes
    de calcular nada — útil para que un bloque reutilizado (p.ej. un
    subtítulo clonado para servir de cuerpo de una lista) arranque más
    grande que el texto donante, en vez de partir de su tamaño pensado para
    una sola línea corta.

    `preset_height_pt`, si se indica, agranda la caja a (al menos) ese alto
    ANTES de medir si el texto cabe — así el cálculo de encogido se hace
    contra el espacio real disponible en la diapositiva, no contra la caja
    original del donante (normalmente pensada para una sola línea). Sin
    esto, el motor encogería primero hasta el suelo de su rol para intentar
    caber en la caja pequeña, y solo crecería la caja como último recurso —
    quedando innecesariamente diminuto incluso habiendo hueco de sobra.

    `grow_ceiling_pt`, si se indica, limita lo que el paso 2 puede crecer la
    caja (por debajo de `text_fit.MAX_GROW_MULTIPLIER` si hace falta) —
    imprescindible cuando detrás de la caja de texto hay una forma
    decorativa de tamaño FIJO (p.ej. la tarjeta redondeada de un servicio):
    sin este tope, el texto podría crecer más allá del borde visible de esa
    tarjeta. Con el tope puesto, si el texto no cabe ni creciendo hasta ahí,
    se prioriza encogerlo más (paso 3) antes que salirse de la tarjeta.

    `max_usable_height_pt` limita el alto que se considera "disponible"
    para el texto, aunque la caja declarada en la plantilla sea más alta.
    Hace falta cuando una decoración (p.ej. uno de los círculos
    traslúcidos de fondo) se solapa con la parte baja de la caja de
    fábrica: si el texto ocupara toda la caja, su parte final quedaría
    tapada por esa decoración aunque la caja NUNCA haya crecido. Con este
    tope, el texto se encoge antes de llegar a esa zona en vez de
    confirmarlo visualmente solo cuando ya se ha generado el contenido.
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

        if base_scale != 1.0:
            replaced = _apply_sz_scale(replaced, base_scale)

        if preset_height_pt is not None:
            replaced = _grow_shape_height(replaced, text_fit.pt_to_emu(preset_height_pt))

        usable_area = _shape_usable_area_pt(replaced)
        if usable_area is None:
            return replaced
        usable_w_pt, usable_h_pt = usable_area
        if max_usable_height_pt is not None:
            usable_h_pt = min(usable_h_pt, max_usable_height_pt)

        paragraphs = [
            info for para in _PARAGRAPH_RE.findall(replaced)
            if (info := _paragraph_fit_info(para)) is not None
        ]
        if not paragraphs:
            return replaced

        result = text_fit.fit_scale(
            paragraphs, usable_w_pt, usable_h_pt, min_scale=min_scale, min_font_pt=min_font_pt,
        )
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
        if grow_ceiling_pt is not None:
            max_box_h_pt = min(max_box_h_pt, grow_ceiling_pt)
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


def _set_open_tag_attr(attrs: str, name: str, value: str) -> str:
    if re.search(rf'\b{name}="[^"]*"', attrs):
        return re.sub(rf'\b{name}="[^"]*"', f'{name}="{value}"', attrs)
    return f' {name}="{value}"' + attrs


def _shape_contains_text(shape_xml: str, marker_text: str) -> bool:
    """
    Comprueba si `marker_text` aparece en el texto VISIBLE de la forma —
    no en su XML en crudo. El texto puede estar repartido en varios
    <a:r> (p.ej. por una palabra en negrita a mitad de frase), así que un
    simple ``marker_text in shape_xml`` falla justo en esos casos; hay que
    consolidar los <a:t> de cada párrafo antes de comparar, igual que hace
    `replace_text_in_xml`.
    """
    for para in re.findall(r'<a:p\b[^>]*>.*?</a:p>', shape_xml, re.DOTALL):
        runs_text = ''.join(re.findall(_A_T_OPEN + r'(.*?)</a:t>', para, re.DOTALL))
        if marker_text in runs_text:
            return True
    return False


def center_paragraph_align(slide_xml: str, marker_text: str) -> str:
    """
    Centra horizontalmente (``algn="ctr"``) todos los párrafos de la forma
    que contiene `marker_text`, sin tocar nada más de su contenido. Útil
    para diapositivas clonadas que heredan la alineación a la izquierda
    del donante original.
    """
    def process_shape(m):
        shape_xml = m.group(0)
        if not _shape_contains_text(shape_xml, marker_text):
            return shape_xml
        return re.sub(
            r'<a:pPr\b([^>]*)>',
            lambda pm: f'<a:pPr{_set_open_tag_attr(pm.group(1), "algn", "ctr")}>',
            shape_xml,
        )
    return _SHAPE_RE.sub(process_shape, slide_xml)


def center_vertical_anchor(slide_xml: str, marker_text: str) -> str:
    """
    Centra verticalmente (``anchor="ctr"`` en `bodyPr`) el contenido de la
    forma que contiene `marker_text` dentro de su caja — útil cuando esa
    caja es mucho más alta que el texto que contiene (p.ej. tras un
    `preset_height_pt` generoso) y, si no, el texto queda todo pegado
    arriba dejando un hueco vacío grande debajo.
    """
    def process_shape(m):
        shape_xml = m.group(0)
        if not _shape_contains_text(shape_xml, marker_text):
            return shape_xml
        return re.sub(
            r'<a:bodyPr\b([^>]*)>',
            lambda pm: f'<a:bodyPr{_set_open_tag_attr(pm.group(1), "anchor", "ctr")}>',
            shape_xml, count=1,
        )
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


def grow_shape_by(slide_xml: str, shape_id: int, delta_cy_emu: int) -> str:
    """Aumenta el alto (<a:ext cy="...">) de la forma `shape_id` en
    `delta_cy_emu`, sin tocar su posición ni su ancho — hace falta cuando
    el contenido de una columna ha crecido más de lo habitual y la tarjeta
    de fondo (una forma fija, aparte del texto) tiene que crecer con él."""
    if delta_cy_emu <= 0:
        return slide_xml
    ext = get_shape_ext_emu(slide_xml, shape_id)
    if not ext:
        return slide_xml
    cx, cy = ext
    pattern = re.compile(
        r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:sp>).)*?</p:sp>' % shape_id,
        re.DOTALL,
    )
    m = pattern.search(slide_xml)
    if not m:
        return slide_xml
    shape_xml = m.group(0)
    new_shape_xml = re.sub(
        r'<a:ext\s+cx="\d+"\s+cy="\d+"\s*/>', f'<a:ext cx="{cx}" cy="{cy + delta_cy_emu}"/>', shape_xml, count=1,
    )
    return slide_xml[:m.start()] + new_shape_xml + slide_xml[m.end():]


def set_shape_pos(slide_xml: str, shape_id: int, x_emu: int | None = None, y_emu: int | None = None) -> str:
    """
    Fija la posición ABSOLUTA (en vez de desplazar, como `shift_shape_y`) de
    la forma `shape_id` — cualquier tipo (texto, imagen...), a diferencia de
    `shift_shape_y` que solo entiende <p:sp>. Si `x_emu` o `y_emu` es None,
    esa coordenada se deja como estaba. Hace falta para recolocar bloques
    de persona (foto + texto) en el hueco que les toque según cuántas
    personas haya seleccionadas en el equipo.
    """
    for tag in _REMOVABLE_TAGS:
        pattern = re.compile(
            r'<%s\b[^>]*>(?:(?!</%s>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</%s>).)*?</%s>'
            % (tag, tag, shape_id, tag, tag),
            re.DOTALL,
        )
        m = pattern.search(slide_xml)
        if not m:
            continue
        shape_xml = m.group(0)
        off = re.search(r'<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"\s*/>', shape_xml)
        if not off:
            return slide_xml
        x = x_emu if x_emu is not None else int(off.group(1))
        y = y_emu if y_emu is not None else int(off.group(2))
        new_shape_xml = shape_xml.replace(off.group(0), f'<a:off x="{x}" y="{y}"/>', 1)
        return slide_xml[:m.start()] + new_shape_xml + slide_xml[m.end():]
    return slide_xml


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


# Cualquiera de estos puede llevar el <p:cNvPr id="..."> que identifica una
# forma. Un elemento de este tipo nunca anida otro del MISMO tipo dentro
# (salvo <p:grpSp>, que sí puede anidar cualquiera de los demás), así que
# basta con exigir que no aparezca un cierre de su propia etiqueta antes de
# encontrar el id buscado.
_REMOVABLE_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp")


def remove_shape_by_id(slide_xml: str, shape_id: int) -> str:
    """
    Quita por completo el elemento (forma, imagen, gráfico agrupado…) cuyo
    <p:cNvPr id="shape_id" .../> coincide — sea cual sea su tipo de
    contenedor. Las decoraciones clonadas de otra diapositiva (p.ej. el
    gráfico de fondo con forma de onda) son casi siempre <p:pic>, no
    <p:sp>, así que hace falta probar varios tipos de etiqueta en vez de
    asumir siempre <p:sp> como antes.
    """
    for tag in _REMOVABLE_TAGS:
        pattern = re.compile(
            r'<%s>(?:(?!</%s>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</%s>).)*?</%s>'
            % (tag, tag, shape_id, tag, tag),
            re.DOTALL,
        )
        new_xml, n = pattern.subn('', slide_xml, count=1)
        if n:
            return new_xml
    return slide_xml


_SRGB_CLR_RE = re.compile(r'<a:srgbClr\s+val="[0-9A-Fa-f]{6}"\s*/>')


def set_shape_text_color(slide_xml: str, shape_id: int, hex_color: str) -> str:
    """
    Fuerza el color de TODO el texto de la forma `shape_id` a `hex_color`
    (sin "#"). Hace falta cuando una forma clonada de otra diapositiva trae
    un color pensado para el fondo ORIGINAL (p.ej. un gris casi blanco,
    legible sobre un fondo oscuro que esta diapositiva ya no tiene) — sin
    esto, el texto queda prácticamente invisible sobre el nuevo fondo.
    """
    pattern = re.compile(
        r'<p:sp>(?:(?!</p:sp>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:sp>).)*?</p:sp>' % shape_id,
        re.DOTALL,
    )
    m = pattern.search(slide_xml)
    if not m:
        return slide_xml
    shape_xml = m.group(0)
    new_shape_xml = _SRGB_CLR_RE.sub(f'<a:srgbClr val="{hex_color}"/>', shape_xml)
    return slide_xml[:m.start()] + new_shape_xml + slide_xml[m.end():]


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


def find_pic_embed_rid(slide_xml: str, shape_id: int) -> str | None:
    """
    Dado el id de una forma <p:pic> (su <p:cNvPr id="...">), devuelve el
    r:id de la relación de imagen que usa (<a:blip r:embed="...">) — con
    eso y `resolve_rel_target` sobre el .rels de la diapositiva se llega al
    archivo de media real a sustituir, sin tocar el tamaño/posición del
    hueco en la diapositiva.
    """
    pic_re = re.compile(
        r'<p:pic>(?:(?!</p:pic>).)*?<p:cNvPr\b[^>]*\bid="%d"(?:(?!</p:pic>).)*?</p:pic>' % shape_id,
        re.DOTALL,
    )
    m = pic_re.search(slide_xml)
    if not m:
        return None
    embed = re.search(r'<a:blip\b[^>]*\br:embed="(rId\d+)"', m.group(0))
    return embed.group(1) if embed else None


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
