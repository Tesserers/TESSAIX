"""Manipulación del XML interno de un .pptx (ppt/slides/slideN.xml).

Todo lo que toca el marcado en crudo del OOXML vive aquí, aislado del resto
de la app, porque es la parte más delicada: un regex demasiado permisivo
puede dejar el XML mal formado y que PowerPoint se niegue a abrir el archivo.
"""
from __future__ import annotations

import re

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
