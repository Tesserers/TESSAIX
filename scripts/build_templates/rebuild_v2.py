"""
Script de construcción (uso puntual, no se ejecuta en runtime): aplica a
plantilla_hc.pptx y plantilla_hc_finance.pptx los cambios estructurales de
la 2ª fase del rebrand 2026 (ver conversación): equipo con/sin Eduardo
Serrano como alternativas según deck_type, tabla comparativa "No somos un
proveedor de CVs" (siempre presente), "Próximos pasos" (solo con fee),
honorarios con el nuevo modelo "Modalidad Partner", hueco de logo de
cliente en portada y cierre, y en HC+Finance la retirada completa del
bloque de términos/honorarios (ya no tiene flujo de fee).

Se apoya en las utilidades ya probadas de tessaix.xml_utils /
tessaix.pptx_builder (mismo mecanismo que usa la generación en runtime:
desempaquetar, editar XML a mano, volver a empaquetar) en vez de tocar los
.pptx con python-pptx, para no arriesgarse a que un re-guardado renombre
las partes o rompa algo que ya funciona.

Ejecutar desde la raíz del repo: python scripts/build_templates/rebuild_v2.py
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from tessaix.pptx_builder import _duplicate_slide, _read, _remove_slides, _write  # noqa: E402
from tessaix.xml_utils import remove_shape_by_id, replace_text_in_xml  # noqa: E402

EMU_PER_PT = 12700


def pt(v: float) -> int:
    return int(round(v * EMU_PER_PT))


def xml_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ─── plantillas de forma (extraídas de shapes reales de plantilla_hc.pptx,
# ver conversación — se clonan y se les cambia posición/texto) ────────────
def rect_sp(shape_id: int, x_pt: float, y_pt: float, w_pt: float, h_pt: float, fill="FFFFFF") -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="Rectangle {shape_id-1}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>'
        f'<a:off x="{pt(x_pt)}" y="{pt(y_pt)}"/><a:ext cx="{pt(w_pt)}" cy="{pt(h_pt)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr>'
        f'<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
        f'<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
        f'<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>'
        f'<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'
        f'<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody></p:sp>'
    )


def text_sp(
    shape_id: int, x_pt: float, y_pt: float, w_pt: float, h_pt: float,
    lines: list[tuple[str, int, str, str]], algn: str = "l",
) -> str:
    """lines: [(texto, tamaño_en_centésimas_de_pt, tipografía, color_hex), ...]
    cada tupla es un párrafo independiente (se renderiza como línea aparte)."""
    paras = []
    for i, (text, sz, font, color) in enumerate(lines):
        pPr = f'<a:pPr algn="{algn}">' + ("<a:spcAft><a:spcPts val=\"600\"/></a:spcAft>" if i == 0 and len(lines) > 1 else "") + "</a:pPr>"
        paras.append(
            f'<a:p>{pPr}<a:r><a:rPr sz="{sz}" b="0" dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/></a:rPr><a:t>{xml_escape(text)}</a:t></a:r></a:p>'
        )
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="TextBox {shape_id-1}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>'
        f'<a:off x="{pt(x_pt)}" y="{pt(y_pt)}"/><a:ext cx="{pt(w_pt)}" cy="{pt(h_pt)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:spAutoFit/></a:bodyPr>'
        f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>'
    )


def insert_shapes_before_sptree_close(slide_xml: str, shapes_xml: list[str]) -> str:
    return slide_xml.replace("</p:spTree>", "".join(shapes_xml) + "</p:spTree>")


def max_shape_id(slide_xml: str) -> int:
    ids = [int(m) for m in re.findall(r'<p:cNvPr\s+id="(\d+)"', slide_xml)]
    return max(ids) if ids else 1


# Transparent 1x1 PNG (placeholder hasta que se suba/encuentre un logo real).
_BLANK_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a4944415478"
    "9c6360000002000155009ba7560000000049454e44ae426082"
)


def add_logo_placeholder(unpack: Path, slide_file: str, media_name: str, x_pt: float, y_pt: float, w_pt: float, h_pt: float) -> int:
    """Añade un <p:pic> nuevo (imagen transparente de relleno) a slide_file,
    con su propia relación de media — devuelve el shape_id asignado."""
    slides_dir = unpack / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    media_dir = unpack / "ppt" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    (media_dir / media_name).write_bytes(_BLANK_PNG)

    rels_path = rels_dir / f"{slide_file}.rels"
    rels_xml = _read(rels_path)
    existing_rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    new_rid = f"rId{max(existing_rids) + 1}"
    rel = (
        f'<Relationship Id="{new_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="../media/{media_name}"/>'
    )
    _write(rels_path, rels_xml.replace("</Relationships>", rel + "</Relationships>"))

    slide_path = slides_dir / slide_file
    slide_xml = _read(slide_path)
    shape_id = max_shape_id(slide_xml) + 1
    pic_xml = (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{shape_id}" name="Logo cliente"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{new_rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{pt(x_pt)}" y="{pt(y_pt)}"/><a:ext cx="{pt(w_pt)}" cy="{pt(h_pt)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    _write(slide_path, insert_shapes_before_sptree_close(slide_xml, [pic_xml]))
    return shape_id


def repack(unpack: Path, out_path: Path) -> None:
    if out_path.exists():
        out_path.unlink()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for fp in sorted(unpack.rglob("*")):
            if fp.is_file():
                zout.write(fp, fp.relative_to(unpack))


def unpack_to(pptx_path: Path, dest: Path) -> None:
    if dest.exists():
        shutil.rmtree(dest)
    dest.mkdir(parents=True)
    with zipfile.ZipFile(pptx_path, "r") as z:
        z.extractall(dest)


# ─── HC: equipo "sin Eduardo" (slide2b.xml) ───────────────────────────────
def build_hc_team_sin_eduardo(unpack: Path) -> None:
    slides = unpack / "ppt" / "slides"
    _duplicate_slide(unpack, "slide2.xml", "slide2b.xml", "slide2.xml")
    s = _read(slides / "slide2b.xml")
    for sid in (10, 11, 12, 26):  # Eduardo: nombre, rol, bio, foto
        s = remove_shape_by_id(s, sid)
    # Manuel Pina (texto ids 15/16/17, foto id28) pasa de la columna 2 (x=347pt) a la 1 (x=60pt).
    for sid in (15, 16, 17):
        s = re.sub(
            r'(<p:cNvPr\s+id="%d"[^/]*/>.*?<a:off\s+x=")\d+(")' % sid,
            lambda m: m.group(1) + str(pt(60.0)) + m.group(2),
            s, count=1, flags=re.DOTALL,
        )
    s = re.sub(
        r'(<p:cNvPr\s+id="28"[^/]*/>.*?<a:off\s+x=")\d+(")',
        lambda m: m.group(1) + str(pt(116.5)) + m.group(2),
        s, count=1, flags=re.DOTALL,
    )
    _write(slides / "slide2b.xml", s)


# ─── HC: tabla comparativa (slide5b.xml) ──────────────────────────────────
def build_hc_comparison_table(unpack: Path) -> None:
    slides = unpack / "ppt" / "slides"
    _duplicate_slide(unpack, "slide5.xml", "slide5b.xml", "slide5.xml")
    s = _read(slides / "slide5b.xml")
    for sid in (7, 9, 10, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24):
        s = remove_shape_by_id(s, sid)
    s = replace_text_in_xml(s, "POR QUÉ TESSERA", "CÓMO TRABAJAMOS")
    s = replace_text_in_xml(s, "Por qué Tessera", "No somos un proveedor de CVs")
    s = replace_text_in_xml(
        s, "Procesos claros y compromisos que se cumplen: así movemos el talento.",
        "Lo que cambia cuando Tessera trabaja con tu equipo.",
    )

    left_body = "\n".join([
        "- Envía CVs en volumen y espera.",
        "- Aparece solo cuando hay una vacante.",
        "- Cambia de interlocutor en cada proceso.",
        "- Conoce el puesto, no tu negocio.",
        "- Desaparece tras la incorporación.",
    ])
    right_body = "\n".join([
        "Presenta solo a quien encaja, en 72 horas.",
        "Seguimiento semanal, con o sin vacantes.",
        "Un partner senior fijo, siempre localizable.",
        "Conoce tu negocio, tu cultura y tus retos.",
        "Sigue contigo tras la incorporación.",
    ])
    new_shapes = [
        text_sp(30, 60, 180, 380, 20, [("LA CONSULTORA TÍPICA", 900, "Raleway Medium", "8AA3A5")]),
        text_sp(31, 60, 212, 380, 190, [(left_body, 1100, "Raleway", "202031")]),
        text_sp(32, 460, 270, 50, 30, [("VS", 1200, "Raleway SemiBold", "587579")], algn="ctr"),
        text_sp(33, 520, 180, 380, 20, [("TESSERA, TU PARTNER", 900, "Raleway Medium", "587579")]),
        text_sp(34, 520, 212, 380, 190, [(right_body, 1100, "Raleway", "202031")]),
    ]
    s = insert_shapes_before_sptree_close(s, new_shapes)
    _write(slides / "slide5b.xml", s)


# ─── HC: Próximos pasos (slide10b.xml) ────────────────────────────────────
def build_hc_next_steps(unpack: Path) -> None:
    slides = unpack / "ppt" / "slides"
    _duplicate_slide(unpack, "slide5.xml", "slide10b.xml", "slide10.xml")
    s = _read(slides / "slide10b.xml")
    for sid in (7, 9, 10, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24):
        s = remove_shape_by_id(s, sid)
    s = replace_text_in_xml(s, "POR QUÉ TESSERA", "PRÓXIMOS PASOS")
    s = replace_text_in_xml(s, "Por qué Tessera", "Próximos pasos")
    s = replace_text_in_xml(
        s, "Procesos claros y compromisos que se cumplen: así movemos el talento.",
        "Desde la firma, así arrancamos juntos.",
    )

    steps = [
        ("01. La firma", "Cerramos condiciones y te asignamos tu partner senior."),
        ("02. Inmersión", "Sesión con tus responsables: cultura, retos y perfiles."),
        ("03. Primer shortlist", "Candidatos filtrados, con nuestra valoración de cada uno, en 72 horas."),
        ("04. Seguimiento", "Reunión semanal y acompañamiento tras cada incorporación."),
    ]
    card_w, gap, x0, y0, card_h = 195.0, 20.0, 60.0, 172.0, 190.0
    new_shapes = []
    next_id = 30
    for i, (title, body) in enumerate(steps):
        cx = x0 + i * (card_w + gap)
        new_shapes.append(rect_sp(next_id, cx, y0, card_w, card_h))
        new_shapes.append(text_sp(
            next_id + 1, cx + 20, y0 + 29, card_w - 40, card_h - 50,
            [(title, 1300, "Raleway SemiBold", "202031"), (body, 1000, "Raleway", "202031")],
        ))
        next_id += 2
    s = insert_shapes_before_sptree_close(s, new_shapes)
    _write(slides / "slide10b.xml", s)


# ─── HC: honorarios — Modalidad Partner (slide13.xml, en el sitio) ────────
def build_hc_honorarios(unpack: Path) -> None:
    path = unpack / "ppt" / "slides" / "slide13.xml"
    s = _read(path)
    s = replace_text_in_xml(s, "Modalidad a éxito en Headhunting", "Modalidad Partner")
    detail_old = "\n".join([
        "12% sobre un SBA de referencia de 36.000 € (4.320 €)",
        "1/3 - 1.440 € a la firma (anticipo fijo)",
        "2/3 - 2.880 € ajustable, a la incorporación",
    ])
    s = replace_text_in_xml(s, "Salario Bruto Anual", detail_old)
    # Ensanchar/agrandar la caja del detalle (id=21) y la banda de fondo
    # (id=18) para que quepan las 3 líneas nuevas sin solaparse con el "13%".
    s = re.sub(
        r'(<p:cNvPr\s+id="21"[^/]*/>.*?<a:ext\s+cx=")\d+(" cy=")\d+(")',
        lambda m: m.group(1) + str(pt(560.0)) + m.group(2) + str(pt(90.0)) + m.group(3),
        s, count=1, flags=re.DOTALL,
    )
    s = re.sub(
        r'(<p:cNvPr\s+id="18"[^/]*/>.*?<a:ext\s+cx="\d+" cy=")\d+(")',
        lambda m: m.group(1) + str(pt(170.0)) + m.group(2),
        s, count=1, flags=re.DOTALL,
    )
    _write(path, s)


def build_plantilla_hc() -> None:
    src = ROOT / "plantilla_hc.pptx"
    unpack = ROOT / "scripts" / "build_templates" / "_work_hc"
    unpack_to(src, unpack)

    build_hc_team_sin_eduardo(unpack)
    build_hc_comparison_table(unpack)
    build_hc_next_steps(unpack)
    build_hc_honorarios(unpack)
    add_logo_placeholder(unpack, "slide1.xml", "logo_cliente_portada.png", 560, 400, 320, 70)
    add_logo_placeholder(unpack, "slide15.xml", "logo_cliente_cierre.png", 560, 250, 320, 70)

    out = ROOT / "plantilla_hc.pptx"
    repack(unpack, out)
    print(f"OK: {out}")


def build_plantilla_hc_finance() -> None:
    src = ROOT / "plantilla_hc_finance.pptx"
    unpack = ROOT / "scripts" / "build_templates" / "_work_hcf"
    unpack_to(src, unpack)

    _remove_slides(unpack, ["slide10.xml", "slide11.xml", "slide12.xml"])
    add_logo_placeholder(unpack, "slide1.xml", "logo_cliente_portada.png", 560, 400, 320, 70)
    add_logo_placeholder(unpack, "slide15.xml", "logo_cliente_cierre.png", 560, 250, 320, 70)

    out = ROOT / "plantilla_hc_finance.pptx"
    repack(unpack, out)
    print(f"OK: {out}")


if __name__ == "__main__":
    build_plantilla_hc()
    build_plantilla_hc_finance()
