"""
Script de construcción (uso puntual, no se ejecuta en runtime) — SEGUNDA
PASADA: a diferencia de rebuild_v2.py (que parcheaba la plantilla antigua),
este construye plantilla_hc.pptx y plantilla_hc_finance.pptx DIRECTAMENTE
a partir de los archivos reales que dio el usuario (misma identidad visual
2026 real, no una reconstrucción a mano): "nueva versi con fee parter"
(CONFEE) para Human Capital, "Catalogo_Servicios_Tessera_Ejecutivo"
(CATALOGO) para Human Capital + Finance. Así no hay que reinventar ningún
diseño de tarjeta: se reutilizan las diapositivas reales tal cual, y solo
se les aplican los cambios estructurales acordados (ver conversación):
- Human Capital: las dos variantes de equipo (con/sin Eduardo Serrano) ya
  existen de fábrica en el archivo (slide2.xml/slide3.xml) — no hay que
  clonar nada. Se añade la alternativa de sedes "solo Madrid" (no existe
  de fábrica) y el hueco de logo de cliente en portada y cierre.
- Human Capital + Finance: el archivo ya no tiene ningún flujo de fee —
  se usa tal cual, solo con el hueco de logo añadido.

Ejecutar desde la raíz del repo: python scripts/build_templates/rebuild_v3.py
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from tessaix.pptx_builder import _duplicate_slide, _read, _write  # noqa: E402
from tessaix.xml_utils import remove_shape_by_id  # noqa: E402

REBRAND_DIR = Path(r"C:\Users\david\Documents\rebrand")
CONFEE_SRC = REBRAND_DIR / "nueva versi con fee parter - Queremos ser tu partner - Nueva imagen 2026 v2 (1).pptx"
CATALOGO_SRC = REBRAND_DIR / "Catalogo_Servicios_Tessera_Ejecutivo - Nueva imagen 2026 (1).pptx"

EMU_PER_PT = 12700


def pt(v: float) -> int:
    return int(round(v * EMU_PER_PT))


_BLANK_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a4944415478"
    "9c6360000002000155009ba7560000000049454e44ae426082"
)


def max_shape_id(slide_xml: str) -> int:
    ids = [int(m) for m in re.findall(r'<p:cNvPr\s+id="(\d+)"', slide_xml)]
    return max(ids) if ids else 1


def insert_shapes_before_sptree_close(slide_xml: str, shapes_xml: list[str]) -> str:
    return slide_xml.replace("</p:spTree>", "".join(shapes_xml) + "</p:spTree>")


def add_logo_placeholder(unpack: Path, slide_file: str, media_name: str, x_pt: float, y_pt: float, w_pt: float, h_pt: float) -> int:
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


def build_plantilla_hc() -> None:
    unpack = ROOT / "scripts" / "build_templates" / "_work_hc3"
    unpack_to(CONFEE_SRC, unpack)
    slides = unpack / "ppt" / "slides"

    # "Dónde estamos" (slide4.xml, 3 ciudades) no trae de fábrica una
    # alternativa "solo Madrid" — se clona y se quitan Bilbao/Oviedo y los
    # dos datos que solo tienen sentido con las 3 sedes.
    _duplicate_slide(unpack, "slide4.xml", "slide4b.xml", "slide4.xml")
    s = _read(slides / "slide4b.xml")
    for sid in (9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20):
        s = remove_shape_by_id(s, sid)
    _write(slides / "slide4b.xml", s)

    # Hueco de logo de cliente: portada (slide1.xml) y cierre (slide17.xml).
    add_logo_placeholder(unpack, "slide1.xml", "logo_cliente_portada.png", 560, 390, 320, 70)
    add_logo_placeholder(unpack, "slide17.xml", "logo_cliente_cierre.png", 560, 250, 320, 70)

    out = ROOT / "plantilla_hc.pptx"
    repack(unpack, out)
    print(f"OK: {out}")


def build_plantilla_hc_finance() -> None:
    unpack = ROOT / "scripts" / "build_templates" / "_work_hcf3"
    unpack_to(CATALOGO_SRC, unpack)

    add_logo_placeholder(unpack, "slide1.xml", "logo_cliente_portada.png", 560, 390, 320, 70)
    add_logo_placeholder(unpack, "slide11.xml", "logo_cliente_cierre.png", 560, 250, 320, 70)

    out = ROOT / "plantilla_hc_finance.pptx"
    repack(unpack, out)
    print(f"OK: {out}")


if __name__ == "__main__":
    build_plantilla_hc()
    build_plantilla_hc_finance()
