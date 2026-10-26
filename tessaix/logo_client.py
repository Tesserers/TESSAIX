"""Búsqueda, validación de calidad y composición del logo del cliente.

El hueco de la portada donde va el logo tiene un tamaño FIJO (viene de la
plantilla). Si simplemente se pega ahí el logo tal cual se descarga —con su
propia relación de aspecto— PowerPoint lo ESTIRA para llenar el hueco y el
logo sale deformado. Este módulo evita eso: el logo se compone siempre sobre
un lienzo con la MISMA relación de aspecto que el hueco de destino, centrado
y sin recortar, así que el estiramiento final de PowerPoint es uniforme en
ambos ejes y el logo no se deforma.

También expone una comprobación de calidad (resolución mínima) para poder
avisar al usuario ANTES de generar el PowerPoint, en vez de que se dé cuenta
al abrir el archivo y tenga que arreglarlo a mano.
"""
from __future__ import annotations

import io

import requests as rq
from PIL import Image, ImageDraw, ImageOps

# Por debajo de esto, un logo se considera "de baja calidad" para el hueco
# habitual de portada (~4.2 x 2.7 cm a resolución de impresión). Se sigue
# pudiendo usar si el usuario lo acepta, pero se le avisa.
MIN_LOGO_DIM = 220

# Resolución del lienzo de salida (lado más largo), para que el logo se vea
# nítido aunque la fuente original sea pequeña.
CANVAS_LONG_SIDE = 1000

_UA = {"User-Agent": "Mozilla/5.0"}


def normalize_domain(website: str, fallback_name: str = "") -> str:
    domain = (website or "").strip()
    if not domain and fallback_name:
        domain = f"{fallback_name.lower().replace(' ', '')}.com"
    domain = domain.replace("https://", "").replace("http://", "").replace("www.", "")
    return domain.split("/")[0].strip()


def fetch_client_logo(domain: str, timeout: int = 6) -> dict | None:
    """
    Intenta descargar el logo del dominio dado probando varias fuentes, de
    mayor a menor calidad esperada. Devuelve un dict con los bytes crudos,
    el tamaño nativo y si supera el umbral de calidad, o None si no se
    encontró nada usable.
    """
    if not domain:
        return None

    sources = [
        (f"https://logo.clearbit.com/{domain}?size=512&format=png", "Clearbit"),
        (f"https://www.google.com/s2/favicons?domain={domain}&sz=256", "Google Favicons"),
    ]
    for url, label in sources:
        try:
            resp = rq.get(url, timeout=timeout, headers=_UA)
        except Exception:
            continue
        if resp.status_code != 200 or len(resp.content) < 300:
            continue
        try:
            img = Image.open(io.BytesIO(resp.content))
            img.load()
        except Exception:
            continue
        w, h = img.size
        if w < 16 or h < 16:
            continue
        return {
            "bytes": resp.content,
            "width": w,
            "height": h,
            "source": label,
            "domain": domain,
            "ok": min(w, h) >= MIN_LOGO_DIM,
        }
    return None


def logo_quality_label(meta: dict | None) -> tuple[str, str]:
    """Devuelve (nivel, mensaje) para mostrar como salvaguarda en la UI."""
    if not meta:
        return "missing", "No se ha encontrado ningún logo automáticamente para este dominio."
    if meta["ok"]:
        return "ok", f"Logo encontrado en {meta['source']} ({meta['width']}×{meta['height']}px). Buena calidad."
    return (
        "low",
        f"Logo encontrado en {meta['source']} pero de baja resolución "
        f"({meta['width']}×{meta['height']}px). Puede verse borroso o pixelado en la portada — "
        f"te recomendamos subir uno de mejor calidad manualmente.",
    )


def _corners_match(img: Image.Image, tolerance: int) -> bool:
    w, h = img.size
    px = img.load()
    corners = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
    ref = px[corners[0]]
    if ref[3] == 0:
        return False  # ya transparente: nada que quitar
    return all(
        all(abs(a - b) <= tolerance for a, b in zip(ref, px[c]))
        for c in corners[1:]
    )


def strip_flat_background(logo_bytes: bytes, tolerance: int = 20) -> bytes:
    """
    Si el logo de origen trae su propio fondo de color PLANO pegado detrás
    (p.ej. una caja de color sólido, como suelen traer los logos
    descargados automáticamente de favicons/Clearbit), lo convierte a
    transparente. Comprueba que las 4 esquinas sean del mismo color — si lo
    son, aplica un relleno por contigüidad (flood fill) desde cada esquina,
    así solo se vacía la región de fondo conectada a los bordes, nunca el
    interior del propio logotipo aunque comparta ese color en algún punto.

    Si el fondo NO es uniforme (ya es transparente, o es una foto/textura),
    no se toca nada — devuelve los bytes originales.
    """
    img = Image.open(io.BytesIO(logo_bytes))
    img = ImageOps.exif_transpose(img).convert("RGBA")
    if not _corners_match(img, tolerance):
        return logo_bytes
    w, h = img.size
    for seed in [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]:
        ImageDraw.floodfill(img, seed, (0, 0, 0, 0), thresh=tolerance)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def compose_logo_for_placeholder(
    logo_bytes: bytes,
    target_w_emu: int,
    target_h_emu: int,
    pad_frac: float = 0.06,
    bg=(0, 0, 0, 0),
) -> bytes:
    """
    Encaja `logo_bytes` (cualquier PNG/JPG) en un lienzo cuya relación de
    aspecto es EXACTAMENTE la del hueco de destino (target_w_emu /
    target_h_emu), centrado y sin deformar. Como PowerPoint rellena el hueco
    estirando la imagen (fillRect), y el lienzo ya tiene la proporción
    correcta, el estiramiento resultante es uniforme y el logo conserva su
    forma original.

    El lienzo es TRANSPARENTE por defecto (no blanco): el logo se pega tal
    cual es, sin añadirle una caja de color alrededor — se adapta a su
    propia silueta y deja ver el fondo de la diapositiva a su alrededor. Si
    el logo de origen trae su propio fondo de color PLANO (p.ej. una caja
    sólida detrás del nombre de la empresa), se quita primero con
    `strip_flat_background` — si el fondo no es plano (foto, textura), se
    deja tal cual, no hay forma automática segura de recortarlo.

    Devuelve bytes PNG listos para sustituir el media del placeholder.
    """
    logo_bytes = strip_flat_background(logo_bytes)
    img = Image.open(io.BytesIO(logo_bytes))
    img = ImageOps.exif_transpose(img).convert("RGBA")

    target_ratio = target_w_emu / target_h_emu
    if target_ratio >= 1:
        canvas_w = CANVAS_LONG_SIDE
        canvas_h = max(1, round(CANVAS_LONG_SIDE / target_ratio))
    else:
        canvas_h = CANVAS_LONG_SIDE
        canvas_w = max(1, round(CANVAS_LONG_SIDE * target_ratio))

    canvas = Image.new("RGBA", (canvas_w, canvas_h), bg)

    avail_w = max(1, int(canvas_w * (1 - pad_frac * 2)))
    avail_h = max(1, int(canvas_h * (1 - pad_frac * 2)))
    scale = min(avail_w / img.width, avail_h / img.height)
    # No hacemos "upscale" agresivo de logos ya pequeños más allá de 2x para
    # no generar artefactos raros; el aviso de baja calidad ya cubre ese caso.
    scale = min(scale, 4.0)
    new_w = max(1, round(img.width * scale))
    new_h = max(1, round(img.height * scale))
    resized = img.resize((new_w, new_h), Image.LANCZOS)

    x = (canvas_w - new_w) // 2
    y = (canvas_h - new_h) // 2
    canvas.alpha_composite(resized, (x, y))

    buf = io.BytesIO()
    canvas.save(buf, "PNG")
    return buf.getvalue()


def checkerboard_preview(composed_png_bytes: bytes, tile: int = 16) -> bytes:
    """
    Igual que `composed_png_bytes` pero con un patrón de damero pegado
    DEBAJO, solo para que la vista previa en la app deje claro que el fondo
    es transparente (y no una caja blanca) — el archivo que se inserta en
    el .pptx sigue siendo el PNG transparente original, esto es solo para
    que se vea en pantalla.
    """
    logo = Image.open(io.BytesIO(composed_png_bytes)).convert("RGBA")
    w, h = logo.size
    checker = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    light, dark = (240, 240, 240, 255), (214, 214, 214, 255)
    for y in range(0, h, tile):
        for x in range(0, w, tile):
            color = dark if ((x // tile) + (y // tile)) % 2 else light
            checker.paste(color, (x, y, min(x + tile, w), min(y + tile, h)))
    checker.alpha_composite(logo)
    buf = io.BytesIO()
    checker.save(buf, "PNG")
    return buf.getvalue()
