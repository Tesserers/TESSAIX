"""Generación de PNGs con la tipografía Ailerons y utilidades de imagen."""
import base64
import io

from PIL import Image, ImageDraw, ImageFont

from .config import AILERONS, LOGO_B


def make_ailerons_png(text, color=(255, 255, 255), size=64, width=520, height=90) -> str:
    """Renderiza `text` con la fuente Ailerons y devuelve un PNG en base64."""
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(str(AILERONS), size)
    except Exception:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((width - tw) // 2, (height - th) // 2 - bbox[1]), text, font=font, fill=color)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode()


def logo_b64() -> str:
    try:
        return base64.b64encode(LOGO_B.read_bytes()).decode()
    except Exception:
        return ""
