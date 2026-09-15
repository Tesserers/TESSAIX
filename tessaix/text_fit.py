"""Estimación de ajuste de texto a una caja de tamaño fijo, sin depender de
que PowerPoint recalcule nada al abrir el archivo.

Contexto: la plantilla usa cajas de texto de tamaño FIJO para el contenido
que genera la IA. El fix anterior (`xml_utils.enable_shrink_autofit`) marca
esas cajas con `<a:normAutofit/>` para que PowerPoint reduzca la fuente si
no cabe — pero esa señal solo funciona si el programa que abre el archivo
la recalcula, y no todos lo hacen de forma fiable al abrir. Por eso seguían
viéndose solapes.

Este módulo hace el cálculo NOSOTROS, en Python, antes de escribir el
archivo. La estrategia tiene dos niveles:

1. Un encogido MODESTO del tamaño de fuente (nunca por debajo de un suelo
   legible) cuando con eso basta para que el texto quepa en su caja
   original.
2. Si ni encogiendo hasta el suelo cabe, no se sigue reduciendo el texto
   (quedaría ilegible) — en su lugar, quien llama (`pptx_builder`) hace
   crecer la caja y desplaza lo que tenga debajo. Este módulo expone
   `fit_scale`, que dice cuál de los dos casos aplica (`fits=True/False`) y
   cuánto alto hace falta de verdad.

No es un motor de layout real (no conocemos las métricas exactas de
Raleway sin incrustar la fuente), así que se usa una estimación de ancho
de carácter deliberadamente conservadora: prefiere pedir un poco más de
alto antes que arriesgarse a que el texto se salga de su caja.
"""
from __future__ import annotations

from dataclasses import dataclass

EMU_PER_PT = 12700

# Ancho medio de carácter para una sans-serif geométrica como Raleway,
# expresado como fracción del tamaño de fuente (em). Calibrado contra los
# textos y cajas ORIGINALES de la plantilla (que sabemos que caben bien):
# con este valor, el texto de fábrica no dispara una reducción salvo en las
# dos cajas que ya vienen ajustadas al límite en el propio diseño — y ahí
# el ajuste es de apenas un 6%, imperceptible.
AVG_CHAR_WIDTH_EM = 0.48

# Factor de interlineado base (alto de línea / tamaño de fuente) para una
# sola línea a espaciado "sencillo".
LINE_HEIGHT_FACTOR = 1.2

# Suelo relativo de emergencia cuando no se puede determinar el tamaño de
# fuente original de un párrafo (caso raro). El suelo normal es el
# absoluto (MIN_FONT_SIZE_PT) — ver `fit_scale`.
MIN_SCALE = 0.55

# Nunca se reduce un cuerpo de texto por debajo de este tamaño: a partir de
# aquí se prefiere hacer crecer la caja (y desplazar lo que haya debajo)
# antes que seguir encogiendo. Los subtítulos/títulos usan min_scale=1.0
# (no se tocan nunca) desde `xml_utils`, así que este suelo aplica sobre
# todo al texto de cuerpo.
MIN_FONT_SIZE_PT = 10.0

# Margen de seguridad: se exige que el contenido quepa en un
# SAFETY_MARGIN del alto real de la caja, no en el 100%, para absorber la
# imprecisión de la estimación.
SAFETY_MARGIN = 0.94

# Cuando hace falta crecer una caja, se le da un poco más del mínimo
# estricto para no dejarla al filo del borde.
GROW_BUFFER = 1.05

# Una caja puede crecer como mucho esto respecto a su alto original. Sin
# este tope, un texto verdaderamente desmedido podría inflar la caja hasta
# salirse de la diapositiva — justo lo que se quiere evitar. Con los
# límites de longitud de ai_content.py (60 caracteres por título, ~220 por
# cuerpo) un crecimiento razonable se queda muy por debajo de este tope.
MAX_GROW_MULTIPLIER = 2.5

# Suelo de ÚLTIMO RECURSO: si ni haciendo crecer la caja hasta el tope
# anterior cabe el texto, se prioriza no salirse de la página por encima de
# mantener el tamaño — se permite encoger hasta aquí (relativo al tamaño
# ORIGINAL, no al suelo normal de cada rol) antes que dejar la caja
# desbordar la diapositiva.
ABSOLUTE_MIN_SCALE = 0.5

# Insets por defecto de un <a:bodyPr> en OOXML cuando no se especifican.
DEFAULT_INSET_LR_EMU = 91440   # 0.1"
DEFAULT_INSET_TB_EMU = 45720   # 0.05"


def emu_to_pt(emu: float) -> float:
    return emu / EMU_PER_PT


def pt_to_emu(pt: float) -> int:
    return round(pt * EMU_PER_PT)


def estimate_wrapped_lines(text: str, box_width_pt: float, font_size_pt: float) -> int:
    """Simula un ajuste de línea por palabras (greedy word-wrap) y devuelve
    cuántas líneas ocuparía `text` a `font_size_pt` en una caja de
    `box_width_pt` de ancho."""
    text = (text or "").strip()
    if not text:
        return 1
    if font_size_pt <= 0 or box_width_pt <= 0:
        return 1

    char_width_pt = AVG_CHAR_WIDTH_EM * font_size_pt
    max_chars_per_line = max(1, int(box_width_pt / char_width_pt))

    lines = 0
    for raw_line in text.split("\n"):
        words = raw_line.split()
        lines += 1  # cada línea explícita del texto original empieza como mínimo 1 línea
        col = 0
        for word in words:
            wlen = len(word)
            if wlen > max_chars_per_line:
                # Palabra más larga que una línea entera: se parte igualmente.
                if col > 0:
                    lines += 1
                lines += wlen // max_chars_per_line
                col = wlen % max_chars_per_line
                continue
            needed = wlen if col == 0 else col + 1 + wlen
            if needed > max_chars_per_line:
                lines += 1
                col = wlen
            else:
                col = needed
    return max(1, lines)


def paragraph_height_pt(
    text: str, font_size_pt: float, box_width_pt: float,
    line_spacing_pct: float = 100.0, space_before_pt: float = 0.0,
) -> float:
    """
    El interlineado (`line_spacing_pct`) separa una línea de la siguiente —
    no infla el alto de una línea que va sola. Aplicarlo también a la
    primera línea sobre-estimaría el alto necesario (y encogería sin
    necesidad) cualquier título a una sola línea con interlineado holgado
    (p.ej. 140%) puramente decorativo.
    """
    lines = estimate_wrapped_lines(text, box_width_pt, font_size_pt)
    base_line_height = font_size_pt * LINE_HEIGHT_FACTOR
    if lines <= 1:
        content_height = base_line_height
    else:
        spaced_line_height = base_line_height * (line_spacing_pct / 100.0)
        content_height = base_line_height + (lines - 1) * spaced_line_height
    return space_before_pt + content_height


def total_height_pt(paragraphs: list[dict], box_width_pt: float, scale: float = 1.0) -> float:
    return sum(
        paragraph_height_pt(
            p["text"], p["size_pt"] * scale, box_width_pt,
            p.get("line_spacing_pct", 100.0), p.get("space_before_pt", 0.0),
        )
        for p in paragraphs
    )


@dataclass(frozen=True)
class FitResult:
    scale: float               # factor aplicado a cada tamaño de fuente original
    required_height_pt: float  # alto que ocupa el contenido a ese scale
    fits: bool                 # True si required_height_pt cabe en la caja actual


def fit_scale(
    paragraphs: list[dict],
    box_width_pt: float,
    box_height_pt: float,
    min_scale: float | None = None,
) -> FitResult:
    """
    Dado un conjunto de párrafos (cada uno con "text", "size_pt",
    "line_spacing_pct" y "space_before_pt") que comparten una misma caja:

    - Si el contenido ya cabe a tamaño completo (o encogiendo hasta
      `min_scale`), devuelve ese factor con `fits=True`.
    - Si ni encogiendo hasta `min_scale` cabe, devuelve `min_scale` con
      `fits=False` y el alto que de verdad hace falta — quien llama debe
      entonces hacer crecer la caja en vez de seguir reduciendo el texto.

    `min_scale=None` (por defecto) calcula el suelo a partir de
    `MIN_FONT_SIZE_PT` y el tamaño de fuente más pequeño del grupo — así un
    cuerpo de texto nunca baja de ese tamaño absoluto. Pásalo explícitamente
    a `1.0` para contenido que no debe encogerse nunca (p.ej. subtítulos).
    """
    if not paragraphs or box_width_pt <= 0 or box_height_pt <= 0:
        return FitResult(1.0, 0.0, True)

    if min_scale is None:
        smallest_size = min((p["size_pt"] for p in paragraphs if p.get("size_pt")), default=None)
        min_scale = (MIN_FONT_SIZE_PT / smallest_size) if smallest_size else MIN_SCALE
    min_scale = max(0.0, min(min_scale, 1.0))

    target = box_height_pt * SAFETY_MARGIN

    h_full = total_height_pt(paragraphs, box_width_pt, 1.0)
    if h_full <= target:
        return FitResult(1.0, h_full, True)

    # Búsqueda lineal simple (rango pequeño, no hace falta más finura).
    k = 0.98
    while k >= min_scale:
        h = total_height_pt(paragraphs, box_width_pt, k)
        if h <= target:
            return FitResult(round(k, 3), h, True)
        k -= 0.02

    h_floor = total_height_pt(paragraphs, box_width_pt, min_scale)
    return FitResult(round(min_scale, 3), h_floor, False)
