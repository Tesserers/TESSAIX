"""Estimación de ajuste de texto a una caja de tamaño fijo, sin depender de
que PowerPoint recalcule nada al abrir el archivo.

Contexto: la plantilla usa cajas de texto de tamaño FIJO para el contenido
que genera la IA. El fix anterior (`xml_utils.enable_shrink_autofit`) marca
esas cajas con `<a:normAutofit/>` para que PowerPoint reduzca la fuente si
no cabe — pero esa señal solo funciona si el programa que abre el archivo
la recalcula, y no todos lo hacen de forma fiable al abrir (algunas
versiones de PowerPoint, Google Slides o LibreOffice la ignoran hasta que
alguien edita el texto a mano). Por eso seguían viéndose solapes.

Este módulo hace el cálculo NOSOTROS, en Python, antes de escribir el
archivo: estima cuántas líneas ocupará un texto dado su ancho de caja y
tamaño de fuente (simulando el ajuste de línea por palabras), y si no cabe
en el alto disponible, calcula un factor de reducción para que sí quepa.
El tamaño resultante se escribe DIRECTAMENTE en el `sz` de cada run — así
funciona igual en cualquier visor, sin depender de ningún recálculo.

No es un motor de layout real (no conocemos las métricas exactas de
Raleway sin incrustar la fuente), así que se usa una estimación de ancho
de carácter deliberadamente conservadora: prefiere encoger un poco de más
antes que arriesgarse a que el texto se salga de su caja.
"""
from __future__ import annotations

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

# No se reduce el texto por debajo de este porcentaje del tamaño original,
# para que nunca quede ilegible — a partir de aquí se prefiere dejar que el
# texto roce el borde de la caja antes que encoger más.
MIN_SCALE = 0.55

# ...ni tampoco por debajo de este tamaño absoluto, gane quien gane de los
# dos límites: un título nunca debería acabar más pequeño que un cuerpo de
# texto legible, así que se respeta el que sea más permisivo con el tamaño.
MIN_FONT_SIZE_PT = 8.0

# Margen de seguridad: se exige que el contenido quepa en un
# SAFETY_MARGIN del alto real de la caja, no en el 100%, para absorber la
# imprecisión de la estimación.
SAFETY_MARGIN = 0.94

# Insets por defecto de un <a:bodyPr> en OOXML cuando no se especifican.
DEFAULT_INSET_LR_EMU = 91440   # 0.1"
DEFAULT_INSET_TB_EMU = 45720   # 0.05"


def emu_to_pt(emu: float) -> float:
    return emu / EMU_PER_PT


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


def fit_scale(
    paragraphs: list[dict],
    box_width_pt: float,
    box_height_pt: float,
    min_scale: float = MIN_SCALE,
) -> float:
    """
    Dado un conjunto de párrafos (cada uno con "text", "size_pt",
    "line_spacing_pct" y "space_before_pt") que comparten una misma caja,
    devuelve el mayor factor de escala k <= 1.0 tal que, si todos los
    tamaños de fuente se multiplican por k, el contenido cabe en
    `box_height_pt` (con margen de seguridad). Nunca baja de `min_scale` NI
    dejaría al párrafo más pequeño del grupo por debajo de
    `MIN_FONT_SIZE_PT` — gana el límite que sea más permisivo con el tamaño.
    """
    if box_width_pt <= 0 or box_height_pt <= 0 or not paragraphs:
        return 1.0

    smallest_size = min((p["size_pt"] for p in paragraphs if p.get("size_pt")), default=None)
    if smallest_size:
        min_scale = max(min_scale, MIN_FONT_SIZE_PT / smallest_size)
        min_scale = min(min_scale, 1.0)

    def total_height(k: float) -> float:
        return sum(
            paragraph_height_pt(
                p["text"], p["size_pt"] * k, box_width_pt,
                p.get("line_spacing_pct", 100.0), p.get("space_before_pt", 0.0),
            )
            for p in paragraphs
        )

    target = box_height_pt * SAFETY_MARGIN
    if total_height(1.0) <= target:
        return 1.0

    # Búsqueda lineal simple (rango pequeño, no hace falta más finura).
    k = 1.0
    best = min_scale
    while k >= min_scale:
        if total_height(k) <= target:
            best = k
            break
        k -= 0.02
    return round(best, 3)
