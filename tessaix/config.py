"""Rutas, constantes y catálogos compartidos por toda la app."""
from pathlib import Path

HERE = Path(__file__).parent.parent


def _find(f: str) -> Path:
    """Busca un asset primero en la raíz del proyecto y luego en /assets."""
    for p in [HERE / f, HERE / "assets" / f]:
        if p.exists():
            return p
    return HERE / f


PLANTILLA = _find("plantilla_2.pptx")
AILERONS = _find("Ailerons-Typeface.otf")
LOGO_B = _find("logo_blanco.png")
FAVICON = _find("logo1.png")

PASSWORD = "Sales2026@"

SERVICES_HC = {
    "headhunting": "Headhunting",
    "outsourcing": "Outsourcing Time & Material",
    "rpo": "RPO / Equipo embebido",
    "salary": "Consultoría Salarial",
}

# Pasos del wizard. "Revisión" es el paso intermedio en el que se enseña y
# puede editarse todo lo que va a aparecer en el PowerPoint (incluido el
# logo del cliente) ANTES de generar el archivo final.
STEPS = ["Config", "Cliente", "Servicios", "Presentador", "Contexto", "Revisión", "Generar"]

STEP_CONFIG = 0
STEP_CLIENTE = 1
STEP_SERVICIOS = 2
STEP_PRESENTADOR = 3
STEP_CONTEXTO = 4
STEP_REVIEW = 5
STEP_GENERATE = 6
