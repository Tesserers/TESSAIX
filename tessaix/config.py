"""Rutas, constantes y catálogos compartidos por toda la app."""
from pathlib import Path

HERE = Path(__file__).parent.parent


def _find(f: str) -> Path:
    """Busca un asset primero en la raíz del proyecto y luego en /assets."""
    for p in [HERE / f, HERE / "assets" / f]:
        if p.exists():
            return p
    return HERE / f


AILERONS = _find("Ailerons-Typeface.otf")
LOGO_B = _find("logo_blanco.png")
FAVICON = _find("logo1.png")

# Plantillas maestras por línea de negocio. Cada una es un superconjunto:
# trae ya dentro los bloques opcionales (equipo con/sin Eduardo, sedes
# Madrid-sola/las-3, términos y condiciones + honorarios) y pptx_builder.py
# decide en tiempo de generación cuáles conservar y cuáles quitar.
PLANTILLA_HC = _find("plantilla_hc.pptx")                  # Human Capital solo
PLANTILLA_HC_FINANCE = _find("plantilla_hc_finance.pptx")  # Human Capital + Finance

PASSWORD = "Sales2026@"

BUSINESS_LINE_HC = "hc"
BUSINESS_LINE_HC_FINANCE = "hc_finance"

BUSINESS_LINES = {
    BUSINESS_LINE_HC: "Human Capital",
    BUSINESS_LINE_HC_FINANCE: "Human Capital + Finance",
}

SERVICES_HC = {
    "headhunting": "Headhunting",
    "outsourcing": "Outsourcing Time & Material",
    "salary": "Consultoría Salarial",
    "formacion": "Formación en Recursos Humanos",
}

# Pasos del wizard. "Servicios" solo aplica a Human Capital solo (en HC+Finance
# los 4 servicios van siempre juntos en una diapositiva resumen, sin elegir).
# "Equipo y sedes" agrupa las decisiones estructurales de la plantilla
# (incluir a Eduardo Serrano, qué sedes mostrar, meses de garantía).
# "Revisión" es el paso en el que se enseña y puede editarse todo lo que va
# a aparecer en el PowerPoint antes de generar el archivo final.
STEPS = ["Config", "Cliente", "Servicios", "Equipo y sedes", "Contexto", "Revisión", "Generar"]

STEP_CONFIG = 0
STEP_CLIENTE = 1
STEP_SERVICIOS = 2
STEP_EQUIPO = 3
STEP_CONTEXTO = 4
STEP_REVIEW = 5
STEP_GENERATE = 6
