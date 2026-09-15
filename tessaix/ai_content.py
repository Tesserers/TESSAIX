"""Generación del contenido de la propuesta con Claude."""
from __future__ import annotations

import json
import re

import anthropic
import streamlit as st

from .config import SERVICES_HC

# Límite duro de caracteres por tipo de campo. Es una red de seguridad
# adicional al autofit dinámico del .pptx (ver xml_utils.enable_shrink_autofit):
# el autofit encoge la fuente si el texto no cabe, pero si el modelo se
# desmadra y devuelve un párrafo larguísimo, es mejor recortarlo con criterio
# (por palabra completa) que dejar que la fuente se reduzca hasta ser
# ilegible. Se aplica por prefijo de la clave del campo.
_LIMITS_BY_SUFFIX = [
    ("_title", 60),
    ("headline", 90),
    ("_body", 220),
    ("body", 260),
]


def _limit_for_key(key: str) -> int | None:
    for suffix, limit in _LIMITS_BY_SUFFIX:
        if key.endswith(suffix):
            return limit
    return None


def _truncate_wordwise(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(",.;:")
    return cut + "…"


def enforce_length_limits(content: dict) -> dict:
    """Recorta (por palabra completa) cualquier campo de texto anormalmente
    largo, para que el autofit del .pptx nunca tenga que encoger la fuente
    de forma extrema."""

    def _walk(node):
        if isinstance(node, dict):
            return {k: _walk(v) if not isinstance(v, str) else _clip(k, v) for k, v in node.items()}
        return node

    def _clip(key, value):
        limit = _limit_for_key(key)
        return _truncate_wordwise(value, limit) if limit else value

    return _walk(content)


SYSTEM_ES = """Eres consultor comercial senior de Tessera Human Capital (tesseraservices.com).
Generas contenido para propuestas comerciales en PowerPoint. Devuelves ÚNICAMENTE JSON válido.
REGLAS ABSOLUTAS:
- Sin guiones largos
- Nunca uses "contexto" ni "criterio"
- Español de España natural, directo, tono de negocio cercano, tono neutral de España
- Comprueba que no existan faltas de ortografía antes de responder
- Utiliza palabras corrientes y de uso común, nunca palabras rebuscadas o poco usuales
  (por ejemplo, nunca digas "embebido/a"; di "integrado" o "dedicado")
- ESPECÍFICO al sector del cliente — nada de mencionar "publicidad" si el cliente es de otro sector
- Los textos tienen que REEMPLAZAR los existentes completamente, sin mezclar con el original
- Cada texto debe ser completo y autónomo, listo para aparecer en el slide tal cual
- SÉ CONCISO: los títulos y titulares aparecen en cajas de tamaño fijo en el slide.
  Títulos: máximo 4-6 palabras. Cuerpos de texto: máximo 2-3 líneas (unos 150-200 caracteres)."""

SYSTEM_EN = """Senior commercial consultant at Tessera Human Capital. Return ONLY valid JSON.
ABSOLUTE RULES: no em dashes, sector-specific content only, complete standalone texts.
Neutral, professional business tone. Check there are no spelling or grammar mistakes
before answering. Use plain, everyday words — never obscure or overly formal vocabulary
(for example, never say "utilize"; say "use" instead).
BE CONCISE: titles and headlines sit in fixed-size boxes on the slide.
Titles: max 4-6 words. Body copy: max 2-3 lines (about 150-200 characters)."""


def generate_content(data: dict) -> dict | None:
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    is_en = data["lang"] == "en"
    svcs = ", ".join([SERVICES_HC.get(s, s) for s in data["services"]])
    cn, sec = data["client_name"], data["sector"]

    system = SYSTEM_ES if not is_en else SYSTEM_EN

    prompt = f"""Genera JSON para propuesta comercial de Tessera Human Capital para:
CLIENTE: {cn}
SECTOR: {sec}
PAÍS: {data.get('country', 'España')}
SERVICIOS: {svcs}
PAIN POINTS: {data.get('pain_points', '')}
PERFILES NECESARIOS: {data.get('roles_needed', '')}
FEE: {data.get('fee_rate', 'XX')}%
INFO ADICIONAL: {data.get('extra_info', '')}

IMPORTANTE: Todos los textos deben ser específicos para el sector "{sec}".
NO mencionar publicidad, AdTech, programmatic ni nada relacionado con marketing digital
a no ser que el cliente sea de ese sector.

Devuelve este JSON con contenido REAL y COMPLETO (no placeholders):
{{
  "why_tessera": {{
    "d1_title": "título diferenciador 1 específico para {sec} (máx 4 palabras)",
    "d1_body": "descripción 2-3 líneas específica para {sec}. Sin mencionar otros sectores.",
    "d2_title": "título diferenciador 2 específico para {sec} (máx 4 palabras)",
    "d2_body": "descripción 2-3 líneas específica para {sec}",
    "d3_title": "título diferenciador 3 específico para {sec} (máx 4 palabras)",
    "d3_body": "descripción 2-3 líneas específica para {sec}"
  }},
  "services": {{
    {"\"headhunting\": {\"why_col_title1\": \"TÍTULO RAZÓN 1\", \"why_col_body1\": \"descripción específica {sec}\", \"why_col_title2\": \"TÍTULO RAZÓN 2\", \"why_col_body2\": \"descripción específica {sec}\", \"why_col_title3\": \"TÍTULO RAZÓN 3\", \"why_col_body3\": \"descripción específica {sec}\", \"benefits_title1\": \"CERO ESTRÉS\", \"benefits_body1\": \"descripción específica {sec}\", \"benefits_title2\": \"TALENTO A MEDIDA\", \"benefits_body2\": \"descripción específica {sec}\", \"benefits_title3\": \"RELACIÓN A LARGO PLAZO\", \"benefits_body3\": \"descripción específica {sec}\", \"how_title1\": \"NOS CONOCEMOS\", \"how_body1\": \"descripción específica {sec}\", \"how_title2\": \"VAMOS MÁS ALLÁ\", \"how_body2\": \"descripción específica {sec}\", \"how_title3\": \"ELEGIMOS CON PRECISIÓN\", \"how_body3\": \"descripción específica {sec}\", \"how_title4\": \"SOLO LO MEJOR\", \"how_body4\": \"descripción específica {sec}\", \"how_title5\": \"ACOMPAÑAMOS EL PROCESO\", \"how_body5\": \"descripción específica {sec}\"}" if "headhunting" in data.get("services", []) else ""},
    {"\"outsourcing\": {\"headline\": \"frase impacto para {sec}\", \"body\": \"2 líneas para {sec}\", \"card1_title\": \"Coste variable\", \"card1_body\": \"Pagas solo por horas reales de trabajo\", \"card2_title\": \"Cobertura total\", \"card2_body\": \"Cubrimos bajas y vacaciones sin interrupciones\", \"card3_title\": \"Cero gestión\", \"card3_body\": \"Nóminas y trámites a cargo nuestro\", \"card4_title\": \"Flexibilidad\", \"card4_body\": \"Escalas el equipo cuando lo necesites\"}" if "outsourcing" in data.get("services", []) else ""},
    {"\"salary\": {\"headline\": \"Inteligencia retributiva para {sec}\", \"body\": \"2 líneas para {cn}\"}" if "salary" in data.get("services", []) else ""},
    {"\"formacion\": {\"body\": \"2 líneas sobre formación en RRHH específicas para {sec}\"}" if "formacion" in data.get("services", []) else ""}
  }}
}}"""

    with st.spinner("Generando contenido adaptado al cliente..."):
        msg = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=4000,
            system=system, messages=[{"role": "user", "content": prompt}]
        )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    # Limpia claves vacías que a veces cuela el modelo
    raw = re.sub(r',\s*""\s*:', ',"_empty":', raw)
    raw = re.sub(r':\s*,', ': null,', raw)
    try:
        content = json.loads(raw)
    except Exception as e:
        st.error(f"Error parseando JSON: {e}")
        st.code(raw[:800])
        return None

    return enforce_length_limits(content)
