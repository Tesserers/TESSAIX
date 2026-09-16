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


def _restore_client_name(text: str, client_name: str) -> str:
    """
    El nombre de la empresa cliente NUNCA debe cambiar. A veces el modelo
    trunca o "muerde" el final de un nombre propio dentro de una frase más
    larga (p.ej. "Alimerka" -> "Alimerk") — esto corrige cualquier variante
    del nombre a la que le falten hasta 3 caracteres al final, devolviéndolo
    exactamente como lo escribió el usuario.
    """
    if not text or not client_name or len(client_name) < 3:
        return text
    for cut in (1, 2, 3):
        if cut >= len(client_name):
            break
        variant = client_name[:-cut]
        if len(variant) < 2 or variant == client_name:
            continue
        text = re.sub(r'\b' + re.escape(variant) + r'\b', client_name, text)
    return text


def _restore_client_name_everywhere(node, client_name: str):
    if isinstance(node, dict):
        return {k: _restore_client_name_everywhere(v, client_name) for k, v in node.items()}
    if isinstance(node, str):
        return _restore_client_name(node, client_name)
    return node


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


# Regla común a todos los prompts: el nombre del cliente es intocable. Se
# repite tal cual (con el nombre real interpolado) en cada system prompt
# para que quede lo más arriba y explícito posible.
def _name_rule_es(cn: str) -> str:
    return (
        f'- El nombre de la empresa cliente es EXACTAMENTE "{cn}". Escríbelo siempre así, '
        f'letra por letra, en cualquier sitio donde aparezca — nunca lo abrevies, acortes, '
        f'completes ni le cambies una sola letra.'
    )


def _name_rule_en(cn: str) -> str:
    return (
        f'- The client company name is EXACTLY "{cn}". Always write it exactly like that, '
        f'letter for letter, wherever it appears — never abbreviate, shorten, complete, or '
        f'change a single letter of it.'
    )


SYSTEM_ES = """Eres consultor comercial senior de Tessera Human Capital (tesseraservices.com).
Generas contenido para propuestas comerciales en PowerPoint. Devuelves ÚNICAMENTE JSON válido.
REGLAS ABSOLUTAS:
{name_rule}
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
ABSOLUTE RULES:
{name_rule}
- No em dashes, sector-specific content only, complete standalone texts.
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

    system = (SYSTEM_EN if is_en else SYSTEM_ES).format(
        name_rule=_name_rule_en(cn) if is_en else _name_rule_es(cn)
    )

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
a no ser que el cliente sea de ese sector. El nombre de cliente "{cn}" debe escribirse
siempre exactamente igual, sin cambiar ni una letra.

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

    content = enforce_length_limits(content)
    return _restore_client_name_everywhere(content, cn)


def generate_context_paragraph(data: dict) -> str | None:
    """
    Párrafo de la diapositiva de "Contexto": una introducción breve a la
    situación del cliente, en el estilo de una propuesta real (inspirado en
    el formato de propuestas con una diapositiva de contexto dedicada).
    Se usa igual en Human Capital y en Human Capital + Finance — es el
    único contenido generado por IA que llevan las dos líneas de negocio.
    """
    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
    is_en = data.get("lang") == "en"
    cn, sec = data["client_name"], data.get("sector", "")

    if is_en:
        system = (
            "You write the context paragraph of a commercial proposal for Tessera. "
            "Return ONLY the paragraph's plain text — no quotes, no JSON, no markdown, "
            "no title.\n" + _name_rule_en(cn) + "\n"
            "Neutral, professional tone, no spelling mistakes, plain everyday words.\n"
            "3-5 sentences, one single paragraph, no bullet points."
        )
        prompt = (
            f"Client: {cn}\nSector: {sec}\nCountry: {data.get('country', 'Spain')}\n"
            f"Challenges: {data.get('pain_points', '')}\n"
            f"Additional info: {data.get('extra_info', '')}\n\n"
            f"Write the context paragraph: briefly introduce {cn}, its sector and "
            f"situation, and why it makes sense to talk about this proposal now."
        )
    else:
        system = (
            "Escribes el párrafo de contexto de una propuesta comercial de Tessera. "
            "Devuelve ÚNICAMENTE el texto del párrafo — sin comillas, sin JSON, sin "
            "markdown, sin título.\n" + _name_rule_es(cn) + "\n"
            "Español de España natural, tono neutral y profesional, sin faltas de "
            "ortografía, palabras corrientes (nunca 'embebido' ni vocabulario rebuscado).\n"
            "3-5 frases, un único párrafo, sin viñetas."
        )
        prompt = (
            f"Cliente: {cn}\nSector: {sec}\nPaís: {data.get('country', 'España')}\n"
            f"Retos: {data.get('pain_points', '')}\n"
            f"Info adicional: {data.get('extra_info', '')}\n\n"
            f"Escribe el párrafo de contexto: presenta brevemente a {cn}, su sector y "
            f"su situación, y por qué tiene sentido hablar de esta propuesta ahora."
        )

    with st.spinner("Generando contexto del cliente..."):
        msg = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=600,
            system=system, messages=[{"role": "user", "content": prompt}]
        )
    text = msg.content[0].text.strip().strip('"')
    text = _truncate_wordwise(text, 600)
    return _restore_client_name(text, cn)
