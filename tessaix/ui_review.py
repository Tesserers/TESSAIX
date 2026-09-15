"""Paso 5: revisión y edición del contenido ANTES de generar.

En Human Capital, esta pantalla es la salvaguarda: nada se escribe en el
.pptx sin que antes se haya visto en pantalla —y se haya podido corregir a
mano si hace falta— el texto exacto que va a llevar cada slide. En Human
Capital + Finance el contenido de las diapositivas de servicio es fijo (por
diseño, ver conversación con el equipo), así que aquí solo se repasan las
decisiones estructurales (equipo, sedes, fee) antes de generar.
"""
from __future__ import annotations

import streamlit as st

from .ai_content import generate_content
from .config import BUSINESS_LINE_HC_FINANCE, SERVICES_HC, STEP_CONTEXTO, STEP_GENERATE

_WHY_FIELDS = [
    ("d1_title", "Título diferenciador 1", "input"),
    ("d1_body", "Descripción 1", "area"),
    ("d2_title", "Título diferenciador 2", "input"),
    ("d2_body", "Descripción 2", "area"),
    ("d3_title", "Título diferenciador 3", "input"),
    ("d3_body", "Descripción 3", "area"),
]

_SERVICE_FIELDS = {
    "headhunting": [
        ("why_col_title1", "Por qué elegirnos — título 1", "input"),
        ("why_col_body1", "Por qué elegirnos — texto 1", "area"),
        ("why_col_title2", "Por qué elegirnos — título 2", "input"),
        ("why_col_body2", "Por qué elegirnos — texto 2", "area"),
        ("why_col_title3", "Por qué elegirnos — título 3", "input"),
        ("why_col_body3", "Por qué elegirnos — texto 3", "area"),
        ("benefits_body1", "Beneficio — cero estrés", "area"),
        ("benefits_body2", "Beneficio — talento a medida", "area"),
        ("benefits_body3", "Beneficio — relación a largo plazo", "area"),
        ("how_body3", "Cómo lo hacemos — elegimos con precisión", "area"),
        ("how_body4", "Cómo lo hacemos — solo lo mejor", "area"),
        ("how_body5", "Cómo lo hacemos — acompañamos el proceso", "area"),
    ],
    "outsourcing": [
        ("headline", "Titular", "input"),
        ("body", "Texto principal", "area"),
        ("card1_body", "Tarjeta — coste variable", "area"),
        ("card2_body", "Tarjeta — cobertura total", "area"),
        ("card3_body", "Tarjeta — cero gestión", "area"),
        ("card4_body", "Tarjeta — flexibilidad", "area"),
    ],
    "salary": [
        ("headline", "Titular", "input"),
        ("body", "Texto principal", "area"),
    ],
    "formacion": [
        ("body", "Texto principal", "area"),
    ],
}


def _text_field(container: dict, key: str, label: str, kind: str, widget_key: str):
    current = container.get(key, "") or ""
    if kind == "area":
        val = st.text_area(label, value=current, height=80, key=widget_key)
        st.caption(f"{len(val)} caracteres")
    else:
        val = st.text_input(label, value=current, key=widget_key)
    container[key] = val


def _render_content_editor(data: dict, content: dict):
    st.markdown("#### 📝 Contenido de la propuesta")
    st.caption("Generado a partir de los datos del cliente. Puedes editar cualquier texto antes de generar el PowerPoint.")

    why = content.setdefault("why_tessera", {})
    with st.expander("Por qué elegir Tessera", expanded=False):
        for key, label, kind in _WHY_FIELDS:
            _text_field(why, key, label, kind, f"why_{key}")

    services_content = content.setdefault("services", {})
    for svc in data.get("services", []):
        fields = _SERVICE_FIELDS.get(svc)
        if not fields:
            continue
        svc_content = services_content.setdefault(svc, {})
        with st.expander(f"{SERVICES_HC.get(svc, svc)}", expanded=False):
            for key, label, kind in fields:
                _text_field(svc_content, key, label, kind, f"{svc}_{key}")


def _render_structural_summary(data: dict):
    """Para Human Capital + Finance no hay contenido generado por IA que
    revisar — el contenido de las diapositivas de servicio es fijo. Aquí
    solo se repasan las decisiones que sí cambian el .pptx."""
    st.markdown("#### 📋 Resumen antes de generar")
    rows = [
        ("Equipo", "Con Eduardo Serrano" if data.get("include_eduardo") else "Sin Eduardo Serrano"),
        ("Sedes", "Las 3 sedes" if data.get("locations", "all") == "all" else "Solo Madrid"),
        ("Tipo", "Con fee y garantía" if data.get("deck_type") == "propuesta" else "Sin fee"),
    ]
    if data.get("deck_type") == "propuesta":
        rows.append(("Garantía", f"{data.get('warranty_months', 3)} meses"))
        if data.get("fee_rate"):
            rows.append(("Fee", f"{data['fee_rate']}%"))
    html_r = "".join(f'<div class="sr"><span class="sk">{k}</span><span class="sv">{v}</span></div>' for k, v in rows)
    st.markdown(f'<div class="sum">{html_r}</div>', unsafe_allow_html=True)
    st.caption(
        "El contenido de las diapositivas de servicio (Finance, Human Capital, Por qué Tessera) "
        "es fijo en las propuestas de Human Capital + Finance."
    )


def step_review():
    data = st.session_state.form
    is_hc_finance = data.get("business_line") == BUSINESS_LINE_HC_FINANCE

    st.markdown("### Revisión antes de generar")
    st.caption(
        "Aquí se enseña todo lo que va a aparecer en el PowerPoint. Revísalo y edítalo "
        "las veces que haga falta — el archivo no se genera hasta que le des a «Generar PPTX»."
    )

    if is_hc_finance:
        if st.session_state.content is None:
            st.session_state.content = {}
        _render_structural_summary(data)
        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        cb, cn_ = st.columns([1, 4])
        with cb:
            if st.button("← Atrás"):
                st.session_state.step = STEP_CONTEXTO
                st.rerun()
        with cn_:
            if st.button("Continuar → Generar PPTX", use_container_width=True):
                st.session_state.step = STEP_GENERATE
                st.rerun()
        return

    if st.session_state.content is None:
        st.markdown("Todavía no se ha generado contenido para esta propuesta.")
        if st.button("✦ Generar propuesta", use_container_width=True):
            c = generate_content(data)
            if c:
                st.session_state.content = c
                st.rerun()
        cb, _ = st.columns([1, 4])
        with cb:
            if st.button("← Atrás"):
                st.session_state.step = STEP_CONTEXTO
                st.rerun()
        return

    _render_content_editor(data, st.session_state.content)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    cb, cr, cn_ = st.columns([1, 1.4, 2.6])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step = STEP_CONTEXTO
            st.rerun()
    with cr:
        if st.button("↺ Regenerar propuesta"):
            c = generate_content(data)
            if c:
                st.session_state.content = c
                st.rerun()
    with cn_:
        if st.button("Continuar → Generar PPTX", use_container_width=True):
            st.session_state.step = STEP_GENERATE
            st.rerun()
