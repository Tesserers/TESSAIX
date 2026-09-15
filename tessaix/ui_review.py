"""Paso 5: revisión y edición del contenido (incluido el logo) ANTES de generar.

Esta pantalla es la salvaguarda pedida: nada se escribe en el .pptx sin que
antes se haya visto en pantalla —y se haya podido corregir a mano si hace
falta— el texto exacto que va a llevar cada slide y el logo del cliente tal
y como va a quedar compuesto en la portada.
"""
from __future__ import annotations

import streamlit as st

from .ai_content import generate_content
from .config import SERVICES_HC, STEP_CONTEXTO, STEP_GENERATE
from .logo_client import (
    checkerboard_preview,
    compose_logo_for_placeholder,
    fetch_client_logo,
    logo_quality_label,
    normalize_domain,
)
from .pptx_builder import get_client_logo_frame_emu

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
        ("how_body1", "Cómo lo hacemos — nos conocemos", "area"),
        ("how_body2", "Cómo lo hacemos — vamos más allá", "area"),
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
    "rpo": [
        ("headline", "Titular", "input"),
        ("body", "Texto principal", "area"),
        ("p1_title", "Pilar 1 — título", "input"),
        ("p1_body", "Pilar 1 — texto", "area"),
        ("p2_title", "Pilar 2 — título", "input"),
        ("p2_body", "Pilar 2 — texto", "area"),
        ("p3_title", "Pilar 3 — título", "input"),
        ("p3_body", "Pilar 3 — texto", "area"),
    ],
    "salary": [
        ("headline", "Titular", "input"),
        ("body", "Texto principal", "area"),
    ],
}


def _init_logo_state(data: dict):
    if st.session_state.get("logo_meta") is not None or st.session_state.get("logo_checked"):
        return
    st.session_state.logo_checked = True
    domain = normalize_domain(data.get("client_website", ""), data.get("client_name", ""))
    st.session_state.logo_domain = domain
    meta = fetch_client_logo(domain)
    st.session_state.logo_meta = meta
    st.session_state.logo_bytes = meta["bytes"] if meta else None
    st.session_state.logo_source = "auto" if meta else "none"


def _render_logo_section(data: dict):
    st.markdown("#### 🖼️ Logo del cliente")
    st.caption(
        "El logo se compone SIEMPRE respetando su forma original — nunca se estira ni se "
        "deforma para rellenar el hueco de la portada. Revísalo aquí antes de generar: "
        "si la calidad no te convence, sube uno mejor manualmente."
    )

    domain = st.session_state.get("logo_domain", "")
    dcol, bcol = st.columns([3, 1])
    with dcol:
        new_domain_input = st.text_input(
            "Dominio usado para buscar el logo", value=domain,
            help="Se detecta automáticamente a partir de la web del cliente."
        )
    with bcol:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("🔄 Buscar logo", use_container_width=True):
            meta = fetch_client_logo(normalize_domain(new_domain_input))
            st.session_state.logo_domain = normalize_domain(new_domain_input)
            st.session_state.logo_meta = meta
            st.session_state.logo_bytes = meta["bytes"] if meta else None
            st.session_state.logo_source = "auto" if meta else "none"
            st.rerun()

    meta = st.session_state.get("logo_meta")
    level, message = logo_quality_label(meta)
    css_class = {"ok": "ok", "low": "warn", "missing": "warn"}[level]
    icon = {"ok": "✅", "low": "⚠️", "missing": "⚠️"}[level]
    st.markdown(f'<div class="{css_class}">{icon} {message}</div>', unsafe_allow_html=True)

    upload = st.file_uploader(
        "Subir logo manualmente (PNG o JPG, recomendado si la calidad automática es baja)",
        type=["png", "jpg", "jpeg"], key="logo_upload",
    )
    if upload is not None:
        st.session_state.logo_bytes = upload.getvalue()
        st.session_state.logo_source = "manual"

    no_logo = st.checkbox(
        "No incluir logo del cliente en la portada",
        value=st.session_state.get("logo_source") == "none" and st.session_state.get("logo_bytes") is None,
    )
    if no_logo:
        st.session_state.logo_bytes = None
        st.session_state.logo_source = "none"

    # ── Vista previa EXACTA (misma composición que se usará al generar) ──
    frame = get_client_logo_frame_emu()
    if st.session_state.get("logo_bytes") and frame:
        try:
            composed = compose_logo_for_placeholder(st.session_state.logo_bytes, *frame)
            pcol, _ = st.columns([1, 2])
            with pcol:
                with st.container(border=True):
                    st.image(
                        checkerboard_preview(composed), use_container_width=True,
                        caption="Así se verá en la portada (el damero es solo para mostrar la transparencia; el fondo real es el de la diapositiva)",
                    )
        except Exception as e:
            st.markdown(f'<div class="err">No se pudo procesar la imagen del logo: {e}</div>', unsafe_allow_html=True)
    elif not no_logo:
        st.info("Sin logo todavía: sube uno o pulsa «Buscar logo».")


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
    with st.expander("Por qué elegir Tessera (diapositiva 3)", expanded=False):
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


def step_review():
    data = st.session_state.form
    st.markdown("### Revisión antes de generar")
    st.caption(
        "Aquí se enseña todo lo que va a aparecer en el PowerPoint. Revísalo y edítalo "
        "las veces que haga falta — el archivo no se genera hasta que le des a «Generar PPTX»."
    )

    if st.session_state.content is None:
        st.markdown("Todavía no se ha generado contenido para esta propuesta.")
        if st.button("✦ Generar propuesta", use_container_width=True):
            c = generate_content(data)
            if c:
                st.session_state.content = c
                st.session_state.logo_checked = False
                st.rerun()
        cb, _ = st.columns([1, 4])
        with cb:
            if st.button("← Atrás"):
                st.session_state.step = STEP_CONTEXTO
                st.rerun()
        return

    _init_logo_state(data)
    _render_logo_section(data)
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
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
