"""Paso 5: revisión y edición del contenido ANTES de generar.

Esta pantalla es la salvaguarda: nada se escribe en el .pptx sin que antes
se haya visto en pantalla —y se haya podido corregir a mano si hace
falta— el texto exacto que va a llevar cada slide. Human Capital + Finance
genera el mismo tipo de contenido por IA que Human Capital (por qué
Tessera + la frase de apertura de cada servicio) — las tarjetas de
servicio con cifras concretas (93% de éxito, 30% de ahorro…) son
afirmaciones fijas de la empresa y no se tocan.
"""
from __future__ import annotations

import io

import streamlit as st

from .ai_content import generate_content, generate_content_hc_finance
from .config import BUSINESS_LINE_HC_FINANCE, SERVICES_HC, STEP_CONTEXTO, STEP_GENERATE
from .logo_client import MIN_LOGO_DIM, checkerboard_preview, fetch_client_logo, logo_quality_label, normalize_domain
from .pptx_builder import (
    HC_TEAM_PEOPLE_LABELS, HC_TEAM_PEOPLE_ORDER,
    HCF_TEAM_PEOPLE_LABELS, HCF_TEAM_PEOPLE_ORDER,
)

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


_HCF_INTRO_SECTIONS = {
    "para_el_fondo": "Para el fondo",
    "finance": "Finance",
    "human_capital": "Human Capital",
}


def _clear_content_widget_keys():
    """
    Los widgets de `_render_content_editor`/`_render_content_editor_hcf`
    usan `key=` fijas (`why_d1_title`, `headhunting_headline`, …) para que
    Streamlit recuerde lo que el usuario ha tecleado entre reruns. Pero eso
    significa que, al (re)generar contenido nuevo, Streamlit sigue
    mostrando el valor ANTIGUO que ya tenía guardado bajo esa key en vez
    del `value=` recién generado — el cuadro de texto parece "no
    actualizarse" (o aparecer vacío, si es la primera vez que se pulsa
    "Generar" tras haber dejado algún campo en blanco antes). Se borran
    esas keys justo antes de guardar contenido nuevo para que los widgets
    se construyan limpios en el próximo rerun.
    """
    for key, _, _ in _WHY_FIELDS:
        st.session_state.pop(f"why_{key}", None)
        st.session_state.pop(f"hcf_why_{key}", None)
    for svc, svc_fields in _SERVICE_FIELDS.items():
        for key, _, _ in svc_fields:
            st.session_state.pop(f"{svc}_{key}", None)
    for section in _HCF_INTRO_SECTIONS:
        st.session_state.pop(f"hcf_{section}_body", None)


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


def _render_content_editor_hcf(content: dict):
    st.markdown("#### 📝 Contenido de la propuesta")
    st.caption("Generado a partir de los datos del cliente. Puedes editar cualquier texto antes de generar el PowerPoint.")

    why = content.setdefault("why_tessera", {})
    with st.expander("Por qué elegir Tessera", expanded=False):
        for key, label, kind in _WHY_FIELDS:
            _text_field(why, key, label, kind, f"hcf_why_{key}")

    for section, label in _HCF_INTRO_SECTIONS.items():
        sec_content = content.setdefault(section, {})
        with st.expander(label, expanded=False):
            _text_field(sec_content, "body", "Texto", "area", f"hcf_{section}_body")


def _render_structural_summary(data: dict):
    """Las decisiones que no son contenido de texto pero también cambian
    el .pptx (equipo, sedes) — se muestran además del contenido generado
    por IA, no en su lugar."""
    is_hc_finance = data.get("business_line") == BUSINESS_LINE_HC_FINANCE
    people_order = HCF_TEAM_PEOPLE_ORDER if is_hc_finance else HC_TEAM_PEOPLE_ORDER
    people_labels = HCF_TEAM_PEOPLE_LABELS if is_hc_finance else HC_TEAM_PEOPLE_LABELS
    team_people = data.get("team_people", people_order)
    equipo = ", ".join(people_labels[p] for p in people_order if p in team_people) or "Sin diapositiva de equipo"

    st.markdown("#### 📋 Resumen antes de generar")
    rows = [
        ("Equipo", equipo),
        ("Sedes", "Las 3 sedes" if data.get("locations", "all") == "all" else "Solo Madrid"),
    ]
    html_r = "".join(f'<div class="sr"><span class="sk">{k}</span><span class="sv">{v}</span></div>' for k, v in rows)
    st.markdown(f'<div class="sum">{html_r}</div>', unsafe_allow_html=True)


def _render_logo_section(data: dict):
    """
    Logo del cliente para la portada y el cierre — se busca automáticamente
    una vez por propuesta a partir de la web del cliente (campo ya pedido
    en el paso «Cliente»), con opción de subir uno a mano si no se
    encuentra o si la calidad no convence. Se compone sin deformar y sin
    fondo de color al generar el .pptx (ver tessaix/logo_client.py) —
    aquí solo se guardan los bytes elegidos en session_state.
    """
    st.markdown("#### 🖼️ Logo del cliente")
    st.caption("Aparece en la portada y en el cierre de la propuesta.")

    if not st.session_state.logo_autofetch_tried and st.session_state.logo_bytes is None:
        domain = normalize_domain(data.get("client_website", ""), data.get("client_name", ""))
        with st.spinner("Buscando el logo del cliente…"):
            meta = fetch_client_logo(domain)
        st.session_state.logo_autofetch_tried = True
        st.session_state["_logo_meta"] = meta
        if meta:
            st.session_state.logo_bytes = meta["bytes"]

    if st.session_state.logo_bytes:
        level, msg = logo_quality_label(st.session_state.get("_logo_meta"))
        c1, c2 = st.columns([1, 2])
        with c1:
            st.image(checkerboard_preview(st.session_state.logo_bytes), width=200)
        with c2:
            icon = {"ok": "✅", "low": "⚠️"}.get(level, "ℹ️")
            st.caption(f"{icon} {msg}")
    else:
        st.caption("No se ha encontrado ningún logo automáticamente. Puedes subirlo a mano.")

    upload = st.file_uploader("Subir logo manualmente (PNG o JPG)", type=["png", "jpg", "jpeg"], key="logo_upload")
    if upload is not None:
        data_bytes = upload.getvalue()
        if data_bytes != st.session_state.get("_logo_uploaded_bytes"):
            from PIL import Image
            w, h = Image.open(io.BytesIO(data_bytes)).size
            st.session_state.logo_bytes = data_bytes
            st.session_state["_logo_uploaded_bytes"] = data_bytes
            st.session_state["_logo_meta"] = {
                "ok": min(w, h) >= MIN_LOGO_DIM, "source": "subido a mano", "width": w, "height": h,
            }
            st.rerun()
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


def step_review():
    data = st.session_state.form
    is_hc_finance = data.get("business_line") == BUSINESS_LINE_HC_FINANCE
    generate = generate_content_hc_finance if is_hc_finance else generate_content

    st.markdown("### Revisión antes de generar")
    st.caption(
        "Aquí se enseña todo lo que va a aparecer en el PowerPoint. Revísalo y edítalo "
        "las veces que haga falta — el archivo no se genera hasta que le des a «Generar PPTX»."
    )

    _render_logo_section(data)

    _render_structural_summary(data)
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    if st.session_state.content is None:
        st.markdown("Todavía no se ha generado contenido para esta propuesta.")
        if st.button("✦ Generar propuesta", use_container_width=True):
            c = generate(data)
            if c:
                _clear_content_widget_keys()
                st.session_state.content = c
                st.rerun()
        cb, _ = st.columns([1, 4])
        with cb:
            if st.button("← Atrás"):
                st.session_state.step = STEP_CONTEXTO
                st.rerun()
        return

    if is_hc_finance:
        _render_content_editor_hcf(st.session_state.content)
    else:
        _render_content_editor(data, st.session_state.content)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    cb, cr, cn_ = st.columns([1, 1.4, 2.6])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step = STEP_CONTEXTO
            st.rerun()
    with cr:
        if st.button("↺ Regenerar propuesta"):
            c = generate(data)
            if c:
                _clear_content_widget_keys()
                st.session_state.content = c
                st.rerun()
    with cn_:
        if st.button("Continuar → Generar PPTX", use_container_width=True):
            st.session_state.step = STEP_GENERATE
            st.rerun()
