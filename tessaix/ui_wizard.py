"""Pasos del wizard: configuración, cliente, servicios, equipo/sedes, contexto."""
import streamlit as st

from .config import (
    BUSINESS_LINE_HC, BUSINESS_LINE_HC_FINANCE, BUSINESS_LINES, SERVICES_HC,
    STEP_CLIENTE, STEP_CONTEXTO, STEP_EQUIPO, STEP_REVIEW, STEP_SERVICIOS,
)
from .team import WARRANTY_MONTHS_OPTIONS


def step_config():
    st.markdown("### ¿Qué tipo de propuesta?")
    c1, c2 = st.columns(2)
    with c1:
        lang = st.radio("Idioma", ["Español", "English"])
        business_line = st.radio(
            "Línea de negocio", list(BUSINESS_LINES.values()),
            help="Determina qué plantilla se usa: Human Capital sola, o Human Capital + Finance.",
        )
    with c2:
        deck_type = st.radio("Tipo", ["Presentación (sin fees)", "Propuesta completa (con fees)"])
    st.markdown("---")
    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Inicio"):
            st.session_state.screen = "home"
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            business_line_key = BUSINESS_LINE_HC_FINANCE if "Finance" in business_line else BUSINESS_LINE_HC
            st.session_state.form.update({
                "lang": "es" if lang == "Español" else "en",
                "deck_type": "presentacion" if "sin fees" in deck_type else "propuesta",
                "business_line": business_line_key,
            })
            st.session_state.step = STEP_CLIENTE
            st.rerun()


def step_cliente():
    st.markdown("### Datos del cliente")
    c1, c2 = st.columns(2)
    with c1:
        cn = st.text_input("Nombre de la empresa *", placeholder="Ej: Alimerka, Corpay…")
        sec = st.text_input("Sector / industria *", placeholder="Ej: Alimentación, Fintech…")
        ctr = st.text_input("País / mercado", value="España")
    with c2:
        sz = st.text_input("Tamaño de empresa", placeholder="Ej: 50–200 empleados")
        web = st.text_input("Web del cliente", placeholder="alimerka.es")
        ccn = st.text_input("Nombre del contacto", placeholder="Nombre y apellido")
        ccr = st.text_input("Cargo del contacto", placeholder="Ej: HR Director")
    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step -= 1
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            if not cn.strip() or not sec.strip():
                st.error("Empresa y sector obligatorios.")
            else:
                st.session_state.form.update({
                    "client_name": cn, "sector": sec, "country": ctr,
                    "size": sz, "client_website": web,
                    "contact_name": ccn, "contact_role": ccr
                })
                # HC+Finance no tiene paso de servicios: los 4 van siempre
                # juntos en la diapositiva resumen de la plantilla.
                is_hc_finance = st.session_state.form.get("business_line") == BUSINESS_LINE_HC_FINANCE
                st.session_state.step = STEP_EQUIPO if is_hc_finance else STEP_SERVICIOS
                st.rerun()


def step_servicios():
    st.markdown("### ¿Qué servicios incluimos?")
    selected = st.multiselect("Servicios", list(SERVICES_HC.keys()),
                               default=["headhunting"], format_func=lambda x: SERVICES_HC[x])

    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step = STEP_CLIENTE
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            if not selected:
                st.error("Selecciona al menos uno.")
            else:
                st.session_state.form["services"] = selected
                st.session_state.step = STEP_EQUIPO
                st.rerun()


def step_equipo():
    data = st.session_state.form
    is_hc_finance = data.get("business_line") == BUSINESS_LINE_HC_FINANCE

    st.markdown("### Equipo y sedes")

    include_eduardo = data.get("include_eduardo", False)
    if is_hc_finance:
        st.caption(
            "El equipo de esta propuesta siempre incluye a Manuel Pina, "
            "Mónica Mayoral y Edward Manrique."
        )
        include_eduardo = st.checkbox("Incluir también a Eduardo Serrano", value=include_eduardo)
    else:
        st.caption("El equipo de esta propuesta es fijo: Eduardo Serrano, Manuel Pina y Edward Manrique.")

    st.markdown("---")
    st.markdown("#### Sedes a mostrar")
    locations_label = st.radio(
        "Sedes", ["Las 3 sedes (Madrid · Bilbao · Oviedo)", "Solo Madrid"],
        index=0 if data.get("locations", "all") == "all" else 1,
        label_visibility="collapsed",
    )
    locations = "madrid" if "Solo Madrid" in locations_label else "all"

    warranty_months = data.get("warranty_months", WARRANTY_MONTHS_OPTIONS[0])
    if data.get("deck_type") == "propuesta":
        st.markdown("---")
        st.markdown("#### Garantía (términos y condiciones)")
        warranty_months = st.radio(
            "Meses de garantía", WARRANTY_MONTHS_OPTIONS,
            index=WARRANTY_MONTHS_OPTIONS.index(warranty_months) if warranty_months in WARRANTY_MONTHS_OPTIONS else 0,
            format_func=lambda m: f"{m} meses", horizontal=True,
        )

    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step = STEP_SERVICIOS if not is_hc_finance else STEP_CLIENTE
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            st.session_state.form.update({
                "include_eduardo": include_eduardo,
                "locations": locations,
                "warranty_months": warranty_months,
            })
            st.session_state.step = STEP_CONTEXTO
            st.rerun()


def step_contexto():
    is_hc_finance = st.session_state.form.get("business_line") == BUSINESS_LINE_HC_FINANCE

    st.markdown("### Personalización")
    if is_hc_finance:
        st.caption(
            "En las propuestas de Human Capital + Finance el contenido de las diapositivas de "
            "servicio es fijo — esto solo se usa como referencia interna."
        )
    pp = st.text_area("Pain points / retos del cliente", height=80,
                       placeholder="Ej: Crecimiento rápido, rotación alta en tienda, necesitan perfiles con vocación…")
    rn = st.text_area("Perfiles que necesitan cubrir", height=60,
                       placeholder="Ej: Jefes de sección, reponedores, 1 HR Business Partner…")
    fee = ""
    if st.session_state.form.get("deck_type") == "propuesta":
        fee = st.text_input("Fee propuesto (%)", placeholder="Ej: 16")
    ei = st.text_area("Info adicional (web, reunión previa, LinkedIn…)", height=90,
                       placeholder="Pega aquí todo lo que sepas del cliente…")
    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step -= 1
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            st.session_state.form.update({
                "pain_points": pp, "roles_needed": rn,
                "fee_rate": fee, "extra_info": ei
            })
            st.session_state.step = STEP_REVIEW
            st.rerun()
