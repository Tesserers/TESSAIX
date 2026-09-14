"""Pasos 0-4 del wizard: configuración, cliente, servicios, presentador, contexto."""
import streamlit as st

from .config import SERVICES_HC, STEP_CLIENTE, STEP_CONTEXTO, STEP_PRESENTADOR, STEP_REVIEW, STEP_SERVICIOS


def step_config():
    st.markdown("### ¿Qué tipo de propuesta?")
    c1, c2 = st.columns(2)
    with c1:
        lang = st.radio("Idioma", ["Español", "English"])
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
            st.session_state.form.update({
                "lang": "es" if lang == "Español" else "en",
                "deck_type": "presentacion" if "sin fees" in deck_type else "propuesta"
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
                st.session_state.step = STEP_SERVICIOS
                st.rerun()


def step_servicios():
    st.markdown("### ¿Qué servicios incluimos?")
    selected = st.multiselect("Servicios", list(SERVICES_HC.keys()),
                               default=["headhunting"], format_func=lambda x: SERVICES_HC[x])

    st.markdown("---")
    include_finance = st.checkbox(
        "Incluir oferta cruzada de Finance/M&A para este cliente",
        value=False,
        help=(
            "Añade el bloque de diapositivas de Soporte financiero, Fiscalidad, "
            "Auditoría y Tessera Services. Actívalo solo si tiene sentido para "
            "este cliente en concreto — si no, se genera la propuesta sin ese "
            "bloque y con la numeración de páginas ajustada."
        ),
    )

    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step -= 1
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            if not selected:
                st.error("Selecciona al menos uno.")
            else:
                st.session_state.form["services"] = selected
                st.session_state.form["include_finance_crosssell"] = include_finance
                st.session_state.step = STEP_PRESENTADOR
                st.rerun()


def step_presentador():
    st.markdown("### ¿Quién presenta esta propuesta?")
    st.caption("Esta información aparecerá en la portada y en el slide de cierre.")
    c1, c2 = st.columns(2)
    with c1:
        pname = st.text_input("Nombre completo *", value="Manuel García Pina",
                               placeholder="Manuel García Pina")
        pemail = st.text_input("Email *", value="Manuel.garcia@tesseraservices.com",
                                placeholder="email@tesseraservices.com")
    with c2:
        pphone = st.text_input("Teléfono *", value="+34 619 511 155",
                                placeholder="+34 6XX XXX XXX")
        _ = st.text_input("Cargo", value="Partner",
                           placeholder="Partner / Head of Recruitment…")
    cb, cn_ = st.columns([1, 4])
    with cb:
        if st.button("← Atrás"):
            st.session_state.step -= 1
            st.rerun()
    with cn_:
        if st.button("Siguiente →"):
            if not pname.strip() or not pphone.strip():
                st.error("Nombre y teléfono obligatorios.")
            else:
                st.session_state.form.update({
                    "presenter_name": pname,
                    "presenter_phone": pphone,
                    "presenter_email": pemail
                })
                st.session_state.step = STEP_CONTEXTO
                st.rerun()


def step_contexto():
    st.markdown("### Personalización")
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
