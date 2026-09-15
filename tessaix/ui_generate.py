"""Paso 6: generación final del .pptx y descarga."""
import streamlit as st

from .config import BUSINESS_LINE_HC_FINANCE, BUSINESS_LINES, SERVICES_HC, STEP_REVIEW
from .pptx_builder import build_pptx


def step_generate():
    d = st.session_state.form
    is_hc_finance = d.get("business_line") == BUSINESS_LINE_HC_FINANCE

    st.markdown("### Genera tu PowerPoint")
    rows = [
        ("Cliente", d.get("client_name", "—")),
        ("Sector", d.get("sector", "—")),
        ("Idioma", "Español" if d.get("lang") == "es" else "English"),
        ("Línea de negocio", BUSINESS_LINES.get(d.get("business_line"), "—")),
        ("Tipo", "Con fees" if d.get("deck_type") == "propuesta" else "Sin fees"),
        ("Sedes", "Las 3 sedes" if d.get("locations", "all") == "all" else "Solo Madrid"),
    ]
    if is_hc_finance:
        rows.append(("Equipo", "Con Eduardo Serrano" if d.get("include_eduardo") else "Sin Eduardo Serrano"))
    else:
        rows.append(("Servicios", " · ".join([SERVICES_HC.get(s, s) for s in d.get("services", [])])))
    if d.get("deck_type") == "propuesta":
        rows.append(("Garantía", f"{d.get('warranty_months', 3)} meses"))
    if d.get("fee_rate"):
        rows.append(("Fee", f"{d['fee_rate']}%"))
    html_r = "".join(f'<div class="sr"><span class="sk">{k}</span><span class="sv">{v}</span></div>'
                      for k, v in rows)
    st.markdown(f'<div class="sum">{html_r}</div>', unsafe_allow_html=True)

    if st.session_state.pptx is None:
        cb, cg = st.columns([1, 4])
        with cb:
            if st.button("← Revisar contenido"):
                st.session_state.step = STEP_REVIEW
                st.rerun()
        with cg:
            if st.button("✦ Generar PowerPoint"):
                with st.spinner("Montando el PowerPoint sobre la plantilla de Tessera…"):
                    try:
                        st.session_state.pptx = build_pptx(d, st.session_state.content)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error generando PPTX: {e}")
                        import traceback
                        st.code(traceback.format_exc())
    else:
        st.markdown('<div class="ok">✅ <strong>Propuesta generada.</strong> Descarga el PowerPoint con la plantilla visual de Tessera.</div>',
                    unsafe_allow_html=True)
        safe = d.get("client_name", "Cliente").replace(" ", "_")
        st.download_button(
            "⬇ Descargar PowerPoint",
            data=st.session_state.pptx,
            file_name=f"Tessera_{safe}_Propuesta.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("← Revisar contenido de nuevo"):
                st.session_state.pptx = None
                st.session_state.step = STEP_REVIEW
                st.rerun()
        with c2:
            if st.button("+ Nueva propuesta"):
                for k, v in [("step", 0), ("form", {}), ("content", None), ("pptx", None)]:
                    st.session_state[k] = v
                st.rerun()
