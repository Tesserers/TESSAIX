"""Paso 6: generación final del .pptx y descarga."""
import streamlit as st

from .config import SERVICES_HC, STEP_REVIEW
from .pptx_builder import build_pptx


def step_generate():
    d = st.session_state.form
    st.markdown("### Genera tu PowerPoint")
    rows = [
        ("Cliente", d.get("client_name", "—")),
        ("Sector", d.get("sector", "—")),
        ("Idioma", "Español" if d.get("lang") == "es" else "English"),
        ("Tipo", "Con fees" if d.get("deck_type") == "propuesta" else "Sin fees"),
        ("Servicios", " · ".join([SERVICES_HC.get(s, s) for s in d.get("services", [])])),
        ("Oferta cruzada Finance/M&A", "Sí" if d.get("include_finance_crosssell") else "No"),
        ("Presentado por", d.get("presenter_name", "—")),
        ("Teléfono", d.get("presenter_phone", "—")),
        ("Logo cliente", {"auto": "Encontrado automáticamente", "manual": "Subido manualmente",
                           "none": "Sin logo"}.get(st.session_state.get("logo_source"), "—")),
    ]
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
                        st.session_state.pptx = build_pptx(
                            d, st.session_state.content, logo_bytes=st.session_state.get("logo_bytes")
                        )
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
                for k, v in [
                    ("step", 0), ("form", {}), ("content", None), ("pptx", None),
                    ("logo_meta", None), ("logo_bytes", None), ("logo_source", None),
                    ("logo_checked", False), ("logo_domain", ""),
                ]:
                    st.session_state[k] = v
                st.rerun()
