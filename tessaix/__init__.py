"""TESSAIX — generación automática de propuestas comerciales en PowerPoint.

Paquete dividido por responsabilidad para que cada pieza se pueda tocar de
forma aislada:

- ``config``        rutas (incluidas las dos plantillas maestras), constantes y catálogo de servicios.
- ``team``           datos de equipo y opciones de garantía/sedes.
- ``assets``         utilidades de imágenes/branding (PNG de Ailerons, etc).
- ``styles``         CSS inyectado en Streamlit y cabecera/pasos del wizard.
- ``auth``           pantalla de login por contraseña.
- ``xml_utils``      manipulación segura del XML interno del .pptx.
- ``text_fit``       estimación de ajuste de texto a cajas de tamaño fijo.
- ``ai_content``     generación de contenido con Claude (solo Human Capital).
- ``pptx_builder``   ensamblado final del .pptx a partir de la plantilla que corresponda.
- ``ui_home``        pantalla de selección de línea de negocio.
- ``ui_wizard``      pasos del formulario (config, cliente, servicios, equipo y sedes...).
- ``ui_review``      paso de revisión/edición de contenido antes de generar.
- ``ui_generate``    paso final: generación y descarga del .pptx.

``logo_client`` sigue en el repo (composición de logo sin deformar, con
salvaguarda de calidad) pero no se usa en el flujo actual: las plantillas
nuevas (Human Capital / Human Capital + Finance) no traen un hueco de logo
de cliente en la portada.
"""
