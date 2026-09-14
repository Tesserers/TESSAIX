"""TESSAIX — generación automática de propuestas comerciales en PowerPoint.

Paquete dividido por responsabilidad para que cada pieza se pueda tocar de
forma aislada:

- ``config``        rutas, constantes y catálogo de servicios.
- ``assets``        utilidades de imágenes/branding (PNG de Ailerons, etc).
- ``styles``        CSS inyectado en Streamlit y cabecera/pasos del wizard.
- ``auth``          pantalla de login por contraseña.
- ``xml_utils``      manipulación segura del XML interno del .pptx.
- ``logo_client``    búsqueda, validación y composición del logo del cliente.
- ``ai_content``     generación de contenido con Claude.
- ``pptx_builder``   ensamblado final del .pptx a partir de la plantilla.
- ``ui_home``        pantalla de selección de línea de negocio.
- ``ui_wizard``      pasos 0-4 del formulario (config, cliente, servicios...).
- ``ui_review``      paso 5: revisión/edición de contenido + logo antes de generar.
- ``ui_generate``    paso 6: generación final y descarga del .pptx.
"""
