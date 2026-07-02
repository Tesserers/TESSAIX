# TESSAIX — Propuestas comerciales inteligentes

App Streamlit independiente para generar propuestas comerciales de Tessera Human Capital.

## Stack
- Streamlit
- Anthropic Claude (generación de contenido)
- python-pptx + Pillow (montaje del PPTX con tipografía Ailerons/Raleway/Manrope)

## Estructura
```
tessaix/
├── tessaix_app/
│   ├── app.py              # App principal
│   └── assets/
│       ├── Ailerons-Typeface.otf
│       ├── logo_azul.png
│       ├── logo_blanco.png
│       └── logo_negro.png
├── .streamlit/
│   └── config.toml
├── requirements.txt
└── README.md
```

## Deploy en Streamlit Cloud

1. Sube este repo a `Tesserers/TESSAIX` en GitHub
2. Ve a share.streamlit.io → New app
3. Repo: `Tesserers/TESSAIX`
4. Branch: `main`
5. Main file: `tessaix_app/app.py`
6. En **Secrets** añade:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
7. Deploy

## Contraseña de acceso
`Sales2026@`

## Tipografía
- **TESSERA / TESSAIX** → Ailerons (renderizado vía PIL desde Ailerons-Typeface.otf)
- **Texto, títulos, cuerpo** → Raleway
- **Números, estadísticas, porcentajes** → Manrope

## Paleta Tessera
- Navy `#202031` — dominante
- Pearl `#F0E8DE`
- Gold `#FBE0A0`
- Teal `#587579`
- Granate `#762D35`
