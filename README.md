# Portal ARIA — Pro Licencias + Pro Membresías

App fusionada de Streamlit con navegación lateral nativa entre dos herramientas.

## Estructura

```
Inicio.py                          ← página de bienvenida con accesos directos
pages/
  1_💼_Pro_Licencias.py            ← antes LicenciasPro/app.py
  2_🎓_Pro_Membresías.py           ← antes MembresiaJunio/app.py
.streamlit/config.toml             ← tema visual compartido
requirements.txt                   ← dependencias fusionadas
```

Streamlit detecta automáticamente la carpeta `pages/` y genera el menú lateral.
El número al inicio de cada archivo controla el orden, el emoji controla el ícono.

## Pasos para subir esto a GitHub (vía web, sin terminal)

1. Crea un repositorio nuevo en GitHub (ej. `portal-aria`).
2. En la página del repo, click en "Add file" → "Upload files".
3. Arrastra **toda la carpeta** `portal-aria` (la mayoría de navegadores permite
   arrastrar carpetas completas y GitHub preserva la estructura interna).
   Si tu navegador no lo permite, sube primero `Inicio.py`, `requirements.txt`,
   y luego entra a "Add file" → "Create new file" escribiendo
   `pages/1_💼_Pro_Licencias.py` como nombre (esto crea la carpeta `pages/`
   automáticamente) y pega el contenido.
4. Commit.

## Despliegue en Streamlit Cloud

1. En share.streamlit.io, "New app" → selecciona el repo `portal-aria`.
2. Main file path: `Inicio.py`
3. En "Secrets" (Settings → Secrets), agrega:

```toml
GEMINI_API_KEY = "tu-clave-de-gemini"
APPS_SCRIPT_URL = "url-del-apps-script-de-licencias"
SHEETS_CSV_URL = "url-del-csv-publicado-de-licencias"
```

Nota: `Pro Membresías` actualmente tiene su URL de Apps Script y de CSV
escritas directamente en el código (líneas ~266-267), no en `st.secrets`.
Funciona igual, pero si quieres mantener el mismo patrón de seguridad que usas
en tus otros proyectos (secrets nunca en el código/GitHub), puedo moverlas a
`st.secrets` también — solo dime.

## Pendiente de tu parte

- Revisar que el dominio personalizado (si lo configuras vía Streamlit Cloud
  o el VPS que estás evaluando) apunte a esta única app en lugar de a las dos
  apps separadas.
- Las dos apps anteriores (`LicenciasPro` y `MembresiaJunio` por separado)
  pueden quedar archivadas o eliminadas de Streamlit Cloud una vez confirmes
  que la fusionada funciona correctamente.
