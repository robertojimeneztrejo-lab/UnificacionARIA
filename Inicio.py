import streamlit as st
import requests
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="Portal ARIA",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Configuración de saludo dinámico ─────────────────────────────────────────
UMBRAL_FRIO = 14  # °C

# Zonas horarias mexicanas (para el aviso de día festivo, que solo aplica a México)
ZONAS_MEXICO = {
    "America/Mexico_City", "America/Tijuana", "America/Hermosillo",
    "America/Chihuahua", "America/Mazatlan", "America/Cancun",
    "America/Bahia_Banderas", "America/Merida", "America/Monterrey",
    "America/Matamoros",
}

# Ciudad de referencia (lat, lon, nombre) por zona horaria, para el aviso de clima
TZ_COORDS = {
    "America/Mexico_City": (19.4326, -99.1332, "Ciudad de México"),
    "America/Tijuana": (32.5149, -117.0382, "Tijuana"),
    "America/Cancun": (21.1619, -86.8515, "Cancún"),
    "America/Monterrey": (25.6866, -100.3161, "Monterrey"),
    "America/Bogota": (4.7110, -74.0721, "Bogotá"),
    "America/Lima": (-12.0464, -77.0428, "Lima"),
    "America/Santiago": (-33.4489, -70.6693, "Santiago"),
    "America/Argentina/Buenos_Aires": (-34.6037, -58.3816, "Buenos Aires"),
    "America/Sao_Paulo": (-23.5505, -46.6333, "São Paulo"),
    "America/New_York": (40.7128, -74.0060, "Nueva York"),
    "America/Chicago": (41.8781, -87.6298, "Chicago"),
    "America/Denver": (39.7392, -104.9903, "Denver"),
    "America/Los_Angeles": (34.0522, -118.2437, "Los Ángeles"),
    "America/Toronto": (43.6532, -79.3832, "Toronto"),
    "Europe/Madrid": (40.4168, -3.7038, "Madrid"),
    "Europe/London": (51.5074, -0.1278, "Londres"),
    "Europe/Paris": (48.8566, 2.3522, "París"),
    "Europe/Berlin": (52.5200, 13.4050, "Berlín"),
    "Europe/Rome": (41.9028, 12.4964, "Roma"),
    "Europe/Lisbon": (38.7223, -9.1393, "Lisboa"),
    "Europe/Amsterdam": (52.3676, 4.9041, "Ámsterdam"),
    "Europe/Moscow": (55.7558, 37.6173, "Moscú"),
    "Africa/Cairo": (30.0444, 31.2357, "El Cairo"),
    "Africa/Johannesburg": (-26.2041, 28.0473, "Johannesburgo"),
    "Asia/Tokyo": (35.6762, 139.6503, "Tokio"),
    "Asia/Shanghai": (31.2304, 121.4737, "Shanghái"),
    "Asia/Hong_Kong": (22.3193, 114.1694, "Hong Kong"),
    "Asia/Singapore": (1.3521, 103.8198, "Singapur"),
    "Asia/Kolkata": (28.6139, 77.2090, "Nueva Delhi"),
    "Asia/Dubai": (25.2048, 55.2708, "Dubái"),
    "Australia/Sydney": (-33.8688, 151.2093, "Sídney"),
    "Pacific/Auckland": (-36.8485, 174.7633, "Auckland"),
}


def _n_esimo_lunes(year: int, month: int, n: int) -> date:
    """Devuelve la fecha del n-ésimo lunes de un mes (1 = primer lunes)."""
    primer_dia = date(year, month, 1)
    dias_hasta_lunes = (0 - primer_dia.weekday()) % 7
    primer_lunes = primer_dia + timedelta(days=dias_hasta_lunes)
    return primer_lunes + timedelta(weeks=n - 1)


def dia_festivo(fecha: date) -> str | None:
    """Días festivos oficiales en México (Art. 74 LFT) para el año de la fecha dada."""
    year = fecha.year
    festivos = {
        date(year, 1, 1): "Año Nuevo",
        _n_esimo_lunes(year, 2, 1): "Día de la Constitución",
        _n_esimo_lunes(year, 3, 3): "Natalicio de Benito Juárez",
        date(year, 5, 1): "Día del Trabajo",
        date(year, 9, 16): "Día de la Independencia",
        _n_esimo_lunes(year, 11, 3): "Día de la Revolución Mexicana",
        date(year, 12, 25): "Navidad",
    }
    return festivos.get(fecha)


@st.cache_data(ttl=1800)
def temperatura_en(lat: float, lon: float):
    """Temperatura actual (°C) en unas coordenadas. Devuelve None si la consulta falla."""
    try:
        resp = requests.get(
            "https://api.open-meteo.com/v1/forecast",
            params={"latitude": lat, "longitude": lon, "current": "temperature_2m"},
            timeout=4,
        )
        return resp.json()["current"]["temperature_2m"]
    except Exception:
        return None


def mensaje_bienvenida() -> str:
    # Zona horaria real del navegador del visitante (puede estar vacía justo
    # después de un cambio de página; en ese caso usamos CDMX como respaldo).
    tz_name = st.context.timezone or "America/Mexico_City"
    try:
        ahora = datetime.now(ZoneInfo(tz_name))
    except Exception:
        tz_name = "America/Mexico_City"
        ahora = datetime.now(ZoneInfo(tz_name))

    hora = ahora.hour
    if 5 <= hora < 12:
        saludo = "Buenos días"
    elif 12 <= hora < 19:
        saludo = "Buenas tardes"
    else:
        saludo = "Buenas noches"

    if tz_name in ZONAS_MEXICO:
        festivo = dia_festivo(ahora.date())
        if festivo:
            return f"{saludo}. Hoy es {festivo}. Bienvenido a Portal ARIA."

    coords = TZ_COORDS.get(tz_name)
    if coords:
        lat, lon, ciudad = coords
        temp = temperatura_en(lat, lon)
        if temp is not None and temp <= UMBRAL_FRIO:
            return f"{saludo}. El ambiente está fresco en {ciudad} ({temp:.0f}°C). Bienvenido a Portal ARIA."

    return f"{saludo}, bienvenido a Portal ARIA."


st.markdown("""
<style>
:root {
    --navy: #0D2B4E;
    --navy-mid: #1A3F6F;
    --gold: #C8973A;
    --teal: #06b6d4;
}
.hero {
    padding: 3rem 2rem;
    border-radius: 18px;
    background: linear-gradient(135deg, var(--navy) 0%, var(--navy-mid) 100%);
    color: #f1f5f9;
    margin-bottom: 2rem;
    text-align: center;
}
.hero h1 { font-size: 2.2rem; margin-bottom: .4rem; }
.hero .subtitle { color: var(--gold); font-size: .85rem; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; margin-bottom: 1.1rem; }
.hero p { color: #cbd5e1; font-size: 1.05rem; margin: .2rem 0; }
.card {
    border: 1px solid #1e2d45;
    border-radius: 14px;
    padding: 1.6rem;
    height: 100%;
}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="hero">
    <h1>🧭 Portal ARIA</h1>
    <p class="subtitle">Alliance Recognition & Intelligence Architecture</p>
    <p>{mensaje_bienvenida()}</p>
    <p>Elige la herramienta que necesitas para tu búsqueda académica.</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card">
        <h3>💼 Pro Licencias</h3>
        <p>Busca software con licencias para uso académico e institucional.</p>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/1_Pro_Licencias.py", label="Ir a Pro Licencias", icon="💼")

with col2:
    st.markdown("""
    <div class="card">
        <h3>🎓 Pro Membresías</h3>
        <p>Busca membresías profesionales e institucionales para programas de posgrado.</p>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/2_Pro_Membresias.py", label="Ir a Pro Membresías", icon="🎓")
