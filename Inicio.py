import streamlit as st

st.set_page_config(
    page_title="Portal ARIA",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
}
.hero h1 { font-size: 2.2rem; margin-bottom: .4rem; }
.hero p { color: #cbd5e1; font-size: 1.05rem; }
.card {
    border: 1px solid #1e2d45;
    border-radius: 14px;
    padding: 1.6rem;
    height: 100%;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>🧭 Portal ARIA</h1>
    <p>Elige la herramienta que necesitas para tu búsqueda académica.</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card">
        <h3>💼 Pro Licencias</h3>
        <p>Busca software con licencias 100% gratuitas para uso académico e institucional.</p>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/1_💼_Pro_Licencias.py", label="Ir a Pro Licencias", icon="💼")

with col2:
    st.markdown("""
    <div class="card">
        <h3>🎓 Pro Membresías</h3>
        <p>Busca membresías profesionales e institucionales gratuitas para programas de posgrado.</p>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/2_🎓_Pro_Membresías.py", label="Ir a Pro Membresías", icon="🎓")
