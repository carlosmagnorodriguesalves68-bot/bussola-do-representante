import streamlit as st

st.set_page_config(
    page_title="Bússola do Representante",
    page_icon="🧭",
    layout="wide",
)

# ── CSS global do hub ──────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #0D1B2A;
}
[data-testid="stSidebar"] * {
    color: white !important;
}
[data-testid="stSidebarNav"] a {
    padding: 10px 16px;
    border-radius: 8px;
    margin-bottom: 4px;
    display: block;
    font-weight: 500;
}
[data-testid="stSidebarNav"] a:hover {
    background-color: rgba(255,255,255,0.08);
}
</style>
""", unsafe_allow_html=True)

# ── Páginas ────────────────────────────────────────────────────
home        = st.Page("pages/home.py",    title="🧭 Início",               default=True)
radar       = st.Page("pages/radar.py",   title="📡 Radar Comercial")
cotabot     = st.Page("pages/cotabot.py", title="💰 CotaBot — Cotação")
visitas     = st.Page("pages/visitas.py", title="📍 Bússola de Visitas")

# ── Navegação ──────────────────────────────────────────────────
nav = st.navigation(
    {
        "🧭 Bússola": [home],
        "📊 Análise":  [radar],
        "🛒 Cotação":  [cotabot],
        "🗺️ Visitas":  [visitas],
    }
)

nav.run()
