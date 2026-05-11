"""app.py — Bússola do Representante v3.0 — Hub principal"""
import streamlit as st
from utils import CSS

st.set_page_config(
    page_title="Bússola do Representante",
    page_icon="🧭",
    layout="wide",
)

st.markdown(CSS, unsafe_allow_html=True)

# ── Estado de navegação ──────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state["pagina"] = "dashboard"

# ── Navbar ───────────────────────────────────────────────────
st.markdown(
    '<div class="navbar"><div class="navbar-brand">🧭 Bússola do <span>Representante</span></div></div>',
    unsafe_allow_html=True,
)

PAGINAS = [
    ("dashboard", "🧭 Dashboard"),
    ("cliente",   "👤 Cliente a Cliente"),
    ("cotabot",   "💰 CotaBot"),
    ("cobranca",  "🧾 Cobrança"),
    ("visitas",   "📍 Painel de Visitas"),
]

cols = st.columns([2.2] + [1.4] * len(PAGINAS))
for i, (key, label) in enumerate(PAGINAS):
    with cols[i + 1]:
        tipo = "primary" if st.session_state["pagina"] == key else "secondary"
        if st.button(label, use_container_width=True, type=tipo, key=f"nav_{key}"):
            st.session_state["pagina"] = key
            st.rerun()

st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)

# ── Roteamento de páginas ────────────────────────────────────
pagina = st.session_state["pagina"]

if pagina == "dashboard":
    from pages.dashboard import render
    render()

elif pagina == "cliente":
    from pages.cliente import render
    render()

elif pagina == "cotabot":
    from pages.cotabot import render
    render()

elif pagina == "cobranca":
    from pages.cobranca import render
    render()

elif pagina == "visitas":
    from pages.visitas import render
    render()
