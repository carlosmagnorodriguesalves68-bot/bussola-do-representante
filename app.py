# ==========================================
# BÚSSOLA DO REPRESENTANTE - V3.5
# Estrutura Base Profissional
# ==========================================

import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path
import sqlite3

# ==========================================
# CONFIG
# ==========================================

st.set_page_config(
    page_title="Bússola do Representante",
    page_icon="🧭",
    layout="wide"
)

# ==========================================
# PASTAS
# ==========================================

Path("database").mkdir(exist_ok=True)
Path("uploads").mkdir(exist_ok=True)

# ==========================================
# DATABASE
# ==========================================

conn = sqlite3.connect("database/bussola.db", check_same_thread=False)

def salvar_historico(df, tabela):
    try:
        df["data_importacao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.to_sql(tabela, conn, if_exists="append", index=False)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar histórico: {e}")
        return False

# ==========================================
# CSS
# ==========================================

st.markdown("""
<style>

section[data-testid="stSidebar"] {
    display:none;
}

.block-container{
    padding-top:1rem;
    max-width:1500px;
}

.navbar{
    background: linear-gradient(90deg,#0D1B2A,#13293D,#0D1B2A);
    padding:14px 24px;
    border-radius:14px;
    margin-bottom:18px;
    border-bottom:3px solid #C62828;
}

.nav-title{
    color:white;
    font-size:24px;
    font-weight:800;
}

.kpi{
    background:white;
    border-radius:16px;
    padding:18px;
    border:1px solid #E5E7EB;
    box-shadow:0 2px 6px rgba(0,0,0,0.05);
}

.sec{
    background:white;
    border-radius:16px;
    padding:18px;
    border:1px solid #E5E7EB;
    margin-bottom:16px;
}

</style>
""", unsafe_allow_html=True)

# ==========================================
# NAVBAR
# ==========================================

if "pagina" not in st.session_state:
    st.session_state["pagina"] = "dashboard"

st.markdown("""
<div class='navbar'>
<div class='nav-title'>🧭 Bússola do Representante</div>
</div>
""", unsafe_allow_html=True)

c1,c2,c3,c4,c5,c6 = st.columns([2,1,1,1,1,1])

with c2:
    if st.button("🧭 Dashboard", use_container_width=True):
        st.session_state["pagina"] = "dashboard"

with c3:
    if st.button("👤 Cliente a Cliente", use_container_width=True):
        st.session_state["pagina"] = "cliente"

with c4:
    if st.button("💰 CotaBot", use_container_width=True):
        st.session_state["pagina"] = "cotabot"

with c5:
    if st.button("🧾 Cobrança", use_container_width=True):
        st.session_state["pagina"] = "cobranca"

with c6:
    if st.button("📍 Painel de Visitas", use_container_width=True):
        st.session_state["pagina"] = "visitas"

pagina = st.session_state["pagina"]

# ==========================================
# IMPORTAÇÃO
# ==========================================

with st.expander("📤 CENTRAL DE IMPORTAÇÃO", expanded=True):

    u1,u2,u3,u4 = st.columns(4)

    with u1:
        arquivo_csv = st.file_uploader(
            "CSV Power BI",
            type=["csv"],
            key="csv"
        )

    with u2:
        arquivo_excel = st.file_uploader(
            "Excel Complementar",
            type=["xlsx"],
            key="excel"
        )

    with u3:
        arquivo_meta = st.file_uploader(
            "Planilha Metas",
            type=["xlsx"],
            key="meta"
        )

    with u4:
        arquivo_cmk = st.file_uploader(
            "Planilha CMK",
            type=["xlsx"],
            key="cmk"
        )

    if st.button("🚀 Processar Dados", use_container_width=True):

        try:

            if arquivo_csv:
                df_csv = pd.read_csv(arquivo_csv)
                salvar_historico(df_csv, "historico_powerbi")

            if arquivo_excel:
                df_excel = pd.read_excel(arquivo_excel)
                salvar_historico(df_excel, "historico_excel")

            if arquivo_meta:
                df_meta = pd.read_excel(arquivo_meta)
                salvar_historico(df_meta, "historico_meta")

            if arquivo_cmk:
                df_cmk = pd.read_excel(arquivo_cmk)
                salvar_historico(df_cmk, "historico_cmk")

            st.success("✅ Dados processados com sucesso!")

        except Exception as e:
            st.error(f"Erro: {e}")

# ==========================================
# DASHBOARD
# ==========================================

if pagina == "dashboard":

    st.subheader("🧭 Dashboard Executivo")

    k1,k2,k3,k4,k5 = st.columns(5)

    with k1:
        st.markdown("""
        <div class='kpi'>
        <h5>Meta Total</h5>
        <h2>R$ 1.250.000</h2>
        </div>
        """, unsafe_allow_html=True)

    with k2:
        st.markdown("""
        <div class='kpi'>
        <h5>Realizado</h5>
        <h2>R$ 920.000</h2>
        </div>
        """, unsafe_allow_html=True)

    with k3:
        st.markdown("""
        <div class='kpi'>
        <h5>Projeção</h5>
        <h2>R$ 1.080.000</h2>
        </div>
        """, unsafe_allow_html=True)

    with k4:
        st.markdown("""
        <div class='kpi'>
        <h5>GAP</h5>
        <h2 style='color:#C62828'>-R$ 170.000</h2>
        </div>
        """, unsafe_allow_html=True)

    with k5:
        st.markdown("""
        <div class='kpi'>
        <h5>Urgentes</h5>
        <h2>18</h2>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    a1,a2 = st.columns(2)

    with a1:
        st.markdown("""
        <div class='sec'>
        <h4>🔴 TOP 5 URGENTES</h4>
        <p>• Cliente A - GAP R$ 12 mil</p>
        <p>• Cliente B - GAP R$ 9 mil</p>
        <p>• Cliente C - GAP R$ 8 mil</p>
        </div>
        """, unsafe_allow_html=True)

    with a2:
        st.markdown("""
        <div class='sec'>
        <h4>📍 Sugestão de Visitas</h4>
        <p>• Itaperuna</p>
        <p>• Santo Antônio de Pádua</p>
        <p>• Miracema</p>
        </div>
        """, unsafe_allow_html=True)

    b1,b2 = st.columns(2)

    with b1:
        st.markdown("""
        <div class='sec'>
        <h4>💰 O que vender hoje</h4>
        <p>• Linha Diabetes</p>
        <p>• Higiene</p>
        <p>• Vitaminas</p>
        </div>
        """, unsafe_allow_html=True)

    with b2:
        st.markdown("""
        <div class='sec'>
        <h4>📉 Laboratórios Negativos</h4>
        <p>• EMS</p>
        <p>• Eurofarma</p>
        <p>• Neo Química</p>
        </div>
        """, unsafe_allow_html=True)

# ==========================================
# CLIENTE A CLIENTE
# ==========================================

elif pagina == "cliente":

    st.subheader("👤 Cliente a Cliente")

    st.info("""
    Aqui ficará:
    - histórico do cliente
    - GAP
    - crescimento
    - produtos que comprava e parou
    - plano de ataque
    """)

# ==========================================
# COTABOT
# ==========================================

elif pagina == "cotabot":

    st.subheader("💰 CotaBot")

    st.info("""
    Integrar aqui:
    - upload cotação
    - cruzamento EAN
    - preenchimento automático
    """)

# ==========================================
# COBRANÇA
# ==========================================

elif pagina == "cobranca":

    st.subheader("🧾 Cobrança Inteligente")

    st.info("""
    Integrar:
    - títulos vencidos
    - dias em atraso
    - prioridade
    """)

# ==========================================
# VISITAS
# ==========================================

elif pagina == "visitas":

    st.subheader("📍 Painel de Visitas")

    st.info("""
    Painel inteligente:
    - ranking por cidade
    - clientes prioritários
    - melhor rota
    - GAP da cidade
    """)
