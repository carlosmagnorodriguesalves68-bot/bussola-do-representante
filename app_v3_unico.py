
import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Bússola do Representante", page_icon="🧭", layout="wide")

Path("database").mkdir(exist_ok=True)

conn = sqlite3.connect("database/bussola.db", check_same_thread=False)

def ler_csv_automatico(arquivo):
    separadores = [";", ",", "|", "\t"]
    for sep in separadores:
        try:
            arquivo.seek(0)
            df = pd.read_csv(
                arquivo,
                sep=sep,
                encoding="utf-8",
                on_bad_lines="skip",
                low_memory=False
            )
            if len(df.columns) > 2:
                return df
        except:
            pass

    arquivo.seek(0)

    return pd.read_csv(
        arquivo,
        sep=None,
        engine="python",
        encoding="utf-8",
        on_bad_lines="skip"
    )

def salvar_historico(df, tabela):
    df["data_importacao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df.to_sql(tabela, conn, if_exists="append", index=False)

st.title("🧭 Bússola do Representante V4.0")

with st.expander("📤 CENTRAL DE IMPORTAÇÃO", expanded=True):

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        arquivo_csv = st.file_uploader("CSV Power BI", type=["csv"])

    with c2:
        arquivo_excel = st.file_uploader("Excel Complementar", type=["xlsx","xls"])

    with c3:
        arquivo_meta = st.file_uploader("Planilha Metas", type=["xlsx","xls"])

    with c4:
        arquivo_cmk = st.file_uploader("Planilha CMK", type=["xlsx","xls"])

    if st.button("🚀 Processar Dados"):

        if arquivo_csv:
            df = ler_csv_automatico(arquivo_csv)
            salvar_historico(df, "historico_powerbi")
            st.success(f"CSV OK - {len(df)} linhas")

        if arquivo_excel:
            df2 = pd.read_excel(arquivo_excel)
            salvar_historico(df2, "historico_excel")
            st.success(f"Excel OK - {len(df2)} linhas")

        if arquivo_meta:
            df3 = pd.read_excel(arquivo_meta)
            salvar_historico(df3, "historico_meta")
            st.success(f"Meta OK - {len(df3)} linhas")

        if arquivo_cmk:
            df4 = pd.read_excel(arquivo_cmk)
            salvar_historico(df4, "historico_cmk")
            st.success(f"CMK OK - {len(df4)} linhas")

st.subheader("🧭 Dashboard Executivo")

k1,k2,k3,k4,k5 = st.columns(5)

k1.metric("Meta Total", "R$ 1.250.000")
k2.metric("Realizado", "R$ 920.000")
k3.metric("Projeção", "R$ 1.080.000")
k4.metric("GAP", "-R$ 170.000")
k5.metric("Urgentes", "18")
