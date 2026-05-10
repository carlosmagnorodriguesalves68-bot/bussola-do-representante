import streamlit as st
import pandas as pd
import unicodedata
import re
import io
from openpyxl import load_workbook

st.set_page_config(
    page_title="Bússola do Representante",
    page_icon="🧭",
    layout="wide",
)

# ── CSS GLOBAL ─────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #0D1B2A; }
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: white !important; }
.stMetric {
    background-color: #FFFFFF;
    border: 1px solid #E5E7EB;
    padding: 14px;
    border-radius: 14px;
    box-shadow: 0 2px 8px rgba(13,27,42,0.06);
}
div[data-testid="stMetricLabel"] { color: #475569; font-weight: 600; }
div[data-testid="stMetricValue"] { color: #0D1B2A; font-weight: 800; }
.stDataFrame { border-radius: 12px; overflow: hidden; }
button { border-radius: 10px !important; font-weight: 600 !important; }
.app-header {
    background: linear-gradient(90deg, #0D1B2A 0%, #13293D 60%, #C62828 100%);
    padding: 24px 28px; border-radius: 18px; margin-bottom: 22px; color: white;
}
.app-header h1 { color: white; margin-bottom: 4px; font-size: 34px; }
.app-header p  { color: #E5E7EB; font-size: 16px; margin: 0; }
.section-card {
    background-color: white; border: 1px solid #E5E7EB;
    border-radius: 16px; padding: 18px; margin-bottom: 18px;
    box-shadow: 0 2px 10px rgba(13,27,42,0.05);
}
.card {
    background: white; border: 1px solid #E5E7EB; border-radius: 16px;
    padding: 24px; margin-bottom: 12px;
}
.card-icon  { font-size: 32px; margin-bottom: 8px; }
.card-title { font-size: 18px; font-weight: 700; color: #0D1B2A; margin-bottom: 4px; }
.card-desc  { font-size: 14px; color: #64748B; }
</style>
""", unsafe_allow_html=True)

# ── MENU LATERAL ───────────────────────────────────────────────
st.sidebar.markdown("## 🧭 Bússola do Representante")
st.sidebar.markdown("---")
modulo = st.sidebar.radio(
    "Módulos",
    ["🏠 Início", "📡 Radar Comercial", "💰 CotaBot — Cotação"],
    label_visibility="collapsed"
)

# ══════════════════════════════════════════════════════════════
# MÓDULO: INÍCIO
# ══════════════════════════════════════════════════════════════
if modulo == "🏠 Início":
    st.markdown("""
    <div class="app-header">
        <h1>🧭 Bússola do Representante</h1>
        <p>Central inteligente — tudo em um só lugar.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class="card">
            <div class="card-icon">📡</div>
            <div class="card-title">Radar Comercial</div>
            <div class="card-desc">Metas, GAP, prioridade de clientes e roteiro inteligente do dia.</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card" style="border-left:4px solid #C62828;">
            <div class="card-icon">📍</div>
            <div class="card-title">Bússola de Visitas</div>
            <div class="card-desc">Em breve — rota otimizada e plano de ataque.</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
            <div class="card-icon">💰</div>
            <div class="card-title">CotaBot</div>
            <div class="card-desc">Upload de cotação, cruzamento por EAN e preenchimento automático de preços.</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.caption("Use o menu lateral para navegar entre os módulos.")

# ══════════════════════════════════════════════════════════════
# MÓDULO: RADAR COMERCIAL
# ══════════════════════════════════════════════════════════════
elif modulo == "📡 Radar Comercial":

    # ── funções ──
    def limpar_filtros():
        for key in list(st.session_state.keys()):
            if key.startswith("filtro_"):
                del st.session_state[key]

    def limpar_nome(txt):
        txt = str(txt).strip().upper()
        txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("utf-8")
        return txt

    def limpar_cnpj(valor):
        return re.sub(r"\D", "", str(valor))

    def numero(valor):
        if pd.isna(valor): return 0.0
        if isinstance(valor, (int, float)): return float(valor)
        valor = str(valor).strip().replace("R$","").replace("%","")
        valor = valor.replace(".", "").replace(",", ".")
        valor = valor.replace("-", "0") if valor == "-" else valor
        try: return float(valor)
        except: return 0.0

    def moeda(valor):
        try: return f"R$ {float(valor):,.0f}".replace(",", ".")
        except: return "R$ 0"

    def pct(valor):
        try: return f"{float(valor)*100:,.1f}%".replace(",","X").replace(".",",").replace("X",".")
        except: return "0,0%"

    def pct_cmv(valor):
        try: return f"{float(valor)*100:,.2f}%".replace(",","X").replace(".",",").replace("X",".")
        except: return "0,00%"

    def achar_coluna(df, opcoes, exata=False):
        for opcao in opcoes:
            for col in df.columns:
                if limpar_nome(col) == limpar_nome(opcao): return col
        if not exata:
            for opcao in opcoes:
                for col in df.columns:
                    if limpar_nome(opcao) in limpar_nome(col): return col
        return None

    def setor_limpo(valor):
        achou = re.search(r"\d+", str(valor))
        return achou.group() if achou else ""

    def calc_prioridade(row):
        if row["GAP"] >= 0: return "🟢 AUMENTAR VENDA"
        if row["PERC_GAP"] <= -0.20: return "🔴 URGENTE"
        return "🟡 ATENÇÃO"

    def calc_acao(row):
        if row["GAP"] >= 0: return "Aumentar venda"
        if row["PERC_GAP"] <= -0.20: return "Recuperar venda urgente"
        return "Acompanhar e recuperar venda"

    def cor_prioridade(valor):
        if "URGENTE" in str(valor): return "background-color:#F8D7DA;color:#842029;font-weight:bold"
        if "ATENÇÃO" in str(valor): return "background-color:#FFF3CD;color:#664D03;font-weight:bold"
        if "AUMENTAR" in str(valor): return "background-color:#D1E7DD;color:#0F5132;font-weight:bold"
        return ""

    def cor_numero(valor):
        try:
            v = float(valor)
            if v < 0: return "color:#C62828;font-weight:bold"
            if v > 0: return "color:#1B8A3D;font-weight:bold"
        except: pass
        return ""

    def cor_prazo(valor):
        try:
            v = float(valor)
            if v <= 0: return "color:#1B8A3D;font-weight:bold"
            return "color:#C62828;font-weight:bold"
        except: return ""

    def cor_cmv(valor):
        try:
            v = float(valor)
            if v <= 0: return "color:#1B8A3D;font-weight:bold"
            return "color:#C62828;font-weight:bold"
        except: return ""

    # ── header ──
    st.markdown("""
    <div class="app-header">
        <h1>📡 Radar Comercial Inteligente</h1>
        <p>Sistema inteligente para decidir quem visitar, corrigir desvios e acelerar suas vendas.</p>
    </div>
    """, unsafe_allow_html=True)

    # ── upload ──
    st.markdown("## 📁 Carregar bases")
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        curva_file = st.file_uploader("📊 Subir Curva semanal", type=["xlsx"], key="r_curva")
    with col_up2:
        cmk_file = st.file_uploader("📍 Subir CMK/endereço", type=["xlsx"], key="r_cmk")

    if not curva_file or not cmk_file:
        st.info("Suba a Curva semanal e o CMK para iniciar a análise.")
        st.stop()

    try:
        curva = pd.read_excel(curva_file, sheet_name="DADOS")
    except:
        curva = pd.read_excel(curva_file)
    cmk = pd.read_excel(cmk_file)

    curva.columns = [limpar_nome(c) for c in curva.columns]
    cmk.columns  = [limpar_nome(c) for c in cmk.columns]

    col_cliente      = achar_coluna(curva, ["CLIENTE"], exata=True)
    col_cnpj_curva   = achar_coluna(curva, ["CNPJ"], exata=True)
    col_cnpj_cmk     = achar_coluna(cmk,   ["CNPJ"], exata=True)
    col_bandeira     = achar_coluna(curva, ["BANDEIRA"], exata=True)
    col_supervisor   = achar_coluna(curva, ["SUPERVISOR"])
    col_setor        = achar_coluna(curva, ["COD.SETOR","COD SETOR","SETOR"])
    col_meta         = achar_coluna(curva, ["META"], exata=True)
    col_real         = achar_coluna(curva, ["REAL"], exata=True)
    col_proj         = achar_coluna(curva, ["REAL PROJ","REAL PROJ AC"])
    col_gap_planilha = achar_coluna(curva, ["DESVIO PROJ"], exata=True)
    col_perc_gap     = achar_coluna(curva, ["% DESVIO"], exata=True)
    col_2025         = achar_coluna(curva, ["2025"], exata=True)
    col_cresc        = achar_coluna(curva, ["% CRESC PROJ","% CRESC PROJ AC","CRESC"])
    col_meta_pv      = achar_coluna(curva, ["META PV"])
    col_pven         = achar_coluna(curva, ["P.VEN","P VEN","PVEN"])
    col_desvio_pv    = achar_coluna(curva, ["DESVIO PV"])
    col_meta_cmv     = achar_coluna(curva, ["META CMV %","META CMV"])
    col_cmv_real     = achar_coluna(curva, ["CMV %","CMV REAL","CMV"])
    col_desvio_cmv   = achar_coluna(curva, ["DESVIO CMV"])

    df = curva.copy()
    df["META"]      = df[col_meta].apply(numero)      if col_meta      else 0
    df["REALIZADO"] = df[col_real].apply(numero)      if col_real      else 0
    df["PROJECAO"]  = df[col_proj].apply(numero)      if col_proj      else 0
    df["GAP"]       = df[col_gap_planilha].apply(numero) if col_gap_planilha else df["PROJECAO"] - df["META"]
    df["PERC_GAP"]  = df[col_perc_gap].apply(numero)  if col_perc_gap  else 0
    df["VALOR_2025"]  = df[col_2025].apply(numero)    if col_2025      else 0
    df["CRESCIMENTO"] = df[col_cresc].apply(numero)   if col_cresc     else 0
    df["META_PV"]     = df[col_meta_pv].apply(numero) if col_meta_pv   else 0
    df["PRAZO_REAL"]  = df[col_pven].apply(numero)    if col_pven      else 0
    df["DESVIO_PRAZO"]= df[col_desvio_pv].apply(numero) if col_desvio_pv else 0
    df["META_CMV"]    = df[col_meta_cmv].apply(numero) if col_meta_cmv  else 0
    df["CMV_REAL"]    = df[col_cmv_real].apply(numero) if col_cmv_real  else 0
    df["DESVIO_CMV"]  = df[col_desvio_cmv].apply(numero) if col_desvio_cmv else 0

    df["CLIENTE_FINAL"]  = df[col_cliente].astype(str).str.strip() if col_cliente else "SEM NOME"
    df["BANDEIRA_FINAL"] = df[col_bandeira].astype(str).str.strip() if col_bandeira else ""
    df["SETOR_FINAL"]    = df[col_setor].apply(setor_limpo) if col_setor else ""

    if col_cnpj_curva and col_cnpj_cmk:
        df["_CNPJ_"]   = df[col_cnpj_curva].apply(limpar_cnpj)
        cmk["_CNPJ_"]  = cmk[col_cnpj_cmk].apply(limpar_cnpj)
        col_cidade_cmk  = achar_coluna(cmk, ["MUNICIPIO","CIDADE","CITY"])
        col_end_cmk     = achar_coluna(cmk, ["ENDERECO","ENDEREÇO","END","LOGRADOURO"])
        campos = {"_CNPJ_": "_CNPJ_"}
        if col_cidade_cmk: campos[col_cidade_cmk] = "CIDADE"
        if col_end_cmk:    campos[col_end_cmk]    = "ENDERECO"
        cmk_merge = cmk[list(campos.keys())].drop_duplicates("_CNPJ_")
        cmk_merge = cmk_merge.rename(columns=campos)
        df = df.merge(cmk_merge, on="_CNPJ_", how="left")
    if "CIDADE"   not in df.columns: df["CIDADE"]   = ""
    if "ENDERECO" not in df.columns: df["ENDERECO"] = ""

    df["PRIORIDADE"] = df.apply(calc_prioridade, axis=1)
    df["ACAO"]       = df.apply(calc_acao, axis=1)

    # ── filtros sidebar ──
    st.sidebar.markdown("## 📡 Radar Comercial")
    st.sidebar.caption("Filtros de decisão")
    if st.sidebar.button("🔄 Limpar filtros", key="r_limpar"):
        limpar_filtros()

    busca = st.sidebar.text_input("Buscar cliente", key="filtro_busca")
    if busca:
        df = df[df["CLIENTE_FINAL"].str.contains(busca, case=False, na=False)]

    if col_supervisor:
        supervisores = ["Todos"] + sorted(df[col_supervisor].dropna().astype(str).unique().tolist())
        f_sup = st.sidebar.selectbox("Supervisor", supervisores, key="filtro_supervisor")
        if f_sup != "Todos":
            df = df[df[col_supervisor].astype(str) == f_sup]

    setores = ["Todos"] + sorted(df["SETOR_FINAL"].dropna().astype(str).unique().tolist())
    f_setor = st.sidebar.selectbox("Setor", setores, key="filtro_setor")
    if f_setor != "Todos":
        df = df[df["SETOR_FINAL"] == f_setor]

    bandeiras = ["Todas"] + sorted(df["BANDEIRA_FINAL"].dropna().astype(str).unique().tolist())
    f_bandeira = st.sidebar.selectbox("Bandeira", bandeiras, key="filtro_bandeira")
    if f_bandeira != "Todas":
        df = df[df["BANDEIRA_FINAL"] == f_bandeira]

    cidades = ["Todas"] + sorted(df["CIDADE"].dropna().astype(str).unique().tolist())
    f_cidade = st.sidebar.selectbox("Cidade", cidades, key="filtro_cidade")
    if f_cidade != "Todas":
        df = df[df["CIDADE"].astype(str) == f_cidade]

    prioridades_opts = ["Todas","🔴 URGENTE","🟡 ATENÇÃO","🟢 AUMENTAR VENDA"]
    f_prio = st.sidebar.selectbox("Prioridade", prioridades_opts, key="filtro_prioridade")
    if f_prio != "Todas":
        df = df[df["PRIORIDADE"] == f_prio]

    if df.empty:
        st.warning("Nenhum cliente encontrado com os filtros aplicados.")
        st.stop()

    ordem = {"🔴 URGENTE":1,"🟡 ATENÇÃO":2,"🟢 AUMENTAR VENDA":3}
    df["ORDEM"] = df["PRIORIDADE"].map(ordem).fillna(9)
    df = df.sort_values(by=["ORDEM","GAP"], ascending=[True,True])

    # ── dashboard ──
    st.markdown("## 📊 Visão Executiva")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Clientes analisados", len(df))
    c2.metric("Clientes positivos", len(df[df["GAP"]>=0]))
    c3.metric("Clientes negativos", len(df[df["GAP"]<0]))
    c4.metric("Urgentes", len(df[df["PRIORIDADE"]=="🔴 URGENTE"]))
    urgentes = df[df["PRIORIDADE"]=="🔴 URGENTE"]
    if not urgentes.empty:
        cidade_critica = urgentes["CIDADE"].value_counts().idxmax()
        qtd_cidade     = urgentes["CIDADE"].value_counts().max()
        c5.metric("Cidade crítica", f"{cidade_critica} ({qtd_cidade})")
    else:
        c5.metric("Cidade crítica", "-")

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Meta total",  moeda(df["META"].sum()))
    m2.metric("Realizado",   moeda(df["REALIZADO"].sum()))
    m3.metric("Projeção",    moeda(df["PROJECAO"].sum()))
    m4.metric("Gap total",   moeda(df["GAP"].sum()))
    st.divider()

    # ── tabela principal ──
    st.markdown("## 🧭 Base Completa de Decisão")
    st.caption("Use esta visão para priorizar clientes, cidades e ações comerciais.")

    tabela = df[["CLIENTE_FINAL","CIDADE","PRIORIDADE","META","REALIZADO","PROJECAO",
                 "GAP","PERC_GAP","VALOR_2025","CRESCIMENTO","META_PV","PRAZO_REAL",
                 "DESVIO_PRAZO","META_CMV","CMV_REAL","DESVIO_CMV","ACAO"]].copy()
    tabela.columns = ["CLIENTE","CIDADE","PRIORIDADE","META R$","REALIZADO R$","PROJEÇÃO R$",
                      "GAP R$","% GAP","2025 R$","CRESC. VS 2025","META PV","PRAZO REAL",
                      "DESVIO PRAZO","META CMV","CMV REAL","DESVIO CMV","AÇÃO"]

    style = tabela.style.format({
        "META R$": moeda, "REALIZADO R$": moeda, "PROJEÇÃO R$": moeda, "GAP R$": moeda,
        "% GAP": pct, "2025 R$": moeda, "CRESC. VS 2025": pct,
        "META PV": "{:.1f}", "PRAZO REAL": "{:.1f}", "DESVIO PRAZO": "{:.1f}",
        "META CMV": pct_cmv, "CMV REAL": pct_cmv, "DESVIO CMV": pct_cmv,
    }).applymap(cor_prioridade, subset=["PRIORIDADE"]) \
      .applymap(cor_numero, subset=["GAP R$","CRESC. VS 2025"]) \
      .applymap(cor_prazo,  subset=["DESVIO PRAZO"]) \
      .applymap(cor_cmv,    subset=["DESVIO CMV"])

    st.dataframe(style, use_container_width=True, hide_index=True, height=520)
    st.divider()

    # ── roteiro ──
    st.markdown("## 📍 Roteiro Inteligente do Dia")
    st.caption("Gera uma sugestão de rota com os clientes mais importantes para visitar primeiro.")

    if st.button("🚀 Gerar roteiro do dia", key="r_roteiro"):
        df_urg = df[df["PRIORIDADE"]=="🔴 URGENTE"].copy().sort_values(by=["CIDADE","GAP"])
        df_ate = df[df["PRIORIDADE"]=="🟡 ATENÇÃO"].copy().sort_values(by=["CIDADE","GAP"])
        roteiro = pd.concat([df_urg, df_ate]).head(10)

        if roteiro.empty:
            st.warning("Nenhum cliente urgente ou em atenção encontrado para hoje.")
        else:
            st.success("Roteiro gerado com sucesso!")
            for cidade in roteiro["CIDADE"].dropna().unique():
                bloco = roteiro[roteiro["CIDADE"]==cidade]
                st.markdown(f"### 📍 {cidade}")
                for i, row in enumerate(bloco.itertuples(), start=1):
                    st.markdown(f"""
**{i}. {row.CLIENTE_FINAL}**  
Prioridade: **{row.PRIORIDADE}**  
Gap: **{moeda(row.GAP)}**  
% Gap: **{pct(row.PERC_GAP)}**  
Ação: **{row.ACAO}**
""")

    # ── resumo por cidade ──
    st.divider()
    st.markdown("## 📍 Resumo por Cidade")

    resumo_cidade = df.groupby("CIDADE", dropna=False).agg(
        CLIENTES=("CLIENTE_FINAL","count"),
        POSITIVOS=("GAP", lambda x: (x>=0).sum()),
        NEGATIVOS=("GAP", lambda x: (x<0).sum()),
        URGENTES=("PRIORIDADE", lambda x: (x=="🔴 URGENTE").sum()),
        ATENCAO=("PRIORIDADE",  lambda x: (x=="🟡 ATENÇÃO").sum()),
        META_TOTAL=("META","sum"), REALIZADO=("REALIZADO","sum"),
        PROJECAO=("PROJECAO","sum"), GAP_TOTAL=("GAP","sum"),
    ).reset_index().sort_values(by=["URGENTES","ATENCAO","GAP_TOTAL"], ascending=[False,False,True])

    resumo_style = resumo_cidade.style.format({
        "META_TOTAL": moeda, "REALIZADO": moeda, "PROJECAO": moeda, "GAP_TOTAL": moeda,
    }).applymap(cor_numero, subset=["GAP_TOTAL"])

    st.dataframe(resumo_style, use_container_width=True, hide_index=True)

    # ── análise individual ──
    st.divider()
    st.markdown("## 🎯 Resumo Objetivo do Cliente")

    cliente_escolhido = st.selectbox("Escolha o cliente", df["CLIENTE_FINAL"].astype(str).unique(), key="r_cliente")
    d = df[df["CLIENTE_FINAL"].astype(str)==cliente_escolhido].iloc[0]

    r1,r2,r3,r4 = st.columns(4)
    r1.metric("Meta",      moeda(d["META"]))
    r2.metric("Realizado", moeda(d["REALIZADO"]))
    r3.metric("Projeção",  moeda(d["PROJECAO"]))
    r4.metric("Gap",       moeda(d["GAP"]))

    st.markdown(f"""
<div class="section-card">

### 🏪 {d["CLIENTE_FINAL"]}

**Bandeira:** {d["BANDEIRA_FINAL"]}  
**Cidade:** {d.get("CIDADE", "")}  
**Endereço:** {d.get("ENDERECO", "")}  

**Prioridade:** {d["PRIORIDADE"]}  
**Ação sugerida:** {d["ACAO"]}  

---

### 📌 Leitura simples

- Meta do mês: **{moeda(d["META"])}**
- Realizado até agora: **{moeda(d["REALIZADO"])}**
- Projeção: **{moeda(d["PROJECAO"])}**
- Gap: **{moeda(d["GAP"])}**
- % Gap: **{pct(d["PERC_GAP"])}**
- Valor 2025: **{moeda(d["VALOR_2025"])}**
- Crescimento vs 2025: **{pct(d["CRESCIMENTO"])}**
- Prazo: meta **{d["META_PV"]:.1f} dias**, realizado **{d["PRAZO_REAL"]:.1f} dias**, desvio **{d["DESVIO_PRAZO"]:.1f} dias**
- CMV: meta **{pct_cmv(d["META_CMV"])}**, realizado **{pct_cmv(d["CMV_REAL"])}**, desvio **{pct_cmv(d["DESVIO_CMV"])}**

---

### 💬 Abordagem sugerida

"Passei aqui porque sua meta do mês é **{moeda(d["META"])}**, você realizou **{moeda(d["REALIZADO"])}** e está projetando **{moeda(d["PROJECAO"])}**.  
Hoje temos um gap de **{moeda(d["GAP"])}**, com % gap de **{pct(d["PERC_GAP"])}**, contra **{moeda(d["VALOR_2025"])}** em 2025.  
Quero te ajudar a ajustar isso com uma compra mais estratégica."

</div>
""", unsafe_allow_html=True)

    primeiro = df.iloc[0]
    st.error(f"👉 Comece por: {primeiro['CLIENTE_FINAL']} | {primeiro['PRIORIDADE']} | {primeiro['ACAO']}")


# ══════════════════════════════════════════════════════════════
# MÓDULO: COTABOT
# ══════════════════════════════════════════════════════════════
elif modulo == "💰 CotaBot — Cotação":

    def normalizar_texto(texto):
        texto = str(texto).strip().lower()
        trocas = {"ç":"c","ã":"a","á":"a","à":"a","â":"a","é":"e","ê":"e",
                  "í":"i","ó":"o","ô":"o","õ":"o","ú":"u"}
        for k,v in trocas.items():
            texto = texto.replace(k,v)
        return texto

    def limpar_ean(serie):
        return (serie.astype(str).str.strip()
                .str.replace(".0","",regex=False)
                .str.replace(" ","",regex=False)
                .str.replace("-","",regex=False))

    def converter_valor_monetario(serie):
        if pd.api.types.is_numeric_dtype(serie):
            return pd.to_numeric(serie, errors="coerce")
        s = serie.astype(str).str.strip().str.replace("R$","",regex=False).str.replace(" ","",regex=False)
        tem_virgula = s.str.contains(",",regex=False,na=False)
        s.loc[tem_virgula] = s.loc[tem_virgula].str.replace(".",",",regex=False).str.replace(",",".",regex=False)
        return pd.to_numeric(s, errors="coerce")

    def formatar_preco_brl(valor):
        if pd.isna(valor) or valor == "": return ""
        try: return f"{float(valor):.2f}".replace(".",",")
        except: return ""

    def carregar_excel_normal(uploaded_file, sheet_name=0):
        nome = uploaded_file.name.lower()
        uploaded_file.seek(0)
        if nome.endswith(".xlsx"): return pd.read_excel(uploaded_file, engine="openpyxl", sheet_name=sheet_name)
        if nome.endswith(".xls"):  return pd.read_excel(uploaded_file, engine="xlrd",    sheet_name=sheet_name)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)

    def carregar_excel_bruto(uploaded_file, sheet_name=0):
        nome = uploaded_file.name.lower()
        uploaded_file.seek(0)
        if nome.endswith(".xlsx"): return pd.read_excel(uploaded_file, engine="openpyxl", sheet_name=sheet_name, header=None)
        if nome.endswith(".xls"):  return pd.read_excel(uploaded_file, engine="xlrd",    sheet_name=sheet_name, header=None)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

    def detectar_linha_cabecalho_cotacao(df_bruto):
        palavras = ["ean","codigo ean","código ean","cod barra","cód barra",
                    "codigo de barras","produto","descricao","descrição",
                    "fabricante","qtd","qt","preco","preço","preco un",
                    "preço un","% desc","preco c/ desc","preço c/ desc"]
        limite = min(len(df_bruto), 25)
        melhor_linha, melhor_score = 0, -1
        for i in range(limite):
            linha = df_bruto.iloc[i].fillna("").astype(str).tolist()
            linha_norm = [normalizar_texto(x) for x in linha]
            score = sum(1 for cel in linha_norm for p in palavras if p in cel)
            if score > melhor_score:
                melhor_score = score
                melhor_linha = i
        return melhor_linha

    def detectar_linha_cabecalho_base(df_bruto):
        palavras = ["codigo ean","código ean","descricao","descrição",
                    "laboratorio","laboratório","st","preco nf","preço nf","estoque"]
        limite = min(len(df_bruto), 15)
        melhor_linha, melhor_score = 0, -1
        for i in range(limite):
            linha = df_bruto.iloc[i].fillna("").astype(str).tolist()
            linha_norm = [normalizar_texto(x) for x in linha]
            score = sum(1 for cel in linha_norm for p in palavras if p in cel)
            if score > melhor_score:
                melhor_score = score
                melhor_linha = i
        return melhor_linha

    def construir_dataframe_com_cabecalho(df_bruto, header_row):
        cab   = df_bruto.iloc[header_row].fillna("").astype(str).tolist()
        dados = df_bruto.iloc[header_row+1:].copy().reset_index(drop=True)
        dados.columns = cab
        return dados

    def encontrar_coluna_por_nomes(colunas, nomes_alvo):
        mapa = {normalizar_texto(c): c for c in colunas}
        for alvo in nomes_alvo:
            alvo_norm = normalizar_texto(alvo)
            for col_norm, col_original in mapa.items():
                if alvo_norm == col_norm: return col_original
            for col_norm, col_original in mapa.items():
                if alvo_norm in col_norm: return col_original
        return None

    def sugerir_coluna_ean(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["ean","codigo ean","código ean","cod barra","cód barra","codigo de barras","gtin"])

    def sugerir_coluna_preco_real(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["preço real","preco real","preço final","preco final","valor final","preço venda","preco venda"])

    def sugerir_coluna_st(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["st","valor st","substituicao tributaria","substituição tributária"])

    def sugerir_coluna_preco_nf(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["preço nf","preco nf","valor nf","preço nota","preco nota","nf"])

    def sugerir_coluna_estoque(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["estoque","saldo","qtd estoque","quantidade estoque","disponivel","disponível"])

    def sugerir_coluna_preco_cotacao(df):
        return encontrar_coluna_por_nomes(df.columns,
            ["preço","preco","price","valor","preço un","preco un","preço unit","preco unit",
             "preço unitário","preco unitario","preço c/ desc","preco c/ desc"])

    def escrever_precos_em_xlsx_original(uploaded_file, aba_nome, header_row_zero_based,
                                          preco_col_idx_zero_based, precos_numericos):
        uploaded_file.seek(0)
        wb = load_workbook(uploaded_file)
        ws = wb[aba_nome]
        data_start_row = header_row_zero_based + 2
        for i, preco in enumerate(precos_numericos):
            row_excel = data_start_row + i
            col_excel = preco_col_idx_zero_based + 1
            if pd.notna(preco):
                ws.cell(row=row_excel, column=col_excel, value=float(preco))
            else:
                ws.cell(row=row_excel, column=col_excel, value=None)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.getvalue()

    # ── header ──
    st.markdown("""
    <div class="app-header">
        <h1>💰 CotaBot</h1>
        <p>Sua cotação pronta em segundos — cruzamento automático por EAN.</p>
    </div>
    """, unsafe_allow_html=True)

    # ── base da empresa ──
    st.markdown("### 1. Base da empresa")
    base_file = st.file_uploader("Subir base de produtos (EAN + preço + estoque)", type=["xlsx","xls"], key="c_base")

    base_df   = None
    base_bruto = None

    if base_file:
        base_bruto = carregar_excel_bruto(base_file)
        base_header_row = detectar_linha_cabecalho_base(base_bruto)
        base_df = construir_dataframe_com_cabecalho(base_bruto, base_header_row)
        st.markdown("### 2. Prévia da base")
        st.dataframe(base_df.head(10), use_container_width=True)

    # ── cotação do cliente ──
    st.markdown("---")
    st.markdown("### 3. Cotação do cliente")

    cotacao_file = st.file_uploader("Subir planilha de cotação do cliente", type=["xlsx","xls"], key="c_cotacao")

    cotacao_df   = None
    cotacao_bruto = None
    aba_escolhida = None

    if cotacao_file:
        nome_arq = cotacao_file.name.lower()
        if nome_arq.endswith(".xlsx"):
            cotacao_file.seek(0)
            wb_temp = load_workbook(cotacao_file, read_only=True)
            abas = wb_temp.sheetnames
            wb_temp.close()
        else:
            abas = [0]

        if len(abas) > 1:
            aba_escolhida = st.selectbox("Aba da cotação", abas, key="c_aba")
        else:
            aba_escolhida = abas[0]

        cotacao_bruto = carregar_excel_bruto(cotacao_file, sheet_name=aba_escolhida)
        cotacao_header_row_detectado = detectar_linha_cabecalho_cotacao(cotacao_bruto)

        header_row_manual = st.number_input(
            "Linha do cabeçalho da cotação", min_value=1,
            max_value=max(1, len(cotacao_bruto)),
            value=int(cotacao_header_row_detectado+1), step=1,
            key="c_header"
        )
        header_row  = int(header_row_manual - 1)
        cotacao_df  = construir_dataframe_com_cabecalho(cotacao_bruto, header_row)

        st.markdown("### Prévia da cotação")
        st.dataframe(cotacao_df.head(10), use_container_width=True)

    # ── associação ──
    if base_df is not None and cotacao_df is not None:
        st.markdown("---")
        st.subheader("4. Associação das colunas")

        col_base_ean_sug        = sugerir_coluna_ean(base_df)
        col_base_preco_real_sug = sugerir_coluna_preco_real(base_df)
        col_base_st_sug         = sugerir_coluna_st(base_df)
        col_base_preco_nf_sug   = sugerir_coluna_preco_nf(base_df)
        col_base_estoque_sug    = sugerir_coluna_estoque(base_df)
        col_cot_ean_sug         = sugerir_coluna_ean(cotacao_df)
        col_cot_preco_sug       = sugerir_coluna_preco_cotacao(cotacao_df)

        estoque_minimo = st.number_input("Estoque mínimo para incluir produto", min_value=0, value=1, step=1, key="c_estoque")

        st.info("O sistema sugere automaticamente, mas o representante confirma antes de processar.")

        s1, s2 = st.columns(2)

        with s1:
            st.markdown("#### Base da empresa")
            opcoes_base = ["-- Selecionar --"] + list(base_df.columns)
            idx_base_ean = opcoes_base.index(col_base_ean_sug) if col_base_ean_sug in opcoes_base else 0
            col_base_ean = st.selectbox("Coluna EAN da base", opcoes_base, index=idx_base_ean, key="c_base_ean")

            modo_preco = st.radio("Forma de preço da base",
                options=["Usar PREÇO REAL","Calcular ST + PREÇO NF"],
                index=0 if col_base_preco_real_sug else 1, key="c_modo_preco")

            if modo_preco == "Usar PREÇO REAL":
                idx_pr = opcoes_base.index(col_base_preco_real_sug) if col_base_preco_real_sug in opcoes_base else 0
                col_base_preco_real = st.selectbox("Coluna PREÇO REAL", opcoes_base, index=idx_pr, key="c_preco_real")
                col_base_st = None
                col_base_preco_nf = None
            else:
                idx_st = opcoes_base.index(col_base_st_sug) if col_base_st_sug in opcoes_base else 0
                idx_nf = opcoes_base.index(col_base_preco_nf_sug) if col_base_preco_nf_sug in opcoes_base else 0
                col_base_st       = st.selectbox("Coluna ST",       opcoes_base, index=idx_st, key="c_st")
                col_base_preco_nf = st.selectbox("Coluna PREÇO NF", opcoes_base, index=idx_nf, key="c_preco_nf")
                col_base_preco_real = None

            col_base_estoque = col_base_estoque_sug
            if col_base_estoque_sug:
                st.caption(f"Coluna ESTOQUE detectada: **{col_base_estoque_sug}**")
            else:
                st.warning("Não consegui detectar automaticamente a coluna ESTOQUE.")

        with s2:
            st.markdown("#### Cotação do cliente")
            opcoes_cot = ["-- Selecionar --"] + list(cotacao_df.columns)
            idx_cot_ean   = opcoes_cot.index(col_cot_ean_sug)   if col_cot_ean_sug   in opcoes_cot else 0
            idx_cot_preco = opcoes_cot.index(col_cot_preco_sug) if col_cot_preco_sug in opcoes_cot else 0
            col_cot_ean   = st.selectbox("Coluna EAN da cotação",   opcoes_cot, index=idx_cot_ean,   key="c_cot_ean")
            col_cot_preco = st.selectbox("Coluna PREÇO da cotação", opcoes_cot, index=idx_cot_preco, key="c_cot_preco")

        st.markdown("---")
        processar = st.button("5. Processar cotação", use_container_width=True, key="c_processar")

        if processar:
            try:
                erros = []
                if col_base_ean == "-- Selecionar --": erros.append("Selecione a coluna EAN da base.")
                if not col_base_estoque:               erros.append("Coluna ESTOQUE não detectada.")
                if col_cot_ean == "-- Selecionar --":  erros.append("Selecione a coluna EAN da cotação.")
                if col_cot_preco == "-- Selecionar --":erros.append("Selecione a coluna PREÇO da cotação.")
                if modo_preco == "Usar PREÇO REAL" and col_base_preco_real == "-- Selecionar --":
                    erros.append("Selecione a coluna PREÇO REAL.")
                if modo_preco != "Usar PREÇO REAL":
                    if col_base_st == "-- Selecionar --":       erros.append("Selecione a coluna ST.")
                    if col_base_preco_nf == "-- Selecionar --": erros.append("Selecione a coluna PREÇO NF.")
                if erros:
                    for e in erros: st.error(e)
                    st.stop()

                base_proc = base_df.copy()
                cot_proc  = cotacao_df.copy()

                base_proc[col_base_ean] = limpar_ean(base_proc[col_base_ean])
                cot_proc[col_cot_ean]   = limpar_ean(cot_proc[col_cot_ean])
                base_proc[col_base_estoque] = pd.to_numeric(base_proc[col_base_estoque], errors="coerce").fillna(0)

                if modo_preco == "Usar PREÇO REAL":
                    base_proc["_PRECO_FINAL_"] = converter_valor_monetario(base_proc[col_base_preco_real])
                else:
                    base_proc["_ST_"]       = converter_valor_monetario(base_proc[col_base_st])
                    base_proc["_PRECO_NF_"] = converter_valor_monetario(base_proc[col_base_preco_nf])
                    base_proc["_PRECO_FINAL_"] = base_proc["_ST_"].fillna(0) + base_proc["_PRECO_NF_"].fillna(0)

                base_filtrada = base_proc[base_proc[col_base_estoque] >= estoque_minimo].copy()
                base_merge    = base_filtrada[[col_base_ean,"_PRECO_FINAL_"]].drop_duplicates(subset=[col_base_ean])

                resultado = cot_proc.merge(base_merge, left_on=col_cot_ean, right_on=col_base_ean, how="left")
                precos_numericos = resultado["_PRECO_FINAL_"].tolist()
                precos_preview   = [formatar_preco_brl(x) for x in precos_numericos]

                total_itens    = len(cot_proc)
                encontrados    = int(pd.notna(resultado["_PRECO_FINAL_"]).sum())
                nao_encontrados= int(pd.isna(resultado["_PRECO_FINAL_"]).sum())

                c1,c2,c3 = st.columns(3)
                c1.metric("Itens na cotação", total_itens)
                c2.metric("Encontrados",      encontrados)
                c3.metric("Não encontrados",  nao_encontrados)

                preview = cot_proc.copy()
                preview[col_cot_preco] = preview[col_cot_preco].astype("object")
                preview[col_cot_preco] = precos_preview

                st.markdown("### 6. Prévia final")
                st.dataframe(preview.head(30), use_container_width=True)

                linha_cab = cotacao_bruto.iloc[header_row].fillna("").astype(str).tolist()
                preco_col_idx = next((idx for idx,v in enumerate(linha_cab)
                                      if str(v).strip()==str(col_cot_preco).strip()), None)

                if preco_col_idx is None:
                    st.error("Não consegui localizar a coluna de preço na planilha original.")
                    st.stop()

                if cotacao_file.name.lower().endswith(".xlsx"):
                    arquivo_saida = escrever_precos_em_xlsx_original(
                        cotacao_file, aba_escolhida, header_row, preco_col_idx, precos_numericos)
                    st.download_button(
                        label="⬇️ Baixar cotação preenchida",
                        data=arquivo_saida,
                        file_name="cotacao_preenchida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="c_download")
                else:
                    resultado_bruto = cotacao_bruto.copy()
                    resultado_bruto[preco_col_idx] = resultado_bruto[preco_col_idx].astype("object")
                    for i, preco in enumerate(precos_numericos):
                        linha_real = header_row + 1 + i
                        resultado_bruto.iat[linha_real, preco_col_idx] = float(preco) if pd.notna(preco) else None
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        resultado_bruto.to_excel(writer, index=False, header=False, sheet_name="Cotacao_Preenchida")
                    output.seek(0)
                    st.download_button(
                        label="⬇️ Baixar cotação preenchida",
                        data=output.getvalue(),
                        file_name="cotacao_preenchida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="c_download_xls")
                    st.info("Arquivo .xls: saída em .xlsx, estrutura preservada.")

            except Exception as e:
                st.error(f"Erro ao processar: {e}")
    else:
        st.markdown("---")
        st.info("Envie a base e a cotação para configurar o preenchimento.")
