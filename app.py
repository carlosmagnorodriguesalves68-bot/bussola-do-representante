import streamlit as st
import pandas as pd
import unicodedata
import re
import io
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(
    page_title="Bússola do Representante",
    page_icon="🧭",
    layout="wide",
)

# ══════════════════════════════════════════════════════════════
# CSS GLOBAL — LAYOUT PROFISSIONAL
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
section[data-testid="stSidebar"] { display: none !important; }
header[data-testid="stHeader"] { display: none !important; }
.block-container { padding-top: 0 !important; padding-bottom: 2rem; max-width: 1480px; }
body { background:#F8FAFC; }
.navbar {
    background: linear-gradient(90deg, #0D1B2A 0%, #13293D 72%, #0D1B2A 100%);
    padding: 0 28px; display: flex; align-items: center;
    border-bottom: 3px solid #C62828; height: 58px; margin-bottom: 0;
}
.navbar-brand { font-size: 18px; font-weight: 900; color: white; margin-right: 32px; white-space: nowrap; }
.navbar-brand span { color: #E53935; }
.app-header {
    background: linear-gradient(90deg,#0D1B2A 0%,#13293D 62%,#C62828 100%);
    padding: 20px 28px; border-radius: 18px; margin: 16px 0 18px; color: white;
    box-shadow: 0 10px 24px rgba(15,23,42,.12);
}
.app-header h1 { color: white; margin-bottom: 5px; font-size: 27px; font-weight:900; }
.app-header p { color: #E5E7EB; font-size: 13px; margin: 0; }
.upload-box {
    background:white; border:1px solid #E5E7EB; border-radius:18px; padding:18px;
    box-shadow: 0 6px 18px rgba(15,23,42,.05); margin-bottom:18px;
}
.kpi-card { background: white; border: 1px solid #E5E7EB; border-radius: 16px; padding: 16px 18px; height: 100%; box-shadow: 0 6px 18px rgba(15,23,42,.04); }
.kpi-card.red { border-left: 5px solid #C62828; }
.kpi-card.green { border-left: 5px solid #1B8A3D; }
.kpi-card.amber { border-left: 5px solid #F9A825; }
.kpi-card.blue { border-left: 5px solid #1D4ED8; }
.kpi-card.navy { border-left: 5px solid #0D1B2A; }
.kpi-label { font-size: 11px; color: #64748B; font-weight: 800; text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 7px; }
.kpi-value { font-size: 25px; font-weight: 900; color: #0D1B2A; line-height: 1.1; }
.kpi-sub { font-size: 11px; color: #94A3B8; margin-top: 4px; }
.sec-title { font-size: 15px; font-weight: 900; color: #0D1B2A; margin: 18px 0 10px; padding-bottom: 7px; border-bottom: 2px solid #E5E7EB; }
.section-card { background: white; border: 1px solid #E5E7EB; border-radius: 18px; padding: 18px; margin-bottom: 18px; box-shadow: 0 6px 18px rgba(15,23,42,.04); }
.cli-card { background: white; border: 1px solid #E5E7EB; border-radius: 14px; padding: 12px 14px; margin-bottom: 9px; display: flex; justify-content: space-between; gap:12px; align-items: center; }
.cli-card.urgente { border-left: 5px solid #C62828; }
.cli-card.atencao { border-left: 5px solid #F9A825; }
.cli-card.ok { border-left: 5px solid #1B8A3D; }
.cli-name { font-size: 13px; font-weight: 900; color: #0D1B2A; }
.cli-sub { font-size: 11px; color: #64748B; margin-top: 2px; }
.badge { font-size: 11px; font-weight: 800; padding: 4px 10px; border-radius: 20px; white-space: nowrap; }
.badge-red { background:#FEF2F2; color:#C62828; }
.badge-amber { background:#FFFBEB; color:#B45309; }
.badge-green { background:#F0FDF4; color:#1B8A3D; }
.badge-blue { background:#EFF6FF; color:#1D4ED8; }
div[data-testid="stMetricLabel"] { color: #475569; font-weight: 700; }
div[data-testid="stMetricValue"] { color: #0D1B2A; font-weight: 900; }
.small-note { color:#64748B; font-size:12px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ESTADO + NAVBAR
# ══════════════════════════════════════════════════════════════
if "pagina" not in st.session_state:
    st.session_state["pagina"] = "dashboard"

st.markdown('<div class="navbar"><div class="navbar-brand">🧭 Bússola do <span>Representante</span></div></div>', unsafe_allow_html=True)

nc1,nc2,nc3,nc4,nc5,nc6 = st.columns([2.2,1.45,1.7,1.35,1.45,1.7])
with nc2:
    if st.button("🧭 Dashboard", use_container_width=True, type="primary" if st.session_state["pagina"]=="dashboard" else "secondary"):
        st.session_state["pagina"]="dashboard"; st.rerun()
with nc3:
    if st.button("👤 Cliente a Cliente", use_container_width=True, type="primary" if st.session_state["pagina"]=="cliente" else "secondary"):
        st.session_state["pagina"]="cliente"; st.rerun()
with nc4:
    if st.button("💰 CotaBot", use_container_width=True, type="primary" if st.session_state["pagina"]=="cotabot" else "secondary"):
        st.session_state["pagina"]="cotabot"; st.rerun()
with nc5:
    if st.button("🧾 Cobrança", use_container_width=True, type="primary" if st.session_state["pagina"]=="cobranca" else "secondary"):
        st.session_state["pagina"]="cobranca"; st.rerun()
with nc6:
    if st.button("📍 Painel de Visitas", use_container_width=True, type="primary" if st.session_state["pagina"]=="visitas" else "secondary"):
        st.session_state["pagina"]="visitas"; st.rerun()

pagina = st.session_state["pagina"]

# ══════════════════════════════════════════════════════════════
# FUNÇÕES GLOBAIS
# ══════════════════════════════════════════════════════════════
def _ln(txt):
    txt = str(txt).strip().upper()
    return unicodedata.normalize("NFKD", txt).encode("ASCII","ignore").decode("utf-8")

def _num(v):
    if pd.isna(v): return 0.0
    if isinstance(v,(int,float)): return float(v)
    v = str(v).strip().replace("R$","").replace("%","").replace(" ","")
    if "," in v:
        v = v.replace(".","").replace(",",".")
    try: return float(v)
    except: return 0.0

def _m(v):
    try: return f"R$ {float(v):,.0f}".replace(",",".")
    except: return "R$ 0"

def _p(v):
    try:
        x = float(v)
        if abs(x) <= 1: x *= 100
        return f"{x:,.1f}%".replace(",","X").replace(".",",").replace("X",".")
    except: return "0,0%"

def _ac(df, ops, exata=False):
    if df is None or df.empty: return None
    for o in ops:
        for c in df.columns:
            if _ln(c)==_ln(o): return c
    if not exata:
        for o in ops:
            for c in df.columns:
                if _ln(o) in _ln(c): return c
    return None

def limpar_cnpj(v):
    return re.sub(r"\D", "", str(v))

def ler_arquivo(file):
    if file is None: return pd.DataFrame()
    nome = file.name.lower()
    file.seek(0)
    if nome.endswith(".csv"):
        try:
            return pd.read_csv(file, sep=None, engine="python", encoding="utf-8")
        except Exception:
            file.seek(0)
            return pd.read_csv(file, sep=";", encoding="latin1")
    return pd.read_excel(file)

def prioridade(row):
    gap = row.get("GAP", 0)
    dias = row.get("DIAS_SEM_COMPRA", 0)
    perc = row.get("PERC_GAP", 0)
    if gap < 0 and (abs(perc) >= 0.20 or dias >= 15): return "🔴 URGENTE"
    if gap < 0 or dias >= 10: return "🟡 ATENÇÃO"
    return "🟢 AUMENTAR VENDA"

def acao_sugerida(row):
    if row.get("PRIORIDADE","") == "🔴 URGENTE": return "Recuperar hoje"
    if row.get("DIAS_SEM_COMPRA",0) >= 15: return "Ligar e ofertar reposição"
    if row.get("GAP",0) < 0: return "Atacar meta do cliente"
    return "Aumentar mix"

def montar_base(powerbi_file=None, metas_file=None, cmk_file=None):
    df = ler_arquivo(powerbi_file)
    metas = ler_arquivo(metas_file)
    cmk = ler_arquivo(cmk_file)

    if df.empty:
        return pd.DataFrame(), metas, cmk

    # Normaliza nomes principais sem destruir nomes originais
    col_cliente = _ac(df,["CLIENTE","RAZAO SOCIAL","RAZÃO SOCIAL","NOME CLIENTE","FANTASIA"])
    col_cnpj = _ac(df,["CNPJ","CPF/CNPJ"])
    col_cidade = _ac(df,["CIDADE","MUNICIPIO","MUNICÍPIO"])
    col_lab = _ac(df,["LABORATORIO","LABORATÓRIO","INDUSTRIA","INDÚSTRIA","FORNECEDOR"])
    col_prod = _ac(df,["PRODUTO","DESCRICAO","DESCRIÇÃO","ITEM"])
    col_meta = _ac(df,["META","OBJETIVO","OBJETIVO BRUTO","META BRUTA"])
    col_real = _ac(df,["REALIZADO","VENDA","VENDAS","FATURAMENTO","VALOR"])
    col_proj = _ac(df,["PROJECAO","PROJEÇÃO","REAL PROJ","PROJETADO"])
    col_gap = _ac(df,["GAP","DESVIO","DESVIO PROJ","SALDO"])
    col_perc = _ac(df,["% GAP","% DESVIO","PERC GAP","PERCENTUAL"])
    col_dias = _ac(df,["DIAS SEM COMPRA","DIAS_SEM_COMPRA","DIAS S/ COMPRA","DIAS"])
    col_ult = _ac(df,["ULTIMA COMPRA","ÚLTIMA COMPRA","DT ULTIMA COMPRA","DATA ULTIMA COMPRA"])
    col_qtd = _ac(df,["QTD","QTDE","QUANTIDADE"])

    base = pd.DataFrame()
    base["CLIENTE"] = df[col_cliente].astype(str).str.strip() if col_cliente else "SEM CLIENTE"
    base["CNPJ"] = df[col_cnpj].apply(limpar_cnpj) if col_cnpj else ""
    base["CIDADE"] = df[col_cidade].astype(str).str.strip() if col_cidade else ""
    base["LABORATORIO"] = df[col_lab].astype(str).str.strip() if col_lab else ""
    base["PRODUTO"] = df[col_prod].astype(str).str.strip() if col_prod else ""
    base["META"] = df[col_meta].apply(_num) if col_meta else 0
    base["REALIZADO"] = df[col_real].apply(_num) if col_real else 0
    base["PROJECAO"] = df[col_proj].apply(_num) if col_proj else base["REALIZADO"]
    base["GAP"] = df[col_gap].apply(_num) if col_gap else base["PROJECAO"] - base["META"]
    base["PERC_GAP"] = df[col_perc].apply(_num) if col_perc else base.apply(lambda r: (r["GAP"] / r["META"]) if r["META"] else 0, axis=1)
    base["DIAS_SEM_COMPRA"] = df[col_dias].apply(_num) if col_dias else 0
    base["ULTIMA_COMPRA"] = df[col_ult].astype(str) if col_ult else ""
    base["QTD"] = df[col_qtd].apply(_num) if col_qtd else 0

    # Metas por cliente: sobrescreve/atualiza objetivo bruto quando achar cliente ou CNPJ
    if not metas.empty:
        mc = _ac(metas,["CLIENTE","RAZAO SOCIAL","RAZÃO SOCIAL","NOME CLIENTE","FANTASIA"])
        mn = _ac(metas,["CNPJ","CPF/CNPJ"])
        mm = _ac(metas,["OBJETIVO BRUTO","META","META BRUTA","OBJETIVO","VALOR META"])
        if mm:
            metas_aux = metas.copy()
            metas_aux["_META_OBJ_"] = metas_aux[mm].apply(_num)
            if mn and base["CNPJ"].astype(str).str.len().max() > 0:
                metas_aux["CNPJ"] = metas_aux[mn].apply(limpar_cnpj)
                base = base.merge(metas_aux[["CNPJ","_META_OBJ_"]].drop_duplicates("CNPJ"), on="CNPJ", how="left")
                base["META"] = base["_META_OBJ_"].fillna(base["META"])
                base.drop(columns=["_META_OBJ_"], inplace=True)
            elif mc:
                metas_aux["CLIENTE"] = metas_aux[mc].astype(str).str.strip()
                base = base.merge(metas_aux[["CLIENTE","_META_OBJ_"]].drop_duplicates("CLIENTE"), on="CLIENTE", how="left")
                base["META"] = base["_META_OBJ_"].fillna(base["META"])
                base.drop(columns=["_META_OBJ_"], inplace=True)
            base["GAP"] = base["PROJECAO"] - base["META"]
            base["PERC_GAP"] = base.apply(lambda r: (r["GAP"] / r["META"]) if r["META"] else 0, axis=1)

    # CMK: endereço/cidade/bairro/telefone
    if not cmk.empty:
        cc = _ac(cmk,["CNPJ","CPF/CNPJ"])
        ccli = _ac(cmk,["CLIENTE","RAZAO SOCIAL","RAZÃO SOCIAL","NOME CLIENTE","FANTASIA"])
        cid = _ac(cmk,["CIDADE","MUNICIPIO","MUNICÍPIO"])
        end = _ac(cmk,["ENDERECO","ENDEREÇO","LOGRADOURO","RUA"])
        bai = _ac(cmk,["BAIRRO"])
        tel = _ac(cmk,["TELEFONE","TEL","CELULAR","WHATSAPP"])
        sup = _ac(cmk,["SUPERVISOR"])
        cmk_aux = pd.DataFrame()
        if cc:
            cmk_aux["CNPJ"] = cmk[cc].apply(limpar_cnpj)
        elif ccli:
            cmk_aux["CLIENTE"] = cmk[ccli].astype(str).str.strip()
        if cid: cmk_aux["CIDADE_CMK"] = cmk[cid].astype(str).str.strip()
        if end: cmk_aux["ENDERECO"] = cmk[end].astype(str).str.strip()
        if bai: cmk_aux["BAIRRO"] = cmk[bai].astype(str).str.strip()
        if tel: cmk_aux["TELEFONE"] = cmk[tel].astype(str).str.strip()
        if sup: cmk_aux["SUPERVISOR"] = cmk[sup].astype(str).str.strip()
        if "CNPJ" in cmk_aux.columns and base["CNPJ"].astype(str).str.len().max() > 0:
            base = base.merge(cmk_aux.drop_duplicates("CNPJ"), on="CNPJ", how="left")
        elif "CLIENTE" in cmk_aux.columns:
            base = base.merge(cmk_aux.drop_duplicates("CLIENTE"), on="CLIENTE", how="left")
        if "CIDADE_CMK" in base.columns:
            base["CIDADE"] = base["CIDADE"].replace("", pd.NA).fillna(base["CIDADE_CMK"])
        for col in ["ENDERECO","BAIRRO","TELEFONE","SUPERVISOR"]:
            if col not in base.columns: base[col] = ""
    else:
        for col in ["ENDERECO","BAIRRO","TELEFONE","SUPERVISOR"]:
            if col not in base.columns: base[col] = ""

    base["PRIORIDADE"] = base.apply(prioridade, axis=1)
    base["ACAO"] = base.apply(acao_sugerida, axis=1)
    ordem = {"🔴 URGENTE":1,"🟡 ATENÇÃO":2,"🟢 AUMENTAR VENDA":3}
    base["ORDEM"] = base["PRIORIDADE"].map(ordem).fillna(9)
    return base.sort_values(["ORDEM","GAP","DIAS_SEM_COMPRA"], ascending=[True, True, False]), metas, cmk

def carregar_central():
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    st.markdown("### 🔄 Central de Importação")
    st.markdown('<div class="small-note">Suba os arquivos que alimentam o Power BI e a Bússola. O dashboard será recalculado automaticamente.</div>', unsafe_allow_html=True)
    u1,u2,u3 = st.columns(3)
    with u1:
        powerbi_f = st.file_uploader("📊 CSV / Excel Power BI", type=["csv","xlsx","xls"], key="central_powerbi")
    with u2:
        metas_f = st.file_uploader("🎯 Planilha de Metas", type=["xlsx","xls","csv"], key="central_metas")
    with u3:
        cmk_f = st.file_uploader("📍 Planilha CMK / Endereço", type=["xlsx","xls","csv"], key="central_cmk")

    if powerbi_f:
        df, metas, cmk = montar_base(powerbi_f, metas_f, cmk_f)
        st.session_state["base_bussola"] = df
        st.session_state["metas_bussola"] = metas
        st.session_state["cmk_bussola"] = cmk
        st.session_state["ultima_att"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        st.success(f"✅ Dados processados com sucesso. Última atualização: {st.session_state['ultima_att']}")
    else:
        st.info("Suba pelo menos o arquivo principal do Power BI para ativar a Bússola.")
    st.markdown('</div>', unsafe_allow_html=True)


def obter_base():
    return st.session_state.get("base_bussola", pd.DataFrame())

# ══════════════════════════════════════════════════════════════
# DASHBOARD INICIAL
# ══════════════════════════════════════════════════════════════
if pagina == "dashboard":
    st.markdown('<div class="app-header"><h1>🧭 Dashboard Inicial</h1><p>Abra a Bússola e entenda em segundos quem visitar, quem cobrar e o que vender hoje.</p></div>', unsafe_allow_html=True)
    carregar_central()
    df = obter_base()

    if df.empty:
        st.markdown("""
        <div style="background:#F8FAFC;border:2px dashed #CBD5E1;border-radius:18px;padding:58px;text-align:center;margin-top:20px;">
            <div style="font-size:54px;margin-bottom:12px;">🧭</div>
            <div style="font-size:24px;font-weight:900;color:#0D1B2A;margin-bottom:8px;">Bússola do Representante</div>
            <div style="font-size:14px;color:#64748B;">Suba o arquivo principal para iniciar a central de inteligência.</div>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # CARDS PRINCIPAIS — MANTIDOS
    total = df["CLIENTE"].nunique()
    meta_t = df.groupby("CLIENTE")["META"].max().sum() if "META" in df.columns else 0
    real_t = df["REALIZADO"].sum()
    proj_t = df.groupby("CLIENTE")["PROJECAO"].sum().sum() if "PROJECAO" in df.columns else real_t
    gap_t = proj_t - meta_t
    urgentes = df[df["PRIORIDADE"]=="🔴 URGENTE"]["CLIENTE"].nunique()
    pct_meta = (real_t/meta_t*100) if meta_t>0 else 0

    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: st.markdown(f'<div class="kpi-card navy"><div class="kpi-label">Meta total</div><div class="kpi-value">{_m(meta_t)}</div><div class="kpi-sub">{total} clientes</div></div>', unsafe_allow_html=True)
    with k2: st.markdown(f'<div class="kpi-card green"><div class="kpi-label">Realizado</div><div class="kpi-value">{_m(real_t)}</div><div class="kpi-sub">{pct_meta:.1f}% da meta</div></div>', unsafe_allow_html=True)
    with k3: st.markdown(f'<div class="kpi-card blue"><div class="kpi-label">Projeção</div><div class="kpi-value">{_m(proj_t)}</div><div class="kpi-sub">até fim do mês</div></div>', unsafe_allow_html=True)
    with k4: st.markdown(f'<div class="kpi-card {"red" if gap_t<0 else "green"}"><div class="kpi-label">GAP total</div><div class="kpi-value">{_m(gap_t)}</div><div class="kpi-sub">{"déficit" if gap_t<0 else "superávit"}</div></div>', unsafe_allow_html=True)
    with k5: st.markdown(f'<div class="kpi-card red"><div class="kpi-label">🔴 Urgentes</div><div class="kpi-value">{urgentes}</div><div class="kpi-sub">clientes críticos</div></div>', unsafe_allow_html=True)

    # Agregação por cliente para blocos estratégicos
    cli = df.groupby(["CLIENTE","CIDADE"], dropna=False).agg(
        META=("META","max"), REALIZADO=("REALIZADO","sum"), PROJECAO=("PROJECAO","sum"),
        GAP=("GAP","sum"), DIAS_SEM_COMPRA=("DIAS_SEM_COMPRA","max"),
        ULTIMA_COMPRA=("ULTIMA_COMPRA","max"), ENDERECO=("ENDERECO","max"),
        PRIORIDADE=("PRIORIDADE","min"), ACAO=("ACAO","max")
    ).reset_index()
    cli["PERC_GAP"] = cli.apply(lambda r: (r["GAP"]/r["META"]) if r["META"] else 0, axis=1)
    cli["PRIORIDADE"] = cli.apply(prioridade, axis=1)
    cli["ACAO"] = cli.apply(acao_sugerida, axis=1)
    cli = cli.sort_values(["PRIORIDADE","GAP","DIAS_SEM_COMPRA"], ascending=[True, True, False])

    c1,c2,c3 = st.columns([1.05, 1.25, 1.05])

    with c1:
        st.markdown('<div class="sec-title">🔴 Top 5 urgentes</div>', unsafe_allow_html=True)
        top5 = cli[cli["PRIORIDADE"]=="🔴 URGENTE"].sort_values("GAP").head(5)
        if top5.empty: st.success("Nenhum cliente urgente no momento.")
        for _,r in top5.iterrows():
            st.markdown(f'''<div class="cli-card urgente"><div><div class="cli-name">{r['CLIENTE']}</div><div class="cli-sub">{r['CIDADE']} · GAP {_m(r['GAP'])} · {int(r['DIAS_SEM_COMPRA'])} dias sem comprar</div></div><span class="badge badge-red">Urgente</span></div>''', unsafe_allow_html=True)

        st.markdown('<div class="sec-title">⏰ Clientes sem comprar</div>', unsafe_allow_html=True)
        parados = cli.sort_values("DIAS_SEM_COMPRA", ascending=False).head(5)
        for _,r in parados.iterrows():
            st.markdown(f'''<div class="cli-card atencao"><div><div class="cli-name">{r['CLIENTE']}</div><div class="cli-sub">{r['CIDADE']} · última compra: {r['ULTIMA_COMPRA']}</div></div><span class="badge badge-amber">{int(r['DIAS_SEM_COMPRA'])} dias</span></div>''', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="sec-title">📍 Sugestão de visitas por município</div>', unsafe_allow_html=True)
        municipios = ["Todos"] + sorted([x for x in cli["CIDADE"].dropna().astype(str).unique() if x and x.lower() != "nan"])
        mun = st.selectbox("Filtrar município", municipios, key="dash_mun")
        vis = cli.copy()
        if mun != "Todos": vis = vis[vis["CIDADE"].astype(str)==mun]
        vis = vis.sort_values(["PRIORIDADE","GAP","DIAS_SEM_COMPRA"], ascending=[True, True, False]).head(8)
        for i,(_,r) in enumerate(vis.iterrows(),1):
            badge = "badge-red" if r["PRIORIDADE"]=="🔴 URGENTE" else "badge-amber" if r["PRIORIDADE"]=="🟡 ATENÇÃO" else "badge-green"
            cls = "urgente" if r["PRIORIDADE"]=="🔴 URGENTE" else "atencao" if r["PRIORIDADE"]=="🟡 ATENÇÃO" else "ok"
            st.markdown(f'''<div class="cli-card {cls}"><div><div class="cli-name">#{i} · {r['CLIENTE']}</div><div class="cli-sub">{r['CIDADE']} · {_m(r['GAP'])} · {r['ACAO']}</div></div><span class="badge {badge}">{r['PRIORIDADE'].replace('🔴 ','').replace('🟡 ','').replace('🟢 ','')}</span></div>''', unsafe_allow_html=True)

        st.markdown('<div class="sec-title">💰 O que vender/oferecer hoje</div>', unsafe_allow_html=True)
        vend = df.copy()
        if mun != "Todos": vend = vend[vend["CIDADE"].astype(str)==mun]
        if "PRODUTO" in vend.columns:
            prod_rank = vend[vend["PRODUTO"].astype(str).str.strip()!=""].groupby(["PRODUTO","LABORATORIO"], dropna=False).agg(VALOR=("REALIZADO","sum"), QTD=("QTD","sum"), CLIENTES=("CLIENTE","nunique")).reset_index().sort_values("VALOR", ascending=False).head(5)
            if prod_rank.empty:
                st.info("Quando o arquivo trouxer produto/laboratório, este bloco mostrará as ofertas do dia.")
            else:
                st.dataframe(prod_rank.rename(columns={"VALOR":"VENDIDO R$"}), use_container_width=True, hide_index=True)

    with c3:
        st.markdown('<div class="sec-title">🏆 Ranking dos mais vendidos</div>', unsafe_allow_html=True)
        labprod = df[df["PRODUTO"].astype(str).str.strip()!=""].groupby(["PRODUTO","LABORATORIO"], dropna=False).agg(VALOR=("REALIZADO","sum"), CLIENTES=("CLIENTE","nunique")).reset_index().sort_values("VALOR", ascending=False).head(8)
        if labprod.empty: st.info("Sem produto identificado no arquivo.")
        else: st.dataframe(labprod, use_container_width=True, hide_index=True)

        st.markdown('<div class="sec-title">📉 Ranking laboratório negativo</div>', unsafe_allow_html=True)
        labneg = df[df["LABORATORIO"].astype(str).str.strip()!=""].groupby("LABORATORIO", dropna=False).agg(META=("META","sum"), REALIZADO=("REALIZADO","sum"), PROJECAO=("PROJECAO","sum"), GAP=("GAP","sum"), CLIENTES=("CLIENTE","nunique")).reset_index().sort_values("GAP").head(8)
        if labneg.empty: st.info("Sem laboratório identificado no arquivo.")
        else: st.dataframe(labneg, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# CLIENTE A CLIENTE
# ══════════════════════════════════════════════════════════════
elif pagina == "cliente":
    st.markdown('<div class="app-header"><h1>👤 Cliente a Cliente</h1><p>Ficha inteligente para entender problema, oportunidade e plano de ataque.</p></div>', unsafe_allow_html=True)
    df = obter_base()
    if df.empty:
        st.warning("Volte ao Dashboard e suba os arquivos primeiro."); st.stop()
    clientes = sorted(df["CLIENTE"].dropna().astype(str).unique())
    escolha = st.selectbox("Selecione o cliente", clientes)
    d = df[df["CLIENTE"].astype(str)==escolha].copy()
    resumo = {
        "meta": d["META"].max(), "real": d["REALIZADO"].sum(), "proj": d["PROJECAO"].sum(),
        "gap": d["PROJECAO"].sum() - d["META"].max(), "dias": d["DIAS_SEM_COMPRA"].max(),
        "cidade": d["CIDADE"].dropna().astype(str).iloc[0] if len(d) else "", "end": d["ENDERECO"].dropna().astype(str).iloc[0] if "ENDERECO" in d.columns and len(d) else ""
    }
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Meta", _m(resumo["meta"])); k2.metric("Realizado", _m(resumo["real"])); k3.metric("Projeção", _m(resumo["proj"])); k4.metric("GAP", _m(resumo["gap"])); k5.metric("Dias sem compra", int(resumo["dias"]))
    st.markdown(f"""<div class="section-card"><b>📍 Localização:</b> {resumo['cidade']} · {resumo['end']}<br><b>Plano de ataque:</b> Cliente com GAP de {_m(resumo['gap'])}. Priorizar produtos/laboratórios com maior oportunidade e recuperar frequência de compra.</div>""", unsafe_allow_html=True)
    st.markdown("### 💰 O que oferecer para este cliente")
    ofertas = d[d["PRODUTO"].astype(str).str.strip()!=""].groupby(["PRODUTO","LABORATORIO"], dropna=False).agg(VALOR=("REALIZADO","sum"), QTD=("QTD","sum")).reset_index().sort_values("VALOR", ascending=False).head(15)
    if ofertas.empty: st.info("Sem produtos identificados para este cliente no arquivo.")
    else: st.dataframe(ofertas, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════
# COTABOT — MÓDULO SIMPLIFICADO E FUNCIONAL
# ══════════════════════════════════════════════════════════════
elif pagina == "cotabot":
    st.markdown('<div class="app-header"><h1>💰 CotaBot</h1><p>Cruzamento por EAN · preenchimento automático · Excel pronto para envio.</p></div>', unsafe_allow_html=True)
    st.info("Este módulo continua separado para você plugar sua versão atual do CotaBot. A estrutura principal da Bússola já está pronta para receber o código completo.")
    base_f = st.file_uploader("Base de produtos", type=["xlsx","xls","csv"], key="cotabot_base")
    cot_f = st.file_uploader("Planilha de cotação", type=["xlsx","xls"], key="cotabot_cot")
    if base_f and cot_f:
        st.success("Arquivos carregados. Aqui entra a lógica atual do seu CotaBot já validado.")

# ══════════════════════════════════════════════════════════════
# COBRANÇA
# ══════════════════════════════════════════════════════════════
elif pagina == "cobranca":
    st.markdown('<div class="app-header"><h1>🧾 Cobrança Inteligente</h1><p>Clientes vencidos, prioridade de cobrança e ação comercial.</p></div>', unsafe_allow_html=True)
    cob_f = st.file_uploader("📤 Subir planilha de cobrança", type=["xlsx","xls","csv"], key="cobranca_file")
    if cob_f:
        cob = ler_arquivo(cob_f)
        st.success("Planilha de cobrança carregada.")
        st.dataframe(cob.head(100), use_container_width=True)
    else:
        st.info("Suba sua planilha de cobrança para ativar este painel.")

# ══════════════════════════════════════════════════════════════
# PAINEL DE VISITAS
# ══════════════════════════════════════════════════════════════
elif pagina == "visitas":
    st.markdown('<div class="app-header"><h1>📍 Painel de Visitas</h1><p>Ranking inteligente por cidade para saber quem realmente visitar hoje.</p></div>', unsafe_allow_html=True)
    df = obter_base()
    if df.empty:
        st.warning("Volte ao Dashboard e suba os arquivos primeiro."); st.stop()

    cli = df.groupby(["CLIENTE","CIDADE"], dropna=False).agg(
        META=("META","max"), REALIZADO=("REALIZADO","sum"), PROJECAO=("PROJECAO","sum"),
        GAP=("GAP","sum"), DIAS_SEM_COMPRA=("DIAS_SEM_COMPRA","max"),
        ENDERECO=("ENDERECO","max"), BAIRRO=("BAIRRO","max"), TELEFONE=("TELEFONE","max")
    ).reset_index()
    cli["PERC_GAP"] = cli.apply(lambda r: (r["GAP"]/r["META"]) if r["META"] else 0, axis=1)
    cli["PRIORIDADE"] = cli.apply(prioridade, axis=1)
    cli["ACAO"] = cli.apply(acao_sugerida, axis=1)
    cidades = ["Todas"] + sorted([x for x in cli["CIDADE"].dropna().astype(str).unique() if x and x.lower() != "nan"])
    f1,f2,f3 = st.columns(3)
    with f1: cidade = st.selectbox("Município", cidades, key="vis_cidade")
    with f2: prio = st.selectbox("Prioridade", ["Todas","🔴 URGENTE","🟡 ATENÇÃO","🟢 AUMENTAR VENDA"], key="vis_prio")
    with f3: dias_min = st.number_input("Dias sem compra mínimo", min_value=0, value=0, step=1, key="vis_dias")

    basev = cli.copy()
    if cidade != "Todas": basev = basev[basev["CIDADE"].astype(str)==cidade]
    if prio != "Todas": basev = basev[basev["PRIORIDADE"]==prio]
    basev = basev[basev["DIAS_SEM_COMPRA"] >= dias_min]

    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Clientes na rota", len(basev))
    k2.metric("Urgentes", len(basev[basev["PRIORIDADE"]=="🔴 URGENTE"]))
    k3.metric("GAP da rota", _m(basev["GAP"].sum()))
    k4.metric("Sem compra", len(basev[basev["DIAS_SEM_COMPRA"]>0]))
    k5.metric("Potencial", _m(abs(basev[basev["GAP"]<0]["GAP"].sum())))

    st.markdown("### 🧠 Ranking sugerido de visitas")
    rota = basev.sort_values(["PRIORIDADE","GAP","DIAS_SEM_COMPRA"], ascending=[True, True, False]).head(30)
    tabela = rota[["CLIENTE","CIDADE","BAIRRO","ENDERECO","TELEFONE","PRIORIDADE","META","REALIZADO","PROJECAO","GAP","DIAS_SEM_COMPRA","ACAO"]].copy()
    st.dataframe(tabela, use_container_width=True, hide_index=True, height=560)
