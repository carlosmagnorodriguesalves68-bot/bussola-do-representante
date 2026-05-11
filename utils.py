"""utils.py — funções globais da Bússola do Representante"""
import unicodedata, re
import pandas as pd

# ── Normalização ─────────────────────────────────────────────
def ln(txt):
    txt = str(txt).strip().upper()
    return unicodedata.normalize("NFKD", txt).encode("ASCII","ignore").decode("utf-8")

def norm_pt(t):
    t = str(t).strip().lower()
    for k,v in {"ç":"c","ã":"a","á":"a","à":"a","â":"a","é":"e","ê":"e",
                "í":"i","ó":"o","ô":"o","õ":"o","ú":"u"}.items():
        t = t.replace(k, v)
    return t

def limpar_cnpj(v):
    return re.sub(r"\D","", str(v))

def setor_num(v):
    m = re.search(r"\d+", str(v))
    return m.group() if m else ""

# ── Conversão numérica ───────────────────────────────────────
def num(v):
    if pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).strip().replace("R$","").replace("%","").replace(".","").replace(",",".")
    try: return float(v)
    except: return 0.0

def conv_moeda(s):
    if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s, errors="coerce")
    s2 = s.astype(str).str.strip().str.replace("R$","",regex=False).str.replace(" ","",regex=False)
    tv = s2.str.contains(",", regex=False, na=False)
    s2.loc[tv] = s2.loc[tv].str.replace(".",",",regex=False).str.replace(",",".",regex=False)
    return pd.to_numeric(s2, errors="coerce")

# ── Formatação ───────────────────────────────────────────────
def m(v):
    try: return f"R$ {float(v):,.0f}".replace(",",".")
    except: return "R$ 0"

def p(v):
    try: return f"{float(v)*100:,.1f}%".replace(",","X").replace(".",",").replace("X",".")
    except: return "0,0%"

def p2(v):
    try: return f"{float(v)*100:,.2f}%".replace(",","X").replace(".",",").replace("X",".")
    except: return "0,00%"

def fmt_brl(v):
    if pd.isna(v) or v == "": return ""
    try: return f"{float(v):.2f}".replace(".",",")
    except: return ""

def leia_ean(s):
    return (s.astype(str).str.strip()
            .str.replace(".0","",regex=False)
            .str.replace(" ","",regex=False)
            .str.replace("-","",regex=False))

# ── Busca de coluna ──────────────────────────────────────────
def ac(df, ops, exata=False):
    for o in ops:
        for c in df.columns:
            if ln(c) == ln(o): return c
    if not exata:
        for o in ops:
            for c in df.columns:
                if ln(o) in ln(c): return c
    return None

def fc(cols, names):
    mp = {norm_pt(c): c for c in cols}
    for n in names:
        nn = norm_pt(n)
        for cn, co in mp.items():
            if nn == cn: return co
        for cn, co in mp.items():
            if nn in cn: return co
    return None

# ── Prioridade ───────────────────────────────────────────────
def prio(row):
    if row["GAP"] >= 0: return "🟢 AUMENTAR VENDA"
    if row["PERC_GAP"] <= -0.20: return "🔴 URGENTE"
    return "🟡 ATENÇÃO"

def acao(row):
    if row["GAP"] >= 0: return "Aumentar venda"
    if row["PERC_GAP"] <= -0.20: return "Recuperar urgente"
    return "Acompanhar"

# ── Estilos condicional ──────────────────────────────────────
def cor_prio(v):
    if "URGENTE"  in str(v): return "background-color:#F8D7DA;color:#842029;font-weight:bold"
    if "ATENÇÃO"  in str(v): return "background-color:#FFF3CD;color:#664D03;font-weight:bold"
    if "AUMENTAR" in str(v): return "background-color:#D1E7DD;color:#0F5132;font-weight:bold"
    return ""

def cor_num(v):
    try:
        x = float(v)
        if x < 0: return "color:#C62828;font-weight:bold"
        if x > 0: return "color:#1B8A3D;font-weight:bold"
    except: pass
    return ""

# ── Detecção de cabeçalho (cotabot) ─────────────────────────
def det_header(df_bruto, words, lim=25):
    best_l, best_s = 0, -1
    for i in range(min(len(df_bruto), lim)):
        row = [norm_pt(x) for x in df_bruto.iloc[i].fillna("").astype(str).tolist()]
        sc = sum(1 for c in row for w in words if w in c)
        if sc > best_s: best_s = sc; best_l = i
    return best_l

def mk_df(df_bruto, h):
    cab = df_bruto.iloc[h].fillna("").astype(str).tolist()
    d   = df_bruto.iloc[h+1:].copy().reset_index(drop=True)
    d.columns = cab
    return d

def xl_bruto(f, s=0):
    f.seek(0); n = f.name.lower()
    if n.endswith(".xlsx"): return pd.read_excel(f, engine="openpyxl", sheet_name=s, header=None)
    if n.endswith(".xls"):  return pd.read_excel(f, engine="xlrd",     sheet_name=s, header=None)
    return pd.read_excel(f, sheet_name=s, header=None)

# ── CSS global ───────────────────────────────────────────────
CSS = """
<style>
section[data-testid="stSidebar"] { display:none !important; }
header[data-testid="stHeader"]   { display:none !important; }
.block-container { padding-top:0 !important; padding-bottom:2rem; max-width:1400px; }

.navbar {
    background: linear-gradient(90deg,#0D1B2A 0%,#13293D 70%,#0D1B2A 100%);
    padding: 0 28px; display:flex; align-items:center;
    border-bottom: 3px solid #C62828; height:56px; margin-bottom:0;
}
.navbar-brand { font-size:17px; font-weight:800; color:white; margin-right:32px; white-space:nowrap; }
.navbar-brand span { color:#C62828; }

.kpi-card { background:white; border:1px solid #E5E7EB; border-radius:14px; padding:16px 18px; }
.kpi-card.red   { border-left:4px solid #C62828; }
.kpi-card.green { border-left:4px solid #1B8A3D; }
.kpi-card.amber { border-left:4px solid #F9A825; }
.kpi-card.blue  { border-left:4px solid #1D4ED8; }
.kpi-card.navy  { border-left:4px solid #0D1B2A; }
.kpi-label { font-size:11px; color:#64748B; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; margin-bottom:6px; }
.kpi-value { font-size:24px; font-weight:800; color:#0D1B2A; line-height:1.1; }
.kpi-sub   { font-size:11px; color:#94A3B8; margin-top:3px; }

.cli-card { background:white; border:1px solid #E5E7EB; border-radius:12px; padding:12px 14px;
            margin-bottom:8px; display:flex; justify-content:space-between; align-items:center; }
.cli-card.urgente { border-left:4px solid #C62828; }
.cli-card.atencao { border-left:4px solid #F9A825; }
.cli-card.positivo{ border-left:4px solid #1B8A3D; }
.cli-name { font-size:13px; font-weight:700; color:#0D1B2A; }
.cli-sub  { font-size:11px; color:#64748B; margin-top:2px; }

.badge { font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; white-space:nowrap; }
.badge-red   { background:#FEF2F2; color:#C62828; }
.badge-amber { background:#FFFBEB; color:#B45309; }
.badge-green { background:#F0FDF4; color:#1B8A3D; }
.badge-blue  { background:#EFF6FF; color:#1D4ED8; }

.sec-title { font-size:14px; font-weight:700; color:#0D1B2A; margin:20px 0 10px;
             padding-bottom:6px; border-bottom:2px solid #E5E7EB; }
.app-header { background:linear-gradient(90deg,#0D1B2A 0%,#13293D 60%,#C62828 100%);
              padding:20px 28px; border-radius:16px; margin-bottom:20px; color:white; }
.app-header h1 { color:white; margin-bottom:4px; font-size:26px; }
.app-header p  { color:#E5E7EB; font-size:13px; margin:0; }
.section-card  { background:white; border:1px solid #E5E7EB; border-radius:16px;
                 padding:18px; margin-bottom:18px; }
div[data-testid="stMetricLabel"] { color:#475569; font-weight:600; }
div[data-testid="stMetricValue"] { color:#0D1B2A; font-weight:800; }
</style>
"""
