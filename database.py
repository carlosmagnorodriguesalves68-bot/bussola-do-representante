"""database.py — histórico incremental com SQLite"""
import sqlite3, os
import pandas as pd
from datetime import datetime

DB_PATH = "bussola_historico.db"

def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def inicializar_banco():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS importacoes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo        TEXT,
            arquivo     TEXT,
            importado_em TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS vendas (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            cliente      TEXT,
            cnpj         TEXT,
            cidade       TEXT,
            bandeira     TEXT,
            meta         REAL,
            realizado    REAL,
            projecao     REAL,
            gap          REAL,
            perc_gap     REAL,
            crescimento  REAL,
            valor_2025   REAL,
            prioridade   TEXT,
            acao         TEXT,
            importado_em TEXT,
            FOREIGN KEY(importacao_id) REFERENCES importacoes(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS metas (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            cliente      TEXT,
            cnpj         TEXT,
            meta_bruta   REAL,
            meta_mensal  REAL,
            laboratorio  TEXT,
            importado_em TEXT,
            FOREIGN KEY(importacao_id) REFERENCES importacoes(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS cmk (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            cnpj         TEXT,
            cliente      TEXT,
            cidade       TEXT,
            endereco     TEXT,
            bairro       TEXT,
            telefone     TEXT,
            supervisor   TEXT,
            rota         TEXT,
            importado_em TEXT,
            FOREIGN KEY(importacao_id) REFERENCES importacoes(id)
        )
    """)

    conn.commit()
    conn.close()

def registrar_importacao(tipo, arquivo):
    conn = get_conn()
    c = conn.cursor()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO importacoes (tipo, arquivo, importado_em) VALUES (?,?,?)",
              (tipo, arquivo, agora))
    imp_id = c.lastrowid
    conn.commit(); conn.close()
    return imp_id, agora

def salvar_vendas(df_norm, imp_id, agora):
    conn = get_conn()
    rows = []
    for _, row in df_norm.iterrows():
        rows.append((
            imp_id,
            str(row.get("CLIENTE_FINAL","")),
            str(row.get("CNPJ_LIMPO","")),
            str(row.get("CIDADE","")),
            str(row.get("BANDEIRA_FINAL","")),
            float(row.get("META",0)),
            float(row.get("REALIZADO",0)),
            float(row.get("PROJECAO",0)),
            float(row.get("GAP",0)),
            float(row.get("PERC_GAP",0)),
            float(row.get("CRESCIMENTO",0)),
            float(row.get("VALOR_2025",0)),
            str(row.get("PRIORIDADE","")),
            str(row.get("ACAO","")),
            agora,
        ))
    conn.executemany("""
        INSERT INTO vendas
        (importacao_id,cliente,cnpj,cidade,bandeira,meta,realizado,projecao,
         gap,perc_gap,crescimento,valor_2025,prioridade,acao,importado_em)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, rows)
    conn.commit(); conn.close()

def salvar_cmk(df_cmk, imp_id, agora):
    conn = get_conn()
    rows = []
    for _, row in df_cmk.iterrows():
        rows.append((
            imp_id,
            str(row.get("CNPJ","")),
            str(row.get("CLIENTE","")),
            str(row.get("CIDADE","")),
            str(row.get("ENDERECO","")),
            str(row.get("BAIRRO","")),
            str(row.get("TELEFONE","")),
            str(row.get("SUPERVISOR","")),
            str(row.get("ROTA","")),
            agora,
        ))
    conn.executemany("""
        INSERT INTO cmk
        (importacao_id,cnpj,cliente,cidade,endereco,bairro,telefone,supervisor,rota,importado_em)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, rows)
    conn.commit(); conn.close()

# ── Leitura ──────────────────────────────────────────────────
def ler_vendas_mais_recente():
    conn = get_conn()
    df = pd.read_sql("""
        SELECT v.* FROM vendas v
        INNER JOIN (
            SELECT MAX(id) as max_id FROM importacoes WHERE tipo='vendas'
        ) i ON v.importacao_id = i.max_id
    """, conn)
    conn.close()
    return df

def ler_cmk_mais_recente():
    conn = get_conn()
    df = pd.read_sql("""
        SELECT c.* FROM cmk c
        INNER JOIN (
            SELECT MAX(id) as max_id FROM importacoes WHERE tipo='cmk'
        ) i ON c.importacao_id = i.max_id
    """, conn)
    conn.close()
    return df

def ler_historico_cliente(nome_cliente):
    conn = get_conn()
    df = pd.read_sql("""
        SELECT cliente, meta, realizado, projecao, gap, perc_gap,
               crescimento, prioridade, importado_em
        FROM vendas
        WHERE cliente LIKE ?
        ORDER BY importado_em ASC
    """, conn, params=(f"%{nome_cliente}%",))
    conn.close()
    return df

def ler_importacoes():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM importacoes ORDER BY id DESC LIMIT 20", conn)
    conn.close()
    return df

def banco_tem_dados():
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM vendas")
        count = c.fetchone()[0]
        conn.close()
        return count > 0
    except:
        return False
