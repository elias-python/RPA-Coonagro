"""
Força a reatualização do Base_Dados_Coonagro.xlsm a partir do banco SQLite,
sem precisar reiniciar o XML Monitoring.py.

Uso: python forcar_atualizacao.py
"""

import sqlite3
import sys
import os

# Importa funções do monitoramento principal (nome com espaço requer importlib)
import importlib
import importlib.util
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
mon = importlib.util.spec_from_file_location(
    "xml_monitoring",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "XML Monitoring.py"),
)
mod = importlib.util.module_from_spec(mon)
mon.loader.exec_module(mod)

caminho_db = mod.caminho_db
atualizar_excel_com_espacos = mod.atualizar_excel_com_espacos

# Recarregar traz TODAS as notas do banco, sem filtro de data nem de nfs_lancadas.txt.

print("[FORÇAR ATUALIZAÇÃO] Lendo banco de dados...")

# ── Migração: garante que a coluna data_importacao existe ─────────────────────
conn_mig = sqlite3.connect(caminho_db)
try:
    conn_mig.execute("ALTER TABLE notas_itens ADD COLUMN data_importacao TEXT")
    conn_mig.commit()
    print("[MIGRAÇÃO] Coluna data_importacao adicionada ao banco.")
except Exception:
    pass  # já existe
conn_mig.close()
# ─────────────────────────────────────────────────────────────────────────────

# ── Limpeza: remove registros com CFOP diferente de 5902 e 5124 ──────────────
conn = sqlite3.connect(caminho_db)
cursor = conn.execute(
    "SELECT COUNT(*) FROM notas_itens WHERE cfop NOT IN ('5902', '5124')"
)
total_sujos = cursor.fetchone()[0]
if total_sujos > 0:
    conn.execute("DELETE FROM notas_itens WHERE cfop NOT IN ('5902', '5124')")
    conn.commit()
    print(
        f"[LIMPEZA] {total_sujos} registro(s) com CFOP fora de 5902/5124 removidos do banco."
    )
else:
    print("[LIMPEZA] Banco já limpo — nenhum CFOP inválido encontrado.")
conn.close()
# ─────────────────────────────────────────────────────────────────────────────

conn = sqlite3.connect(caminho_db)
query = f"""
    SELECT
        nf                  as 'Nº Nota Fiscal',
        seq                 as 'Seq.',
        cod_material        as 'Código Material',
        descricao           as 'Descrição do Material',
        ordem_producao      as 'Ordem de Produção',
        qtd                 as 'Qtd.',
        un                  as 'UN.',
        vlr_unit            as 'Vlr. Unitário',
        vlr_total_item      as 'Vlr. Total Item',
        cfop                as 'CFOP',
        nf_vinculada_5124   as 'NF Vinculada (5124)',
        nf_origem           as 'NF Origem do Material',
        serie_origem        as 'Série (Origem)',
        vlr_total_nf        as 'Vlr. Total (NF)',
        emissao             as 'Emissão',
        chave_acesso        as 'Chave de Acesso'
    FROM notas_itens
"""
df = pd.read_sql_query(query, conn)
conn.close()

if df.empty:
    print("[AVISO] Banco vazio — nenhum dado para exportar.")
    sys.exit(0)

# Filtra NFs já lançadas diretamente aqui (path resolvido localmente, sem depender do módulo importado)
caminho_lancadas_local = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "nfs_lancadas.txt"
)
nfs_lancadas_set = set()
if os.path.exists(caminho_lancadas_local):
    with open(caminho_lancadas_local, "r", encoding="utf-8") as _f:
        for _linha in _f:
            _nf = _linha.strip()
            if _nf:
                nfs_lancadas_set.add(_nf)

if nfs_lancadas_set and "Nº Nota Fiscal" in df.columns:
    antes = len(df)
    df = df[~df["Nº Nota Fiscal"].astype(str).str.strip().isin(nfs_lancadas_set)].copy()
    print(
        f"[FILTRO] {antes - len(df)} linha(s) excluidas por nfs_lancadas.txt ({len(nfs_lancadas_set)} NF(s) registradas)."
    )
else:
    print("[FILTRO] nfs_lancadas.txt vazio ou nao encontrado — sem exclusoes.")

if df.empty:
    print("[INFO] Todas as NFs do banco ja foram lancadas. Nada a exibir.")
    sys.exit(0)

print(f"[INFO] {len(df)} linhas para exibir. Atualizando Excel...")
ok = atualizar_excel_com_espacos(df, ignorar_lancadas=True)  # filtragem ja feita acima
if ok:
    print("[OK] Excel atualizado com sucesso.")
else:
    print("[ERRO] Falha ao atualizar o Excel. Verifique o log.")
