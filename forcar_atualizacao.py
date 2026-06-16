"""
forcar_atualizacao.py - versao autonoma (sem dependencia do XML Monitoring.py)
Funciona tanto como script Python quanto compilado como .exe via PyInstaller.
Acionado pelo botao Recarregar no Excel.
"""

import os
import sys
import json
import sqlite3
import pandas as pd
import pythoncom
import win32com.client


def _pasta_script() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _auto_detectar_pasta() -> str:
    base = os.path.join(os.environ.get("USERPROFILE", ""), "The Mosaic Company")
    candidato = os.path.join(base, "Controladoria PGA1 (Arquivos) - RPA - Coonagro")
    if os.path.isdir(candidato):
        return candidato
    if os.path.isdir(base):
        for sub in os.listdir(base):
            if "RPA - Coonagro" in sub and os.path.isdir(os.path.join(base, sub)):
                return os.path.join(base, sub)
    return ""


def _ler_config() -> dict:
    cfg = os.path.join(_pasta_script(), "config.json")
    if os.path.exists(cfg):
        try:
            with open(cfg, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[AVISO] Falha ao ler config.json: {e}")
    return {}


def _pasta_runtime_compartilhada() -> str:
    pasta = os.path.join(pasta_trabalho, "system")
    os.makedirs(pasta, exist_ok=True)
    return pasta


def _resolver_arquivo_compartilhado(nome_arquivo: str) -> str:
    caminho_novo = os.path.join(_pasta_runtime_compartilhada(), nome_arquivo)
    caminho_legado = os.path.join(pasta_trabalho, nome_arquivo)

    if os.path.exists(caminho_novo):
        return caminho_novo

    if os.path.exists(caminho_legado):
        try:
            os.replace(caminho_legado, caminho_novo)
            return caminho_novo
        except Exception:
            return caminho_legado

    return caminho_novo


_cfg = _ler_config()
_cfg_pasta = _cfg.get("pasta_trabalho", "")
if "SEU_USUARIO" in _cfg_pasta or not os.path.isdir(_cfg_pasta):
    _cfg_pasta = _auto_detectar_pasta()
pasta_trabalho = (
    _cfg_pasta
    or r"C:\Users\esantan3\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro"
)
caminho_db = _resolver_arquivo_compartilhado("dados_rpa_coonagro.db")
caminho_nfs_lancadas = os.path.join(_pasta_script(), "nfs_lancadas.txt")
caminho_xmls_pendentes_recarregar = os.path.join(
    _pasta_script(), "xmls_pendentes_recarregar.txt"
)
NOME_XLSX_PREFERIDO = "Base_Operacional_Toll_Coonagro.xlsm"
NOME_XLSX_LEGADO = "Base_Dados_Coonagro.xlsm"


def _nomes_xlsx_suportados() -> tuple[str, ...]:
    return (NOME_XLSX_PREFERIDO, NOME_XLSX_LEGADO)


def _resolver_nome_xlsx() -> str:
    for nome in _nomes_xlsx_suportados():
        if os.path.exists(os.path.join(pasta_trabalho, nome)):
            return nome
    return NOME_XLSX_PREFERIDO


NOME_XLSX = _resolver_nome_xlsx()


def _limpar_xmls_pendentes_recarregar():
    try:
        if os.path.exists(caminho_xmls_pendentes_recarregar):
            os.remove(caminho_xmls_pendentes_recarregar)
    except Exception as e:
        print(f"[AVISO] Falha ao limpar pendencia de XMLs: {e}")


EXCEL_UI_PROTECTION_PASSWORD = "RPA_Coonagro_UI"


def _desproteger_base_operacional_com(ws):
    protegido = False
    try:
        protegido = bool(ws.ProtectContents)
    except Exception:
        protegido = False

    if protegido:
        ws.Unprotect(EXCEL_UI_PROTECTION_PASSWORD)

    return protegido


def _proteger_base_operacional_com(ws):
    ws.Protect(
        Password=EXCEL_UI_PROTECTION_PASSWORD,
        DrawingObjects=True,
        Contents=True,
        Scenarios=True,
        UserInterfaceOnly=True,
        AllowFiltering=True,
        AllowSorting=True,
    )


def _valor_excel_limpo(valor, como_texto: bool = False):
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    if isinstance(valor, str):
        texto = valor.strip()
        if texto.lower() in {"nan", "none", "nat", "<na>", "n/a"}:
            return ""
        return texto

    return str(valor) if como_texto else valor


def _atualizar_via_com(df_out: pd.DataFrame) -> bool:
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetObject(Class="Excel.Application")
        wb = None
        alvos = {nome.lower() for nome in _nomes_xlsx_suportados()}
        for w in excel.Workbooks:
            try:
                if os.path.basename(str(w.Name)).lower() in alvos:
                    wb = w
                    break
            except Exception:
                continue
        if wb is None:
            print(f"[COM] {NOME_XLSX} nao esta aberto no Excel.")
            return False
        ws = wb.Sheets("Base")
        header_row = 2
        for r in range(1, 6):
            for c in range(1, 30):
                try:
                    if ws.Cells(r, c).Value == "Nr Nota Fiscal":
                        header_row = r
                        raise StopIteration
                except StopIteration:
                    raise
                except Exception:
                    pass
    except StopIteration:
        pass
    except Exception as e:
        print(f"[COM] Erro ao acessar aba Base: {e}")
        return False

    sucesso = False
    reprotecao_ok = True
    aba_estava_protegida = False

    try:
        aba_estava_protegida = _desproteger_base_operacional_com(ws)
        data_start = header_row + 1
        used = ws.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        last_col = max(used.Column + used.Columns.Count - 1, len(df_out.columns))
        template_last_col = max(last_col, len(df_out.columns))
        template_row_height = None

        # Antes de limpar, preserva o mapa NF → cor da coluna A (status VBA)
        # _norm_nf normaliza 112110 / 112110.0 / "112110" → "112110" (Excel retorna float)
        def _norm_nf(v):
            try:
                return str(int(float(str(v).strip())))
            except Exception:
                return str(v).strip()

        nf_cor_map = {}
        if last_row >= data_start:
            for r in range(data_start, last_row + 1):
                try:
                    cell = ws.Cells(r, 1)
                    nf_val = cell.Value
                    if nf_val is not None and str(nf_val).strip():
                        nf_str = _norm_nf(nf_val)
                        if nf_str not in nf_cor_map:
                            c_idx = cell.Interior.ColorIndex
                            if c_idx is not None and c_idx != -4142:
                                nf_cor_map[nf_str] = (
                                    cell.Interior.Color,
                                    cell.Font.Color,
                                )
                except Exception:
                    pass

        try:
            template_row_height = ws.Rows(data_start).RowHeight
        except Exception:
            template_row_height = None

        if last_row >= data_start:
            ws.Range(
                ws.Cells(data_start, 1), ws.Cells(last_row, last_col)
            ).ClearContents()
            # Limpa cores obsoletas da coluna A (fundo E fonte)
            _col_a = ws.Range(ws.Cells(data_start, 1), ws.Cells(last_row, 1))
            _col_a.Interior.ColorIndex = -4142  # xlColorIndexNone
            _col_a.Font.ColorIndex = -4105  # xlColorIndexAutomatic
        col_chave = None
        for idx, col in enumerate(df_out.columns, start=1):
            ws.Cells(header_row, idx).Value = col
            if col == "Chave de Acesso":
                col_chave = idx
        rows_data = []
        for _, serie in df_out.iterrows():
            row_vals = []
            for idx, val in enumerate(serie.tolist(), start=1):
                row_vals.append(_valor_excel_limpo(val, como_texto=(idx == col_chave)))
            rows_data.append(row_vals)
        if rows_data:
            n_rows, n_cols = len(rows_data), len(rows_data[0])
            format_start_row = max(data_start, last_row + 1)

            try:
                if format_start_row <= data_start + n_rows - 1:
                    ws.Range(
                        ws.Cells(data_start, 1),
                        ws.Cells(data_start, template_last_col),
                    ).Copy()
                    ws.Range(
                        ws.Cells(format_start_row, 1),
                        ws.Cells(data_start + n_rows - 1, n_cols),
                    ).PasteSpecial(
                        Paste=-4122
                    )  # xlPasteFormats
                    if template_row_height is not None:
                        ws.Range(
                            ws.Cells(format_start_row, 1),
                            ws.Cells(data_start + n_rows - 1, 1),
                        ).EntireRow.RowHeight = template_row_height
                    ws.Application.CutCopyMode = False
            except Exception:
                pass

            if col_chave:
                ws.Range(
                    ws.Cells(data_start, col_chave),
                    ws.Cells(data_start + n_rows - 1, col_chave),
                ).NumberFormat = "@"
            ws.Range(
                ws.Cells(data_start, 1), ws.Cells(data_start + n_rows - 1, n_cols)
            ).Value = rows_data

            # Re-aplica cores de status na coluna A
            # Propaga a cor para TODAS as linhas do grupo (igual ao VBA).
            if nf_cor_map:
                _cur_cor = None
                _cur_font = None
                for row_offset, row_vals in enumerate(rows_data):
                    _is_sep = all(v == "" for v in row_vals)
                    if _is_sep:
                        _cur_cor = None
                        _cur_font = None
                    else:
                        nf_val = row_vals[0]
                        if nf_val is not None and str(nf_val).strip():
                            _key = _norm_nf(nf_val)
                            if _key in nf_cor_map:
                                _cur_cor, _cur_font = nf_cor_map[_key]
                    if _cur_cor is not None and not _is_sep:
                        try:
                            cell = ws.Cells(data_start + row_offset, 1)
                            cell.Interior.Color = _cur_cor
                            cell.Font.Color = _cur_font
                        except Exception:
                            pass
        wb.Save()
        sucesso = True
    except Exception as e:
        print(f"[COM] Falha ao gravar: {e}")
    finally:
        if aba_estava_protegida:
            try:
                _proteger_base_operacional_com(ws)
            except Exception as e:
                reprotecao_ok = False
                print(f"[COM] Falha ao reproteger aba Base: {e}")

    if sucesso and reprotecao_ok:
        print("[COM] Excel atualizado com sucesso.")

    return sucesso and reprotecao_ok


def _preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    if "Serie (Origem)" in df.columns:

        def norm(v):
            try:
                return int(v)
            except Exception:
                return v

        df["Serie (Origem)"] = df["Serie (Origem)"].apply(norm)

    def grupo(row):
        if row["CFOP"] == "5124":
            return row["Nr Nota Fiscal"]
        if row["CFOP"] == "5902" and row["NF Vinculada (5124)"] != "N/A":
            return row["NF Vinculada (5124)"]
        return row["Nr Nota Fiscal"]

    df["Grupo"] = df.apply(grupo, axis=1)
    df = df.sort_values(by=["Grupo", "CFOP", "Seq."], ascending=[False, False, True])
    lista = []
    prev = None
    for _, row in df.iterrows():
        if prev is not None and row["Grupo"] != prev:
            lista.append({c: None for c in df.columns})
        lista.append(row.to_dict())
        prev = row["Grupo"]
    df_out = pd.DataFrame(lista).drop(columns=["Grupo"])
    return df_out.where(pd.notna(df_out), None)


print("[RECARREGAR] Iniciando...")

conn_mig = sqlite3.connect(caminho_db)
try:
    conn_mig.execute("ALTER TABLE notas_itens ADD COLUMN data_importacao TEXT")
    conn_mig.commit()
except Exception:
    pass
conn_mig.close()

conn = sqlite3.connect(caminho_db)
sujos = conn.execute(
    "SELECT COUNT(*) FROM notas_itens WHERE cfop NOT IN ('5902','5124')"
).fetchone()[0]
if sujos > 0:
    conn.execute("DELETE FROM notas_itens WHERE cfop NOT IN ('5902','5124')")
    conn.commit()
    print(f"[LIMPEZA] {sujos} registro(s) CFOP invalido removidos.")
conn.close()

conn = sqlite3.connect(caminho_db)
df = pd.read_sql_query(
    """
    SELECT nf as 'Nr Nota Fiscal', seq as 'Seq.', cod_material as 'Codigo Material',
           descricao as 'Descricao do Material', ordem_producao as 'Ordem de Producao',
           qtd as 'Qtd.', un as 'UN.', vlr_unit as 'Vlr. Unitario',
           vlr_total_item as 'Vlr. Total Item', cfop as 'CFOP',
           nf_vinculada_5124 as 'NF Vinculada (5124)', nf_origem as 'NF Origem do Material',
           serie_origem as 'Serie (Origem)', vlr_total_nf as 'Vlr. Total (NF)',
           emissao as 'Emissao',
           CASE
               WHEN lower(trim(COALESCE(chave_acesso, ''))) IN ('nan', 'none', 'nat', '<na>', 'n/a') THEN ''
               ELSE COALESCE(chave_acesso, '')
           END as 'Chave de Acesso'
    FROM notas_itens
""",
    conn,
)
conn.close()

if df.empty:
    print("[AVISO] Banco vazio.")
    _limpar_xmls_pendentes_recarregar()
    sys.exit(0)

nfs_lancadas: set = set()
if os.path.exists(caminho_nfs_lancadas):
    with open(caminho_nfs_lancadas, "r", encoding="utf-8") as f:
        for linha in f:
            nf = linha.strip()
            if nf:
                nfs_lancadas.add(nf)

if nfs_lancadas and "Nr Nota Fiscal" in df.columns:
    antes = len(df)
    df = df[~df["Nr Nota Fiscal"].astype(str).str.strip().isin(nfs_lancadas)].copy()
    print(f"[FILTRO] {antes - len(df)} linha(s) excluidas (nfs_lancadas.txt).")

if df.empty:
    print("[INFO] Todas as NFs ja foram lancadas.")
    _limpar_xmls_pendentes_recarregar()
    sys.exit(0)

print(f"[INFO] {len(df)} linhas. Atualizando Excel...")
ok = _atualizar_via_com(_preparar_df(df))
if not ok:
    print("[ERRO] Nao foi possivel atualizar. Verifique se o Excel esta aberto.")
    sys.exit(1)

_limpar_xmls_pendentes_recarregar()
print("[INFO] Pendencia de XMLs para recarregar limpa.")
