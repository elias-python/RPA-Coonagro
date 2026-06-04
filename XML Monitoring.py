import win32com.client
import time
import os
import re
import ctypes
import pandas as pd
import pythoncom
import threading
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from copy import copy

# Tenta importar openpyxl para a estética visual do Excel
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows

    FORMATACAO_ATIVA = True
except ImportError:
    FORMATACAO_ATIVA = False
    print("[AVISO] Instale 'openpyxl' para ter a formatação profissional no Excel.")

try:
    import pystray
    from PIL import Image, ImageDraw

    TRAY_ATIVO = True
except ImportError:
    TRAY_ATIVO = False
    print("[AVISO] Instale 'pystray' e 'Pillow' para ativar o ícone de bandeja.")

# --- CONFIGURAÇÕES ---
pasta_trabalho = r"C:\Users\esantan3\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro"
caminho_db = os.path.join(pasta_trabalho, "dados_rpa_coonagro.db")
caminho_excel = os.path.join(pasta_trabalho, "Base_Dados_Coonagro.xlsm")
# NFs já exportadas e deletadas — Python não as re-adiciona à planilha
caminho_nfs_lancadas = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "nfs_lancadas.txt"
)
pasta_xml = os.path.join(
    pasta_trabalho, "XMLs_Recebidos"
)  # pode ser esvaziada sem perda de dados

os.makedirs(pasta_trabalho, exist_ok=True)
os.makedirs(pasta_xml, exist_ok=True)

# ==========================================
# LOG E ESTADO GLOBAL
# ==========================================
_log_path = os.path.join(pasta_trabalho, "rpa_coonagro.log")
_status_global = "Iniciando..."
_nf_count_global = 0
_notas_novas_global = 0  # NFs novas encontradas nesta sessão (para lançamento no SAP)
_icon_ref = None  # referência ao ícone de bandeja para notificações
_aguardando_fechamento_xlsm = (
    False  # evita spam de alerta quando arquivo está bloqueado
)
_proximo_alerta_fechamento_xlsm = 0.0  # epoch para novo alerta após adiamento


def _mostrar_card_fechamento_ou_adiar_5min() -> bool:
    """
    Exibe card de decisão quando o xlsm está bloqueado.
    Retorna True quando o usuário escolhe adiar 5 minutos.
    """
    # Tenta abrir card customizado com botoes explicitos.
    try:
        import tkinter as tk

        escolha = {"adiar": False}

        root = tk.Tk()
        root.withdraw()

        win = tk.Toplevel(root)
        win.title("RPA Coonagro - Atualizacao da Base")
        win.attributes("-topmost", True)
        win.resizable(False, False)

        frame = tk.Frame(win, padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        lbl = tk.Label(
            frame,
            text=(
                "Ha dados novos para atualizar a base XLSM.\n\n"
                "Feche o arquivo agora para concluir a atualizacao\n"
                "ou adie este aviso por 5 minutos."
            ),
            justify="left",
            anchor="w",
        )
        lbl.pack(fill="x", pady=(0, 10))

        btns = tk.Frame(frame)
        btns.pack(fill="x")

        def _fechar_agora():
            escolha["adiar"] = False
            win.destroy()

        def _adiar_5min():
            escolha["adiar"] = True
            win.destroy()

        b1 = tk.Button(btns, text="Fechar agora", width=16, command=_fechar_agora)
        b1.pack(side="left")

        b2 = tk.Button(btns, text="Adiar 5 min", width=16, command=_adiar_5min)
        b2.pack(side="right")

        def _on_close():
            # Fechar no X conta como adiar para evitar bloqueio de UX.
            escolha["adiar"] = True
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)

        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f"+{x}+{y}")

        win.grab_set()
        b1.focus_set()
        root.wait_window(win)
        root.destroy()

        return escolha["adiar"]
    except Exception:
        # Fallback: caixa padrao Sim/Nao.
        msg = (
            "Ha dados novos para atualizar a base XLSM.\n\n"
            "Clique em 'Sim' para fechar o arquivo agora e atualizar.\n"
            "Clique em 'Nao' para adiar este aviso por 5 minutos."
        )
        title = "RPA Coonagro - Atualizacao da Base"
        flags = (
            0x00000004  # MB_YESNO
            | 0x00000030  # MB_ICONWARNING
            | 0x00040000  # MB_TOPMOST
            | 0x00010000  # MB_SETFOREGROUND
        )
        resp = ctypes.windll.user32.MessageBoxW(0, msg, title, flags)
        return resp == 7  # IDNO = adiar


def _tentar_fechar_base_xlsm() -> bool:
    """Tenta fechar automaticamente APENAS o Base_Dados_Coonagro.xlsm alvo."""

    def _norm(path: str) -> str:
        return os.path.normcase(os.path.normpath(os.path.abspath(path)))

    def _norm_loose(path: str) -> str:
        return str(path or "").replace("/", "\\").strip().lower()

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetObject(Class="Excel.Application")
        alvo = _norm(caminho_excel)
        alvo_nome = os.path.basename(alvo)
        alvo_nome_loose = alvo_nome.lower()

        fechado = False

        for wb in excel.Workbooks:
            try:
                nome_wb = os.path.basename(_norm(str(getattr(wb, "Name", ""))))
                full_wb_raw = str(getattr(wb, "FullName", ""))
                caminho_wb = _norm(full_wb_raw) if full_wb_raw else ""
                caminho_wb_loose = _norm_loose(full_wb_raw)

                # Assinaturas permitidas para o caminho da nossa base no ambiente.
                assinaturas = [
                    "rpa - coonagro",
                    "base_dados_coonagro.xlsm",
                ]
                assinatura_ok = all(s in caminho_wb_loose for s in assinaturas)

                # Regra de segurança: fecha somente quando o caminho completo casar.
                if caminho_wb and caminho_wb == alvo:
                    try:
                        wb.Close(SaveChanges=True)
                    except Exception:
                        wb.Close(SaveChanges=False)
                    _log(f"[EXCEL] Fechamento seletivo OK: {caminho_wb}")
                    fechado = True
                    break

                # Regra de contingencia (OneDrive/SharePoint):
                # Mesmo nome + assinaturas de pasta/projeto no FullName.
                if str(nome_wb).lower() == alvo_nome_loose and assinatura_ok:
                    try:
                        wb.Close(SaveChanges=True)
                    except Exception:
                        wb.Close(SaveChanges=False)
                    _log(
                        f"[EXCEL] Fechamento seletivo (match flexivel) OK: {full_wb_raw}"
                    )
                    fechado = True
                    break

                # Mesmo nome em outro caminho: explicitamente ignorado.
                if nome_wb == alvo_nome and caminho_wb != alvo:
                    _log(f"[EXCEL] Ignorado (mesmo nome, outro caminho): {caminho_wb}")
            except Exception:
                continue

        if not fechado:
            _log("[EXCEL] Arquivo alvo não encontrado aberto nesta instância do Excel.")
        return fechado
    except Exception as e:
        _log(f"[AVISO] Falha ao tentar fechar xlsm automaticamente: {e}")
    return False


def _contar_pendentes_lancamento():
    """
    Conta pendências reais para lançamento no SAP com base na planilha:
    grupos (NF vinculada 5124, col K) que ainda estão sem cor na coluna A.
    """
    if not os.path.exists(caminho_excel):
        return 0

    try:
        wb = openpyxl.load_workbook(caminho_excel, data_only=True, keep_vba=True)
        ws = wb["Base"] if "Base" in wb.sheetnames else wb.active
        grupos_pendentes = set()

        for row in ws.iter_rows(min_row=3):
            cel_a = row[0]
            nf_a = str(cel_a.value).strip() if cel_a.value is not None else ""
            nf_vinc = str(row[10].value).strip() if row[10].value is not None else ""
            sem_cor = cel_a.fill is None or cel_a.fill.patternType is None

            if nf_a and nf_vinc and sem_cor:
                grupos_pendentes.add(nf_vinc)

        wb.close()
        return len(grupos_pendentes)
    except Exception:
        # Em caso de bloqueio/erro pontual, mantém o indicador legado da sessão.
        return _notas_novas_global


def _notificar(titulo: str, msg: str):
    """Exibe notificação na bandeja do Windows (requer pystray)."""
    global _icon_ref
    if _icon_ref is not None:
        try:
            _icon_ref.notify(msg, titulo)
        except Exception:
            pass


def _texto_pendentes_tray():
    pend = _contar_pendentes_lancamento()
    if pend > 0:
        return f"{pend} grupo(s) pendente(s) para lançamento"
    return "Nenhum grupo pendente para lançamento"


def _log(msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{ts}] {msg}"
    print(linha)
    try:
        with open(_log_path, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass


# ==========================================
# GESTÃO DO BANCO DE DADOS (SQLITE)
# ==========================================


def inicializar_db():
    conn = sqlite3.connect(caminho_db)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notas_itens (
            chave_acesso TEXT,
            nf TEXT,
            seq INTEGER,
            cod_material TEXT,
            descricao TEXT,
            ordem_producao TEXT,
            qtd REAL,
            un TEXT,
            vlr_unit REAL,
            vlr_total_item REAL,
            cfop TEXT,
            nf_vinculada_5124 TEXT,
            nf_origem TEXT,
            serie_origem TEXT,
            vlr_total_nf REAL,
            emissao TEXT,
            data_importacao TEXT,
            PRIMARY KEY (chave_acesso, seq)
        )
    """)
    # Migra banco existente (adiciona coluna se ainda nao existir)
    try:
        conn.execute("ALTER TABLE notas_itens ADD COLUMN data_importacao TEXT")
        conn.commit()
    except Exception:
        pass  # coluna ja existe
    conn.commit()
    conn.close()


# ==========================================
# EXTRAÇÃO E FORMATAÇÃO (ESTILO MOSAIC)
# ==========================================


def safe_float(element):
    """Proteção contra tags vazias ou ausentes no XML"""
    try:
        return float(element.text) if element is not None and element.text else 0.0
    except ValueError:
        return 0.0


def extrair_detalhes_xml(caminho_arquivo):
    itens = []
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()

        # Limpeza de Namespace
        for elem in root.iter():
            if "}" in elem.tag:
                elem.tag = elem.tag.split("}", 1)[1]

        n_nf = root.find(".//nNF").text if root.find(".//nNF") is not None else "N/A"
        chave = (
            root.find(".//chNFe").text if root.find(".//chNFe") is not None else "N/A"
        )

        # Data de emissão
        data_emi_tag = root.find(".//dhEmi")
        data_emissao = (
            data_emi_tag.text[:10]
            if data_emi_tag is not None and data_emi_tag.text
            else "N/A"
        )

        v_total_nf = safe_float(root.find(".//vNF"))

        # Extração de Referências (Mãe e Sequenciais)
        ref_mae = "N/A"
        refs_origem = []
        inf_cpl = root.find(".//infCpl")
        if inf_cpl is not None and inf_cpl.text:
            m_mae = re.search(r"Ref\. NF-e (\d+)", inf_cpl.text, re.IGNORECASE)
            if m_mae:
                ref_mae = str(int(m_mae.group(1)))
            m_orig = re.findall(
                r"Nota de referência\s+(\d+)\s*-\s*(\d+)", inf_cpl.text, re.IGNORECASE
            )
            if m_orig:
                refs_origem = [(str(int(n)), s) for n, s in m_orig]

        # Loop pelos itens da nota
        for idx, item in enumerate(root.findall(".//det"), start=1):
            prod = item.find("prod")
            if prod is not None:
                cfop = (
                    prod.find("CFOP").text if prod.find("CFOP") is not None else "N/A"
                )

                # Filtra: somente CFOP 5902 (materiais) e 5124 (NF mãe)
                if cfop not in ("5902", "5124"):
                    continue
                ref_item = (
                    refs_origem[idx - 1][0]
                    if (cfop == "5902" and idx - 1 < len(refs_origem))
                    else "N/A"
                )
                ser_item = (
                    str(
                        int(refs_origem[idx - 1][1])
                    )  # remove zeros a esquerda (ex: "013" -> "13")
                    if (
                        cfop == "5902"
                        and idx - 1 < len(refs_origem)
                        and refs_origem[idx - 1][1].isdigit()
                    )
                    else (
                        refs_origem[idx - 1][1]
                        if (cfop == "5902" and idx - 1 < len(refs_origem))
                        else "N/A"
                    )
                )

                itens.append(
                    (
                        chave,
                        n_nf,
                        idx,
                        (
                            prod.find("cProd").text
                            if prod.find("cProd") is not None
                            else "N/A"
                        ),
                        (
                            prod.find("xProd").text
                            if prod.find("xProd") is not None
                            else "N/A"
                        ),
                        (
                            prod.find("xPed").text
                            if prod.find("xPed") is not None
                            else "N/A"
                        ),
                        safe_float(prod.find("qCom")),
                        (
                            prod.find("uCom").text
                            if prod.find("uCom") is not None
                            else "N/A"
                        ),
                        safe_float(prod.find("vUnCom")),
                        safe_float(prod.find("vProd")),
                        cfop,
                        ref_mae if cfop == "5902" else "N/A",
                        ref_item,
                        ser_item,
                        v_total_nf,
                        data_emissao,
                    )
                )
        return itens
    except Exception as e:
        _log(f"[ERRO EXTRAÇÃO] Falha no XML {os.path.basename(caminho_arquivo)}: {e}")
        return []


def _tentar_substituir_excel(tmp_path, caminho_destino):
    """Move tmp_path para caminho_destino com até 3 tentativas (5s entre elas).
    Se ainda bloqueado, descarta o temp e loga aviso — será tentado no próximo ciclo."""
    import shutil

    global _aguardando_fechamento_xlsm, _status_global, _proximo_alerta_fechamento_xlsm

    if not os.path.exists(tmp_path):
        _log(f"[ERRO] Arquivo temporário não encontrado: {tmp_path}")
        return False

    if not zipfile.is_zipfile(tmp_path):
        _log(
            f"[ERRO] Arquivo temporário inválido (não é um XLSX/XLSM válido): {tmp_path}"
        )
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        return False

    if os.path.exists(caminho_destino):
        destino_tem_vba = _tem_projeto_vba(caminho_destino)
        tmp_tem_vba = _tem_projeto_vba(tmp_path)
        if destino_tem_vba and not tmp_tem_vba:
            _log(
                "[ERRO] Temp sem VBA enquanto destino possui macros. Atualização cancelada para evitar perda de macros."
            )
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return False

    for tentativa in range(3):
        try:
            if os.path.exists(caminho_destino):
                base_dir = os.path.dirname(caminho_destino)
                nome_base = os.path.splitext(os.path.basename(caminho_destino))[0]
                stamp = time.strftime("%Y%m%d_%H%M%S")
                caminho_backup = os.path.join(
                    base_dir, f"{nome_base}_backup_{stamp}.xlsm"
                )
                shutil.copy2(caminho_destino, caminho_backup)
                _log(f"[EXCEL] Backup criado: {caminho_backup}")
                os.remove(caminho_destino)
            shutil.move(tmp_path, caminho_destino)
            _log("[EXCEL] Relatório atualizado com sucesso.")
            # Se antes estava bloqueado e o usuário fechou, reabre já atualizado.
            if _aguardando_fechamento_xlsm:
                _notificar(
                    "RPA Coonagro — Base Atualizada",
                    "Base atualizada com sucesso. Reabrindo o arquivo atualizado...",
                )
                try:
                    os.startfile(caminho_destino)
                    _log("[EXCEL] Arquivo reaberto automaticamente após atualização.")
                except Exception as e:
                    _log(
                        f"[AVISO] Não foi possível reabrir o xlsm automaticamente: {e}"
                    )
                _aguardando_fechamento_xlsm = False
                _proximo_alerta_fechamento_xlsm = 0.0
                _status_global = "Base atualizada"
            return True
        except PermissionError:
            agora = time.time()
            precisa_alertar = (not _aguardando_fechamento_xlsm) or (
                _proximo_alerta_fechamento_xlsm > 0
                and agora >= _proximo_alerta_fechamento_xlsm
            )

            if precisa_alertar and agora >= _proximo_alerta_fechamento_xlsm:
                _aguardando_fechamento_xlsm = True
                _status_global = "Feche o xlsm para atualizar a base"
                _notificar(
                    "RPA Coonagro — Feche a Base XLSM",
                    "Novos dados chegaram. Feche 'Base_Dados_Coonagro.xlsm' para atualizar a base.",
                )

                if _mostrar_card_fechamento_ou_adiar_5min():
                    _proximo_alerta_fechamento_xlsm = agora + 300
                    _status_global = "Atualizacao adiada por 5 minutos"
                    _log("[AVISO] Atualizacao adiada pelo usuario por 5 minutos.")
                else:
                    _proximo_alerta_fechamento_xlsm = 0.0
                    if _tentar_fechar_base_xlsm():
                        _status_global = "Fechando base para atualizar"
                    else:
                        _notificar(
                            "RPA Coonagro — Fechamento Manual Necessario",
                            "Nao consegui fechar a base automaticamente. Feche o arquivo manualmente para atualizar.",
                        )
                    _log(
                        "[AVISO] Usuario optou por fechar o xlsm para atualizar agora."
                    )

            if tentativa < 2:
                _log(
                    f"[AVISO] Excel bloqueado. Nova tentativa em 5s ({tentativa + 2}/3)..."
                )
                time.sleep(5)

    # Esgotou as tentativas — descarta o temp e avisa (próximo ciclo tenta de novo)
    _log("[AVISO] Excel em uso. Feche o arquivo para receber a atualização.")
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass
    return False


def _detectar_linha_cabecalho(ws, max_linhas=5):
    """Detecta em qual linha está o cabeçalho da Base. Retorna None se não encontrar."""
    limite = min(max_linhas, ws.max_row)
    for r in range(1, limite + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if "Nº Nota Fiscal" in vals and "Seq." in vals:
            return r
    return None


def _mapa_cabecalho(ws, header_row):
    return {
        ws.cell(header_row, c).value: c
        for c in range(1, ws.max_column + 1)
        if ws.cell(header_row, c).value is not None
    }


def _tem_projeto_vba(path_arquivo):
    """Retorna True se o pacote Office contém projeto VBA."""
    try:
        with zipfile.ZipFile(path_arquivo, "r") as zf:
            return "xl/vbaProject.bin" in zf.namelist()
    except Exception:
        return False


def _ler_nfs_lancadas() -> set:
    """Retorna conjunto de NF numbers já exportadas e deletadas da planilha."""
    nfs: set = set()
    if os.path.exists(caminho_nfs_lancadas):
        try:
            with open(caminho_nfs_lancadas, "r", encoding="utf-8") as f:
                for linha in f:
                    nf = linha.strip()
                    if nf:
                        nfs.add(nf)
        except Exception:
            pass
    return nfs


def _wb_aberto_no_excel():
    """
    Verifica se Base_Dados_Coonagro.xlsm está aberto no Excel.
    Retorna o objeto COM do Workbook se encontrado, ou None.
    """
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetObject(Class="Excel.Application")
        alvo = os.path.basename(caminho_excel).lower()
        for wb in excel.Workbooks:
            try:
                nome = os.path.basename(str(getattr(wb, "Name", ""))).lower()
                if nome == alvo:
                    return wb
            except Exception:
                continue
    except Exception:
        pass
    return None


def _atualizar_via_com(wb_com, df_out):
    """
    Atualiza dados da aba Base usando Excel COM enquanto o arquivo está aberto.
    Preserva 100%: macros, botões, shapes, personalização, formatação.
    Retorna True se bem-sucedido.
    """
    try:
        ws_com = wb_com.Sheets("Base")

        # Detecta linha de cabeçalho (procura "Nº Nota Fiscal")
        header_row = 2
        for r in range(1, 6):
            for c in range(1, 30):
                try:
                    val = ws_com.Cells(r, c).Value
                    if val == "Nº Nota Fiscal":
                        header_row = r
                        raise StopIteration
                except StopIteration:
                    raise
                except Exception:
                    pass

    except StopIteration:
        pass
    except Exception as e:
        _log(f"[COM] Erro ao acessar aba Base: {e}")
        return False

    try:
        data_start_row = header_row + 1

        # Calcula área usada real
        used = ws_com.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        last_col = max(used.Column + used.Columns.Count - 1, len(df_out.columns))

        # Limpa somente os dados (mantém shapes/formatação)
        if last_row >= data_start_row:
            ws_com.Range(
                ws_com.Cells(data_start_row, 1),
                ws_com.Cells(last_row, last_col),
            ).ClearContents()

        # Escreve cabeçalho e detecta coluna da Chave de Acesso
        col_chave = None
        for idx, col_name in enumerate(df_out.columns, start=1):
            ws_com.Cells(header_row, idx).Value = col_name
            if col_name == "Chave de Acesso":
                col_chave = idx

        # Escreve dados em bloco (muito mais rápido que célula a célula)
        rows_data = []
        for _, serie in df_out.iterrows():
            row_vals = []
            for val in serie.tolist():
                if val is None:
                    row_vals.append("")
                elif isinstance(val, float) and pd.isna(val):
                    row_vals.append("")
                else:
                    row_vals.append(val)
            rows_data.append(row_vals)

        if rows_data:
            n_rows = len(rows_data)
            n_cols = len(rows_data[0])

            # Formata coluna Chave de Acesso como texto ANTES de escrever
            # (evita que o Excel converta para notação científica)
            if col_chave:
                ws_com.Range(
                    ws_com.Cells(data_start_row, col_chave),
                    ws_com.Cells(data_start_row + n_rows - 1, col_chave),
                ).NumberFormat = "@"

            ws_com.Range(
                ws_com.Cells(data_start_row, 1),
                ws_com.Cells(data_start_row + n_rows - 1, n_cols),
            ).Value = rows_data

        wb_com.Save()
        _log(
            "[COM] Dados atualizados via Excel COM — "
            "macros, botões e personalização 100% preservados."
        )
        return True

    except Exception as e:
        _log(f"[COM] Falha ao gravar dados via COM: {e}")
        return False


def atualizar_excel_com_espacos(df, ignorar_lancadas=False):
    """Atualiza a aba Base preservando macros, botões e personalização.

    ignorar_lancadas=True: ignora o filtro de nfs_lancadas.txt e traz todos os
    registros do df (usado pelo botão Recarregar para mostrar tudo que está no banco).
    """
    if df.empty:
        return False

    # Remove zeros à esquerda da Série (Origem).
    def normalizar_serie(val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return val

    if "Série (Origem)" in df.columns:
        df["Série (Origem)"] = df["Série (Origem)"].apply(normalizar_serie)

    # Exclui NFs já lançadas e removidas da planilha (evita que voltem após exportação)
    # Quando ignorar_lancadas=True (Recarregar completo), este filtro é ignorado.
    if not ignorar_lancadas:
        _nfs_lancadas = _ler_nfs_lancadas()
        if _nfs_lancadas and "Nº Nota Fiscal" in df.columns:
            df = df[
                ~df["Nº Nota Fiscal"].astype(str).str.strip().isin(_nfs_lancadas)
            ].copy()
            if df.empty:
                _log(
                    "[INFO] Todas as NFs do ciclo atual já foram lançadas. Nada a atualizar."
                )
                return True

    def definir_grupo(row):
        if row["CFOP"] == "5124":
            return row["Nº Nota Fiscal"]
        if row["CFOP"] == "5902" and row["NF Vinculada (5124)"] != "N/A":
            return row["NF Vinculada (5124)"]
        return row["Nº Nota Fiscal"]

    df["Grupo"] = df.apply(definir_grupo, axis=1)
    df = df.sort_values(by=["Grupo", "CFOP", "Seq."], ascending=[False, False, True])

    lista_final = []
    prev_g = None
    for _, row in df.iterrows():
        if prev_g is not None and row["Grupo"] != prev_g:
            lista_final.append({c: None for c in df.columns})
        lista_final.append(row.to_dict())
        prev_g = row["Grupo"]

    df_out = pd.DataFrame(lista_final).drop(columns=["Grupo"])

    # ── Caminho primário: COM ──────────────────────────────────────────────────
    # Se o arquivo estiver aberto no Excel, atualiza via COM.
    # Isso preserva 100%: macros, botões, shapes, personalização e formatação.
    # Só cai no caminho openpyxl se o Excel não tiver o arquivo aberto.
    wb_com_aberto = _wb_aberto_no_excel()
    if wb_com_aberto is not None:
        if _atualizar_via_com(wb_com_aberto, df_out):
            return True
        _log("[COM] Caminho COM falhou — tentando via arquivo (openpyxl).")
    # ── Caminho secundário: openpyxl ──────────────────────────────────────────

    if not os.path.exists(caminho_excel):
        _log(
            "[ERRO] Base XLSM não encontrada. Atualização abortada para preservar macros/personalização."
        )
        return False

    if not zipfile.is_zipfile(caminho_excel):
        _log(
            "[ERRO] Base XLSM inválida/corrompida. Atualização abortada para evitar perda de macros."
        )
        return False

    caminho_tmp = os.path.join(os.path.dirname(caminho_excel), "~Base_Dados_temp.xlsm")
    try:
        import shutil

        try:
            shutil.copy2(caminho_excel, caminho_tmp)
        except Exception as e:
            _log(f"[ERRO] Falha ao copiar base XLSM para temp: {e}")
            return False

        wb = openpyxl.load_workbook(caminho_tmp, keep_vba=True)
        if "Base" in wb.sheetnames:
            ws = wb["Base"]
            ws_foi_criada = False
        else:
            ws = wb.create_sheet("Base", 0)
            ws_foi_criada = True

        # Respeita a linha detectada (pode ser 1 se o usuário personalizou lá).
        # Se não encontrar, padrão é linha 2.
        header_row = _detectar_linha_cabecalho(ws) or 2
        data_start_row = header_row + 1

        cores_existentes = {}
        cab_old = _mapa_cabecalho(ws, header_row)
        col_nf_old = cab_old.get("Nº Nota Fiscal")
        col_seq_old = cab_old.get("Seq.")
        if col_nf_old and col_seq_old:
            for row in ws.iter_rows(min_row=data_start_row):
                nf_val = row[col_nf_old - 1].value
                seq_val = row[col_seq_old - 1].value
                cel_a = row[col_nf_old - 1]
                if (
                    nf_val
                    and seq_val
                    and cel_a.fill is not None
                    and cel_a.fill.patternType is not None
                ):
                    cores_existentes[(str(nf_val), str(seq_val))] = (
                        copy(cel_a.fill),
                        copy(cel_a.font),
                    )

        max_col = max(ws.max_column, len(df_out.columns))
        max_row = ws.max_row

        # Limpa somente valores da área de dados, preservando layout e shapes.
        for r in range(data_start_row, max_row + 1):
            for c in range(1, max_col + 1):
                ws.cell(r, c).value = None

        for idx, col_name in enumerate(df_out.columns, start=1):
            ws.cell(header_row, idx).value = col_name

        # Índice da coluna Chave de Acesso para forçar formato texto
        cab_new = _mapa_cabecalho(ws, header_row)
        col_chave_idx = cab_new.get("Chave de Acesso")

        for ridx, (_, serie) in enumerate(df_out.iterrows(), start=data_start_row):
            for cidx, val in enumerate(serie.tolist(), start=1):
                cell = ws.cell(ridx, cidx)
                if cidx == col_chave_idx and val is not None and val != "":
                    cell.value = str(val)
                    cell.number_format = "@"
                else:
                    cell.value = val

        cab_new = _mapa_cabecalho(ws, header_row)
        col_nf_new = cab_new.get("Nº Nota Fiscal")
        col_seq_new = cab_new.get("Seq.")
        if col_nf_new and col_seq_new and cores_existentes:
            for row in ws.iter_rows(min_row=data_start_row, max_row=ws.max_row):
                nf_v = str(row[col_nf_new - 1].value or "")
                sq_v = str(row[col_seq_new - 1].value or "")
                chave = (nf_v, sq_v)
                if chave in cores_existentes:
                    fill_a, font_a = cores_existentes[chave]
                    row[col_nf_new - 1].fill = copy(fill_a)
                    row[col_nf_new - 1].font = copy(font_a)

        if FORMATACAO_ATIVA:
            centro = Alignment(horizontal="center", vertical="center")
            esquerda = Alignment(horizontal="left", vertical="center")

            if ws_foi_criada:
                header_style = PatternFill(
                    start_color="1F4E78", end_color="1F4E78", fill_type="solid"
                )
                for cell in ws[header_row]:
                    cell.fill = header_style
                    cell.font = Font(color="FFFFFF", bold=True)
                    cell.alignment = centro

            # Aplica formato em toda atualização para evitar desalinhamento na 1a linha de dados.
            for row in ws.iter_rows(min_row=data_start_row, max_row=ws.max_row):
                if not row[0].value:
                    continue
                for idx, cell in enumerate(row):
                    col_name = ws.cell(row=header_row, column=idx + 1).value
                    if col_name in ["Qtd."]:
                        cell.number_format = "#,##0.000"
                    elif col_name in [
                        "Vlr. Unitário",
                        "Vlr. Total Item",
                        "Vlr. Total (NF)",
                    ]:
                        cell.number_format = '"R$" #,##0.00'

                    if col_name == "Descrição do Material":
                        cell.alignment = esquerda
                    else:
                        cell.alignment = centro

        wb.save(caminho_tmp)
        wb.close()
        return _tentar_substituir_excel(caminho_tmp, caminho_excel)
    except Exception as e:
        _log(f"[ERRO] Falha ao gerar Excel temporário: {e}")
        try:
            if os.path.exists(caminho_tmp):
                os.remove(caminho_tmp)
        except Exception:
            pass
        return False


# ==========================================
# THREADS DE EXECUÇÃO
# ==========================================


def tarefa_processamento():
    global _status_global, _nf_count_global, _notas_novas_global
    _log("[SISTEMA] Motor de Dados e Relatórios iniciado.")
    inicializar_db()

    # Carrega chaves já no banco para evitar reprocessar XMLs já conhecidos
    conn = sqlite3.connect(caminho_db)
    chaves_processadas = set(
        row[0] for row in conn.execute("SELECT DISTINCT chave_acesso FROM notas_itens")
    )
    conn.close()
    _nf_count_global = len(chaves_processadas)
    _log(f"[SISTEMA] {_nf_count_global} NF(s) já no banco de dados.")
    _status_global = f"Aguardando — {_nf_count_global} NFs no banco"

    # Geração inicial do Excel ao iniciar (preserva status VBA existentes)
    if _nf_count_global > 0:
        _log("[SISTEMA] Regenerando Excel a partir do banco de dados...")
        conn = sqlite3.connect(caminho_db)
        query = """
            SELECT
                nf as 'Nº Nota Fiscal',
                seq as 'Seq.',
                cod_material as 'Código Material',
                descricao as 'Descrição do Material',
                ordem_producao as 'Ordem de Produção',
                qtd as 'Qtd.',
                un as 'UN.',
                vlr_unit as 'Vlr. Unitário',
                vlr_total_item as 'Vlr. Total Item',
                cfop as 'CFOP',
                nf_vinculada_5124 as 'NF Vinculada (5124)',
                nf_origem as 'NF Origem do Material',
                serie_origem as 'Série (Origem)',
                vlr_total_nf as 'Vlr. Total (NF)',
                emissao as 'Emissão',
                chave_acesso as 'Chave de Acesso'
            FROM notas_itens
        """
        df_inicial = pd.read_sql_query(query, conn)
        conn.close()
        ok_inicial = atualizar_excel_com_espacos(df_inicial)
        IfMsg = (
            "[SISTEMA] Excel atualizado na inicialização."
            if ok_inicial
            else "[SISTEMA] Excel pendente (arquivo em uso)."
        )
        _log(IfMsg)
        _status_global = f"Aguardando — {_nf_count_global} NFs no banco"

    while True:
        arquivos = [f for f in os.listdir(pasta_xml) if f.lower().endswith(".xml")]
        novos_inseridos = 0

        if arquivos:
            conn = sqlite3.connect(caminho_db)
            for arq in arquivos:
                dados = extrair_detalhes_xml(os.path.join(pasta_xml, arq))
                if not dados:
                    # XML sem itens CFOP 5902/5124 — ignora silenciosamente
                    _log(f"[IGNORADO] {arq} — sem itens CFOP 5902 ou 5124.")
                    chaves_processadas.add(arq)  # evita reprocessar a cada ciclo
                    continue
                chave = dados[0][0]  # chave_acesso é o primeiro campo
                if chave in chaves_processadas:
                    continue  # já está no banco, ignora
                for item in dados:
                    conn.execute(
                        """
                        INSERT OR REPLACE INTO notas_itens
                            (chave_acesso, nf, seq, cod_material, descricao,
                             ordem_producao, qtd, un, vlr_unit, vlr_total_item,
                             cfop, nf_vinculada_5124, nf_origem, serie_origem,
                             vlr_total_nf, emissao, data_importacao)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        item + (time.strftime("%Y-%m-%d %H:%M:%S"),),
                    )
                chaves_processadas.add(chave)
                novos_inseridos += 1
                _log(f"[PROCESSAMENTO] Nova NF inserida: {arq}")
            conn.commit()

            if novos_inseridos > 0:
                _nf_count_global = len(chaves_processadas)
                _notas_novas_global += novos_inseridos
                _status_global = f"Atualizando Excel — {_nf_count_global} NFs"
                _log(
                    f"[PROCESSAMENTO] {novos_inseridos} NF(s) nova(s). Atualizando Excel..."
                )
                query = """
                    SELECT 
                        nf as 'Nº Nota Fiscal', 
                        seq as 'Seq.', 
                        cod_material as 'Código Material', 
                        descricao as 'Descrição do Material', 
                        ordem_producao as 'Ordem de Produção', 
                        qtd as 'Qtd.', 
                        un as 'UN.', 
                        vlr_unit as 'Vlr. Unitário', 
                        vlr_total_item as 'Vlr. Total Item', 
                        cfop as 'CFOP', 
                        nf_vinculada_5124 as 'NF Vinculada (5124)', 
                        nf_origem as 'NF Origem do Material', 
                        serie_origem as 'Série (Origem)', 
                        vlr_total_nf as 'Vlr. Total (NF)', 
                        emissao as 'Emissão', 
                        chave_acesso as 'Chave de Acesso' 
                    FROM notas_itens
                """
                df_atual = pd.read_sql_query(query, conn)
                conn.close()
                ok_atualizacao = atualizar_excel_com_espacos(df_atual)
                if ok_atualizacao:
                    _notificar(
                        "RPA Coonagro — Base Atualizada",
                        f"{novos_inseridos} NF(s) nova(s) adicionada(s) à base. Total: {_nf_count_global}.",
                    )
                    _status_global = f"Aguardando — {_nf_count_global} NFs no banco"
                else:
                    _status_global = "Aguardando fechamento do xlsm para atualizar"
            else:
                conn.close()

        time.sleep(30)


def _buscar_pasta_outlook(folder, nome, profundidade=0, max_prof=5):
    """Busca recursiva pela pasta com o nome dado."""
    if profundidade > max_prof:
        return None
    try:
        for sub in folder.Folders:
            if sub.Name == nome:
                return sub
            encontrado = _buscar_pasta_outlook(sub, nome, profundidade + 1, max_prof)
            if encontrado:
                return encontrado
    except Exception:
        pass
    return None


def tarefa_outlook():
    _log("[SISTEMA] Coletor Outlook iniciado. Aguardando e-mails...")
    pythoncom.CoInitialize()

    arquivos_salvos = set(
        f.lower() for f in os.listdir(pasta_xml) if f.lower().endswith(".xml")
    )

    while True:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
                "MAPI"
            )
            root_folder = outlook.DefaultStore.GetRootFolder()

            pasta = _buscar_pasta_outlook(root_folder, "XML Coonagro")

            if not pasta:
                _log(
                    "[AVISO] Pasta 'XML Coonagro' não encontrada no Outlook. "
                    "Verifique se a pasta existe e se o Outlook está aberto."
                )
            else:
                itens = pasta.Items
                total = itens.Count
                novos = 0
                for i in range(total, 0, -1):  # reverso para não perder índice
                    try:
                        msg = itens.Item(i)  # Item() usa índice 1-based (padrão COM)
                        # Ignora itens que não são e-mails (reuniões, tarefas, etc.)
                        # MailItem = classe 43
                        if getattr(msg, "Class", None) != 43:
                            continue
                        for anexo in msg.Attachments:
                            nome_arq = anexo.FileName.lower()
                            if (
                                nome_arq.endswith(".xml")
                                and nome_arq not in arquivos_salvos
                            ):
                                destino = os.path.join(pasta_xml, anexo.FileName)
                                anexo.SaveAsFile(destino)
                                arquivos_salvos.add(nome_arq)
                                novos += 1
                                _log(f"[OUTLOOK] XML salvo: {anexo.FileName}")
                        msg.UnRead = False
                    except Exception as e_msg:
                        _log(f"[AVISO] Erro ao processar e-mail #{i}: {e_msg}")

                if novos == 0:
                    _log(f"[OUTLOOK] Verificado ({total} e-mails). Nenhum XML novo.")
                else:
                    _log(f"[OUTLOOK] {novos} XML(s) novo(s) salvo(s).")
                    _notificar(
                        "RPA Coonagro — XMLs Recebidos",
                        f"{novos} XML(s) baixado(s) do Outlook. Processando...",
                    )

        except Exception as e:
            _log(f"[ERRO OUTLOOK] {e}")

        time.sleep(15)


if __name__ == "__main__":
    _log("RPA COONAGRO - VERSÃO PRODUÇÃO iniciada.")

    t1 = threading.Thread(target=tarefa_outlook, daemon=True)
    t2 = threading.Thread(target=tarefa_processamento, daemon=True)
    t1.start()
    t2.start()

    if TRAY_ATIVO:

        def _criar_icone():
            img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.ellipse([2, 2, 62, 62], fill=(26, 120, 56))
            draw.rectangle([18, 14, 46, 50], fill=(255, 255, 255))
            draw.rectangle([22, 20, 42, 24], fill=(26, 120, 56))
            draw.rectangle([22, 29, 42, 33], fill=(26, 120, 56))
            draw.rectangle([22, 38, 42, 42], fill=(26, 120, 56))
            return img

        def _abrir_log(icon, item):
            try:
                os.startfile(_log_path)
            except Exception:
                pass

        def _sair(icon, item):
            _log("Encerrado pelo usuário.")
            icon.stop()

        menu = pystray.Menu(
            pystray.MenuItem(
                lambda item: f"Status: {_status_global}  |  {_nf_count_global} NFs",
                None,
                enabled=False,
            ),
            pystray.MenuItem(
                lambda item: _texto_pendentes_tray(),
                None,
                enabled=False,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Ver Log", _abrir_log),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Encerrar", _sair),
        )
        icon = pystray.Icon(
            "RPA Coonagro",
            _criar_icone(),
            "RPA Coonagro — XML Monitor",
            menu,
        )
        _icon_ref = icon  # permite que as threads enviem notificações
        icon.run()  # bloqueia thread principal — roda em segundo plano
    else:
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            _log("Desligando...")
