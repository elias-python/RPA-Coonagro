import win32com.client
import time
import os
import re
import pandas as pd
import pythoncom
import threading
import sqlite3
import xml.etree.ElementTree as ET

# Tenta importar openpyxl para a estética visual do Excel
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    FORMATACAO_ATIVA = True
except ImportError:
    FORMATACAO_ATIVA = False
    print("[AVISO] Instale 'openpyxl' para ter a formatação profissional no Excel.")

# --- CONFIGURAÇÕES ---
pasta_trabalho = r"C:\Users\esantan3\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro"
caminho_db = os.path.join(pasta_trabalho, "dados_rpa_coonagro.db")
caminho_excel = os.path.join(pasta_trabalho, "Base_Dados_Coonagro.xlsx")

if not os.path.exists(pasta_trabalho):
    os.makedirs(pasta_trabalho)

# ==========================================
# GESTÃO DO BANCO DE DADOS (SQLITE)
# ==========================================

def inicializar_db():
    conn = sqlite3.connect(caminho_db)
    cursor = conn.cursor()
    cursor.execute('''
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
            PRIMARY KEY (chave_acesso, seq)
        )
    ''')
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
            if '}' in elem.tag: elem.tag = elem.tag.split('}', 1)[1]
        
        n_nf = root.find('.//nNF').text if root.find('.//nNF') is not None else "N/A"
        chave = root.find('.//chNFe').text if root.find('.//chNFe') is not None else "N/A"
        
        # Data de emissão
        data_emi_tag = root.find('.//dhEmi')
        data_emissao = data_emi_tag.text[:10] if data_emi_tag is not None and data_emi_tag.text else "N/A"
        
        v_total_nf = safe_float(root.find('.//vNF'))
        
        # Extração de Referências (Mãe e Sequenciais)
        ref_mae = "N/A"
        refs_origem = []
        inf_cpl = root.find('.//infCpl')
        if inf_cpl is not None and inf_cpl.text:
            m_mae = re.search(r'Ref\. NF-e (\d+)', inf_cpl.text, re.IGNORECASE)
            if m_mae: ref_mae = str(int(m_mae.group(1)))
            m_orig = re.findall(r'Nota de referência\s+(\d+)\s*-\s*(\d+)', inf_cpl.text, re.IGNORECASE)
            if m_orig: refs_origem = [(str(int(n)), s) for n, s in m_orig]

        # Loop pelos itens da nota
        for idx, item in enumerate(root.findall('.//det'), start=1):
            prod = item.find('prod')
            if prod is not None:
                cfop = prod.find('CFOP').text if prod.find('CFOP') is not None else "N/A"
                ref_item = refs_origem[idx-1][0] if (cfop == '5902' and idx-1 < len(refs_origem)) else "N/A"
                ser_item = refs_origem[idx-1][1] if (cfop == '5902' and idx-1 < len(refs_origem)) else "N/A"

                itens.append((
                    chave, 
                    n_nf, 
                    idx, 
                    prod.find('cProd').text if prod.find('cProd') is not None else "N/A", 
                    prod.find('xProd').text if prod.find('xProd') is not None else "N/A",
                    prod.find('xPed').text if prod.find('xPed') is not None else "N/A",
                    safe_float(prod.find('qCom')), 
                    prod.find('uCom').text if prod.find('uCom') is not None else "N/A",
                    safe_float(prod.find('vUnCom')), 
                    safe_float(prod.find('vProd')),
                    cfop, 
                    ref_mae if cfop == '5902' else "N/A",
                    ref_item, 
                    ser_item, 
                    v_total_nf, 
                    data_emissao
                ))
        return itens
    except Exception as e:
        print(f"[ERRO EXTRAÇÃO] Falha no XML {os.path.basename(caminho_arquivo)}: {e}")
        return []

def atualizar_excel_com_espacos(df):
    """Gera o Excel com agrupamento inteligente e linha em branco"""
    if df.empty: return
    
    # Lógica de Agrupamento
    def definir_grupo(row):
        if row['CFOP'] == '5124':
            return row['Nº Nota Fiscal']
        elif row['CFOP'] == '5902' and row['NF Vinculada (5124)'] != "N/A":
            return row['NF Vinculada (5124)']
        return row['Nº Nota Fiscal']
        
    df['Grupo'] = df.apply(definir_grupo, axis=1)
    
    # Ordena pelo Grupo, depois CFOP (5902 primeiro), e Sequência
    df = df.sort_values(by=['Grupo', 'CFOP', 'Seq.'], ascending=[False, False, True])

    lista_final = []
    prev_g = None
    for _, row in df.iterrows():
        if prev_g is not None and row['Grupo'] != prev_g:
            lista_final.append({c: None for c in df.columns}) # Linha em branco
        lista_final.append(row.to_dict())
        prev_g = row['Grupo']

    # Remove a coluna temporária de Grupo
    df_out = pd.DataFrame(lista_final).drop(columns=['Grupo'])
    
    try:
        df_out.to_excel(caminho_excel, index=False)
        if FORMATACAO_ATIVA:
            wb = openpyxl.load_workbook(caminho_excel)
            ws = wb.active
            
            # Estilos
            header_style = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            centro = Alignment(horizontal="center", vertical="center")
            esquerda = Alignment(horizontal="left", vertical="center")
            
            # Cabeçalho
            for cell in ws[1]:
                cell.fill = header_style
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = centro
                
            # Dados
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                if not row[0].value: # Se a linha for o espaçamento
                    continue
                for idx, cell in enumerate(row):
                    col_name = ws.cell(row=1, column=idx+1).value
                    if col_name in ["Qtd."]:
                        cell.number_format = '#,##0.000'
                    elif col_name in ["Vlr. Unitário", "Vlr. Total Item", "Vlr. Total (NF)"]:
                        cell.number_format = '"R$" #,##0.00'
                    
                    if col_name == "Descrição do Material":
                        cell.alignment = esquerda
                    else:
                        cell.alignment = centro
            
            # Largura das colunas
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value: max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column].width = min(max_length + 2, 50)
                
            wb.save(caminho_excel)
    except PermissionError:
        print("[AVISO] Excel aberto. Feche para atualizar o relatório.")

# ==========================================
# THREADS DE EXECUÇÃO
# ==========================================

def tarefa_processamento():
    print("[SISTEMA] Motor de Dados e Relatórios iniciado.")
    inicializar_db()
    
    while True:
        arquivos = [f for f in os.listdir(pasta_trabalho) if f.lower().endswith('.xml')]
        if arquivos:
            conn = sqlite3.connect(caminho_db)
            for arq in arquivos:
                dados = extrair_detalhes_xml(os.path.join(pasta_trabalho, arq))
                for item in dados:
                    conn.execute('INSERT OR REPLACE INTO notas_itens VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', item)
            conn.commit()
            
            # Exporta para o DataFrame e depois Excel
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
            
            atualizar_excel_com_espacos(df_atual)
            
        time.sleep(30)

def tarefa_outlook():
    print("[SISTEMA] Coletor Outlook em espera...")
    pythoncom.CoInitialize()
    while True:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            root = outlook.DefaultStore.GetRootFolder()
            pasta = next((f for f in root.Folders if f.Name == "XML Coonagro"), None)
            if not pasta:
                inbox = outlook.GetDefaultFolder(6)
                pasta = next((f for f in inbox.Folders if f.Name == "XML Coonagro"), None)

            if pasta:
                nao_lidas = pasta.Items.Restrict("[Unread] = true")
                for msg in list(nao_lidas):
                    for anexo in msg.Attachments:
                        if anexo.FileName.lower().endswith(".xml"):
                            anexo.SaveAsFile(os.path.join(pasta_trabalho, anexo.FileName))
                    msg.UnRead = False
        except: pass
        time.sleep(15)

if __name__ == "__main__":
    print("="*60 + "\nRPA COONAGRO - VERSÃO PRODUÇÃO (DB SQLite + Excel)\n" + "="*60)
    t1 = threading.Thread(target=tarefa_outlook, daemon=True)
    t2 = threading.Thread(target=tarefa_processamento, daemon=True)
    t1.start()
    t2.start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        print("\nDesligando...")