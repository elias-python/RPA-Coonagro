import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Cockpit RPA - Logística",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# CSS PREMIUM - ENTERPRISE DASHBOARD
# =========================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Poppins:wght@500;600;700&display=swap');

:root {
    --bg: #0B0F19;
    --bg-soft: #0f172a;
    --surface: rgba(17, 24, 39, 0.72);
    --surface-strong: rgba(15, 23, 42, 0.92);
    --surface-hover: rgba(30, 41, 59, 0.85);
    --stroke: rgba(148, 163, 184, 0.18);
    --stroke-strong: rgba(148, 163, 184, 0.28);
    --text: #E5E7EB;
    --text-soft: #94A3B8;
    --title: #F8FAFC;
    --primary: #2563EB;
    --primary-2: #38BDF8;
    --success: #22C55E;
    --warning: #F59E0B;
    --danger: #EF4444;
    --shadow: 0 20px 60px rgba(0,0,0,0.35);
    --radius-lg: 24px;
    --radius-md: 18px;
    --radius-sm: 14px;
}

html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(37, 99, 235, 0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(56, 189, 248, 0.10), transparent 24%),
        linear-gradient(180deg, #0B0F19 0%, #0A0F1A 35%, #0B1220 100%);
    color: var(--text);
}

.block-container {
    padding-top: 1.4rem;
    padding-bottom: 2rem;
    max-width: 1600px;
}

/* Remove visual padrão excessivo */
#MainMenu, footer, header {
    visibility: hidden;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background:
        linear-gradient(180deg, rgba(15, 23, 42, 0.98) 0%, rgba(11, 15, 25, 0.98) 100%);
    border-right: 1px solid var(--stroke);
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1rem;
}

.sidebar-shell {
    padding: 0.5rem 0.25rem 1rem 0.25rem;
}

.sidebar-brand {
    background: linear-gradient(135deg, rgba(37,99,235,0.20), rgba(56,189,248,0.10));
    border: 1px solid rgba(56,189,248,0.18);
    border-radius: 20px;
    padding: 18px 18px;
    margin-bottom: 18px;
    box-shadow: var(--shadow);
    backdrop-filter: blur(10px);
}

.sidebar-brand h2 {
    margin: 0;
    color: white;
    font-size: 1.05rem;
    font-weight: 700;
    letter-spacing: -0.02em;
}

.sidebar-brand p {
    margin: 6px 0 0 0;
    color: var(--text-soft);
    font-size: 0.84rem;
}

.sidebar-section-title {
    color: #CBD5E1;
    font-size: 0.74rem;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    margin: 18px 0 10px 4px;
    opacity: 0.85;
    font-weight: 700;
}

.nav-card {
    background: rgba(255,255,255,0.03);
    border: 1px solid var(--stroke);
    border-radius: 16px;
    padding: 12px 14px;
    margin-bottom: 10px;
    transition: all 0.25s ease;
    backdrop-filter: blur(8px);
}

.nav-card:hover {
    transform: translateY(-2px);
    border-color: rgba(56,189,248,0.35);
    background: rgba(255,255,255,0.05);
    box-shadow: 0 12px 24px rgba(0,0,0,0.18);
}

.nav-label {
    color: #E2E8F0;
    font-size: 0.92rem;
    font-weight: 600;
}

.nav-sub {
    color: #94A3B8;
    font-size: 0.78rem;
    margin-top: 2px;
}

.system-mini {
    background: linear-gradient(180deg, rgba(17,24,39,0.92), rgba(15,23,42,0.92));
    border: 1px solid var(--stroke);
    border-radius: 18px;
    padding: 14px;
    margin-top: 18px;
}

.system-mini .row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin: 7px 0;
}

.system-mini .label {
    color: #94A3B8;
    font-size: 0.82rem;
}

.system-mini .value {
    color: #F8FAFC;
    font-size: 0.82rem;
    font-weight: 600;
}

/* Header principal */
.hero {
    position: relative;
    overflow: hidden;
    border-radius: 28px;
    padding: 28px 30px;
    background:
        linear-gradient(135deg, rgba(37,99,235,0.20), rgba(15,23,42,0.88) 38%, rgba(17,24,39,0.94) 100%);
    border: 1px solid rgba(148,163,184,0.18);
    box-shadow: var(--shadow);
    backdrop-filter: blur(16px);
    margin-bottom: 20px;
}

.hero::before {
    content: "";
    position: absolute;
    top: -60px;
    right: -40px;
    width: 240px;
    height: 240px;
    background: radial-gradient(circle, rgba(56,189,248,0.18), transparent 65%);
    pointer-events: none;
}

.hero-topline {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 6px 12px;
    border-radius: 999px;
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.08);
    color: #BFDBFE;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 14px;
}

.hero-title {
    color: var(--title);
    font-family: 'Poppins', 'Inter', sans-serif;
    font-size: 2rem;
    line-height: 1.1;
    font-weight: 700;
    letter-spacing: -0.03em;
    margin: 0;
}

.hero-subtitle {
    color: #A5B4FC;
    font-size: 0.98rem;
    margin-top: 10px;
    margin-bottom: 0;
    max-width: 880px;
}

.status-pill {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    padding: 10px 14px;
    border-radius: 999px;
    background: rgba(34,197,94,0.10);
    color: #DCFCE7;
    border: 1px solid rgba(34,197,94,0.24);
    font-size: 0.86rem;
    font-weight: 700;
    white-space: nowrap;
    box-shadow: 0 0 0 1px rgba(34,197,94,0.04) inset;
}

.status-dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    background: #22C55E;
    box-shadow: 0 0 0 0 rgba(34,197,94,0.7);
    animation: pulse 2s infinite;
}

@keyframes pulse {
    0% { box-shadow: 0 0 0 0 rgba(34,197,94,0.55); }
    70% { box-shadow: 0 0 0 10px rgba(34,197,94,0.0); }
    100% { box-shadow: 0 0 0 0 rgba(34,197,94,0.0); }
}

/* Cards / painéis */
.panel {
    background: var(--surface);
    border: 1px solid var(--stroke);
    border-radius: 24px;
    padding: 20px;
    box-shadow: var(--shadow);
    backdrop-filter: blur(14px);
    transition: all 0.25s ease;
}

.panel:hover {
    border-color: rgba(56,189,248,0.24);
    transform: translateY(-2px);
}

.section-title {
    color: #F8FAFC;
    font-size: 1.05rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    margin-bottom: 4px;
}

.section-subtitle {
    color: #94A3B8;
    font-size: 0.88rem;
    margin-bottom: 16px;
}

/* KPI cards */
.kpi-card {
    position: relative;
    overflow: hidden;
    background:
        linear-gradient(180deg, rgba(17,24,39,0.88), rgba(15,23,42,0.92));
    border: 1px solid rgba(148,163,184,0.16);
    border-radius: 22px;
    padding: 18px 18px 16px 18px;
    min-height: 132px;
    box-shadow: var(--shadow);
    transition: all 0.28s ease;
    backdrop-filter: blur(12px);
}

.kpi-card:hover {
    transform: translateY(-4px);
    border-color: rgba(56,189,248,0.35);
    box-shadow: 0 20px 40px rgba(2, 6, 23, 0.45);
}

.kpi-card::after {
    content: "";
    position: absolute;
    inset: auto -30px -30px auto;
    width: 120px;
    height: 120px;
    background: radial-gradient(circle, rgba(37,99,235,0.18), transparent 60%);
    pointer-events: none;
}

.kpi-top {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 10px;
}

.kpi-icon {
    width: 42px;
    height: 42px;
    border-radius: 14px;
    display: grid;
    place-items: center;
    font-size: 1.1rem;
    background: linear-gradient(135deg, rgba(37,99,235,0.22), rgba(56,189,248,0.12));
    border: 1px solid rgba(56,189,248,0.22);
}

.kpi-title {
    color: #94A3B8;
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.10em;
    font-weight: 700;
    margin-bottom: 12px;
}

.kpi-value {
    color: #F8FAFC;
    font-size: 1.9rem;
    font-weight: 800;
    line-height: 1.1;
    margin-bottom: 6px;
    letter-spacing: -0.03em;
}

.kpi-foot {
    color: #A5B4FC;
    font-size: 0.82rem;
    font-weight: 500;
}

.kpi-success .kpi-value { color: #86EFAC; }
.kpi-warning .kpi-value { color: #FCD34D; }
.kpi-primary .kpi-value { color: #93C5FD; }
.kpi-accent .kpi-value  { color: #67E8F9; }

/* Botões */
.stButton > button {
    width: 100%;
    min-height: 46px;
    border-radius: 14px;
    border: 1px solid rgba(96,165,250,0.26);
    background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 45%, #0EA5E9 100%);
    color: white;
    font-weight: 700;
    font-size: 0.94rem;
    letter-spacing: 0.01em;
    box-shadow: 0 14px 26px rgba(37,99,235,0.24);
    transition: all 0.22s ease;
}

.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 18px 30px rgba(37,99,235,0.34);
    border-color: rgba(125,211,252,0.36);
}

.stButton > button:active {
    transform: translateY(0px) scale(0.99);
}

.action-button-wrap {
    margin-top: 10px;
}

/* Divisor premium */
.premium-divider {
    height: 1px;
    border: none;
    background: linear-gradient(90deg, transparent, rgba(148,163,184,0.35), transparent);
    margin: 18px 0 22px 0;
}

/* Alertas */
[data-testid="stAlert"] {
    border-radius: 16px;
    border: 1px solid rgba(148,163,184,0.18);
    backdrop-filter: blur(12px);
}

/* Expander */
details {
    background: rgba(15,23,42,0.68) !important;
    border: 1px solid rgba(148,163,184,0.14) !important;
    border-radius: 18px !important;
    margin-bottom: 12px !important;
    overflow: hidden;
    transition: all 0.25s ease;
}

details:hover {
    border-color: rgba(56,189,248,0.30) !important;
    transform: translateY(-1px);
}

summary {
    padding: 14px 18px !important;
    background: linear-gradient(180deg, rgba(30,41,59,0.42), rgba(15,23,42,0.36)) !important;
    color: #F8FAFC !important;
    font-weight: 700 !important;
    border-radius: 18px !important;
}

/* DataFrame / tabelas */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(148,163,184,0.16);
    border-radius: 18px;
    overflow: hidden;
    box-shadow: var(--shadow);
    background: rgba(15,23,42,0.70);
}

[data-testid="stDataFrame"] div[role="grid"] {
    border-radius: 18px;
}

[data-testid="stDataFrame"] [role="columnheader"] {
    background: linear-gradient(180deg, rgba(30,41,59,0.96), rgba(15,23,42,0.98)) !important;
    color: #E2E8F0 !important;
    font-weight: 700 !important;
    border-bottom: 1px solid rgba(148,163,184,0.18) !important;
}

[data-testid="stDataFrame"] [role="gridcell"] {
    background: rgba(15,23,42,0.68) !important;
    color: #E5E7EB !important;
    border-bottom: 1px solid rgba(148,163,184,0.08) !important;
}

[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] {
    background: rgba(17,24,39,0.92) !important;
}

[data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"] {
    background: rgba(30,41,59,0.92) !important;
}

/* Texto geral */
h1, h2, h3, h4 {
    color: #F8FAFC;
    letter-spacing: -0.02em;
}

p, span, label, div {
    color: var(--text);
}

/* Espaços */
.small-gap { height: 8px; }
.medium-gap { height: 16px; }

/* Badge utilitário */
.badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 7px 12px;
    border-radius: 999px;
    background: rgba(56,189,248,0.10);
    border: 1px solid rgba(56,189,248,0.22);
    color: #BAE6FD;
    font-size: 0.8rem;
    font-weight: 700;
    margin-top: 6px;
}

/* Responsividade */
@media (max-width: 1200px) {
    .hero-title {
        font-size: 1.6rem;
    }
}

@media (max-width: 768px) {
    .block-container {
        padding-top: 1rem;
    }
    .hero {
        padding: 22px 18px;
    }
    .hero-title {
        font-size: 1.35rem;
    }
    .kpi-card {
        min-height: 118px;
    }
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# CAMINHOS
# =========================================================
pasta_trabalho = r"C:\Users\esantan3\The Mosaic Company\Controladoria PGA1 (Arquivos) - RPA - Coonagro"
caminho_db = os.path.join(pasta_trabalho, "dados_rpa_coonagro.db")

# =========================================================
# FUNÇÃO DE LEITURA
# =========================================================
@st.cache_data(ttl=5)
def carregar_dados():
    if not os.path.exists(caminho_db):
        return pd.DataFrame()

    conn = sqlite3.connect(caminho_db)

    query = """
        SELECT 
            nf as 'Nº Nota Fiscal', 
            cod_material as 'Código Material', 
            descricao as 'Descrição do Material', 
            ordem_producao as 'Ordem de Produção', 
            qtd as 'Qtd.', 
            un as 'UN', 
            cfop as 'CFOP', 
            nf_vinculada_5124 as 'NF Vinculada (5124)', 
            emissao as 'Emissão'
        FROM notas_itens
    """

    df = pd.read_sql_query(query, conn)
    conn.close()

    if not df.empty:
        notas_com_par = df[df['CFOP'] == '5902']['NF Vinculada (5124)'].unique()

        def definir_status(row):
            if row['CFOP'] == '5902' and row['NF Vinculada (5124)'] in df['Nº Nota Fiscal'].values:
                return "🟢 Pronto para Lançamento"
            elif row['CFOP'] == '5124' and row['Nº Nota Fiscal'] in notas_com_par:
                return "🟢 Pronto para Lançamento"
            return "🟡 Em Espera (Falta Par)"

        df['Status'] = df.apply(definir_status, axis=1)

    return df

# =========================================================
# HELPERS VISUAIS
# =========================================================
def render_kpi_card(title, value, icon, footer="", variant="kpi-primary"):
    st.markdown(
        f"""
        <div class="kpi-card {variant}">
            <div class="kpi-top">
                <div>
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-foot">{footer}</div>
                </div>
                <div class="kpi-icon">{icon}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def estilo_tabela(df):
    return (
        df.style
        .set_properties(**{
            "background-color": "rgba(15,23,42,0.80)",
            "color": "#E5E7EB",
            "border-color": "rgba(148,163,184,0.08)",
            "font-size": "0.92rem"
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background", "linear-gradient(180deg, rgba(30,41,59,0.96), rgba(15,23,42,0.98))"),
                    ("color", "#F8FAFC"),
                    ("font-weight", "700"),
                    ("border-bottom", "1px solid rgba(148,163,184,0.16)"),
                    ("text-align", "left"),
                ]
            },
            {
                "selector": "td",
                "props": [
                    ("border-bottom", "1px solid rgba(148,163,184,0.06)")
                ]
            }
        ])
    )

# =========================================================
# SIDEBAR ELEGANTE
# =========================================================
agora = datetime.now().strftime("%d/%m/%Y %H:%M")

with st.sidebar:
    st.markdown("""
        <div class="sidebar-shell">
            <div class="sidebar-brand">
                <h2>⚙️ Cockpit RPA</h2>
                <p>Painel executivo de industrialização logística</p>
            </div>

            <div class="sidebar-section-title">Navegação</div>
            <div class="nav-card">
                <div class="nav-label">Visão Geral Operacional</div>
                <div class="nav-sub">Monitoramento e indicadores em tempo real</div>
            </div>
            <div class="nav-card">
                <div class="nav-label">Fila de Extração</div>
                <div class="nav-sub">Notas aguardando pareamento</div>
            </div>
            <div class="nav-card">
                <div class="nav-label">Lançamentos SAP</div>
                <div class="nav-sub">Execução e conferência de itens liberados</div>
            </div>

            <div class="sidebar-section-title">Status do Sistema</div>
        </div>
    """, unsafe_allow_html=True)

# =========================================================
# DADOS
# =========================================================
df_painel = carregar_dados()

if df_painel.empty:
    status_robo = "Aguardando base"
    status_cor = "warning"
    total_notas = 0
    total_prontas = 0
    total_espera = 0
else:
    status_robo = "Operacional"
    status_cor = "success"
    total_notas = df_painel['Nº Nota Fiscal'].nunique()
    total_prontas = df_painel[df_painel['Status'] == "🟢 Pronto para Lançamento"]['Nº Nota Fiscal'].nunique()
    total_espera = df_painel[df_painel['Status'] == "🟡 Em Espera (Falta Par)"]['Nº Nota Fiscal'].nunique()

with st.sidebar:
    st.markdown(
        f"""
        <div class="system-mini">
            <div class="row">
                <div class="label">Robô</div>
                <div class="value">{status_robo}</div>
            </div>
            <div class="row">
                <div class="label">Última leitura</div>
                <div class="value">{agora}</div>
            </div>
            <div class="row">
                <div class="label">Notas monitoradas</div>
                <div class="value">{total_notas}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# HEADER EXECUTIVO
# =========================================================
header_col1, header_col2 = st.columns([8, 2])

with header_col1:
    st.markdown(
        """
        <div class="hero">
            <div class="hero-topline">Enterprise Control Tower</div>
            <h1 class="hero-title">Centro de Controle de Industrialização</h1>
            <p class="hero-subtitle">
                Painel operacional premium para monitoramento da fila de notas fiscais,
                conferência de pareamentos CFOP e orquestração de lançamentos SAP.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

with header_col2:
    st.markdown("<div class='small-gap'></div>", unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="status-pill">
            <span class="status-dot"></span>
            Robô SAP • {status_robo}
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<div class='small-gap'></div>", unsafe_allow_html=True)

    if st.button("🔄 Atualizar Fila"):
        st.cache_data.clear()
        st.rerun()

st.markdown("<div class='premium-divider'></div>", unsafe_allow_html=True)

# =========================================================
# TRATAMENTO DE DADOS
# =========================================================
if df_painel.empty:
    st.warning("Banco de dados SQLite não encontrado ou vazio.")
else:
    df_espera = df_painel[df_painel['Status'] == "🟡 Em Espera (Falta Par)"]
    df_prontos = df_painel[df_painel['Status'] == "🟢 Pronto para Lançamento"]

    qtd_itens_prontos = len(df_prontos)
    qtd_itens_espera = len(df_espera)
    total_ops = df_prontos['Ordem de Produção'].nunique() if 'Ordem de Produção' in df_prontos.columns else 0

    # =====================================================
    # KPIs
    # =====================================================
    k1, k2, k3, k4 = st.columns(4)

    with k1:
        render_kpi_card(
            "Leitura de XMLs",
            "ATIVA",
            "📡",
            "Pipeline monitorado em tempo real",
            "kpi-success"
        )

    with k2:
        render_kpi_card(
            "Notas Monitoradas",
            total_notas,
            "🧾",
            "Volume total em observação",
            "kpi-primary"
        )

    with k3:
        render_kpi_card(
            "Aguardando Par",
            df_espera['Nº Nota Fiscal'].nunique(),
            "⏳",
            "Pendências para vinculação",
            "kpi-warning"
        )

    with k4:
        render_kpi_card(
            "Prontas para SAP",
            df_prontos['Nº Nota Fiscal'].nunique(),
            "🚀",
            "Notas liberadas para lançamento",
            "kpi-accent"
        )

    st.markdown("<div class='medium-gap'></div>", unsafe_allow_html=True)

    # =====================================================
    # LAYOUT PRINCIPAL
    # =====================================================
    col_esquerda, col_direita = st.columns(2, gap="large")

    # -----------------------------------------------------
    # LADO ESQUERDO
    # -----------------------------------------------------
    with col_esquerda:
        st.markdown("""
            <div class="panel">
                <div class="section-title">📡 Monitor de Extração e Fila</div>
                <div class="section-subtitle">
                    Acompanhamento das notas fiscais pendentes de pareamento e status do pipeline de leitura.
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='small-gap'></div>", unsafe_allow_html=True)

        sub_kpi1, sub_kpi2 = st.columns(2)

        with sub_kpi1:
            render_kpi_card(
                "Itens em Espera",
                qtd_itens_espera,
                "🟡",
                "Itens aguardando o par correspondente",
                "kpi-warning"
            )

        with sub_kpi2:
            render_kpi_card(
                "Operações Vinculadas",
                total_ops,
                "🏭",
                "Ordens de produção identificadas",
                "kpi-primary"
            )

        st.markdown("<div class='small-gap'></div>", unsafe_allow_html=True)

        if not df_espera.empty:
            tabela_espera = (
                df_espera[['Nº Nota Fiscal', 'CFOP', 'Emissão', 'Status']]
                .drop_duplicates()
                .sort_values(by=['Emissão', 'Nº Nota Fiscal'], ascending=[False, True])
            )
            st.dataframe(
                estilo_tabela(tabela_espera),
                use_container_width=True,
                hide_index=True,
                height=420
            )
        else:
            st.info("Nenhuma nota aguardando. Fila limpa!")

    # -----------------------------------------------------
    # LADO DIREITO
    # -----------------------------------------------------
    with col_direita:
        st.markdown("""
            <div class="panel">
                <div class="section-title">🚀 Prontos para Lançamento</div>
                <div class="section-subtitle">
                    Itens liberados para execução no SAP após validação de pareamento entre CFOPs.
                </div>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("<div class='action-button-wrap'></div>", unsafe_allow_html=True)

        if st.button("▶ REALIZAR LANÇAMENTO NO SAP"):
            st.toast("Comando enviado para o Robô SAP!")

        st.markdown(
            """
            <div class="badge">Execução assistida • Fluxo de lançamento controlado</div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<div class='medium-gap'></div>", unsafe_allow_html=True)

        if not df_prontos.empty:
            maes = df_prontos[df_prontos['CFOP'] == '5124']['Nº Nota Fiscal'].unique()

            for m in maes:
                dados_m = df_prontos[df_prontos['Nº Nota Fiscal'] == m].iloc[0]
                op = f" | OP: {dados_m['Ordem de Produção']}" if dados_m['Ordem de Produção'] != "N/A" else ""

                with st.expander(f"📦 NF de Serviço (5124): {m}{op} • Liberado"):
                    filhas = df_prontos[
                        (df_prontos['CFOP'] == '5902') &
                        (df_prontos['NF Vinculada (5124)'] == m)
                    ]

                    tabela_filhas = filhas[['Código Material', 'Descrição do Material', 'Qtd.', 'UN']]

                    st.dataframe(
                        estilo_tabela(tabela_filhas),
                        use_container_width=True,
                        hide_index=True
                    )
        else:
            st.info("Não há pares completos prontos.")
