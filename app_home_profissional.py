import base64
import html
import importlib.util
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

import streamlit as st


# =========================================================
# CONFIGURAÇÃO DA APLICAÇÃO
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_FILE = Path(__file__).name

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Central Operacional de Análises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# IDENTIDADE VISUAL
# =========================================================
CORES = {
    "azul_escuro": "#0B1F4D",
    "azul_profundo": "#111D4A",
    "azul_medio": "#123B7A",
    "azul_claro": "#2F80ED",
    "azul_ciano": "#1D8ACB",
    "fundo": "#F4F6F9",
    "card": "#FFFFFF",
    "borda": "#E5E7EB",
    "texto": "#1F2937",
    "texto_suave": "#6B7280",
    "sucesso": "#166534",
    "alerta": "#92400E",
}


# =========================================================
# FUNÇÕES DE ARQUIVO / LOGO
# =========================================================
def encontrar_arquivo(candidatos: list[str]) -> Path | None:
    """Procura um arquivo na pasta do app aceitando variações de nome."""
    for nome in candidatos:
        caminho = BASE_DIR / nome
        if caminho.exists():
            return caminho

    nomes_normalizados = {normalizar_nome_arquivo(p.name): p for p in BASE_DIR.iterdir() if p.is_file()}
    for nome in candidatos:
        chave = normalizar_nome_arquivo(nome)
        if chave in nomes_normalizados:
            return nomes_normalizados[chave]

    return None


def normalizar_nome_arquivo(nome: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", nome.lower())


def encontrar_logo() -> Path | None:
    return encontrar_arquivo(
        [
            "logo_nepomuceno.jpeg",
            "logo_nepomuceno.jpg",
            "logo_nepomuceno.png",
            "Logo Nepomuceno.jpeg",
            "Logo Nepomuceno.jpg",
            "Logo Nepomuceno.png",
            "Expresso Nepomuceno.jpeg",
            "Expresso Nepomuceno.jpg",
            "Expresso Nepomuceno.png",
            "WhatsApp Image 2025-08-12 at 15.22.05.jpeg",
        ]
    )


def imagem_base64(caminho: Path | None) -> str:
    if caminho is None or not caminho.exists():
        return ""

    try:
        return base64.b64encode(caminho.read_bytes()).decode("utf-8")
    except Exception:
        return ""


LOGO_PATH = encontrar_logo()
LOGO_B64 = imagem_base64(LOGO_PATH)

if LOGO_B64:
    LOGO_HTML = (
        f'<img src="data:image/jpeg;base64,{LOGO_B64}" '
        'class="company-logo" alt="Expresso Nepomuceno">'
    )
else:
    LOGO_HTML = '<div class="company-logo-fallback">EN</div>'


# =========================================================
# CSS EXECUTIVO
# =========================================================
st.markdown(
    f"""
<style>
:root {{
    --azul-escuro: {CORES['azul_escuro']};
    --azul-profundo: {CORES['azul_profundo']};
    --azul-medio: {CORES['azul_medio']};
    --azul-claro: {CORES['azul_claro']};
    --azul-ciano: {CORES['azul_ciano']};
    --fundo: {CORES['fundo']};
    --card: {CORES['card']};
    --borda: {CORES['borda']};
    --texto: {CORES['texto']};
    --texto-suave: {CORES['texto_suave']};
}}

html, body, [class*="css"] {{
    font-family: "Segoe UI", Roboto, Arial, sans-serif;
}}

.stApp {{
    background:
        radial-gradient(circle at top left, rgba(47, 128, 237, 0.08), transparent 32%),
        linear-gradient(180deg, #F8FAFC 0%, #EEF3F8 100%);
}}

.block-container {{
    padding-top: 1.1rem;
    padding-bottom: 3rem;
    max-width: 1480px;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header {{visibility: hidden;}}

section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #081A3A 0%, #10182E 100%);
    border-right: 1px solid rgba(255, 255, 255, 0.08);
}}

section[data-testid="stSidebar"] > div {{
    padding-top: 1.3rem;
}}

section[data-testid="stSidebar"] * {{
    color: #E5E7EB;
}}

section[data-testid="stSidebar"] img {{
    border-radius: 14px;
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.24);
    margin-bottom: 0.8rem;
}}

.sidebar-title {{
    font-size: 16px;
    font-weight: 850;
    color: white;
    margin-top: 0.2rem;
}}

.sidebar-subtitle {{
    font-size: 12px;
    line-height: 1.45;
    color: #A8B3C7;
    margin-bottom: 1.1rem;
}}

.sidebar-section {{
    font-size: 11px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-weight: 850;
    color: #93C5FD;
    margin-top: 1rem;
    margin-bottom: 0.65rem;
}}

section[data-testid="stSidebar"] div[role="radiogroup"] label {{
    background: rgba(255, 255, 255, 0.055);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 13px;
    padding: 0.75rem 0.78rem;
    margin-bottom: 0.5rem;
    transition: all 0.15s ease-in-out;
}}

section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
    background: rgba(47, 128, 237, 0.18);
    border-color: rgba(147, 197, 253, 0.45);
    transform: translateX(2px);
}}

section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{
    margin-bottom: 0;
}}

.main-header {{
    background: linear-gradient(135deg, #0B1F4D 0%, #123B7A 55%, #1D8ACB 100%);
    border-radius: 24px;
    padding: 26px 30px;
    border: 1px solid rgba(255, 255, 255, 0.14);
    box-shadow: 0 20px 46px rgba(15, 23, 42, 0.17);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 24px;
    margin-bottom: 22px;
}}

.header-left {{
    display: flex;
    align-items: center;
    gap: 22px;
}}

.company-logo {{
    width: 250px;
    max-width: 28vw;
    height: auto;
    border-radius: 15px;
    object-fit: contain;
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.22);
}}

.company-logo-fallback {{
    width: 98px;
    height: 66px;
    border-radius: 15px;
    background: rgba(255, 255, 255, 0.13);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    font-weight: 900;
    color: white;
}}

.header-eyebrow {{
    font-size: 12px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    font-weight: 850;
    color: #BDE7FF;
    margin-bottom: 7px;
}}

.header-title {{
    color: white;
    font-size: 35px;
    line-height: 1.05;
    font-weight: 900;
    margin: 0;
}}

.header-subtitle {{
    color: #D8E3F1;
    font-size: 15px;
    margin-top: 9px;
    max-width: 800px;
}}

.header-badge {{
    min-width: 158px;
    text-align: center;
    padding: 12px 16px;
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.11);
    border: 1px solid rgba(255, 255, 255, 0.22);
    color: white;
    font-weight: 800;
    font-size: 13px;
}}

.page-title-card {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid var(--borda);
    border-radius: 19px;
    padding: 22px 24px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07);
    margin-bottom: 18px;
}}

.page-title-row {{
    display: flex;
    align-items: center;
    gap: 14px;
}}

.page-icon {{
    width: 48px;
    height: 48px;
    border-radius: 15px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #EAF5FF;
    color: var(--azul-medio);
    font-size: 25px;
    font-weight: 900;
}}

.page-title {{
    color: #1E293B;
    font-size: 30px;
    font-weight: 900;
    margin: 0;
}}

.page-subtitle {{
    color: var(--texto-suave);
    margin: 4px 0 0 0;
    font-size: 15px;
}}

.executive-card {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid var(--borda);
    border-radius: 18px;
    padding: 20px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06);
    margin-bottom: 16px;
}}

.card-title {{
    font-size: 17px;
    color: #1E293B;
    font-weight: 850;
    margin-bottom: 6px;
}}

.card-text {{
    color: var(--texto-suave);
    font-size: 14px;
    line-height: 1.55;
}}

.card-kpi {{
    background: white;
    border: 1px solid var(--borda);
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
}}

.kpi-label {{
    color: var(--texto-suave);
    font-size: 13px;
    font-weight: 700;
    margin-bottom: 5px;
}}

.kpi-value {{
    color: #1E293B;
    font-size: 28px;
    font-weight: 900;
}}

.status-pill,
.info-pill,
.warning-pill {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
}}

.status-pill {{
    background: #ECFDF5;
    border: 1px solid #BBF7D0;
    color: #166534;
}}

.info-pill {{
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    color: #1D4ED8;
}}

.warning-pill {{
    background: #FFF7ED;
    border: 1px solid #FED7AA;
    color: #9A3412;
}}

.message-card-success,
.message-card-warning,
.message-card-info {{
    border-radius: 16px;
    padding: 16px 18px;
    margin-bottom: 16px;
    font-weight: 650;
}}

.message-card-success {{
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    color: #166534;
}}

.message-card-warning {{
    background: #FFF7ED;
    border: 1px solid #FED7AA;
    color: #9A3412;
}}

.message-card-info {{
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    color: #1D4ED8;
}}

.stButton > button {{
    border-radius: 13px !important;
    border: 1px solid #CBD5E1 !important;
    font-weight: 800 !important;
    min-height: 44px;
}}

.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, #123B7A 0%, #1D8ACB 100%) !important;
    color: white !important;
    border: 0 !important;
}}

.stDownloadButton > button {{
    border-radius: 13px !important;
    background: linear-gradient(135deg, #123B7A 0%, #1D8ACB 100%) !important;
    color: white !important;
    border: 0 !important;
    font-weight: 850 !important;
    min-height: 45px;
}}

.stProgress > div > div > div > div {{
    background: linear-gradient(90deg, #123B7A 0%, #1D8ACB 100%);
}}

[data-testid="stMetric"] {{
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 17px;
    padding: 15px 16px;
    box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
}}

[data-testid="stFileUploader"] section {{
    border-radius: 16px;
    border: 1px dashed #94A3B8;
    background: #F8FAFC;
}}

hr {{
    margin: 1.25rem 0;
}}


/* =========================================================
   HOME PAGE PREMIUM
   ========================================================= */
.home-hero {{
    position: relative;
    overflow: hidden;
    background:
        radial-gradient(circle at 88% 18%, rgba(29, 138, 203, 0.34), transparent 24%),
        radial-gradient(circle at 14% 12%, rgba(147, 197, 253, 0.22), transparent 26%),
        linear-gradient(135deg, #07162F 0%, #0B1F4D 44%, #123B7A 100%);
    border-radius: 30px;
    padding: 34px 36px;
    min-height: 318px;
    color: white;
    box-shadow: 0 28px 68px rgba(15, 23, 42, 0.22);
    border: 1px solid rgba(255, 255, 255, 0.16);
    margin-bottom: 24px;
}}

.home-hero::after {{
    content: "";
    position: absolute;
    right: -80px;
    bottom: -110px;
    width: 360px;
    height: 360px;
    border-radius: 999px;
    border: 38px solid rgba(255, 255, 255, 0.055);
}}

.home-hero-grid {{
    position: relative;
    z-index: 1;
    display: grid;
    grid-template-columns: minmax(0, 1.45fr) minmax(310px, 0.55fr);
    gap: 28px;
    align-items: stretch;
}}

.home-eyebrow {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 8px 12px;
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.11);
    border: 1px solid rgba(255, 255, 255, 0.18);
    color: #CFFAFE;
    font-size: 12px;
    letter-spacing: 0.13em;
    text-transform: uppercase;
    font-weight: 900;
    margin-bottom: 18px;
}}

.home-title {{
    margin: 0;
    color: #FFFFFF;
    font-size: clamp(34px, 4.4vw, 56px);
    line-height: 0.98;
    letter-spacing: -0.045em;
    font-weight: 950;
    max-width: 860px;
}}

.home-subtitle {{
    color: #DDEBFF;
    font-size: 17px;
    line-height: 1.62;
    max-width: 850px;
    margin-top: 18px;
    margin-bottom: 0;
}}

.home-actions {{
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin-top: 24px;
}}

.home-action-primary,
.home-action-secondary {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 9px;
    min-height: 44px;
    padding: 0 17px;
    border-radius: 999px;
    font-weight: 900;
    font-size: 13px;
}}

.home-action-primary {{
    background: white;
    color: #0B1F4D;
    box-shadow: 0 12px 26px rgba(15, 23, 42, 0.22);
}}

.home-action-secondary {{
    background: rgba(255, 255, 255, 0.10);
    border: 1px solid rgba(255, 255, 255, 0.20);
    color: white;
}}

.home-side-panel {{
    background: rgba(255, 255, 255, 0.11);
    border: 1px solid rgba(255, 255, 255, 0.16);
    border-radius: 24px;
    padding: 22px;
    backdrop-filter: blur(14px);
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.12);
}}

.home-side-title {{
    font-size: 14px;
    color: #BFDBFE;
    font-weight: 900;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 16px;
}}

.home-side-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    padding: 13px 0;
    border-bottom: 1px solid rgba(255,255,255,0.12);
}}

.home-side-row:last-child {{ border-bottom: 0; }}

.home-side-label {{
    color: #E5EDF9;
    font-size: 13px;
    font-weight: 750;
}}

.home-side-value {{
    color: white;
    font-size: 16px;
    font-weight: 950;
    text-align: right;
}}

.home-section-title {{
    display: flex;
    align-items: end;
    justify-content: space-between;
    gap: 16px;
    margin: 26px 0 14px;
}}

.home-section-title h3 {{
    margin: 0;
    font-size: 24px;
    color: #172033;
    letter-spacing: -0.02em;
    font-weight: 950;
}}

.home-section-title p {{
    margin: 4px 0 0;
    color: #64748B;
    font-size: 14px;
}}

.module-card-premium {{
    position: relative;
    min-height: 215px;
    background: rgba(255, 255, 255, 0.97);
    border: 1px solid rgba(226, 232, 240, 0.96);
    border-radius: 24px;
    padding: 24px;
    box-shadow: 0 18px 42px rgba(15, 23, 42, 0.075);
    overflow: hidden;
    transition: all 0.18s ease-in-out;
}}

.module-card-premium:hover {{
    transform: translateY(-3px);
    box-shadow: 0 24px 54px rgba(15, 23, 42, 0.11);
    border-color: rgba(29, 138, 203, 0.35);
}}

.module-card-premium::before {{
    content: "";
    position: absolute;
    inset: 0 0 auto 0;
    height: 5px;
    background: linear-gradient(90deg, #0B1F4D 0%, #1D8ACB 100%);
}}

.module-icon-premium {{
    width: 54px;
    height: 54px;
    border-radius: 18px;
    background: linear-gradient(135deg, #EAF5FF 0%, #DCEBFF 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 25px;
    margin-bottom: 18px;
    box-shadow: inset 0 1px 0 white;
}}

.module-title-premium {{
    font-size: 19px;
    font-weight: 950;
    color: #172033;
    margin-bottom: 9px;
    letter-spacing: -0.015em;
}}

.module-text-premium {{
    color: #64748B;
    line-height: 1.58;
    font-size: 14px;
    margin-bottom: 18px;
}}

.module-footer-premium {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    color: #123B7A;
    font-weight: 900;
    font-size: 13px;
}}

.home-kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
    margin-bottom: 18px;
}}

.home-kpi-card {{
    background: rgba(255, 255, 255, 0.98);
    border: 1px solid #E2E8F0;
    border-radius: 22px;
    padding: 19px 20px;
    box-shadow: 0 14px 32px rgba(15, 23, 42, 0.065);
}}

.home-kpi-top {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin-bottom: 10px;
}}

.home-kpi-label {{
    color: #64748B;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 900;
}}

.home-kpi-dot {{
    width: 10px;
    height: 10px;
    border-radius: 999px;
    background: #1D8ACB;
    box-shadow: 0 0 0 5px rgba(29, 138, 203, 0.12);
}}

.home-kpi-value {{
    color: #0F172A;
    font-size: 30px;
    font-weight: 950;
    letter-spacing: -0.03em;
}}

.home-kpi-caption {{
    color: #94A3B8;
    font-size: 13px;
    margin-top: 3px;
}}

.home-workflow {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #E2E8F0;
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 16px 38px rgba(15, 23, 42, 0.065);
}}

.workflow-row {{
    display: grid;
    grid-template-columns: 44px minmax(0, 1fr);
    gap: 14px;
    padding: 14px 0;
    border-bottom: 1px solid #EDF2F7;
}}

.workflow-row:last-child {{ border-bottom: 0; }}

.workflow-number {{
    width: 44px;
    height: 44px;
    border-radius: 16px;
    background: #EFF6FF;
    color: #123B7A;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 950;
}}

.workflow-title {{
    color: #172033;
    font-size: 15px;
    font-weight: 900;
    margin-bottom: 4px;
}}

.workflow-text {{
    color: #64748B;
    font-size: 13px;
    line-height: 1.5;
}}

.system-banner {{
    background: linear-gradient(135deg, #FFFFFF 0%, #F8FBFF 100%);
    border: 1px solid #DDE7F3;
    border-radius: 24px;
    padding: 20px 22px;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.055);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    margin-top: 20px;
}}

.system-banner-title {{
    color: #172033;
    font-size: 16px;
    font-weight: 950;
    margin-bottom: 4px;
}}

.system-banner-text {{
    color: #64748B;
    font-size: 13px;
    line-height: 1.48;
}}

.system-banner-pill {{
    white-space: nowrap;
    color: #166534;
    background: #ECFDF5;
    border: 1px solid #BBF7D0;
    padding: 9px 13px;
    border-radius: 999px;
    font-weight: 900;
    font-size: 13px;
}}

@media (max-width: 1050px) {{
    .home-hero-grid,
    .home-kpi-grid {{
        grid-template-columns: 1fr 1fr;
    }}
}}

@media (max-width: 760px) {{
    .home-hero {{ padding: 26px 22px; }}
    .home-hero-grid,
    .home-kpi-grid {{
        grid-template-columns: 1fr;
    }}
    .system-banner,
    .home-section-title {{
        display: block;
    }}
    .system-banner-pill {{ display: inline-flex; margin-top: 14px; }}
}}

@media (max-width: 900px) {{
    .main-header {{
        display: block;
    }}

    .header-left {{
        display: block;
    }}

    .company-logo {{
        max-width: 100%;
        width: 100%;
        margin-bottom: 18px;
    }}

    .header-title {{
        font-size: 28px;
    }}

    .header-badge {{
        margin-top: 16px;
        width: fit-content;
    }}
}}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# FUNÇÕES GERAIS
# =========================================================
def atualizar_progresso(barra, status, pct: int, texto: str) -> None:
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.20)


def script_deve_ser_ignorado(caminho: Path) -> bool:
    nome = caminho.name
    nome_lower = nome.lower()

    ignorar_exatos = {
        APP_FILE.lower(),
        "app.py",
        "app_executivo.py",
        "app_executivo_final.py",
        "__init__.py",
    }

    if nome_lower in ignorar_exatos:
        return True

    # Evita que arquivos criados acidentalmente com comandos virem módulos do sistema.
    prefixos_invalidos = (
        "streamlit ",
        "pip ",
        "py ",
        "python ",
    )

    if nome_lower.startswith(prefixos_invalidos):
        return True

    return False


def listar_scripts_python() -> list[str]:
    arquivos: list[str] = []

    for arq in BASE_DIR.iterdir():
        if not arq.is_file():
            continue

        if arq.suffix.lower() != ".py":
            continue

        if script_deve_ser_ignorado(arq):
            continue

        arquivos.append(arq.name)

    return sorted(arquivos, key=lambda x: x.lower())


def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".py", "")
    nome = nome.replace("_", " ").replace("-", " ")
    nome = " ".join(parte.capitalize() for parte in nome.split())
    return nome


def nome_modulo_seguro(caminho_script: Path) -> str:
    base = re.sub(r"\W+", "_", caminho_script.stem, flags=re.UNICODE).strip("_")
    if not base:
        base = "modulo"
    return f"modulo_{base}_{abs(hash(str(caminho_script))) % 10_000_000}"


def set_page_config_noop(*args, **kwargs) -> None:
    """Evita erro quando um módulo interno também chama st.set_page_config()."""
    return None


def carregar_modulo_por_arquivo(caminho_script: Path):
    nome_modulo = nome_modulo_seguro(caminho_script)
    spec = importlib.util.spec_from_file_location(nome_modulo, caminho_script)

    if spec is None or spec.loader is None:
        raise ImportError(f"Não foi possível carregar o módulo: {caminho_script.name}")

    modulo = importlib.util.module_from_spec(spec)

    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop
        spec.loader.exec_module(modulo)
    finally:
        st.set_page_config = original_set_page_config

    return modulo


def executar_modulo_streamlit(modulo) -> None:
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop

        if hasattr(modulo, "main_streamlit"):
            modulo.main_streamlit()
        elif hasattr(modulo, "main"):
            modulo.main()
        else:
            st.warning("Este módulo não possui função main_streamlit() nem main().")
    finally:
        st.set_page_config = original_set_page_config


def validar_funcoes_modulo(modulo, funcoes_obrigatorias: list[str]) -> list[str]:
    return [funcao for funcao in funcoes_obrigatorias if not hasattr(modulo, funcao)]


# =========================================================
# COMPONENTES VISUAIS
# =========================================================
def render_header() -> None:
    data_atual = datetime.now().strftime("%d/%m/%Y")
    st.markdown(
        f"""
        <div class="main-header">
            <div class="header-left">
                {LOGO_HTML}
                <div>
                    <div class="header-eyebrow">Inteligência operacional</div>
                    <h1 class="header-title">Central Operacional de Análises</h1>
                    <div class="header-subtitle">
                        Portal executivo para processamento de bases, indicadores operacionais e geração de arquivos tratados.
                    </div>
                </div>
            </div>
            <div class="header-badge">Ambiente interno<br>{data_atual}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_page_title(icone: str, titulo: str, subtitulo: str) -> None:
    st.markdown(
        f"""
        <div class="page-title-card">
            <div class="page-title-row">
                <div class="page-icon">{html.escape(icone)}</div>
                <div>
                    <h2 class="page-title">{html.escape(titulo)}</h2>
                    <p class="page-subtitle">{html.escape(subtitulo)}</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_card(titulo: str, texto: str) -> None:
    st.markdown(
        f"""
        <div class="executive-card">
            <div class="card-title">{html.escape(titulo)}</div>
            <div class="card-text">{html.escape(texto)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="card-kpi">
            <div class="kpi-label">{html.escape(label)}</div>
            <div class="kpi-value">{html.escape(value)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_message(tipo: str, texto: str) -> None:
    classe = {
        "success": "message-card-success",
        "warning": "message-card-warning",
        "info": "message-card-info",
    }.get(tipo, "message-card-info")

    st.markdown(
        f'<div class="{classe}">{texto}</div>',
        unsafe_allow_html=True,
    )


def render_sidebar(menu: list[str]) -> str:
    with st.sidebar:
        if LOGO_PATH and LOGO_PATH.exists():
            st.image(str(LOGO_PATH), use_container_width=True)
        else:
            st.markdown("### Expresso Nepomuceno")

        st.markdown(
            """
            <div class="sidebar-title">Central de Análises</div>
            <div class="sidebar-subtitle">
                Processamento de bases e relatórios operacionais.
            </div>
            <div class="sidebar-section">Navegação</div>
            """,
            unsafe_allow_html=True,
        )

        if "pagina_atual" not in st.session_state:
            st.session_state.pagina_atual = menu[0]

        if st.session_state.pagina_atual not in menu:
            st.session_state.pagina_atual = menu[0]

        pagina = st.radio(
            "Selecione uma área",
            menu,
            index=menu.index(st.session_state.pagina_atual),
            label_visibility="collapsed",
            key="menu_principal",
        )

        st.session_state.pagina_atual = pagina

        st.markdown("---")
        st.caption("Expresso Nepomuceno SA")
        st.caption(f"Pasta: {BASE_DIR.name}")
        st.caption("Versão visual: executiva")

    return pagina


# =========================================================
# PÁGINA INICIAL
# =========================================================
def pagina_inicio(scripts_detectados: list[str]) -> None:
    total_modulos = len(scripts_detectados) + 2
    total_dinamicos = len(scripts_detectados)
    data_atual = datetime.now().strftime("%d/%m/%Y")

    st.markdown(
        f"""
        <div class="home-hero">
            <div class="home-hero-grid">
                <div>
                    <div class="home-eyebrow">● Portal corporativo de dados</div>
                    <h1 class="home-title">Operação, indicadores e processamento em um só ambiente.</h1>
                    <p class="home-subtitle">
                        Uma central profissional para executar módulos analíticos, tratar bases operacionais,
                        consolidar informações críticas e entregar arquivos finais com padrão executivo.
                    </p>
                    <div class="home-actions">
                        <div class="home-action-primary">🚀 Iniciar processamento</div>
                        <div class="home-action-secondary">📊 Acompanhar módulos ativos</div>
                    </div>
                </div>
                <div class="home-side-panel">
                    <div class="home-side-title">Resumo do ambiente</div>
                    <div class="home-side-row">
                        <div class="home-side-label">Status operacional</div>
                        <div class="home-side-value">Online</div>
                    </div>
                    <div class="home-side-row">
                        <div class="home-side-label">Módulos disponíveis</div>
                        <div class="home-side-value">{total_modulos}</div>
                    </div>
                    <div class="home-side-row">
                        <div class="home-side-label">Scripts detectados</div>
                        <div class="home-side-value">{total_dinamicos}</div>
                    </div>
                    <div class="home-side-row">
                        <div class="home-side-label">Atualização</div>
                        <div class="home-side-value">{data_atual}</div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="home-kpi-grid">
            <div class="home-kpi-card">
                <div class="home-kpi-top">
                    <div class="home-kpi-label">Módulos ativos</div>
                    <div class="home-kpi-dot"></div>
                </div>
                <div class="home-kpi-value">{total_modulos}</div>
                <div class="home-kpi-caption">Áreas disponíveis no portal</div>
            </div>
            <div class="home-kpi-card">
                <div class="home-kpi-top">
                    <div class="home-kpi-label">Layout</div>
                    <div class="home-kpi-dot"></div>
                </div>
                <div class="home-kpi-value">Pro</div>
                <div class="home-kpi-caption">Interface executiva e responsiva</div>
            </div>
            <div class="home-kpi-card">
                <div class="home-kpi-top">
                    <div class="home-kpi-label">Automação</div>
                    <div class="home-kpi-dot"></div>
                </div>
                <div class="home-kpi-value">ON</div>
                <div class="home-kpi-caption">Detecção automática de scripts</div>
            </div>
            <div class="home-kpi-card">
                <div class="home-kpi-top">
                    <div class="home-kpi-label">Ambiente</div>
                    <div class="home-kpi-dot"></div>
                </div>
                <div class="home-kpi-value">Interno</div>
                <div class="home-kpi-caption">Uso operacional controlado</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="home-section-title">
            <div>
                <h3>Módulos principais</h3>
                <p>Selecione uma área no menu lateral para iniciar o processamento.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            """
            <div class="module-card-premium">
                <div class="module-icon-premium">📊</div>
                <div class="module-title-premium">Análise de Permanência</div>
                <div class="module-text-premium">
                    Processa a base de permanência, identifica eventos, classifica tempos e gera um Excel final
                    pronto para acompanhamento operacional.
                </div>
                <div class="module-footer-premium"><span>Base Excel</span><span>→</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            """
            <div class="module-card-premium">
                <div class="module-icon-premium">🚛</div>
                <div class="module-title-premium">Odômetro V12</div>
                <div class="module-text-premium">
                    Consolida combustível, Maxtrack, ativos e produção oficial para gerar o relatório final
                    com cruzamento das quatro bases obrigatórias.
                </div>
                <div class="module-footer-premium"><span>4 bases obrigatórias</span><span>→</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            """
            <div class="module-card-premium">
                <div class="module-icon-premium">🧩</div>
                <div class="module-title-premium">Módulos Dinâmicos</div>
                <div class="module-text-premium">
                    Scripts Python adicionados na pasta do sistema são detectados automaticamente quando possuem
                    main_streamlit() ou main().
                </div>
                <div class="module-footer-premium"><span>Expansível</span><span>→</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        """
        <div class="home-section-title">
            <div>
                <h3>Fluxo recomendado</h3>
                <p>Organização visual para reduzir erro operacional e padronizar a execução.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_fluxo, col_sistema = st.columns([1.15, 0.85])
    with col_fluxo:
        st.markdown(
            """
            <div class="home-workflow">
                <div class="workflow-row">
                    <div class="workflow-number">01</div>
                    <div>
                        <div class="workflow-title">Escolha o módulo operacional</div>
                        <div class="workflow-text">Use o menu lateral para acessar Permanência, Odômetro ou um módulo dinâmico.</div>
                    </div>
                </div>
                <div class="workflow-row">
                    <div class="workflow-number">02</div>
                    <div>
                        <div class="workflow-title">Envie as bases necessárias</div>
                        <div class="workflow-text">Faça upload dos arquivos solicitados e confira o status antes de processar.</div>
                    </div>
                </div>
                <div class="workflow-row">
                    <div class="workflow-number">03</div>
                    <div>
                        <div class="workflow-title">Gere o arquivo final</div>
                        <div class="workflow-text">Execute o processamento e baixe o Excel consolidado com o padrão do sistema.</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_sistema:
        render_card(
            "Governança visual",
            "A home foi pensada para transmitir ambiente corporativo, clareza operacional e confiança para usuários internos.",
        )
        render_card(
            "Escalável",
            "Novos scripts podem entrar no portal sem redesenhar a aplicação principal, mantendo o padrão visual.",
        )

    if scripts_detectados:
        st.markdown("### Scripts Python detectados")
        dados = [{"Arquivo": s, "Nome no menu": nome_amigavel_script(s)} for s in scripts_detectados]
        st.dataframe(dados, use_container_width=True, hide_index=True)

    st.markdown(
        """
        <div class="system-banner">
            <div>
                <div class="system-banner-title">Central pronta para operação</div>
                <div class="system-banner-text">
                    Para adicionar novos módulos, coloque o arquivo .py na mesma pasta do app.py e inclua uma função
                    main_streamlit() ou main(). O menu será atualizado automaticamente.
                </div>
            </div>
            <div class="system-banner-pill">● Sistema online</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# PÁGINA FIXA 1 - ANÁLISE DE PERMANÊNCIA
# =========================================================
def pagina_permanencia() -> None:
    render_page_title(
        "📊",
        "Análise de Permanência",
        "Processamento da base de permanência com classificação por tempo configurável.",
    )

    caminho_permanencia = encontrar_arquivo(
        [
            "Codigo_colado.py",
            "Código_colado.py",
            "Codigo colado.py",
            "Código colado.py",
        ]
    )

    if caminho_permanencia is None:
        st.error("Arquivo Codigo_colado.py não encontrado na pasta do app.")
        st.stop()

    try:
        permanencia = carregar_modulo_por_arquivo(caminho_permanencia)
    except Exception as e:
        st.error(f"Erro ao importar {caminho_permanencia.name}: {e}")
        st.stop()

    funcoes_faltando = validar_funcoes_modulo(
        permanencia,
        [
            "carregar_dados",
            "identificar_eventos_carregamento",
            "montar_ciclos_carregamento",
            "gerar_resumos",
            "salvar_saida",
        ],
    )

    if funcoes_faltando:
        st.error(
            "O arquivo de permanência foi encontrado, mas não possui as funções esperadas: "
            + ", ".join(funcoes_faltando)
        )
        st.stop()

    st.markdown('<span class="info-pill">⚙️ Parâmetros de tratamento</span>', unsafe_allow_html=True)
    st.write("")

    col_param1, col_param2 = st.columns(2)

    with col_param1:
        tempo_minimo = st.number_input(
            "Tempo mínimo aceitável (minutos)",
            min_value=0,
            value=15,
            step=1,
        )

    with col_param2:
        tempo_maximo = st.number_input(
            "Tempo máximo aceitável (minutos)",
            min_value=1,
            value=55,
            step=1,
        )

    if tempo_maximo <= tempo_minimo:
        st.error("O tempo máximo precisa ser maior que o mínimo.")

    st.markdown('<span class="info-pill">📁 Importação da base</span>', unsafe_allow_html=True)
    st.write("")

    arquivo = st.file_uploader(
        "Selecione o Excel de permanência",
        type=["xlsx", "xls"],
        key="upload_permanencia",
    )

    if not arquivo:
        render_message("warning", "Aguardando upload da base de permanência para iniciar o processamento.")
        return

    render_message("success", f"✅ Arquivo carregado: <b>{html.escape(arquivo.name)}</b>")

    if st.button("🚀 Processar Permanência", use_container_width=True, type="primary"):
        if tempo_maximo <= tempo_minimo:
            st.error("Corrija os tempos antes de processar.")
            st.stop()

        inicio = time.time()
        barra = st.progress(0)
        status = st.empty()

        try:
            permanencia.TEMPO_MINIMO = tempo_minimo
            permanencia.TEMPO_MAXIMO = tempo_maximo

            atualizar_progresso(barra, status, 10, "Preparando arquivo")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(arquivo.read())
                caminho = tmp.name

            atualizar_progresso(barra, status, 30, "Lendo base")
            df_base = permanencia.carregar_dados(caminho)

            atualizar_progresso(barra, status, 50, "Identificando eventos")
            eventos = permanencia.identificar_eventos_carregamento(df_base)

            atualizar_progresso(barra, status, 70, "Montando ciclos")
            df_resultado, df_alertas = permanencia.montar_ciclos_carregamento(eventos)

            atualizar_progresso(barra, status, 85, "Gerando resumos")
            resumo_geral, resumo_up, resumo_eq, ranking = permanencia.gerar_resumos(df_resultado)

            atualizar_progresso(barra, status, 95, "Gerando Excel final")
            saida = permanencia.salvar_saida(
                arquivo_entrada=caminho,
                df_base=df_base,
                eventos_brutos=eventos,
                df_resultado=df_resultado,
                df_alertas=df_alertas,
                resumo_geral=resumo_geral,
                resumo_por_up=resumo_up,
                resumo_por_equipamento=resumo_eq,
                ranking_improcedentes=ranking,
            )

            atualizar_progresso(barra, status, 100, "Finalizado")

            tempo_total = round(time.time() - inicio, 2)
            st.success(f"Processo finalizado em {tempo_total} segundos")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Eventos", len(df_resultado))
            c2.metric("Alertas", 0 if df_alertas.empty else len(df_alertas))
            c3.metric("OK", 0 if df_resultado.empty else int((df_resultado["Status"] == "OK").sum()))
            c4.metric("Improcedentes", 0 if df_resultado.empty else int((df_resultado["Status"] == "IMPROCEDENTE").sum()))

            st.markdown("### Prévia dos resultados")
            if not df_resultado.empty:
                st.dataframe(df_resultado.head(100), use_container_width=True)
            else:
                st.warning("Nenhum resultado gerado.")

            with open(saida, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel Permanência",
                    f,
                    file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Erro ao processar permanência: {e}")


# =========================================================
# PÁGINA FIXA 2 - ODÔMETRO
# =========================================================
def pagina_odometro() -> None:
    render_page_title(
        "🚛",
        "Odômetro V12",
        "Consolidação do ODOMETRO_MATCH com as quatro bases obrigatórias.",
    )

    render_message("info", "Envie as quatro bases obrigatórias para gerar o Excel final do odômetro.")

    col1, col2 = st.columns(2)

    with col1:
        comb = st.file_uploader("1 - Base Combustível", type=["xlsx", "xls"], key="comb")
        ativo = st.file_uploader("3 - Base Ativo de Veículos", type=["xlsx", "xls"], key="ativo")

    with col2:
        maxtrack = st.file_uploader("2 - Base Km Rodado Maxtrack", type=["xlsx", "xls"], key="maxtrack")
        producao = st.file_uploader("4 - Produção Oficial / Cliente", type=["xlsx", "xls"], key="producao")

    qtd_carregados = sum(arquivo is not None for arquivo in [comb, maxtrack, ativo, producao])
    st.progress(qtd_carregados / 4)
    st.caption(f"{qtd_carregados} de 4 arquivos carregados")

    if not (comb and maxtrack and ativo and producao):
        render_message("warning", "Aguardando upload das quatro bases para liberar o processamento.")
        return

    render_message("success", "✅ Todas as bases foram carregadas. O processamento já pode ser iniciado.")

    if st.button("🚀 Processar Odômetro V12", use_container_width=True, type="primary"):
        inicio = time.time()
        barra = st.progress(0)
        status = st.empty()
        log_box = st.empty()

        try:
            atualizar_progresso(barra, status, 10, "Salvando arquivos temporários")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f1:
                f1.write(comb.read())
                p1 = f1.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f2:
                f2.write(maxtrack.read())
                p2 = f2.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f3:
                f3.write(ativo.read())
                p3 = f3.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f4:
                f4.write(producao.read())
                p4 = f4.name

            saida = os.path.join(
                tempfile.gettempdir(),
                f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            )

            script = encontrar_arquivo(["odometro_v12_com_percentual.py"])

            if script is None:
                st.error("Arquivo odometro_v12_com_percentual.py não encontrado.")
                st.stop()

            atualizar_progresso(barra, status, 25, "Executando script do odômetro")

            processo = subprocess.Popen(
                [sys.executable, str(script), p1, p2, p3, p4, saida],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            logs: list[str] = []
            progresso = 25

            while True:
                linha = processo.stdout.readline() if processo.stdout else ""

                if linha:
                    logs.append(linha.strip())
                    log_box.code("\n".join(logs[-20:]))

                if processo.poll() is not None:
                    break

                if progresso < 90:
                    progresso += 1
                    barra.progress(progresso)
                    status.info(f"{progresso}% - Processando odômetro V12...")

                time.sleep(0.4)

            if processo.returncode != 0:
                st.error("O processamento retornou erro.")
                st.stop()

            if not os.path.exists(saida):
                st.error("Arquivo final não foi gerado.")
                st.stop()

            atualizar_progresso(barra, status, 100, "Finalizado")

            tempo_total = round(time.time() - inicio, 2)
            st.success(f"Odômetro finalizado em {tempo_total} segundos")

            with open(saida, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel Odômetro V12",
                    f,
                    file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Erro ao processar odômetro: {e}")


# =========================================================
# MÓDULOS DINÂMICOS
# =========================================================
def pagina_modulo_dinamico(script: str) -> None:
    nome = nome_amigavel_script(script)
    caminho_script = BASE_DIR / script

    render_page_title(
        "🧩",
        nome,
        f"Módulo detectado automaticamente: {script}",
    )

    if not caminho_script.exists():
        st.error(f"Arquivo não encontrado: {script}")
        st.stop()

    try:
        modulo = carregar_modulo_por_arquivo(caminho_script)
    except Exception as e:
        st.error(f"Erro ao carregar módulo {script}: {e}")
        st.stop()

    if hasattr(modulo, "main_streamlit") or hasattr(modulo, "main"):
        render_message("success", "✅ Módulo compatível com o site. A tela abaixo pertence ao script selecionado.")
        executar_modulo_streamlit(modulo)
    else:
        render_message(
            "warning",
            "Para este módulo rodar integrado ao site, o arquivo precisa ter uma função <code>main_streamlit()</code> ou <code>main()</code>.",
        )


# =========================================================
# APLICAÇÃO PRINCIPAL
# =========================================================
def main() -> None:
    scripts_detectados = listar_scripts_python()

    scripts_fixos = {
        "Codigo_colado.py",
        "Código_colado.py",
        "Codigo colado.py",
        "Código colado.py",
        "odometro_v12_com_percentual.py",
    }

    modulos_fixos = [
        "🏠 Início",
        "📊 Análise de Permanência",
        "🚛 Odômetro V12",
    ]

    modulos_dinamicos = [
        f"🧩 {nome_amigavel_script(s)}"
        for s in scripts_detectados
        if s not in scripts_fixos
    ]

    menu = modulos_fixos + modulos_dinamicos

    pagina = render_sidebar(menu)
    render_header()

    mapa_dinamico = {
        f"🧩 {nome_amigavel_script(s)}": s
        for s in scripts_detectados
        if s not in scripts_fixos
    }

    if pagina == "🏠 Início":
        pagina_inicio(scripts_detectados)
    elif pagina == "📊 Análise de Permanência":
        pagina_permanencia()
    elif pagina == "🚛 Odômetro V12":
        pagina_odometro()
    elif pagina in mapa_dinamico:
        pagina_modulo_dinamico(mapa_dinamico[pagina])
    else:
        st.error("Página não encontrada.")


if __name__ == "__main__":
    main()
