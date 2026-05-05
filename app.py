import ast
import base64
import html
import importlib.util
import json
import logging
import mimetypes
import os
import re
import runpy
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st

# =========================================================
# CONFIGURACAO
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_FILE = Path(__file__).name
REGISTRY_FILE = BASE_DIR / "modulos_config.json"
HISTORY_FILE = BASE_DIR / "historico_processamentos.json"
LOG_FILE = Path(tempfile.gettempdir()) / "central_operacional.log"

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("central_operacional")

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Central Operacional de Análises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# IDENTIDADE VISUAL - EXPRESSO NEPOMUCENO
# =========================================================
COLORS = {
    "navy": "#061A3A",
    "navy2": "#0A2453",
    "navy3": "#102E68",
    "blue": "#1E5EFF",
    "blue2": "#2F80ED",
    "light_blue": "#EAF2FF",
    "text": "#071B4A",
    "muted": "#667085",
    "border": "#DCE6F5",
    "panel": "#FFFFFF",
    "bg": "#F4F8FF",
    "success": "#16A34A",
    "warning": "#F59E0B",
    "danger": "#DC2626",
}

ADMIN_USER = "pedro admin"
ADMIN_PASSWORD = "admin pedro"

# =========================================================
# UTILITARIOS DE ARQUIVO / LOGO
# =========================================================
def normalizar_nome_arquivo(nome: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", nome.lower())


def encontrar_arquivo(candidatos: list[str]) -> Path | None:
    for nome in candidatos:
        caminho = BASE_DIR / nome
        if caminho.exists():
            return caminho

    try:
        nomes_normalizados = {
            normalizar_nome_arquivo(p.name): p
            for p in BASE_DIR.iterdir()
            if p.is_file()
        }
    except Exception:
        nomes_normalizados = {}

    for nome in candidatos:
        chave = normalizar_nome_arquivo(nome)
        if chave in nomes_normalizados:
            return nomes_normalizados[chave]
    return None


def encontrar_logo() -> Path | None:
    return encontrar_arquivo([
        "logo_nepomuceno.png.jpeg",
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
    ])


@st.cache_data(show_spinner=False)
def imagem_base64_cache(caminho_str: str, mtime: float) -> str:
    caminho = Path(caminho_str)
    return base64.b64encode(caminho.read_bytes()).decode("utf-8")


def imagem_base64(caminho: Path | None) -> str:
    if not caminho or not caminho.exists():
        return ""
    try:
        return imagem_base64_cache(str(caminho), caminho.stat().st_mtime)
    except Exception as exc:
        logger.exception("Falha ao converter imagem para base64: %s", exc)
        return ""


def mime_arquivo(caminho: Path | None) -> str:
    if not caminho:
        return "image/png"
    mime, _ = mimetypes.guess_type(str(caminho))
    return mime or "image/png"


LOGO_PATH = encontrar_logo()
LOGO_B64 = imagem_base64(LOGO_PATH)
LOGO_MIME = mime_arquivo(LOGO_PATH)
LOGO_EMBED_B64 = "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AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigB8fQ/WiiPofrRQA1vvH6mkpW+8fqaSnYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYB8fQ/WinRfdP1oosBG33j9TSUrfeP1NJVAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUASw/dP1ooh+6frRQBG33j9TSUrfeP1NJQIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAlh+6frRRD90/WigCNvvH6mkpzfeP1NJTsSJRS0UWASiloosAlFLRRYBKKWiiwCUUtFFgEopaKLAJRS0UWASiloosAlFLRRYBKKWiiwCUUtFFgEopaKLAJRS0UWASiloosAlFLRRYBKKWiiwCUUtFFgEopaKLAJRS0UWASiloosAlFLRRYBKKWiiwCUUtFFgEopaKLASQ/dP1opYvun60UmAxh8x+ppMU5vvH6mkoKExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAExRilooATFGKWigBMUYpaKAJIR8p+tFLD90/WigCNvvH6mkpW+8fqaSnYYUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCiiiiwBRRRRYAoooosBLD90/WiiH7p+tFJgRt94/U0lK33j9TSVQBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAEsP3T9aKIfun60VLA/9k="
LOGO_EMBED_MIME = "image/jpeg"



def logo_html(css_class: str = "top-logo") -> str:
    """Retorna a logo oficial. Usa arquivo local se existir e fallback embutido se o deploy não levar a imagem."""
    if LOGO_B64:
        return f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="{css_class}" alt="Expresso Nepomuceno">'
    if LOGO_EMBED_B64:
        return f'<img src="data:{LOGO_EMBED_MIME};base64,{LOGO_EMBED_B64}" class="{css_class}" alt="Expresso Nepomuceno">'
    return f'<div class="{css_class} logo-fallback">E.N</div>'


# =========================================================
# HISTORICO / TEMPORARIOS / LOG
# =========================================================
def limpar_arquivo_temporario(caminho: str | Path | None) -> None:
    if not caminho:
        return
    try:
        Path(caminho).unlink(missing_ok=True)
    except Exception as exc:
        logger.warning("Nao foi possivel limpar arquivo temporario %s: %s", caminho, exc)


def carregar_historico() -> list[dict[str, Any]]:
    if not HISTORY_FILE.exists():
        return []
    try:
        dados = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return dados if isinstance(dados, list) else []
    except Exception as exc:
        logger.exception("Falha ao carregar historico: %s", exc)
        return []


def registrar_historico(tipo: str, arquivo: str, status: str, duracao: float, detalhe: str = "") -> None:
    try:
        historico = carregar_historico()
        historico.insert(0, {
            "data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "tipo": tipo,
            "arquivo": arquivo,
            "status": status,
            "duracao_s": round(duracao, 2),
            "detalhe": detalhe,
        })
        HISTORY_FILE.write_text(json.dumps(historico[:300], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.exception("Falha ao registrar historico: %s", exc)


def validar_upload_excel(arquivo, nome_base: str) -> bool:
    if arquivo is None:
        st.error(f"{nome_base}: arquivo não informado.")
        return False
    nome = arquivo.name.lower()
    if not (nome.endswith(".xlsx") or nome.endswith(".xls")):
        st.error(f"{nome_base}: envie um arquivo Excel .xlsx ou .xls.")
        return False
    tamanho = getattr(arquivo, "size", 0) or 0
    if tamanho == 0:
        st.error(f"{nome_base}: arquivo vazio.")
        return False
    return True


def salvar_upload_temporario(arquivo, suffix: str = ".xlsx") -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(arquivo.getbuffer())
        return tmp.name


# =========================================================
# NAVEGACAO
# =========================================================
def pagina_atual_default() -> None:
    if "pagina_atual" not in st.session_state:
        st.session_state["pagina_atual"] = "inicio"


def ir_para(pagina: str) -> None:
    st.session_state["pagina_atual"] = pagina
    st.rerun()


# =========================================================
# CSS - VERSAO APROVADA, BRANCO/AZUL
# =========================================================
def aplicar_css() -> None:
    st.markdown(
        """
<style>
:root {
    --navy: #061A3A;
    --navy2: #0A2453;
    --navy3: #102E68;
    --blue: #1E5EFF;
    --blue2: #2F80ED;
    --bg: #F4F8FF;
    --panel: #FFFFFF;
    --line: #DCE6F5;
    --text: #071B4A;
    --muted: #667085;
    --success: #16A34A;
    --warning: #F59E0B;
    --danger: #DC2626;
}

html, body, [class*="css"], .stApp {
    font-family: "Inter", "Segoe UI", Roboto, Arial, sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 75% 8%, rgba(47,128,237,.13), transparent 22%),
        linear-gradient(135deg, #F8FBFF 0%, #EEF5FF 55%, #FFFFFF 100%);
    color: var(--text);
}

/* Manter o header nativo visivel para o botao < da sidebar funcionar */
header[data-testid="stHeader"] {
    visibility: visible !important;
    background: transparent !important;
    height: 2.5rem !important;
    z-index: 999999 !important;
}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 0.25rem;
    padding-bottom: 3rem;
    max-width: 1420px;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #061A3A 0%, #082453 52%, #061A3A 100%) !important;
    border-right: 1px solid rgba(255,255,255,.12);
}
section[data-testid="stSidebar"] > div {
    padding-top: 2.2rem;
}
section[data-testid="stSidebar"] * {
    color: #EAF2FF;
}
.sidebar-brand {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: .2rem .1rem 1.2rem .1rem;
}
.sidebar-en {
    width: 46px;
    height: 46px;
    border-radius: 999px;
    display:flex;
    align-items:center;
    justify-content:center;
    background: rgba(30,94,255,.16);
    border: 2px solid rgba(80,150,255,.75);
    color:white;
    font-weight:900;
    box-shadow: 0 0 0 6px rgba(30,94,255,.08);
}
.sidebar-title {
    font-size: 15px;
    font-weight: 900;
    line-height: 1.1;
}
.sidebar-subtitle {
    font-size: 11px;
    color: #B8C7E5;
    margin-top: 4px;
}
.sidebar-section {
    margin: 1.1rem 0 .55rem;
    font-size: 10px;
    letter-spacing: .15em;
    text-transform: uppercase;
    color: #83B4FF;
    font-weight: 900;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    min-height: 46px;
    background: rgba(255,255,255,.055);
    border: 1px solid rgba(255,255,255,.08);
    border-radius: 13px;
    padding: .68rem .7rem;
    margin-bottom: .45rem;
    transition: all .18s ease;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background: rgba(30,94,255,.22);
    border-color: rgba(119,174,255,.65);
    transform: translateX(3px);
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
    background: linear-gradient(135deg, rgba(30,94,255,.38), rgba(47,128,237,.24));
    border-color: rgba(119,174,255,.85);
    box-shadow: inset 4px 0 0 #2F80ED;
}
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    margin-bottom: 0;
}

.sidebar-logo-img {
    width: 150px;
    max-width: 100%;
    height: auto;
    object-fit: contain;
    display: block;
}
section[data-testid="stSidebar"] .stButton > button {
    justify-content: flex-start !important;
    text-align: left !important;
    padding-left: 16px !important;
    color: #EAF2FF !important;
    background: rgba(255,255,255,.055) !important;
    border: 1px solid rgba(255,255,255,.10) !important;
    border-radius: 13px !important;
    min-height: 48px !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(30,94,255,.22) !important;
    border-color: rgba(119,174,255,.65) !important;
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, rgba(30,94,255,.78), rgba(47,128,237,.52)) !important;
    border-color: rgba(119,174,255,.90) !important;
    color: #FFFFFF !important;
    box-shadow: inset 4px 0 0 #68B8FF !important;
}
.sidebar-muted {
    color: #89A2CB;
    font-size: 11px;
    margin-top: 1rem;
}

/* Top bar */
.topbar {
    height: 74px;
    border-radius: 0 0 0 0;
    background: linear-gradient(135deg, #061A3A 0%, #0A2E6D 50%, #123B8F 100%);
    color: white;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 32px;
    margin: -0.25rem -1rem 24px -1rem;
    box-shadow: 0 14px 32px rgba(10,36,83,.16);
    position: relative;
    overflow: hidden;
}
.topbar:after {
    content:"";
    position:absolute;
    right:180px; top:-80px;
    width:360px; height:220px;
    background: rgba(30,94,255,.24);
    transform: rotate(38deg);
}
.topbar-left, .topbar-right { position:relative; z-index:2; }
.topbar-left {
    display:flex;
    align-items:center;
    gap:22px;
}
.top-logo {
    width: 175px;
    height: 46px;
    object-fit: contain;
    object-position: left center;
}
.logo-fallback {
    width:52px; height:52px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    background:#0E3B86; color:white; font-weight:900;
}
.top-divider {
    width:1px;
    height:42px;
    background: rgba(255,255,255,.38);
}
.top-system {
    font-size: 14px;
    line-height: 1.3;
    text-transform: uppercase;
    color: #DCEAFF;
    letter-spacing: .02em;
}
.topbar-right {
    display:flex;
    align-items:center;
    gap:17px;
    font-size:13px;
}
.top-icon {
    width:32px; height:32px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    background: rgba(255,255,255,.10);
    border:1px solid rgba(255,255,255,.12);
}
.avatar {
    width:34px; height:34px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    background:#BFD7FF;
    color:#0A2E6D;
    font-weight:900;
}
.userbox { line-height:1.25; }
.userbox strong { display:block; font-size:12px; }
.userbox span { color:#C8D7EF; font-size:11px; }

/* Hero e cards */
.hero {
    position: relative;
    background: linear-gradient(140deg, #FFFFFF 0%, #F9FBFF 58%, #F0F6FF 100%);
    border: 1px solid var(--line);
    border-radius: 18px;
    padding: 36px 42px;
    overflow: hidden;
    min-height: 138px;
    box-shadow: 0 18px 48px rgba(10,36,83,.07);
    margin-bottom: 22px;
}
.hero:after {
    content:"EN";
    position:absolute;
    right:28px;
    top:-10px;
    font-size:112px;
    line-height:1;
    color:#E6ECF6;
    font-weight:900;
    letter-spacing:-8px;
    opacity:.86;
}
.hero h1 {
    position:relative; z-index:2;
    color: var(--text);
    font-size: 28px;
    line-height:1.15;
    margin:0 0 10px 0;
    font-weight: 900;
}
.hero p {
    position:relative; z-index:2;
    margin:0;
    color:#51617C;
    font-size:15px;
}

.module-grid {
    display:grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap:18px;
    margin: 12px 0 16px;
}
.module-card, .wide-card, .system-card, .workflow-card {
    background: rgba(255,255,255,.96);
    border:1px solid var(--line);
    border-radius:16px;
    box-shadow: 0 18px 40px rgba(10,36,83,.075);
    color:var(--text);
}
.module-card {
    padding:22px;
    min-height: 205px;
    position:relative;
}
.module-card .arrow-dot, .wide-card .arrow-dot {
    position:absolute;
    top:20px;
    right:20px;
    width:24px; height:24px;
    border-radius:50%;
    border:1px solid #C8D9F6;
    color:var(--blue);
    display:flex; align-items:center; justify-content:center;
    font-size:14px;
    font-weight:900;
}
.icon-bubble {
    width:50px; height:50px; border-radius:50%;
    background:#EAF2FF;
    display:flex; align-items:center; justify-content:center;
    color:var(--blue);
    font-size:26px;
    margin-bottom:18px;
}
.module-card h3, .wide-card h3 {
    margin:0 0 10px;
    color:var(--text);
    font-size:18px;
    font-weight:900;
}
.module-card p, .wide-card p {
    margin:0 0 18px;
    color:#53647F;
    font-size:13px;
    line-height:1.48;
    min-height:54px;
}
.wide-grid {
    display:grid;
    grid-template-columns: 1fr 1fr;
    gap:18px;
    margin-bottom:18px;
}
.wide-card {
    padding:22px;
    min-height:140px;
    position:relative;
    display:grid;
    grid-template-columns: 190px 1fr;
    gap:18px;
    align-items:center;
}
.timeline-visual span {
    display:block;
    height:8px;
    border-radius:999px;
    background:#E8EEF8;
    margin:14px 0;
}
.timeline-visual span:nth-child(2) { width:70%; background:#2F80ED; }
.bar-visual {
    height:92px;
    display:flex;
    align-items:end;
    gap:12px;
    justify-content:flex-end;
}
.bar-visual span {
    width:24px;
    border-radius:4px 4px 0 0;
    background:#DDE7F5;
}
.bar-visual span:nth-child(2), .bar-visual span:nth-child(3), .bar-visual span:nth-child(5) { background:#2F80ED; }
.system-card {
    padding:18px 22px;
    display:grid;
    grid-template-columns: 180px repeat(4, 1fr);
    gap:20px;
    align-items:center;
}
.system-title {
    color:var(--blue);
    font-weight:900;
}
.system-item small {
    display:block;
    color:#75839B;
    font-size:11px;
}
.system-item strong {
    color:var(--text);
    font-size:13px;
}

h2, h3, h4, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
    color: var(--text) !important;
}
p, label, .stCaption, [data-testid="stCaptionContainer"] {
    color: #53647F !important;
}

.page-head {
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    gap:18px;
    padding:20px 0 18px;
}
.breadcrumb {
    color:#61708A;
    font-size:13px;
    margin-bottom:12px;
}
.page-head h1 {
    margin:0;
    font-size:30px;
    color:var(--text);
    font-weight:900;
}
.page-head p {
    margin:8px 0 0;
    color:#53647F;
}
.status-pill, .info-pill, .warning-pill, .success-pill {
    display:inline-flex;
    align-items:center;
    gap:7px;
    padding:7px 11px;
    border-radius:999px;
    font-size:12px;
    font-weight:800;
}
.status-pill, .success-pill { background:#EAFBF1; color:#15803D; border:1px solid #BBF7D0; }
.info-pill { background:#EEF5FF; color:#1D4ED8; border:1px solid #C7DAFF; }
.warning-pill { background:#FFF7ED; color:#B45309; border:1px solid #FED7AA; }

.tabs-line {
    display:flex;
    gap:28px;
    border-bottom:1px solid var(--line);
    margin:12px 0 28px;
}
.tabs-line span {
    padding:12px 0;
    font-weight:800;
    color:#53647F;
}
.tabs-line span.active {
    color:var(--blue);
    border-bottom:2px solid var(--blue);
}
.workflow-grid {
    display:grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap:18px;
    margin:18px 0;
}
.workflow-card {
    padding:22px 16px 16px;
    text-align:center;
    min-height:260px;
    position:relative;
}
.step-number {
    position:absolute;
    right:14px; top:14px;
    width:26px; height:26px; border-radius:50%;
    background:var(--blue);
    color:white;
    font-weight:900;
    display:flex; align-items:center; justify-content:center;
}
.file-status {
    display:inline-flex;
    font-size:11px;
    font-weight:800;
    padding:5px 10px;
    border-radius:999px;
    background:#FFF7ED;
    color:#EA580C;
    margin:5px 0 14px;
}
.file-status.ok { background:#EAFBF1; color:#15803D; }
.upload-box-note {
    border:1px dashed #9DB7E6;
    background:#FAFCFF;
    border-radius:13px;
    padding:17px 12px;
    color:#354766;
    margin:10px 0;
}

.notice {
    border:1px solid #C7DAFF;
    background:#F4F8FF;
    color:#174EA6;
    padding:14px 16px;
    border-radius:14px;
    margin: 14px 0;
    font-weight:700;
}
.notice.warn { border-color:#FED7AA; background:#FFF7ED; color:#9A3412; }
.notice.ok { border-color:#BBF7D0; background:#F0FDF4; color:#166534; }

/* Streamlit widgets */
.stButton > button, .stDownloadButton > button {
    border-radius: 10px !important;
    min-height: 40px;
    font-weight: 800 !important;
    border: 1px solid #C7D7F0 !important;
    color: #174EA6 !important;
    background: #FFFFFF !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    border-color:#2F80ED !important;
    color:#0A2E6D !important;
    box-shadow: 0 8px 18px rgba(30,94,255,.10);
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1E5EFF 0%, #0646D9 100%) !important;
    color: white !important;
    border: 0 !important;
}
.stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-baseweb="select"] > div {
    background: #FFFFFF !important;
    color: #071B4A !important;
    border-color: #C7D7F0 !important;
}
[data-testid="stFileUploader"] section {
    border-radius: 13px !important;
    border: 1px dashed #9DB7E6 !important;
    background: #FAFCFF !important;
}
[data-testid="stFileUploader"] section * {
    color: #071B4A !important;
}
[data-testid="stFileUploader"] button {
    background: #FFFFFF !important;
    color: #174EA6 !important;
    border: 1px solid #C7D7F0 !important;
}
[data-testid="stProgress"] div div div div {
    background: linear-gradient(90deg, #1E5EFF, #2F80ED) !important;
}
[data-testid="stDataFrame"], .stDataFrame {
    background:#FFFFFF !important;
    border-radius:14px;
}
.stAlert {
    border-radius:14px !important;
}

@media (max-width: 1050px) {
    .module-grid, .workflow-grid { grid-template-columns: 1fr 1fr; }
    .wide-grid { grid-template-columns: 1fr; }
    .system-card { grid-template-columns: 1fr 1fr; }
    .topbar { padding: 0 18px; }
    .top-logo { width:145px; }
}
@media (max-width: 760px) {
    .module-grid, .workflow-grid, .wide-grid { grid-template-columns: 1fr; }
    .system-card { grid-template-columns: 1fr; }
    .hero { padding: 28px 24px; }
    .hero:after { display:none; }
    .topbar-right { display:none; }
    .top-system { display:none; }
    .top-divider { display:none; }
}
</style>
""",
        unsafe_allow_html=True,
    )


# =========================================================
# REGISTRO DE MODULOS
# =========================================================
def ler_registry() -> dict[str, Any]:
    if not REGISTRY_FILE.exists():
        return {"modules": {}}
    try:
        data = json.loads(REGISTRY_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"modules": {}}
        if "modules" not in data or not isinstance(data["modules"], dict):
            data["modules"] = {}
        return data
    except Exception as exc:
        logger.exception("Falha ao ler registry: %s", exc)
        return {"modules": {}}


def salvar_registry(data: dict[str, Any]) -> None:
    REGISTRY_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def script_deve_ser_ignorado(caminho: Path) -> bool:
    nome = caminho.name.lower()
    ignorar = {
        APP_FILE.lower(),
        "app.py",
        "app_home_profissional.py",
        "app_nepomuceno_final.py",
        "app_layout_editor.py",
        "app_dark_glass_corrigido.py",
        "app_roadmap_aplicado.py",
        "app_profissional_final.py",
        "__init__.py",
    }
    if nome in ignorar:
        return True
    if nome.startswith(("streamlit ", "pip ", "python ", "py ")):
        return True
    return False


@st.cache_data(show_spinner=False)
def listar_scripts_python_cache(base: str, app_file: str, snapshot: tuple[tuple[str, float], ...]) -> list[str]:
    base_dir = Path(base)
    arquivos: list[str] = []
    for arq in base_dir.iterdir():
        if not arq.is_file() or arq.suffix.lower() != ".py":
            continue
        if script_deve_ser_ignorado(arq):
            continue
        arquivos.append(arq.name)
    return sorted(arquivos, key=lambda x: x.lower())


def listar_scripts_python() -> list[str]:
    snapshot = tuple(sorted(
        (p.name, p.stat().st_mtime)
        for p in BASE_DIR.iterdir()
        if p.is_file() and p.suffix.lower() == ".py"
    ))
    return listar_scripts_python_cache(str(BASE_DIR), APP_FILE, snapshot)


def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = Path(nome_arquivo).stem.replace("_", " ").replace("-", " ")
    return " ".join(parte.capitalize() for parte in nome.split())


def slugify(texto: str) -> str:
    texto = texto.lower().strip()
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_") or "modulo"


def inspecionar_script(nome_arquivo: str) -> dict[str, Any]:
    caminho = BASE_DIR / nome_arquivo
    info = {
        "arquivo": nome_arquivo,
        "existe": caminho.exists(),
        "parse_ok": False,
        "main_streamlit": False,
        "main": False,
        "modulo_config": False,
        "erro": "",
    }
    if not caminho.exists():
        info["erro"] = "Arquivo não encontrado"
        return info
    try:
        tree = ast.parse(caminho.read_text(encoding="utf-8"), filename=str(caminho))
        info["parse_ok"] = True
        for node in tree.body:
            if isinstance(node, ast.FunctionDef):
                if node.name == "main_streamlit":
                    info["main_streamlit"] = True
                if node.name == "main":
                    info["main"] = True
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id == "MODULO_CONFIG":
                        info["modulo_config"] = True
    except Exception as exc:
        info["erro"] = str(exc)
    return info


def config_padrao_modulo(nome_arquivo: str) -> dict[str, Any]:
    nome = nome_amigavel_script(nome_arquivo)
    return {
        "arquivo": nome_arquivo,
        "nome": nome,
        "icone": "🧩",
        "descricao": "Módulo operacional adicionado ao portal.",
        "categoria": "Módulos adicionais",
        "ativo": False,
        "ordem": 100,
        "modo": "auto",
        "slug": slugify(nome),
    }


def obter_config_modulo(nome_arquivo: str) -> dict[str, Any]:
    registry = ler_registry()
    config = registry.get("modules", {}).get(nome_arquivo)
    if not isinstance(config, dict):
        return config_padrao_modulo(nome_arquivo)
    padrao = config_padrao_modulo(nome_arquivo)
    padrao.update(config)
    padrao["arquivo"] = nome_arquivo
    return padrao


def salvar_config_modulo(nome_arquivo: str, config: dict[str, Any]) -> None:
    registry = ler_registry()
    registry.setdefault("modules", {})[nome_arquivo] = config
    salvar_registry(registry)
    listar_scripts_python_cache.clear()


def modulos_ativos() -> list[dict[str, Any]]:
    mods: list[dict[str, Any]] = []
    for arquivo in listar_scripts_python():
        cfg = obter_config_modulo(arquivo)
        if not cfg.get("ativo", False):
            continue
        info = inspecionar_script(arquivo)
        if not info["parse_ok"]:
            continue
        cfg["_info"] = info
        mods.append(cfg)
    return sorted(mods, key=lambda x: (int(x.get("ordem", 100)), str(x.get("nome", ""))))


# =========================================================
# CARREGAMENTO / EXECUCAO DE MODULOS
# =========================================================
def nome_modulo_seguro(caminho_script: Path) -> str:
    base = re.sub(r"\W+", "_", caminho_script.stem).strip("_") or "modulo"
    return f"modulo_{base}_{abs(hash(str(caminho_script))) % 10000000}"


def set_page_config_noop(*args, **kwargs) -> None:
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


def executar_script_integrado(caminho: Path) -> None:
    """Executa um .py Streamlit legado dentro do portal sem alterar o arquivo original.

    Esse adaptador é usado pelo Editor de Módulos quando o arquivo não possui
    main_streamlit() ou possui código Streamlit solto no corpo do arquivo.
    Ele preserva a lógica do script e apenas bloqueia chamadas incompatíveis
    com o app principal, como st.set_page_config().
    """
    original_set_page_config = st.set_page_config
    cwd_original = Path.cwd()
    argv_original = sys.argv[:]
    try:
        st.set_page_config = set_page_config_noop
        os.chdir(BASE_DIR)
        sys.argv = [str(caminho)]
        runpy.run_path(str(caminho), run_name="__main__")
    finally:
        sys.argv = argv_original
        os.chdir(cwd_original)
        st.set_page_config = original_set_page_config


def recomendar_modo_execucao(info: dict[str, Any]) -> str:
    """Escolhe o modo mais seguro para o app executar um arquivo .py sem alterar a lógica."""
    if info.get("main_streamlit"):
        return "main_streamlit"
    if info.get("main"):
        return "main"
    return "script"


def executar_modulo_config(config: dict[str, Any]) -> None:
    caminho = BASE_DIR / str(config.get("arquivo", ""))
    modo = str(config.get("modo", "auto"))
    if not caminho.exists():
        st.error("Arquivo do módulo não encontrado.")
        return

    info = inspecionar_script(caminho.name)
    if not info.get("parse_ok", False):
        st.error(f"O arquivo {caminho.name} possui erro de sintaxe e não pode ser executado: {info.get('erro', '')}")
        return

    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop

        if modo == "auto":
            modo = recomendar_modo_execucao(info)

        if modo == "script":
            executar_script_integrado(caminho)
            return

        modulo = carregar_modulo_por_arquivo(caminho)
        if modo == "main_streamlit":
            if not hasattr(modulo, "main_streamlit"):
                st.warning("main_streamlit() não foi encontrada. O portal executará o arquivo como script integrado.")
                executar_script_integrado(caminho)
                return
            modulo.main_streamlit()
        elif modo == "main":
            if not hasattr(modulo, "main"):
                st.warning("main() não foi encontrada. O portal executará o arquivo como script integrado.")
                executar_script_integrado(caminho)
                return
            modulo.main()
        else:
            executar_script_integrado(caminho)
    except SystemExit:
        # Alguns scripts antigos chamam sys.exit(). No portal isso não deve quebrar a aplicação inteira.
        st.info("O módulo terminou a execução.")
    except Exception as exc:
        logger.exception("Falha ao executar modulo %s: %s", caminho.name, exc)
        st.error(f"Erro ao executar módulo {caminho.name}: {exc}")
        with st.expander("Detalhes técnicos para correção"):
            st.code(str(exc))
    finally:
        st.set_page_config = original_set_page_config


def validar_funcoes_modulo(modulo, funcoes: list[str]) -> list[str]:
    return [f for f in funcoes if not hasattr(modulo, f)]


# =========================================================
# COMPONENTES VISUAIS
# =========================================================
def render_topbar() -> None:
    st.markdown(
        f"""
        <div class="topbar">
            <div class="topbar-left">
                {logo_html('top-logo')}
                <div class="top-divider"></div>
                <div class="top-system">Central Operacional<br>de Análises</div>
            </div>
            <div class="topbar-right">
                <div class="top-icon">🔔</div>
                <div class="top-icon">?</div>
                <div class="avatar">AD</div>
                <div class="userbox"><strong>Administrador</strong><span>admin@expnepo.com.br</span></div>
                <div>⌄</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar(menu_items: list[dict[str, str]]) -> str:
    pagina_atual_default()
    keys = [m["key"] for m in menu_items]
    if st.session_state["pagina_atual"] not in keys:
        st.session_state["pagina_atual"] = "inicio"

    with st.sidebar:
        st.markdown(
            f"""
            <div class="sidebar-brand">
                {logo_html('sidebar-logo-img')}
            </div>
            <div class="sidebar-title">Central Operacional</div>
            <div class="sidebar-subtitle">Expresso Nepomuceno</div>
            <div class="sidebar-section">Navegação</div>
            """,
            unsafe_allow_html=True,
        )

        for item in menu_items:
            key = item["key"]
            label = item["label"]
            tipo = "primary" if st.session_state["pagina_atual"] == key else "secondary"
            if st.button(label, key=f"nav_{slugify(key)}", type=tipo, use_container_width=True):
                st.session_state["pagina_atual"] = key
                st.rerun()

        st.markdown("<div style='height:180px'></div>", unsafe_allow_html=True)
        st.markdown("<div class='sidebar-muted'>Versão 2.5.4<br>Ambiente interno</div>", unsafe_allow_html=True)
    return st.session_state["pagina_atual"]

def nav_button(label: str, page: str, primary: bool = False, key: str | None = None) -> None:
    safe_key = key or f"navbtn_{slugify(page)}_{slugify(label)}"
    if st.button(label, key=safe_key, use_container_width=True, type="primary" if primary else "secondary"):
        ir_para(page)


def render_notice(text: str, kind: str = "info") -> None:
    cls = "notice"
    if kind == "warn":
        cls += " warn"
    if kind == "ok":
        cls += " ok"
    st.markdown(f"<div class='{cls}'>{text}</div>", unsafe_allow_html=True)


def render_page_head(titulo: str, subtitulo: str, badge: str = "Ativo") -> None:
    st.markdown(
        f"""
        <div class="page-head">
            <div>
                <div class="breadcrumb">Central Operacional de Análises / <b>{html.escape(titulo)}</b></div>
                <h1>{html.escape(titulo)} <span class="status-pill">● {html.escape(badge)}</span></h1>
                <p>{html.escape(subtitulo)}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_module_card(icon: str, title: str, text: str, button: str, page: str) -> None:
    st.markdown(
        f"""
        <div class="module-card">
            <div class="arrow-dot">›</div>
            <div class="icon-bubble">{html.escape(icon)}</div>
            <h3>{html.escape(title)}</h3>
            <p>{html.escape(text)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    nav_button(button, page, key=f"card_{slugify(page)}_{slugify(title)}")


def render_metric_card(label: str, value: str, caption: str = "") -> None:
    st.markdown(
        f"""
        <div class="system-item">
            <small>{html.escape(label)}</small>
            <strong>{html.escape(value)}</strong>
            <small>{html.escape(caption)}</small>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_workflow_card(numero: int, icon: str, title: str, subtitle: str, ok: bool) -> None:
    status = "Enviado" if ok else "Pendente"
    status_cls = "file-status ok" if ok else "file-status"
    st.markdown(
        f"""
        <div class="workflow-card">
            <div class="step-number">{numero}</div>
            <div class="icon-bubble" style="margin:0 auto 16px;">{html.escape(icon)}</div>
            <h3>{html.escape(title)}</h3>
            <p style="min-height:0;margin-bottom:8px;">{html.escape(subtitle)}</p>
            <div class="{status_cls}">{status}</div>
            <div class="upload-box-note">☁️<br>Arraste e solte o arquivo aqui<br><small>ou clique para selecionar</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# HOME
# =========================================================
def pagina_inicio() -> None:
    ativos = modulos_ativos()
    historico = carregar_historico()
    st.markdown(
        """
        <div class="hero">
            <h1>Bem-vindo à Central Operacional de Análises</h1>
            <p>Acesse as ferramentas de análise, histórico e relatórios do sistema.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        render_module_card("◷", "Análise de Permanência", "Analise o tempo de permanência dos veículos nos locais de carga e descarga.", "Acessar módulo  →", "permanencia")
    with c2:
        render_module_card("◴", "Odômetro V12", "Consulte, valide e analise os dados de odômetro dos veículos da frota V12.", "Acessar módulo  →", "odometro")
    with c3:
        render_module_card("🧩", "Editor de Módulos", "Crie, edite e gerencie módulos de regras e parâmetros para análises personalizadas.", "Acessar módulo  →", "editor")

    c4, c5 = st.columns(2)
    with c4:
        st.markdown(
            """
            <div class="wide-card">
                <div>
                    <div class="icon-bubble">▤</div>
                    <h3>Histórico</h3>
                    <p>Visualize o histórico completo de análises realizadas no sistema.</p>
                </div>
                <div class="timeline-visual"><span></span><span></span><span></span><span></span></div>
                <div class="arrow-dot">›</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        nav_button("Acessar histórico  →", "historico", key="home_historico")
    with c5:
        st.markdown(
            """
            <div class="wide-card">
                <div>
                    <div class="icon-bubble">▥</div>
                    <h3>Relatórios</h3>
                    <p>Gere relatórios detalhados e personalizados com base nas análises realizadas.</p>
                </div>
                <div class="bar-visual"><span style="height:34px"></span><span style="height:52px"></span><span style="height:60px"></span><span style="height:78px"></span><span style="height:92px"></span></div>
                <div class="arrow-dot">›</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        nav_button("Acessar relatórios  →", "relatorios", key="home_relatorios")

    st.markdown(
        f"""
        <div class="system-card">
            <div class="system-title">ⓘ &nbsp; Informações do Sistema</div>
            <div class="system-item"><small>Versão do sistema</small><strong>2.5.4</strong></div>
            <div class="system-item"><small>Ambiente</small><strong>Produção</strong></div>
            <div class="system-item"><small>Último acesso</small><strong>{datetime.now().strftime('%d/%m/%Y %H:%M')}</strong></div>
            <div class="system-item"><small>Módulos ativos</small><strong>{len(ativos) + 2}</strong></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# PERMANENCIA
# =========================================================
def pagina_permanencia() -> None:
    render_page_head("Análise de Permanência", "Processamento da base de permanência com classificação por tempo configurável.")
    st.markdown('<div class="tabs-line"><span class="active">Visão Geral</span><span>Configurações</span><span>Histórico de Execuções</span><span>Documentação</span></div>', unsafe_allow_html=True)

    caminho_permanencia = encontrar_arquivo(["Codigo_colado.py", "Código_colado.py", "Codigo colado.py", "Código colado.py"])
    if caminho_permanencia is None:
        st.error("Arquivo Codigo_colado.py não encontrado na pasta do app.")
        return
    try:
        permanencia = carregar_modulo_por_arquivo(caminho_permanencia)
    except Exception as exc:
        logger.exception("Erro ao importar permanencia: %s", exc)
        st.error(f"Erro ao importar {caminho_permanencia.name}: {exc}")
        return
    faltando = validar_funcoes_modulo(permanencia, ["carregar_dados", "identificar_eventos_carregamento", "montar_ciclos_carregamento", "gerar_resumos", "salvar_saida"])
    if faltando:
        st.error("Arquivo de permanência sem funções esperadas: " + ", ".join(faltando))
        return

    col_a, col_b = st.columns(2)
    with col_a:
        tempo_minimo = st.number_input("Tempo mínimo aceitável (minutos)", min_value=0, value=15, step=1, key="perm_tempo_minimo")
    with col_b:
        tempo_maximo = st.number_input("Tempo máximo aceitável (minutos)", min_value=1, value=55, step=1, key="perm_tempo_maximo")
    if tempo_maximo <= tempo_minimo:
        st.error("O tempo máximo precisa ser maior que o mínimo.")

    arquivo = st.file_uploader("Selecione o Excel de permanência", type=["xlsx", "xls"], key="upload_permanencia")
    if not arquivo:
        render_notice("Aguardando upload da base de permanência para iniciar o processamento.", "warn")
        return
    render_notice(f"Arquivo carregado: <b>{html.escape(arquivo.name)}</b>", "ok")

    if st.button("Executar workflow", use_container_width=True, type="primary", key="btn_executar_permanencia"):
        inicio = time.time()
        temp_path = None
        saida = None
        try:
            if tempo_maximo <= tempo_minimo:
                st.error("Corrija os tempos antes de processar.")
                return
            if not validar_upload_excel(arquivo, "Base de permanência"):
                return
            permanencia.TEMPO_MINIMO = tempo_minimo
            permanencia.TEMPO_MAXIMO = tempo_maximo
            barra = st.progress(0)
            status = st.empty()
            barra.progress(10); status.info("10% - Preparando arquivo")
            temp_path = salvar_upload_temporario(arquivo)
            barra.progress(30); status.info("30% - Lendo base")
            df_base = permanencia.carregar_dados(temp_path)
            barra.progress(50); status.info("50% - Identificando eventos")
            eventos = permanencia.identificar_eventos_carregamento(df_base)
            barra.progress(70); status.info("70% - Montando ciclos")
            df_resultado, df_alertas = permanencia.montar_ciclos_carregamento(eventos)
            barra.progress(85); status.info("85% - Gerando resumos")
            resumo_geral, resumo_up, resumo_eq, ranking = permanencia.gerar_resumos(df_resultado)
            barra.progress(95); status.info("95% - Gerando Excel final")
            saida = permanencia.salvar_saida(temp_path, df_base, eventos, df_resultado, df_alertas, resumo_geral, resumo_up, resumo_eq, ranking)
            barra.progress(100); status.success("100% - Finalizado")
            duracao = time.time() - inicio
            registrar_historico("Permanência", arquivo.name, "Concluído", duracao, f"Eventos: {len(df_resultado)}")
            st.success(f"Processo finalizado em {duracao:.2f} segundos")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Eventos", len(df_resultado))
            c2.metric("Alertas", 0 if df_alertas.empty else len(df_alertas))
            c3.metric("OK", 0 if df_resultado.empty else int((df_resultado["Status"] == "OK").sum()))
            c4.metric("Improcedentes", 0 if df_resultado.empty else int((df_resultado["Status"] == "IMPROCEDENTE").sum()))
            if not df_resultado.empty:
                st.dataframe(df_resultado.head(100), use_container_width=True)
            with open(saida, "rb") as f:
                st.download_button("Baixar Excel Permanência", f, file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_permanencia")
        except Exception as exc:
            duracao = time.time() - inicio
            registrar_historico("Permanência", getattr(arquivo, "name", ""), "Erro", duracao, str(exc))
            logger.exception("Erro no processamento de permanencia: %s", exc)
            st.error(f"Erro ao processar permanência: {exc}")
        finally:
            limpar_arquivo_temporario(temp_path)


# =========================================================
# ODOMETRO
# =========================================================
def pagina_odometro() -> None:
    render_page_head("Odômetro V12", "Validação e análise de odômetro com integração de bases e geração de relatórios.")
    st.markdown('<div class="tabs-line"><span class="active">Visão Geral</span><span>Configurações</span><span>Agendamentos</span><span>Histórico de Execuções</span><span>Documentação</span></div>', unsafe_allow_html=True)
    render_notice("Faça upload das quatro bases necessárias para executar o workflow Odômetro V12.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_workflow_card(1, "⛽", "Base Combustível", "CSV, XLSX ou Parquet", st.session_state.get("comb") is not None)
        comb = st.file_uploader("Base Combustível", type=["xlsx", "xls"], key="comb")
    with c2:
        render_workflow_card(2, "◴", "Base Km Rodado Maxtrack", "CSV, XLSX ou Parquet", st.session_state.get("maxtrack") is not None)
        maxtrack = st.file_uploader("Base Km Rodado Maxtrack", type=["xlsx", "xls"], key="maxtrack")
    with c3:
        render_workflow_card(3, "🚛", "Base Ativo de Veículos", "CSV, XLSX ou Parquet", st.session_state.get("ativo") is not None)
        ativo = st.file_uploader("Base Ativo de Veículos", type=["xlsx", "xls"], key="ativo")
    with c4:
        render_workflow_card(4, "👥", "Produção Oficial / Cliente", "CSV, XLSX ou Parquet", st.session_state.get("producao") is not None)
        producao = st.file_uploader("Produção Oficial / Cliente", type=["xlsx", "xls"], key="producao")

    arquivos = [comb, maxtrack, ativo, producao]
    qtd = sum(a is not None for a in arquivos)
    st.progress(qtd / 4)
    st.caption(f"{qtd} de 4 arquivos carregados")
    if qtd < 4:
        render_notice("Próximo passo: após o envio de todas as bases, clique em Executar workflow para iniciar o processamento.", "warn")
        return

    if st.button("Executar workflow", use_container_width=True, type="primary", key="btn_executar_permanencia"):
        inicio = time.time()
        temps: list[str] = []
        saida = None
        try:
            nomes = ["Base Combustível", "Base Km Rodado Maxtrack", "Base Ativo de Veículos", "Produção Oficial / Cliente"]
            for arq, nome in zip(arquivos, nomes):
                if not validar_upload_excel(arq, nome):
                    return
            barra = st.progress(0)
            status = st.empty()
            barra.progress(10); status.info("10% - Salvando arquivos temporários")
            for arq in arquivos:
                temps.append(salvar_upload_temporario(arq))
            saida = os.path.join(tempfile.gettempdir(), f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            script = encontrar_arquivo(["odometro_v12_com_percentual.py"])
            if script is None:
                st.error("Arquivo odometro_v12_com_percentual.py não encontrado.")
                return
            barra.progress(25); status.info("25% - Executando script do odômetro")
            processo = subprocess.Popen([sys.executable, str(script), *temps, saida], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
            logs: list[str] = []
            log_box = st.empty()
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
                time.sleep(0.25)
            if processo.returncode != 0:
                raise RuntimeError("O processamento retornou erro. Consulte os logs exibidos na tela.")
            if not os.path.exists(saida):
                raise FileNotFoundError("Arquivo final não foi gerado.")
            barra.progress(100); status.success("100% - Finalizado")
            duracao = time.time() - inicio
            registrar_historico("Odômetro V12", "4 bases", "Concluído", duracao, "Arquivo final gerado")
            st.success(f"Odômetro finalizado em {duracao:.2f} segundos")
            with open(saida, "rb") as f:
                st.download_button("Baixar Excel Odômetro V12", f, file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_odometro")
        except Exception as exc:
            duracao = time.time() - inicio
            registrar_historico("Odômetro V12", "4 bases", "Erro", duracao, str(exc))
            logger.exception("Erro no processamento de odometro: %s", exc)
            st.error(f"Erro ao processar odômetro: {exc}")
        finally:
            for temp_path in temps:
                limpar_arquivo_temporario(temp_path)


# =========================================================
# EDITOR DE MODULOS - PROTEGIDO POR SENHA
# =========================================================
def editor_autenticado() -> bool:
    if st.session_state.get("editor_auth") is True:
        return True
    render_page_head("Editor de Módulos", "Área administrativa protegida para configurar e criar módulos do portal.", "Protegido")
    render_notice("Acesso restrito. Informe usuário e senha para continuar.", "warn")
    with st.form("login_editor"):
        usuario = st.text_input("Usuário", key="editor_login_usuario")
        senha = st.text_input("Senha", type="password", key="editor_login_senha")
        entrar = st.form_submit_button("Entrar", use_container_width=True)
    if entrar:
        if usuario.strip().lower() == ADMIN_USER and senha == ADMIN_PASSWORD:
            st.session_state["editor_auth"] = True
            st.success("Acesso liberado.")
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")
    return False


def criar_template_modulo(nome_modulo: str, icone: str, descricao: str, arquivo: str) -> str:
    titulo = nome_modulo.strip() or "Novo Módulo"
    safe_file = slugify(arquivo or titulo) + ".py"
    conteudo = f'''"""Módulo criado pelo Editor da Central Operacional de Análises."""
import streamlit as st

MODULO_CONFIG = {{
    "nome": {titulo!r},
    "icone": {icone!r},
    "descricao": {descricao!r},
    "categoria": "Módulos adicionais",
    "ativo": True,
    "ordem": 100,
    "modo": "auto",
}}


def main_streamlit():
    st.title({titulo!r})
    st.info({descricao!r})

    arquivo = st.file_uploader("Envie o arquivo de entrada", type=["xlsx", "xls", "csv"])

    if arquivo is None:
        st.warning("Aguardando arquivo para iniciar o processamento.")
        return

    st.success(f"Arquivo recebido: {{arquivo.name}}")

    if st.button("Processar", type="primary", use_container_width=True):
        # TODO: incluir aqui a regra de processamento específica do módulo.
        st.success("Processamento executado. Personalize esta função conforme a regra do módulo.")


if __name__ == "__main__":
    main_streamlit()
'''
    return safe_file, conteudo


def pagina_editor() -> None:
    if not editor_autenticado():
        return
    render_page_head("Editor de Módulos", "Configure qualquer arquivo Python sem alterar o código original, ou crie um módulo novo já compatível.")
    if st.button("Sair do editor protegido", use_container_width=False, key="btn_sair_editor"):
        st.session_state["editor_auth"] = False
        st.rerun()

    tab1, tab2, tab3 = st.tabs(["Configurar módulo existente", "Criar módulo novo", "Diagnóstico"])
    with tab1:
        scripts = listar_scripts_python()
        if not scripts:
            st.warning("Nenhum arquivo .py adicional encontrado na pasta do app.")
        else:
            selecionado = st.selectbox("Arquivo Python", scripts, key="editor_arquivo")
            info = inspecionar_script(selecionado)
            if info["parse_ok"]:
                modo_recomendado = recomendar_modo_execucao(info)
                st.success(
                    f"Arquivo válido. main_streamlit: {info['main_streamlit']} | "
                    f"main: {info['main']} | MODULO_CONFIG: {info['modulo_config']} | "
                    f"modo recomendado: {modo_recomendado}"
                )
                render_notice(
                    "O Editor não altera a lógica do arquivo original. Ele apenas grava uma configuração externa "
                    "em modulos_config.json para o portal saber como executar esse .py dentro do app.",
                    "info",
                )
            else:
                st.error(f"Arquivo com erro de sintaxe: {info['erro']}")
            cfg = obter_config_modulo(selecionado)
            if st.button("Analisar e preparar automaticamente para rodar no app", use_container_width=True, key=f"btn_auto_preparar_{selecionado}"):
                if not info["parse_ok"]:
                    st.error("Não é possível preparar automaticamente: primeiro corrija o erro de sintaxe informado acima.")
                else:
                    modo_auto = recomendar_modo_execucao(info)
                    novo = cfg.copy()
                    novo.update({
                        "arquivo": selecionado,
                        "nome": str(cfg.get("nome") or nome_amigavel_script(selecionado)),
                        "icone": str(cfg.get("icone") or "🧩"),
                        "descricao": str(cfg.get("descricao") or "Módulo operacional adicionado ao portal."),
                        "categoria": str(cfg.get("categoria") or "Módulos adicionais"),
                        "ativo": True,
                        "ordem": int(cfg.get("ordem", 100)),
                        "modo": modo_auto,
                        "slug": slugify(str(cfg.get("nome") or selecionado)),
                        "adaptador_integrado": True,
                    })
                    salvar_config_modulo(selecionado, novo)
                    st.success(f"Módulo preparado e ativado no menu. Modo definido automaticamente como: {modo_auto}")
                    time.sleep(.5)
                    st.rerun()
            col1, col2 = st.columns(2)
            with col1:
                nome = st.text_input("Nome no menu", value=str(cfg.get("nome", "")), key=f"editor_nome_{selecionado}")
                icone = st.text_input("Ícone", value=str(cfg.get("icone", "🧩")), key=f"editor_icone_{selecionado}")
                ordem = st.number_input("Ordem", min_value=1, max_value=999, value=int(cfg.get("ordem", 100)), step=1, key=f"editor_ordem_{selecionado}")
                ativo = st.checkbox("Ativo no menu", value=bool(cfg.get("ativo", False)), key=f"editor_ativo_{selecionado}")
            with col2:
                descricao = st.text_area("Descrição", value=str(cfg.get("descricao", "")), height=118, key=f"editor_desc_{selecionado}")
                modo = st.selectbox("Modo de execução", ["auto", "main_streamlit", "main", "script"], index=["auto", "main_streamlit", "main", "script"].index(str(cfg.get("modo", "auto"))) if str(cfg.get("modo", "auto")) in ["auto", "main_streamlit", "main", "script"] else 0, key=f"editor_modo_{selecionado}")
                categoria = st.text_input("Categoria", value=str(cfg.get("categoria", "Módulos adicionais")), key=f"editor_categoria_{selecionado}")
            if st.button("Salvar configuração", use_container_width=True, type="primary", key=f"btn_salvar_config_{selecionado}"):
                novo = {
                    "arquivo": selecionado,
                    "nome": nome.strip() or nome_amigavel_script(selecionado),
                    "icone": icone.strip() or "🧩",
                    "descricao": descricao.strip() or "Módulo operacional adicionado ao portal.",
                    "categoria": categoria.strip() or "Módulos adicionais",
                    "ativo": ativo,
                    "ordem": int(ordem),
                    "modo": modo,
                    "slug": slugify(nome or selecionado),
                }
                salvar_config_modulo(selecionado, novo)
                st.success("Configuração salva. O menu será atualizado automaticamente.")
                time.sleep(.5)
                st.rerun()
    with tab2:
        st.markdown("### Criar arquivo Python já pronto para funcionar no app")
        col1, col2 = st.columns(2)
        with col1:
            nome_modulo = st.text_input("Nome do módulo", value="Novo Módulo Operacional", key="novo_modulo_nome")
            icone_modulo = st.text_input("Ícone do módulo", value="🧩", key="novo_modulo_icone")
        with col2:
            arquivo_modulo = st.text_input("Nome do arquivo sem .py", value="novo_modulo_operacional", key="novo_modulo_arquivo")
            desc_modulo = st.text_area("Descrição do módulo", value="Módulo criado pelo Editor para processamento operacional.", height=98, key="novo_modulo_desc")
        if st.button("Criar módulo compatível", use_container_width=True, type="primary", key="btn_criar_modulo_compat"):
            nome_arquivo, conteudo = criar_template_modulo(nome_modulo, icone_modulo, desc_modulo, arquivo_modulo)
            destino = BASE_DIR / nome_arquivo
            if destino.exists():
                st.error(f"Já existe um arquivo chamado {nome_arquivo}.")
            else:
                destino.write_text(conteudo, encoding="utf-8")
                salvar_config_modulo(nome_arquivo, {
                    "arquivo": nome_arquivo,
                    "nome": nome_modulo,
                    "icone": icone_modulo,
                    "descricao": desc_modulo,
                    "categoria": "Módulos adicionais",
                    "ativo": True,
                    "ordem": 100,
                    "modo": "auto",
                    "slug": slugify(nome_modulo),
                })
                st.success(f"Módulo {nome_arquivo} criado e ativado no menu.")
                st.code(conteudo, language="python")
    with tab3:
        dados = []
        for script in listar_scripts_python():
            info = inspecionar_script(script)
            cfg = obter_config_modulo(script)
            dados.append({
                "Arquivo": script,
                "Ativo": bool(cfg.get("ativo", False)),
                "Nome": cfg.get("nome", ""),
                "Modo": cfg.get("modo", "auto"),
                "Sintaxe OK": info["parse_ok"],
                "main_streamlit": info["main_streamlit"],
                "main": info["main"],
                "Erro": info["erro"],
            })
        st.dataframe(dados, use_container_width=True, hide_index=True)


# =========================================================
# HISTORICO / RELATORIOS / CONFIG
# =========================================================
def pagina_historico() -> None:
    render_page_head("Histórico", "Consulte os processamentos realizados e seus resultados.")
    hist = carregar_historico()
    if not hist:
        render_notice("Nenhum processamento registrado ainda.", "warn")
        return
    st.dataframe(hist, use_container_width=True, hide_index=True)


def pagina_relatorios() -> None:
    render_page_head("Relatórios", "Área para acompanhamento consolidado dos processamentos.")
    hist = carregar_historico()
    concluidos = len([h for h in hist if h.get("status") == "Concluído"])
    erros = len([h for h in hist if h.get("status") == "Erro"])
    c1, c2, c3 = st.columns(3)
    c1.metric("Processamentos", len(hist))
    c2.metric("Concluídos", concluidos)
    c3.metric("Erros", erros)
    render_notice("Os relatórios avançados podem ser expandidos conforme novos indicadores forem definidos.")


def pagina_configuracoes() -> None:
    render_page_head("Configurações", "Preferências e informações técnicas do ambiente.")
    st.write("**Pasta do app:**", str(BASE_DIR))
    st.write("**Arquivo principal:**", APP_FILE)
    st.write("**Arquivo de log:**", str(LOG_FILE))
    st.write("**Registry de módulos:**", str(REGISTRY_FILE))


# =========================================================
# MODULOS ADICIONAIS
# =========================================================
def pagina_modulo_dinamico(config: dict[str, Any]) -> None:
    render_page_head(str(config.get("nome", "Módulo")), str(config.get("descricao", "Módulo adicional configurado.")))
    executar_modulo_config(config)


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    aplicar_css()
    pagina_atual_default()
    ativos = modulos_ativos()

    menu_items = [
        {"key": "inicio", "label": "⌂  Início"},
        {"key": "permanencia", "label": "◷  Análise de Permanência"},
        {"key": "odometro", "label": "◴  Odômetro V12"},
        {"key": "editor", "label": "🧩  Editor de Módulos"},
        {"key": "historico", "label": "▤  Histórico"},
        {"key": "relatorios", "label": "▥  Relatórios"},
        {"key": "configuracoes", "label": "⚙  Configurações"},
    ]
    for cfg in ativos:
        menu_items.insert(-1, {"key": "modulo:" + str(cfg.get("arquivo")), "label": f"{cfg.get('icone','🧩')}  {cfg.get('nome','Módulo')}"})

    pagina = render_sidebar(menu_items)
    render_topbar()

    mapa_mods = {"modulo:" + str(cfg.get("arquivo")): cfg for cfg in ativos}

    if pagina == "inicio":
        pagina_inicio()
    elif pagina == "permanencia":
        pagina_permanencia()
    elif pagina == "odometro":
        pagina_odometro()
    elif pagina == "editor":
        pagina_editor()
    elif pagina == "historico":
        pagina_historico()
    elif pagina == "relatorios":
        pagina_relatorios()
    elif pagina == "configuracoes":
        pagina_configuracoes()
    elif pagina in mapa_mods:
        pagina_modulo_dinamico(mapa_mods[pagina])
    else:
        st.error("Página não encontrada.")


if __name__ == "__main__":
    main()
