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


def logo_html(css_class: str = "top-logo") -> str:
    if LOGO_B64:
        return f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="{css_class}" alt="Expresso Nepomuceno">'
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


def executar_modulo_config(config: dict[str, Any]) -> None:
    caminho = BASE_DIR / str(config.get("arquivo", ""))
    modo = str(config.get("modo", "auto"))
    if not caminho.exists():
        st.error("Arquivo do módulo não encontrado.")
        return
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop
        if modo == "script":
            runpy.run_path(str(caminho), run_name="__main__")
            return
        modulo = carregar_modulo_por_arquivo(caminho)
        if modo == "main_streamlit":
            if not hasattr(modulo, "main_streamlit"):
                st.error("Este módulo não possui main_streamlit().")
                return
            modulo.main_streamlit()
        elif modo == "main":
            if not hasattr(modulo, "main"):
                st.error("Este módulo não possui main().")
                return
            modulo.main()
        else:
            if hasattr(modulo, "main_streamlit"):
                modulo.main_streamlit()
            elif hasattr(modulo, "main"):
                modulo.main()
            else:
                st.warning("Este módulo não possui main_streamlit() nem main(). Configure como script ou ajuste pelo Editor de Módulos.")
    except Exception as exc:
        logger.exception("Falha ao executar modulo %s: %s", caminho.name, exc)
        st.error(f"Erro ao executar módulo {caminho.name}: {exc}")
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
            """
            <div class="sidebar-brand">
                <div class="sidebar-en">E.N</div>
                <div>
                    <div class="sidebar-title">Central Operacional</div>
                    <div class="sidebar-subtitle">Expresso Nepomuceno</div>
                </div>
            </div>
            <div class="sidebar-section">Navegação</div>
            """,
            unsafe_allow_html=True,
        )
        labels = [m["label"] for m in menu_items]
        label_by_key = {m["key"]: m["label"] for m in menu_items}
        key_by_label = {m["label"]: m["key"] for m in menu_items}
        current_label = label_by_key.get(st.session_state["pagina_atual"], labels[0])
        selected = st.radio(
            "Menu principal",
            labels,
            index=labels.index(current_label) if current_label in labels else 0,
            label_visibility="collapsed",
            key="menu_radio",
        )
        st.session_state["pagina_atual"] = key_by_label[selected]
        st.markdown("<div style='height:180px'></div>", unsafe_allow_html=True)
        st.markdown("<div class='sidebar-muted'>Versão 2.5.4<br>Ambiente interno</div>", unsafe_allow_html=True)
    return st.session_state["pagina_atual"]


def nav_button(label: str, page: str, primary: bool = False) -> None:
    if st.button(label, use_container_width=True, type="primary" if primary else "secondary"):
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
    nav_button(button, page)


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
        nav_button("Acessar histórico  →", "historico")
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
        nav_button("Acessar relatórios  →", "relatorios")

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
        tempo_minimo = st.number_input("Tempo mínimo aceitável (minutos)", min_value=0, value=15, step=1)
    with col_b:
        tempo_maximo = st.number_input("Tempo máximo aceitável (minutos)", min_value=1, value=55, step=1)
    if tempo_maximo <= tempo_minimo:
        st.error("O tempo máximo precisa ser maior que o mínimo.")

    arquivo = st.file_uploader("Selecione o Excel de permanência", type=["xlsx", "xls"], key="upload_permanencia")
    if not arquivo:
        render_notice("Aguardando upload da base de permanência para iniciar o processamento.", "warn")
        return
    render_notice(f"Arquivo carregado: <b>{html.escape(arquivo.name)}</b>", "ok")

    if st.button("Executar workflow", use_container_width=True, type="primary"):
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
                st.download_button("Baixar Excel Permanência", f, file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
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

    if st.button("Executar workflow", use_container_width=True, type="primary"):
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
                st.download_button("Baixar Excel Odômetro V12", f, file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
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
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
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
    if st.button("Sair do editor protegido", use_container_width=False):
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
                st.success(f"Arquivo válido. main_streamlit: {info['main_streamlit']} | main: {info['main']} | MODULO_CONFIG: {info['modulo_config']}")
            else:
                st.error(f"Arquivo com erro de sintaxe: {info['erro']}")
            cfg = obter_config_modulo(selecionado)
            col1, col2 = st.columns(2)
            with col1:
                nome = st.text_input("Nome no menu", value=str(cfg.get("nome", "")))
                icone = st.text_input("Ícone", value=str(cfg.get("icone", "🧩")))
                ordem = st.number_input("Ordem", min_value=1, max_value=999, value=int(cfg.get("ordem", 100)), step=1)
                ativo = st.checkbox("Ativo no menu", value=bool(cfg.get("ativo", False)))
            with col2:
                descricao = st.text_area("Descrição", value=str(cfg.get("descricao", "")), height=118)
                modo = st.selectbox("Modo de execução", ["auto", "main_streamlit", "main", "script"], index=["auto", "main_streamlit", "main", "script"].index(str(cfg.get("modo", "auto"))) if str(cfg.get("modo", "auto")) in ["auto", "main_streamlit", "main", "script"] else 0)
                categoria = st.text_input("Categoria", value=str(cfg.get("categoria", "Módulos adicionais")))
            if st.button("Salvar configuração", use_container_width=True, type="primary"):
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
            nome_modulo = st.text_input("Nome do módulo", value="Novo Módulo Operacional")
            icone_modulo = st.text_input("Ícone do módulo", value="🧩")
        with col2:
            arquivo_modulo = st.text_input("Nome do arquivo sem .py", value="novo_modulo_operacional")
            desc_modulo = st.text_area("Descrição do módulo", value="Módulo criado pelo Editor para processamento operacional.", height=98)
        if st.button("Criar módulo compatível", use_container_width=True, type="primary"):
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
