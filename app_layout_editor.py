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
LAYOUT_FILE = BASE_DIR / "layout_config.json"

DEFAULT_LAYOUT_CONFIG = {
    "tema": "dark_glass",
    "bg_0": "#020814",
    "bg_1": "#061426",
    "glass": "rgba(9, 20, 38, 0.72)",
    "glass_2": "rgba(16, 34, 60, 0.62)",
    "line": "rgba(143, 178, 225, 0.22)",
    "line_strong": "rgba(116, 169, 255, 0.48)",
    "blue": "#2f80ed",
    "blue_2": "#60a5fa",
    "cyan": "#22d3ee",
    "text": "#f8fbff",
    "muted": "#a9bad3",
    "success": "#22c55e",
    "warning": "#f59e0b",
    "page_width": 1320,
    "radius": 18,
    "card_opacity": 72,
    "show_grid": True,
    "compact_home": True,
    "css_extra": "",
}
logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("central_operacional")

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Central Operacional de Analises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# ARQUIVOS / LOGO
# =========================================================
def normalizar_nome_arquivo(nome: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", nome.lower())


def encontrar_arquivo(candidatos: list[str]) -> Path | None:
    for nome in candidatos:
        caminho = BASE_DIR / nome
        if caminho.exists():
            return caminho

    nomes_normalizados = {
        normalizar_nome_arquivo(p.name): p
        for p in BASE_DIR.iterdir()
        if p.is_file()
    }
    for nome in candidatos:
        chave = normalizar_nome_arquivo(nome)
        if chave in nomes_normalizados:
            return nomes_normalizados[chave]
    return None


def encontrar_logo() -> Path | None:
    return encontrar_arquivo([
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
    if not caminho.exists():
        return ""
    return base64.b64encode(caminho.read_bytes()).decode("utf-8")


def imagem_base64(caminho: Path | None) -> str:
    if not caminho or not caminho.exists():
        return ""
    try:
        return imagem_base64_cache(str(caminho), caminho.stat().st_mtime)
    except Exception as exc:
        logger.exception("Falha ao converter logo para base64: %s", exc)
        return ""


def mime_arquivo(caminho: Path | None) -> str:
    if not caminho:
        return "image/png"
    mime, _ = mimetypes.guess_type(str(caminho))
    return mime or "image/png"


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
        HISTORY_FILE.write_text(json.dumps(historico[:200], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.exception("Falha ao registrar historico: %s", exc)


LOGO_PATH = encontrar_logo()
LOGO_B64 = imagem_base64(LOGO_PATH)
LOGO_MIME = mime_arquivo(LOGO_PATH)

if LOGO_B64:
    BRAND_MARK = f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="brand-img" alt="Expresso Nepomuceno">'
else:
    BRAND_MARK = '<div class="brand-circle">E.N</div>'

# =========================================================
# ESTADO / NAVEGACAO
# =========================================================
def ir_para(pagina: str) -> None:
    st.session_state["pagina_atual"] = pagina
    st.rerun()


def pagina_atual_default() -> None:
    if "pagina_atual" not in st.session_state:
        st.session_state["pagina_atual"] = "inicio"



# =========================================================
# CONFIGURADOR VISUAL DO USUARIO
# =========================================================
def normalizar_hex(cor: str, fallback: str) -> str:
    if isinstance(cor, str) and re.fullmatch(r"#[0-9A-Fa-f]{6}", cor.strip()):
        return cor.strip()
    return fallback


def carregar_layout_config() -> dict[str, Any]:
    cfg = DEFAULT_LAYOUT_CONFIG.copy()
    if LAYOUT_FILE.exists():
        try:
            dados = json.loads(LAYOUT_FILE.read_text(encoding="utf-8"))
            if isinstance(dados, dict):
                cfg.update({k: v for k, v in dados.items() if k in cfg})
        except Exception as exc:
            logger.warning("Nao foi possivel carregar layout_config.json: %s", exc)
    return cfg


def salvar_layout_config(cfg: dict[str, Any]) -> None:
    seguro = DEFAULT_LAYOUT_CONFIG.copy()
    seguro.update({k: v for k, v in cfg.items() if k in seguro})
    LAYOUT_FILE.write_text(json.dumps(seguro, ensure_ascii=False, indent=2), encoding="utf-8")


def rgba_de_hex(hex_color: str, opacity_percent: int) -> str:
    hex_color = normalizar_hex(hex_color, "#091426").lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    a = max(0, min(100, int(opacity_percent))) / 100
    return f"rgba({r}, {g}, {b}, {a:.2f})"


def layout_atual() -> dict[str, Any]:
    if "layout_config" not in st.session_state:
        st.session_state["layout_config"] = carregar_layout_config()
    return dict(st.session_state["layout_config"])


def aplicar_layout_custom_css() -> None:
    cfg = layout_atual()
    card_rgba = rgba_de_hex(cfg.get("bg_1", "#061426"), int(cfg.get("card_opacity", 72)))
    grid_css = "" if cfg.get("show_grid", True) else "background-image: none !important; opacity: 0 !important;"
    home_min_height = "250px" if cfg.get("compact_home", True) else "455px"
    css = f"""
<style>
:root {{
    --bg-0: {normalizar_hex(str(cfg.get('bg_0')), '#020814')};
    --bg-1: {normalizar_hex(str(cfg.get('bg_1')), '#061426')};
    --glass: {cfg.get('glass', 'rgba(9, 20, 38, 0.72)')};
    --glass-2: {cfg.get('glass_2', 'rgba(16, 34, 60, 0.62)')};
    --line: {cfg.get('line', 'rgba(143, 178, 225, 0.22)')};
    --line-strong: {cfg.get('line_strong', 'rgba(116, 169, 255, 0.48)')};
    --blue: {normalizar_hex(str(cfg.get('blue')), '#2f80ed')};
    --blue-2: {normalizar_hex(str(cfg.get('blue_2')), '#60a5fa')};
    --cyan: {normalizar_hex(str(cfg.get('cyan')), '#22d3ee')};
    --text: {normalizar_hex(str(cfg.get('text')), '#f8fbff')};
    --muted: {normalizar_hex(str(cfg.get('muted')), '#a9bad3')};
    --success: {normalizar_hex(str(cfg.get('success')), '#22c55e')};
    --warning: {normalizar_hex(str(cfg.get('warning')), '#f59e0b')};
}}
.block-container {{ max-width: {int(cfg.get('page_width', 1320))}px !important; }}
.hero {{ min-height: {home_min_height} !important; border-radius: {int(cfg.get('radius', 18)) + 6}px !important; }}
.glass-card, .visual-card, .metric-card, .upload-card, .page-shell, .module-glass {{
    border-radius: {int(cfg.get('radius', 18))}px !important;
}}
.glass-card, .visual-card, .upload-card {{ background: linear-gradient(135deg, {card_rgba}, rgba(5, 15, 31, .62)) !important; }}
.stApp::before {{ {grid_css} }}

/* Correção definitiva do upload branco */
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"] {{
    background: rgba(5, 16, 32, .82) !important;
    border: 1px dashed rgba(151,188,240,.45) !important;
}}
[data-testid="stFileUploader"] button,
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploader"] button p,
[data-testid="stFileUploaderDropzone"] button p {{
    background: linear-gradient(135deg, var(--blue), #1d4ed8) !important;
    color: #ffffff !important;
    border: 1px solid rgba(96,165,250,.72) !important;
    opacity: 1 !important;
    font-weight: 900 !important;
}}
[data-testid="stFileUploader"] small,
[data-testid="stFileUploaderDropzone"] small,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploaderDropzone"] span,
[data-testid="stFileUploader"] p,
[data-testid="stFileUploaderDropzone"] p {{
    color: #dbeafe !important;
    opacity: 1 !important;
}}

/* Inputs do editor sempre legíveis */
.stTextInput input, .stNumberInput input, .stTextArea textarea,
div[data-baseweb="select"] > div, div[data-baseweb="select"] span {{
    background-color: rgba(5, 16, 32, .94) !important;
    color: #f8fbff !important;
    -webkit-text-fill-color: #f8fbff !important;
}}
.stCheckbox label p, .stCheckbox label span, label p {{ color: #dbeafe !important; opacity: 1 !important; }}
</style>
"""
    st.markdown(css, unsafe_allow_html=True)

# =========================================================
# CSS - CONCEITO DARK / GLASSMORPHISM
# =========================================================
def aplicar_css() -> None:
    st.markdown(
        """
<style>
:root {
    --bg-0: #020814;
    --bg-1: #061426;
    --glass: rgba(9, 20, 38, 0.72);
    --glass-2: rgba(16, 34, 60, 0.62);
    --line: rgba(143, 178, 225, 0.22);
    --line-strong: rgba(116, 169, 255, 0.48);
    --blue: #2f80ed;
    --blue-2: #60a5fa;
    --cyan: #22d3ee;
    --text: #f8fbff;
    --muted: #a9bad3;
    --success: #22c55e;
    --warning: #f59e0b;
    --danger: #ef4444;
}

html, body, [class*="css"], .stApp {
    font-family: "Segoe UI", Roboto, Arial, sans-serif;
}

.stApp {
    color: var(--text);
    background:
        radial-gradient(circle at 18% 18%, rgba(47,128,237,.24), transparent 26%),
        radial-gradient(circle at 83% 20%, rgba(34,211,238,.13), transparent 25%),
        linear-gradient(rgba(2,8,20,.88), rgba(2,8,20,.92)),
        linear-gradient(135deg, #041025 0%, #061a33 45%, #020814 100%);
    background-attachment: fixed;
}

.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background:
      linear-gradient(90deg, rgba(255,255,255,.035) 1px, transparent 1px),
      linear-gradient(0deg, rgba(255,255,255,.025) 1px, transparent 1px);
    background-size: 72px 72px;
    mask-image: radial-gradient(circle at 50% 20%, black, transparent 75%);
    z-index: 0;
}

.block-container {
    max-width: 1320px;
    padding-top: 1.2rem;
    padding-bottom: 4.8rem;
    position: relative;
    z-index: 1;
}

#MainMenu, footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent !important;}

/* Streamlit widgets em dark mode */
main label, main label p, main .stMarkdown, main .stMarkdown p,
main div[data-testid="stMarkdownContainer"], main div[data-testid="stCaptionContainer"] {
    color: var(--text) !important;
}
main small, main .stCaptionContainer, .muted, .card-desc, .module-desc {
    color: var(--muted) !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: rgba(2, 8, 20, .92);
    border-right: 1px solid var(--line);
}
section[data-testid="stSidebar"] * {color: var(--text) !important;}
section[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    justify-content: flex-start;
    border-radius: 14px !important;
    min-height: 46px;
    border: 1px solid rgba(143,178,225,.16) !important;
    background: rgba(255,255,255,.04) !important;
    color: var(--text) !important;
    font-weight: 780 !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(47,128,237,.25) !important;
    border-color: rgba(96,165,250,.55) !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, rgba(47,128,237,.92), rgba(29,78,216,.78)) !important;
    border-color: rgba(96,165,250,.65) !important;
}

/* Top bar */
.top-shell {
    border: 1px solid var(--line);
    background: rgba(2, 8, 20, .58);
    backdrop-filter: blur(18px);
    border-radius: 20px;
    padding: 18px 22px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 22px;
    box-shadow: 0 24px 70px rgba(0,0,0,.24);
    margin-bottom: 28px;
}
.brand-wrap {display:flex; align-items:center; gap:14px;}
.brand-circle {
    width: 52px; height: 52px; border-radius: 999px;
    display:flex; align-items:center; justify-content:center;
    color:#fff; font-weight:950; letter-spacing:.03em;
    border: 3px solid #2f80ed;
    background: radial-gradient(circle at 35% 35%, rgba(96,165,250,.32), rgba(8,20,40,.92));
    box-shadow: 0 0 0 5px rgba(47,128,237,.08), 0 0 28px rgba(47,128,237,.3);
}
.brand-img {width: 58px; height: 58px; object-fit: contain; border-radius: 14px; background: rgba(255,255,255,.08); padding: 4px;}
.brand-title {font-size: 17px; font-weight: 900; color: #fff; line-height: 1.15;}
.brand-subtitle {font-size: 12px; font-weight: 700; color: var(--muted); margin-top: 4px; letter-spacing: .02em;}
.top-actions {display:flex; align-items:center; gap: 13px;}
.status-chip {
    display:inline-flex; align-items:center; gap:8px;
    padding: 9px 14px; border-radius: 999px;
    border: 1px solid var(--line);
    background: rgba(255,255,255,.045);
    color: var(--text); font-size: 12px; font-weight: 800;
}
.status-dot {width:7px; height:7px; border-radius:999px; background: var(--success); box-shadow: 0 0 12px rgba(34,197,94,.95);}
.user-chip {
    width: 42px; height: 42px; border-radius: 999px; display:flex; align-items:center; justify-content:center;
    border: 1px solid var(--line); background: rgba(255,255,255,.05); font-weight:900;
}

/* Hero */
.hero {
    min-height: 455px;
    border: 1px solid var(--line-strong);
    border-radius: 24px;
    background:
        linear-gradient(rgba(2,8,20,.38), rgba(2,8,20,.84)),
        radial-gradient(circle at 20% 45%, rgba(47,128,237,.30), transparent 24%),
        radial-gradient(circle at 80% 30%, rgba(96,165,250,.18), transparent 26%),
        linear-gradient(135deg, rgba(8,18,34,.90), rgba(5,15,31,.82));
    box-shadow: 0 26px 90px rgba(0,0,0,.36), inset 0 1px 0 rgba(255,255,255,.08);
    overflow: hidden;
    position: relative;
    padding: 44px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    margin-bottom: 20px;
}
.hero::before {
    content:"";
    position:absolute; inset:0;
    background:
      radial-gradient(circle at 25% 72%, rgba(0,0,0,.45), transparent 20%),
      linear-gradient(110deg, rgba(15,23,42,.8), transparent 45%),
      repeating-linear-gradient(90deg, rgba(255,255,255,.025) 0, rgba(255,255,255,.025) 1px, transparent 1px, transparent 80px);
    opacity:.85;
}
.hero-content {position:relative; z-index:1; text-align:center; width:100%;}
.hero-kicker {font-size: 13px; letter-spacing: .42em; text-transform: uppercase; color: #8dbbff; font-weight: 850; margin-bottom: 12px;}
.hero-title {font-size: clamp(36px, 4vw, 58px); color:#fff; font-weight: 300; line-height:1.05; margin: 0;}
.hero-title strong {font-weight: 900; color: #b8d8ff;}
.hero-subtitle {color: #d4e3f8; font-size: 17px; margin-top: 14px;}

.module-glass {
    position:relative; z-index:1;
    width: min(760px, 100%);
    margin: 34px auto 0 auto;
    background: linear-gradient(135deg, rgba(31,47,76,.70), rgba(14,24,43,.66));
    border: 1px solid rgba(151,188,240,.34);
    border-radius: 18px;
    box-shadow: 0 28px 70px rgba(0,0,0,.34), inset 0 1px 0 rgba(255,255,255,.08);
    backdrop-filter: blur(18px);
    overflow: hidden;
}
.module-grid {
    display:grid;
    grid-template-columns: repeat(3, 1fr);
}
.module-item {
    min-height: 150px;
    padding: 24px 18px;
    border-right: 1px solid rgba(151,188,240,.16);
    border-bottom: 1px solid rgba(151,188,240,.16);
    text-align:center;
    transition: .18s ease;
}
.module-item:nth-child(3n) {border-right: 0;}
.module-item:hover {background: rgba(47,128,237,.18); transform: translateY(-1px);}
.module-icon {font-size: 42px; color:#75aeff; margin-bottom: 13px; filter: drop-shadow(0 0 15px rgba(96,165,250,.28));}
.module-title {color:#fff; font-weight: 850; font-size: 15px; margin-bottom: 6px;}
.module-desc {font-size: 12px; line-height: 1.4; max-width: 180px; margin: 0 auto;}
.home-actions {width:min(760px,100%); margin: 10px auto 0 auto; position:relative; z-index:2;}
.home-actions .stButton > button {
    min-height: 42px;
    border-radius: 0 !important;
    background: transparent !important;
    border: 0 !important;
    color: rgba(255,255,255,.0) !important;
    box-shadow:none !important;
}
.home-actions .stButton > button:hover {background: transparent !important;}

/* Paginas internas */
.page-shell {
    border: 1px solid var(--line-strong);
    border-radius: 24px;
    background: rgba(2,8,20,.54);
    backdrop-filter: blur(18px);
    box-shadow: 0 26px 90px rgba(0,0,0,.28), inset 0 1px 0 rgba(255,255,255,.08);
    overflow:hidden;
    margin-bottom: 22px;
}
.page-head {
    padding: 24px 28px;
    border-bottom: 1px solid rgba(151,188,240,.16);
    display:flex; align-items:center; justify-content:space-between; gap:18px;
    background: linear-gradient(135deg, rgba(8,28,55,.82), rgba(9,18,35,.48));
}
.page-title-wrap {display:flex; align-items:center; gap:16px;}
.back-dot {
    width: 42px; height:42px; border-radius:14px; display:flex; align-items:center; justify-content:center;
    color:#9dc6ff; background:rgba(47,128,237,.12); border:1px solid rgba(96,165,250,.28); font-weight:900;
}
.page-kicker {font-size:12px; color:#8dbbff; font-weight:850; letter-spacing:.14em; text-transform:uppercase; margin-bottom:6px;}
.page-title {font-size: 25px; line-height:1.1; margin:0; color:#fff; font-weight:900;}
.page-subtitle {font-size: 14px; color: var(--muted); margin-top:6px;}
.page-body {padding: 26px 28px 28px 28px;}

.glass-card {
    background: linear-gradient(135deg, rgba(12,26,48,.74), rgba(13,23,40,.56));
    border: 1px solid rgba(151,188,240,.18);
    border-radius: 18px;
    padding: 19px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.045), 0 12px 30px rgba(0,0,0,.16);
    margin-bottom: 16px;
}
.card-title {font-size: 17px; color:#fff; font-weight:900; margin-bottom: 5px;}
.card-desc {font-size: 13px; line-height: 1.55;}
.metric-row {display:grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 17px;}
.metric-card {background: rgba(3,12,25,.55); border:1px solid rgba(151,188,240,.18); border-radius:16px; padding:16px;}
.metric-label {font-size:12px; color:var(--muted); font-weight:750;}
.metric-value {font-size:26px; color:#fff; font-weight:950; margin-top:5px;}
.upload-card {background: rgba(3,12,25,.50); border:1px solid rgba(151,188,240,.18); border-radius:18px; padding:17px; margin-bottom:16px;}
.upload-head {display:flex; align-items:flex-start; justify-content:space-between; gap:12px; margin-bottom:10px;}
.step-badge {width:34px; height:34px; border-radius:12px; display:flex; align-items:center; justify-content:center; color:#fff; background:rgba(47,128,237,.26); border:1px solid rgba(96,165,250,.40); font-weight:950;}
.upload-title {font-size:15px; font-weight:900; color:#fff;}
.upload-desc {font-size:12px; color:var(--muted); margin-top:3px;}
.pill {display:inline-flex; align-items:center; gap:7px; padding:7px 10px; border-radius:999px; font-size:12px; font-weight:850; border:1px solid rgba(151,188,240,.18); background:rgba(255,255,255,.04); color:#dceafe;}
.pill-ok {color:#bbf7d0; border-color: rgba(34,197,94,.35); background:rgba(34,197,94,.09);}
.pill-wait {color:#fed7aa; border-color: rgba(245,158,11,.32); background:rgba(245,158,11,.08);}

/* Botoes Streamlit */
.stButton > button, .stDownloadButton > button {
    border-radius: 13px !important;
    border: 1px solid rgba(151,188,240,.25) !important;
    background: rgba(255,255,255,.055) !important;
    color: #f8fbff !important;
    font-weight: 850 !important;
    min-height: 44px;
    box-shadow: none !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    border-color: rgba(96,165,250,.68) !important;
    background: rgba(47,128,237,.22) !important;
}
.stButton > button[kind="primary"], .stDownloadButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2f80ed) !important;
    border-color: rgba(96,165,250,.65) !important;
}
[data-testid="stFileUploader"] section {
    border-radius: 14px;
    border: 1px dashed rgba(151,188,240,.38);
    background: rgba(255,255,255,.035);
}
[data-testid="stFileUploader"] button {border-radius: 11px !important;}
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background: rgba(255,255,255,.06) !important;
    color: #fff !important;
    border-color: rgba(151,188,240,.25) !important;
}
.stProgress > div > div > div > div {background: linear-gradient(90deg, #1d4ed8, #22d3ee) !important;}
[data-testid="stDataFrame"] {border-radius: 16px; overflow:hidden; border:1px solid rgba(151,188,240,.18);}

/* Navegacao rapida */
.quick-nav {margin: 0 0 20px 0;}
.quick-title {font-size: 11px; letter-spacing:.22em; color:#8dbbff; font-weight:900; text-transform:uppercase; margin-bottom:8px;}

@media (max-width: 900px) {
    .top-shell, .page-head {display:block;}
    .top-actions {margin-top:14px;}
    .hero {padding: 28px 18px; min-height: auto;}
    .module-grid {grid-template-columns: 1fr;}
    .module-item, .module-item:nth-child(3n) {border-right:0;}
    .metric-row {grid-template-columns: 1fr 1fr;}
}


/* Ajustes de correção: toolbar, inputs e home clicável */
[data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
    visibility: hidden !important;
    height: 0 !important;
    position: fixed !important;
}

.hero-compact {
    min-height: 265px !important;
    margin-bottom: 24px !important;
}

.home-section-title {
    color: #ffffff !important;
    font-size: 24px;
    font-weight: 900;
    margin: 8px 0 6px 0;
}
.home-section-subtitle {
    color: var(--muted) !important;
    font-size: 15px;
    margin-bottom: 18px;
}
.visual-card {
    min-height: 178px;
    border: 1px solid rgba(151,188,240,.28);
    border-radius: 20px;
    background: linear-gradient(135deg, rgba(31,47,76,.62), rgba(8,18,34,.68));
    box-shadow: 0 18px 55px rgba(0,0,0,.24), inset 0 1px 0 rgba(255,255,255,.06);
    backdrop-filter: blur(16px);
    padding: 24px 22px;
    margin-bottom: 10px;
    text-align: center;
}
.visual-card-muted { opacity: .94; }
.visual-icon {
    font-size: 42px;
    color: #75aeff;
    line-height: 1;
    margin-bottom: 14px;
}
.visual-title {
    color: #ffffff !important;
    font-size: 20px;
    font-weight: 920;
    margin-bottom: 8px;
}
.visual-desc {
    color: #a9bad3 !important;
    font-size: 14px;
    line-height: 1.5;
}

/* Inputs escuros e legíveis */
.stTextInput input,
.stNumberInput input,
.stTextArea textarea,
div[data-baseweb="select"] > div {
    background: rgba(5, 16, 32, .92) !important;
    color: #f8fbff !important;
    border: 1px solid rgba(151,188,240,.32) !important;
    border-radius: 12px !important;
}
.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: rgba(248,251,255,.55) !important;
}
.stCheckbox label, .stCheckbox label p {
    color: #f8fbff !important;
}
.stSelectbox label, .stTextInput label, .stNumberInput label, .stTextArea label, .stFileUploader label,
.stSelectbox label p, .stTextInput label p, .stNumberInput label p, .stTextArea label p, .stFileUploader label p {
    color: #dbeafe !important;
    opacity: 1 !important;
}

</style>
""",
        unsafe_allow_html=True,
    )


aplicar_css()

# =========================================================
# MODULOS DINAMICOS E REGISTRO
# =========================================================
PROTEGIDOS = {
    "codigo_colado.py",
    "código_colado.py",
    "codigo colado.py",
    "código colado.py",
    "odometro_v12_com_percentual.py",
    "app.py",
    APP_FILE.lower(),
    "app_home_profissional.py",
    "app_profissional_final.py",
    "app_clean_editor.py",
    "app_home_limpa_click.py",
    "app_menu_fix.py",
    "app_home_nav_fix.py",
}


def carregar_registro() -> dict[str, Any]:
    if not REGISTRY_FILE.exists():
        return {}
    try:
        dados = json.loads(REGISTRY_FILE.read_text(encoding="utf-8"))
        return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


def salvar_registro(dados: dict[str, Any]) -> None:
    REGISTRY_FILE.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".py", "")
    nome = nome.replace("_", " ").replace("-", " ")
    return " ".join(parte.capitalize() for parte in nome.split())


def listar_arquivos_configuraveis() -> list[Path]:
    arquivos: list[Path] = []
    for caminho in sorted(BASE_DIR.glob("*.py"), key=lambda p: p.name.lower()):
        lower = caminho.name.lower()
        if lower in PROTEGIDOS:
            continue
        if lower.startswith(("_", "test_", "streamlit ", "python ", "pip ", "py ")):
            continue
        arquivos.append(caminho)
    return arquivos


@st.cache_data(show_spinner=False)
def inspecionar_script(caminho: Path, mtime: float = 0.0) -> dict[str, Any]:
    info = {"erro": "", "main_streamlit": False, "main": False}
    try:
        texto = caminho.read_text(encoding="utf-8")
        arvore = ast.parse(texto)
    except UnicodeDecodeError:
        info["erro"] = "Arquivo nao esta em UTF-8."
        return info
    except SyntaxError as exc:
        info["erro"] = f"Erro de sintaxe na linha {exc.lineno}: {exc.msg}"
        return info
    except Exception as exc:
        info["erro"] = str(exc)
        return info

    for node in arvore.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.name == "main_streamlit":
                info["main_streamlit"] = True
            elif node.name == "main":
                info["main"] = True
    return info


def modulos_ativos() -> list[dict[str, Any]]:
    registro = carregar_registro()
    saida: list[dict[str, Any]] = []
    usados: set[str] = set()
    for caminho in listar_arquivos_configuraveis():
        cfg = registro.get(caminho.name, {}) if isinstance(registro.get(caminho.name, {}), dict) else {}
        info = inspecionar_script(caminho, caminho.stat().st_mtime)
        ativo = bool(cfg.get("ativo", False))
        if not ativo:
            continue
        if info["erro"]:
            continue
        slug = re.sub(r"[^a-z0-9]+", "_", str(cfg.get("nome") or caminho.stem).lower()).strip("_")
        if slug in usados:
            continue
        usados.add(slug)
        saida.append({
            "arquivo": caminho.name,
            "caminho": caminho,
            "nome": str(cfg.get("nome") or nome_amigavel_script(caminho.name)),
            "icone": str(cfg.get("icone") or "🧩"),
            "descricao": str(cfg.get("descricao") or "Modulo operacional adicionado ao portal."),
            "ordem": int(cfg.get("ordem", 100)),
            "modo": str(cfg.get("modo", "auto")),
            "slug": slug,
        })
    return sorted(saida, key=lambda m: (m["ordem"], m["nome"]))


# =========================================================
# CARREGAMENTO DE MODULOS PYTHON
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
        raise ImportError(f"Nao foi possivel carregar o modulo: {caminho_script.name}")
    modulo = importlib.util.module_from_spec(spec)
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop
        spec.loader.exec_module(modulo)
    finally:
        st.set_page_config = original_set_page_config
    return modulo


def executar_modulo_dinamico(caminho: Path, modo: str = "auto") -> None:
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop
        if modo == "script":
            runpy.run_path(str(caminho), run_name="__main__")
            return
        modulo = carregar_modulo_por_arquivo(caminho)
        if modo == "main_streamlit":
            if not hasattr(modulo, "main_streamlit"):
                raise AttributeError("O modulo nao possui main_streamlit().")
            modulo.main_streamlit()
        elif modo == "main":
            if not hasattr(modulo, "main"):
                raise AttributeError("O modulo nao possui main().")
            modulo.main()
        else:
            if hasattr(modulo, "main_streamlit"):
                modulo.main_streamlit()
            elif hasattr(modulo, "main"):
                modulo.main()
            else:
                runpy.run_path(str(caminho), run_name="__main__")
    finally:
        st.set_page_config = original_set_page_config


def validar_funcoes_modulo(modulo, funcoes_obrigatorias: list[str]) -> list[str]:
    return [f for f in funcoes_obrigatorias if not hasattr(modulo, f)]


# =========================================================
# COMPONENTES VISUAIS
# =========================================================
def render_topbar() -> None:
    data_atual = datetime.now().strftime("%d/%m/%Y")
    st.markdown(
        f"""
        <div class="top-shell">
            <div class="brand-wrap">
                {BRAND_MARK}
                <div>
                    <div class="brand-title">Central Operacional de Análises</div>
                    <div class="brand-subtitle">Expresso Nepomuceno</div>
                </div>
            </div>
            <div class="top-actions">
                <div class="status-chip"><span class="status-dot"></span> Ambiente interno</div>
                <div class="status-chip">{data_atual}</div>
                <div class="user-chip">EN</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def nav_buttons() -> None:
    st.markdown('<div class="quick-nav"><div class="quick-title">Navegação rápida</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1:
        if st.button("⌂ Início", use_container_width=True):
            ir_para("inicio")
    with c2:
        if st.button("◷ Permanência", use_container_width=True):
            ir_para("permanencia")
    with c3:
        if st.button("⏱ Odômetro", use_container_width=True):
            ir_para("odometro")
    with c4:
        if st.button("⚙ Editor", use_container_width=True):
            ir_para("editor")
    with c5:
        if st.button("▤ Histórico", use_container_width=True):
            ir_para("historico")
    with c6:
        if st.button("✓ Roadmap", use_container_width=True):
            ir_para("roadmap")
    with c7:
        if st.button("🎨 Layout", use_container_width=True):
            ir_para("layout")


def render_sidebar(modulos: list[dict[str, Any]]) -> None:
    with st.sidebar:
        st.markdown(
            f"""
            <div style="padding:8px 2px 18px 2px;">
                <div style="display:flex;align-items:center;gap:12px;">{BRAND_MARK}
                    <div><b>Central Operacional</b><br><span style="font-size:12px;color:#a9bad3;">Expresso Nepomuceno</span></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⌂ Início", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "inicio" else "secondary"):
            ir_para("inicio")
        if st.button("◷ Análise de Permanência", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "permanencia" else "secondary"):
            ir_para("permanencia")
        if st.button("⏱ Odômetro V12", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "odometro" else "secondary"):
            ir_para("odometro")
        if st.button("⚙ Editor de Módulos", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "editor" else "secondary"):
            ir_para("editor")
        if st.button("▤ Histórico", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "historico" else "secondary"):
            ir_para("historico")
        if st.button("✓ Roadmap Técnico", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "roadmap" else "secondary"):
            ir_para("roadmap")
        if st.button("🎨 Editor de Layout", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == "layout" else "secondary"):
            ir_para("layout")
        if modulos:
            st.markdown("---")
            st.caption("Módulos adicionais")
            for m in modulos:
                chave = f"mod_{m['slug']}"
                if st.button(f"{m['icone']} {m['nome']}", key=f"side_{chave}", use_container_width=True, type="primary" if st.session_state.get("pagina_atual") == chave else "secondary"):
                    st.session_state["modulo_ativo"] = m["arquivo"]
                    ir_para(chave)


def page_shell_open(titulo: str, subtitulo: str, icone: str = "◉") -> None:
    st.markdown(
        f"""
        <div class="page-shell">
            <div class="page-head">
                <div class="page-title-wrap">
                    <div class="back-dot">{html.escape(icone)}</div>
                    <div>
                        <div class="page-kicker">Módulo operacional</div>
                        <h1 class="page-title">{html.escape(titulo)}</h1>
                        <div class="page-subtitle">{html.escape(subtitulo)}</div>
                    </div>
                </div>
                <div class="status-chip"><span class="status-dot"></span> Online</div>
            </div>
            <div class="page-body">
        """,
        unsafe_allow_html=True,
    )


def page_shell_close() -> None:
    st.markdown("</div></div>", unsafe_allow_html=True)


def message(tipo: str, texto: str) -> None:
    cores = {
        "info": ("#bfdbfe", "rgba(47,128,237,.10)", "rgba(96,165,250,.28)"),
        "success": ("#bbf7d0", "rgba(34,197,94,.10)", "rgba(34,197,94,.28)"),
        "warning": ("#fed7aa", "rgba(245,158,11,.10)", "rgba(245,158,11,.28)"),
        "error": ("#fecaca", "rgba(239,68,68,.10)", "rgba(239,68,68,.28)"),
    }
    color, bg, border = cores.get(tipo, cores["info"])
    st.markdown(
        f'<div class="glass-card" style="color:{color};background:{bg};border-color:{border};font-weight:800;">{texto}</div>',
        unsafe_allow_html=True,
    )


def upload_card(numero: int, titulo: str, desc: str, key: str):
    st.markdown(
        f"""
        <div class="upload-card">
            <div class="upload-head">
                <div style="display:flex;gap:12px;align-items:flex-start;">
                    <div class="step-badge">{numero}</div>
                    <div>
                        <div class="upload-title">{html.escape(titulo)}</div>
                        <div class="upload-desc">{html.escape(desc)}</div>
                    </div>
                </div>
                <span class="pill pill-wait">Aguardando</span>
            </div>
        """,
        unsafe_allow_html=True,
    )
    arquivo = st.file_uploader("Selecionar arquivo", type=["xlsx", "xls"], key=key, label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)
    return arquivo


# =========================================================
# PAGINA INICIAL
# =========================================================
def home_button(label: str, destino: str, key: str, primary: bool = False) -> None:
    if st.button(label, key=key, use_container_width=True, type="primary" if primary else "secondary"):
        ir_para(destino)


def pagina_inicio(modulos: list[dict[str, Any]]) -> None:
    st.markdown(
        """
        <div class="hero hero-compact">
            <div class="hero-content">
                <div class="hero-kicker">Portal Executivo</div>
                <h1 class="hero-title">Central Operacional de <strong>Análises</strong></h1>
                <div class="hero-subtitle">Escolha uma categoria para acessar os módulos disponíveis.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="home-section-title">Escolha uma categoria</div>'
        '<div class="home-section-subtitle">Clique em uma opção abaixo para abrir a página operacional.</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="visual-card">
            <div class="visual-icon">◷</div>
            <div class="visual-title">Permanência</div>
            <div class="visual-desc">Tempos, eventos e padrões operacionais.</div>
        </div>
        """, unsafe_allow_html=True)
        home_button("Abrir Permanência", "permanencia", "home_btn_permanencia", True)
    with c2:
        st.markdown("""
        <div class="visual-card">
            <div class="visual-icon">⏱</div>
            <div class="visual-title">Odômetro</div>
            <div class="visual-desc">Leituras, bases e consolidação V12.</div>
        </div>
        """, unsafe_allow_html=True)
        home_button("Abrir Odômetro V12", "odometro", "home_btn_odometro", True)
    with c3:
        st.markdown("""
        <div class="visual-card">
            <div class="visual-icon">⚙</div>
            <div class="visual-title">Editor</div>
            <div class="visual-desc">Configure módulos sem alterar o código.</div>
        </div>
        """, unsafe_allow_html=True)
        home_button("Abrir Editor de Módulos", "editor", "home_btn_editor", True)

    c4, c5, c6 = st.columns(3)
    with c4:
        st.markdown("""
        <div class="visual-card visual-card-muted">
            <div class="visual-icon">⛽</div>
            <div class="visual-title">Combustível</div>
            <div class="visual-desc">Módulos adicionais configuráveis.</div>
        </div>
        """, unsafe_allow_html=True)
        if modulos:
            home_button("Ver Módulos Adicionais", f"mod_{modulos[0]['slug']}", "home_btn_extra", False)
        else:
            st.button("Configure no Editor", key="home_extra_disabled", use_container_width=True, disabled=True)
    with c5:
        st.markdown("""
        <div class="visual-card visual-card-muted">
            <div class="visual-icon">▣</div>
            <div class="visual-title">Bases</div>
            <div class="visual-desc">Arquivos de entrada e processamento.</div>
        </div>
        """, unsafe_allow_html=True)
        home_button("Ir para Odômetro", "odometro", "home_btn_bases", False)
    with c6:
        st.markdown("""
        <div class="visual-card visual-card-muted">
            <div class="visual-icon">▤</div>
            <div class="visual-title">Relatórios</div>
            <div class="visual-desc">Saídas tratadas em Excel.</div>
        </div>
        """, unsafe_allow_html=True)
        home_button("Ir para Permanência", "permanencia", "home_btn_relatorios", False)

    st.write("")
    metric_html = f"""
    <div class="metric-row">
        <div class="metric-card"><div class="metric-label">Módulos fixos</div><div class="metric-value">2</div></div>
        <div class="metric-card"><div class="metric-label">Módulos adicionais</div><div class="metric-value">{len(modulos)}</div></div>
        <div class="metric-card"><div class="metric-label">Visual</div><div class="metric-value">Pro</div></div>
        <div class="metric-card"><div class="metric-label">Ambiente</div><div class="metric-value">EN</div></div>
    </div>
    """
    st.markdown(metric_html, unsafe_allow_html=True)


# =========================================================
# PAGINA PERMANENCIA
# =========================================================
def atualizar_progresso(barra, status, pct: int, texto: str) -> None:
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.15)


def validar_upload_excel(arquivo, colunas_obrigatorias: list[str] | None = None) -> tuple[bool, str]:
    """Valida o arquivo de entrada antes de processar sem alterar a regra de negócio."""
    if arquivo is None:
        return False, "Arquivo não enviado."
    nome = getattr(arquivo, "name", "")
    if not nome.lower().endswith((".xlsx", ".xls")):
        return False, "Formato inválido. Envie um arquivo .xlsx ou .xls."
    try:
        # Validação leve: confirma que o arquivo pode ser aberto como Excel.
        import pandas as pd
        pos = arquivo.tell() if hasattr(arquivo, "tell") else 0
        arquivo.seek(0)
        amostra = pd.read_excel(arquivo, nrows=5)
        arquivo.seek(pos)
        if colunas_obrigatorias:
            faltando = [c for c in colunas_obrigatorias if c not in amostra.columns]
            if faltando:
                return False, "Colunas obrigatórias ausentes: " + ", ".join(faltando)
        return True, "Arquivo Excel validado."
    except Exception as exc:
        logger.exception("Falha ao validar upload Excel %s: %s", nome, exc)
        try:
            arquivo.seek(0)
        except Exception:
            pass
        return False, f"Não foi possível abrir o Excel enviado: {exc}"


def pagina_permanencia() -> None:
    page_shell_open("Análise de Permanência", "Processamento da base de permanência com classificação por tempo configurável.", "◷")

    caminho_permanencia = encontrar_arquivo(["Codigo_colado.py", "Código_colado.py", "Codigo colado.py", "Código colado.py"])
    if caminho_permanencia is None:
        message("error", "Arquivo Codigo_colado.py não encontrado na pasta do app.")
        page_shell_close()
        return

    try:
        permanencia = carregar_modulo_por_arquivo(caminho_permanencia)
    except Exception as e:
        message("error", f"Erro ao importar {html.escape(caminho_permanencia.name)}: {html.escape(str(e))}")
        page_shell_close()
        return

    faltando = validar_funcoes_modulo(permanencia, [
        "carregar_dados",
        "identificar_eventos_carregamento",
        "montar_ciclos_carregamento",
        "gerar_resumos",
        "salvar_saida",
    ])
    if faltando:
        message("error", "Funções ausentes no módulo de permanência: " + ", ".join(faltando))
        page_shell_close()
        return

    st.markdown('<div class="glass-card"><div class="card-title">Parâmetros de tratamento</div><div class="card-desc">Defina a faixa aceitável para classificação dos tempos antes de processar a base.</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        tempo_minimo = st.number_input("Tempo mínimo aceitável (minutos)", min_value=0, value=15, step=1)
    with c2:
        tempo_maximo = st.number_input("Tempo máximo aceitável (minutos)", min_value=1, value=55, step=1)
    st.markdown('</div>', unsafe_allow_html=True)

    if tempo_maximo <= tempo_minimo:
        message("error", "O tempo máximo precisa ser maior que o mínimo.")

    st.markdown('<div class="glass-card"><div class="card-title">Importação da base</div><div class="card-desc">Envie o arquivo Excel de permanência para iniciar o processamento.</div>', unsafe_allow_html=True)
    arquivo = st.file_uploader("Selecione o Excel de permanência", type=["xlsx", "xls"], key="upload_permanencia")
    st.markdown('</div>', unsafe_allow_html=True)

    if not arquivo:
        message("warning", "Aguardando upload da base de permanência para iniciar o processamento.")
        page_shell_close()
        return

    valido, msg_validacao = validar_upload_excel(arquivo)
    if not valido:
        message("error", html.escape(msg_validacao))
        page_shell_close()
        return
    message("success", f"Arquivo carregado e validado: <b>{html.escape(arquivo.name)}</b>")

    if st.button("Processar Permanência", use_container_width=True, type="primary"):
        if tempo_maximo <= tempo_minimo:
            st.error("Corrija os tempos antes de processar.")
            page_shell_close()
            return
        inicio = time.time()
        barra = st.progress(0)
        status = st.empty()
        caminho = None
        saida = None
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
            duracao = round(time.time() - inicio, 2)
            registrar_historico("Permanência", arquivo.name, "Sucesso", duracao, f"Eventos: {len(df_resultado)}")
            st.success(f"Processo finalizado em {duracao} segundos")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Eventos", len(df_resultado))
            k2.metric("Alertas", 0 if df_alertas.empty else len(df_alertas))
            k3.metric("OK", 0 if df_resultado.empty else int((df_resultado["Status"] == "OK").sum()))
            k4.metric("Improcedentes", 0 if df_resultado.empty else int((df_resultado["Status"] == "IMPROCEDENTE").sum()))
            st.markdown("### Prévia dos resultados")
            if not df_resultado.empty:
                st.dataframe(df_resultado.head(100), use_container_width=True)
            with open(saida, "rb") as f:
                dados_saida = f.read()
            st.download_button(
                "Baixar Excel Permanência",
                dados_saida,
                file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            logger.exception("Erro ao processar permanencia")
            registrar_historico("Permanência", arquivo.name, "Erro", time.time() - inicio, str(e))
            st.error(f"Erro ao processar permanência: {e}")
        finally:
            limpar_arquivo_temporario(caminho)
            limpar_arquivo_temporario(saida)
    page_shell_close()


# =========================================================
# PAGINA ODOMETRO
# =========================================================
def pagina_odometro() -> None:
    page_shell_open("Odômetro V12", "Consolidação do ODOMETRO_MATCH com as quatro bases obrigatórias.", "⏱")
    message("info", "Envie as quatro bases obrigatórias para gerar o Excel final do odômetro.")

    c1, c2 = st.columns(2)
    with c1:
        comb = upload_card(1, "Base Combustível", "Arquivo de abastecimentos / combustível.", "comb")
        ativo = upload_card(3, "Base Ativo de Veículos", "Cadastro de ativos, frota e veículos.", "ativo")
    with c2:
        maxtrack = upload_card(2, "Base Km Rodado Maxtrack", "Arquivo de leituras e quilometragem.", "maxtrack")
        producao = upload_card(4, "Produção Oficial / Cliente", "Base oficial de produção do cliente.", "producao")

    qtd = sum(a is not None for a in [comb, maxtrack, ativo, producao])
    st.progress(qtd / 4)
    st.caption(f"{qtd} de 4 arquivos carregados")

    if not (comb and maxtrack and ativo and producao):
        message("warning", "Aguardando upload das quatro bases para liberar o processamento.")
        page_shell_close()
        return

    validacoes = []
    for arq in [comb, maxtrack, ativo, producao]:
        valido, msg = validar_upload_excel(arq)
        validacoes.append((valido, msg, arq.name))
    erros_validacao = [f"{nome}: {msg}" for valido, msg, nome in validacoes if not valido]
    if erros_validacao:
        message("error", "<br>".join(html.escape(e) for e in erros_validacao))
        page_shell_close()
        return
    message("success", "Todas as bases foram carregadas e validadas. O processamento já pode ser iniciado.")

    if st.button("Processar Odômetro V12", use_container_width=True, type="primary"):
        inicio = time.time()
        barra = st.progress(0)
        status = st.empty()
        log_box = st.empty()
        p1 = p2 = p3 = p4 = saida = None
        try:
            atualizar_progresso(barra, status, 10, "Salvando arquivos temporários")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f1:
                f1.write(comb.read()); p1 = f1.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f2:
                f2.write(maxtrack.read()); p2 = f2.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f3:
                f3.write(ativo.read()); p3 = f3.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f4:
                f4.write(producao.read()); p4 = f4.name

            saida = os.path.join(tempfile.gettempdir(), f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            script = encontrar_arquivo(["odometro_v12_com_percentual.py"])
            if script is None:
                st.error("Arquivo odometro_v12_com_percentual.py não encontrado.")
                page_shell_close(); return

            atualizar_progresso(barra, status, 25, "Executando script do odômetro")
            proc = subprocess.Popen(
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
                linha = proc.stdout.readline() if proc.stdout else ""
                if linha:
                    logs.append(linha.strip())
                    log_box.code("\n".join(logs[-20:]))
                if proc.poll() is not None:
                    break
                if progresso < 90:
                    progresso += 1
                    barra.progress(progresso)
                    status.info(f"{progresso}% - Processando odômetro V12...")
                time.sleep(0.25)

            if proc.returncode != 0:
                st.error("O processamento retornou erro.")
                page_shell_close(); return
            if not os.path.exists(saida):
                st.error("Arquivo final não foi gerado.")
                page_shell_close(); return

            atualizar_progresso(barra, status, 100, "Finalizado")
            duracao = round(time.time() - inicio, 2)
            registrar_historico("Odômetro V12", ", ".join([comb.name, maxtrack.name, ativo.name, producao.name]), "Sucesso", duracao, "4 bases processadas")
            st.success(f"Odômetro finalizado em {duracao} segundos")
            with open(saida, "rb") as f:
                dados_saida = f.read()
            st.download_button(
                "Baixar Excel Odômetro V12",
                dados_saida,
                file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            logger.exception("Erro ao processar odometro")
            registrar_historico("Odômetro V12", ", ".join([comb.name, maxtrack.name, ativo.name, producao.name]), "Erro", time.time() - inicio, str(e))
            st.error(f"Erro ao processar odômetro: {e}")
        finally:
            for tmp_path in [p1, p2, p3, p4, saida]:
                limpar_arquivo_temporario(tmp_path)
    page_shell_close()


# =========================================================
# EDITOR DE MODULOS
# =========================================================
def pagina_editor() -> None:
    page_shell_open("Editor de Módulos", "Configure arquivos .py para aparecerem no portal sem alterar a lógica original do script.", "⚙")
    arquivos = listar_arquivos_configuraveis()
    registro = carregar_registro()

    if not arquivos:
        message("warning", "Nenhum arquivo Python configurável encontrado na pasta do app.")
        page_shell_close()
        return

    nomes = [p.name for p in arquivos]
    escolhido = st.selectbox("Arquivo Python", nomes)
    caminho = BASE_DIR / escolhido
    info = inspecionar_script(caminho, caminho.stat().st_mtime)
    cfg = registro.get(escolhido, {}) if isinstance(registro.get(escolhido, {}), dict) else {}

    if info["erro"]:
        message("error", f"Este arquivo possui problema e não será exibido no menu: {html.escape(info['erro'])}")
    else:
        message("info", f"Arquivo válido. main_streamlit: <b>{info['main_streamlit']}</b> | main: <b>{info['main']}</b>")

    c1, c2 = st.columns(2)
    with c1:
        nome = st.text_input("Nome no menu", value=str(cfg.get("nome") or nome_amigavel_script(escolhido)))
        icone = st.text_input("Ícone", value=str(cfg.get("icone") or "🧩"))
        ordem = st.number_input("Ordem", min_value=1, value=int(cfg.get("ordem", 100)), step=1)
    with c2:
        descricao = st.text_area("Descrição", value=str(cfg.get("descricao") or "Módulo operacional adicionado ao portal."), height=106)
        modo = st.selectbox("Modo de execução", ["auto", "main_streamlit", "main", "script"], index=["auto", "main_streamlit", "main", "script"].index(str(cfg.get("modo", "auto"))))
        ativo = st.checkbox("Ativo no menu", value=bool(cfg.get("ativo", False)))

    if st.button("Salvar configuração", use_container_width=True, type="primary"):
        registro[escolhido] = {
            "nome": nome.strip() or nome_amigavel_script(escolhido),
            "icone": icone.strip() or "🧩",
            "descricao": descricao.strip(),
            "ordem": int(ordem),
            "modo": modo,
            "ativo": bool(ativo),
        }
        salvar_registro(registro)
        st.success("Configuração salva com sucesso.")
        time.sleep(0.4)
        st.rerun()

    st.markdown("### Módulos adicionais ativos")
    ativos = modulos_ativos()
    if ativos:
        st.dataframe([{"Arquivo": m["arquivo"], "Nome": m["nome"], "Modo": m["modo"], "Ordem": m["ordem"]} for m in ativos], use_container_width=True, hide_index=True)
    else:
        st.caption("Nenhum módulo adicional ativo no momento.")
    page_shell_close()


# =========================================================
# PAGINA MODULO DINAMICO
# =========================================================
def pagina_modulo_dinamico(modulo_cfg: dict[str, Any]) -> None:
    page_shell_open(modulo_cfg["nome"], modulo_cfg["descricao"], modulo_cfg.get("icone", "🧩"))
    try:
        executar_modulo_dinamico(modulo_cfg["caminho"], modulo_cfg.get("modo", "auto"))
    except Exception as e:
        message("error", f"Erro ao executar módulo {html.escape(modulo_cfg['arquivo'])}: {html.escape(str(e))}")
    page_shell_close()


# =========================================================
# HISTORICO E ROADMAP TECNICO
# =========================================================
def pagina_historico() -> None:
    page_shell_open("Histórico de Processamentos", "Registro local dos processamentos executados no portal.", "▤")
    historico = carregar_historico()
    if historico:
        st.dataframe(historico, use_container_width=True, hide_index=True)
    else:
        message("info", "Nenhum processamento registrado ainda.")
    st.caption(f"Log técnico: {LOG_FILE}")
    page_shell_close()



def pagina_editor_layout() -> None:
    page_shell_open("Editor de Layout", "Ajuste o visual do portal sem alterar o código Python.", "🎨")
    cfg = layout_atual()
    st.markdown(
        "<div class='glass-card'><div class='card-title'>Configuração visual</div>"
        "<div class='card-desc'>Altere cores, largura, arredondamento e visual dos cards. Ao salvar, o arquivo layout_config.json será atualizado na pasta do app.</div></div>",
        unsafe_allow_html=True,
    )

    preset = st.selectbox(
        "Preset rápido",
        ["Personalizado", "Dark Glass Azul", "Dark Executivo", "Clean Claro", "Alto Contraste"],
        index=0,
    )
    if preset != "Personalizado":
        presets = {
            "Dark Glass Azul": {
                "bg_0": "#020814", "bg_1": "#061426", "blue": "#2f80ed", "blue_2": "#60a5fa", "cyan": "#22d3ee", "text": "#f8fbff", "muted": "#a9bad3", "card_opacity": 72, "radius": 18, "show_grid": True, "compact_home": True,
            },
            "Dark Executivo": {
                "bg_0": "#050816", "bg_1": "#0f172a", "blue": "#2563eb", "blue_2": "#93c5fd", "cyan": "#38bdf8", "text": "#ffffff", "muted": "#cbd5e1", "card_opacity": 82, "radius": 16, "show_grid": True, "compact_home": True,
            },
            "Clean Claro": {
                "bg_0": "#eaf1fb", "bg_1": "#f8fafc", "blue": "#1d4ed8", "blue_2": "#2563eb", "cyan": "#0891b2", "text": "#ffffff", "muted": "#dbeafe", "card_opacity": 86, "radius": 18, "show_grid": False, "compact_home": True,
            },
            "Alto Contraste": {
                "bg_0": "#000000", "bg_1": "#050505", "blue": "#0ea5e9", "blue_2": "#7dd3fc", "cyan": "#22d3ee", "text": "#ffffff", "muted": "#e5e7eb", "card_opacity": 94, "radius": 12, "show_grid": False, "compact_home": True,
            },
        }
        cfg.update(presets[preset])

    col1, col2, col3 = st.columns(3)
    with col1:
        cfg["bg_0"] = st.color_picker("Fundo principal", cfg.get("bg_0", "#020814"))
        cfg["bg_1"] = st.color_picker("Fundo secundário", cfg.get("bg_1", "#061426"))
        cfg["blue"] = st.color_picker("Azul principal", cfg.get("blue", "#2f80ed"))
    with col2:
        cfg["blue_2"] = st.color_picker("Azul claro", cfg.get("blue_2", "#60a5fa"))
        cfg["cyan"] = st.color_picker("Ciano / destaque", cfg.get("cyan", "#22d3ee"))
        cfg["success"] = st.color_picker("Cor de sucesso", cfg.get("success", "#22c55e"))
    with col3:
        cfg["text"] = st.color_picker("Texto principal", cfg.get("text", "#f8fbff"))
        cfg["muted"] = st.color_picker("Texto secundário", cfg.get("muted", "#a9bad3"))
        cfg["warning"] = st.color_picker("Cor de alerta", cfg.get("warning", "#f59e0b"))

    col4, col5, col6, col7 = st.columns(4)
    with col4:
        cfg["page_width"] = st.slider("Largura máxima", 980, 1680, int(cfg.get("page_width", 1320)), 20)
    with col5:
        cfg["radius"] = st.slider("Arredondamento", 8, 34, int(cfg.get("radius", 18)), 1)
    with col6:
        cfg["card_opacity"] = st.slider("Opacidade dos cards", 35, 100, int(cfg.get("card_opacity", 72)), 1)
    with col7:
        cfg["show_grid"] = st.checkbox("Grade no fundo", bool(cfg.get("show_grid", True)))
        cfg["compact_home"] = st.checkbox("Home compacta", bool(cfg.get("compact_home", True)))

    st.markdown(
        f"""
        <div class="visual-card" style="margin-top:10px;">
            <div class="visual-icon">🎨</div>
            <div class="visual-title">Prévia do Card</div>
            <div class="visual-desc">Fundo: {html.escape(str(cfg.get('bg_0', '')))} • Destaque: {html.escape(str(cfg.get('blue', '')))} • Raio: {int(cfg.get('radius', 18))}px</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    b1, b2, b3 = st.columns([1, 1, 1])
    with b1:
        if st.button("Salvar layout", type="primary", use_container_width=True):
            st.session_state["layout_config"] = cfg
            salvar_layout_config(cfg)
            st.success("Layout salvo. O visual será aplicado no próximo carregamento da página.")
            st.rerun()
    with b2:
        if st.button("Restaurar padrão", use_container_width=True):
            st.session_state["layout_config"] = DEFAULT_LAYOUT_CONFIG.copy()
            salvar_layout_config(DEFAULT_LAYOUT_CONFIG.copy())
            st.rerun()
    with b3:
        st.download_button(
            "Baixar layout_config.json",
            json.dumps(cfg, ensure_ascii=False, indent=2),
            file_name="layout_config.json",
            mime="application/json",
            use_container_width=True,
        )

    with st.expander("CSS adicional avançado"):
        st.caption("Use apenas se precisar de um ajuste fino. Cole CSS puro sem a tag <style>.")
        cfg["css_extra"] = st.text_area("CSS adicional", value=str(cfg.get("css_extra", "")), height=160)
        if st.button("Salvar CSS adicional", use_container_width=True):
            st.session_state["layout_config"] = cfg
            salvar_layout_config(cfg)
            st.rerun()

    css_extra = str(cfg.get("css_extra", "")).strip()
    if css_extra:
        st.markdown(f"<style>{css_extra}</style>", unsafe_allow_html=True)
    page_shell_close()

def pagina_roadmap() -> None:
    page_shell_open("Roadmap Técnico", "Plano organizado de evolução do portal em fases.", "✓")
    fases = [
        ("Fase 1 — Segurança e correções críticas", [
            "Limpar arquivos temporários após uso",
            "Adicionar logging estruturado",
            "Corrigir MIME type da logo",
        ], "Aplicado"),
        ("Fase 2 — Performance e responsividade", [
            "Cache de inspeção de módulos",
            "Cache da logo em base64",
            "Uso de session_state para navegação",
        ], "Aplicado"),
        ("Fase 3 — Refatoração da estrutura", [
            "Separar CSS, componentes e páginas",
            "Criar config.py com constantes",
            "Extrair assets/style.css",
        ], "Próxima etapa"),
        ("Fase 4 — Funcionalidades novas", [
            "Histórico de processamentos",
            "Validação do Excel antes de processar",
            "Autenticação opcional por senha",
        ], "Parcialmente aplicado"),
        ("Fase 5 — Testes e qualidade", [
            "Testes unitários do core",
            "requirements.txt e .env.example",
            "README profissional",
        ], "Planejado"),
    ]
    for titulo, itens, status in fases:
        st.markdown(f'<div class="glass-card"><div class="card-title">{html.escape(titulo)} <span class="pill pill-ok">{html.escape(status)}</span></div>', unsafe_allow_html=True)
        for i, item in enumerate(itens, 1):
            st.markdown(f"<div class='card-desc'><b>{i}.</b> {html.escape(item)}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    page_shell_close()


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    pagina_atual_default()
    modulos = modulos_ativos()
    render_sidebar(modulos)
    render_topbar()
    nav_buttons()

    pagina = st.session_state.get("pagina_atual", "inicio")
    mapa = {f"mod_{m['slug']}": m for m in modulos}

    if pagina == "inicio":
        pagina_inicio(modulos)
    elif pagina == "permanencia":
        pagina_permanencia()
    elif pagina == "odometro":
        pagina_odometro()
    elif pagina == "editor":
        pagina_editor()
    elif pagina == "historico":
        pagina_historico()
    elif pagina == "roadmap":
        pagina_roadmap()
    elif pagina == "layout":
        pagina_editor_layout()
    elif pagina in mapa:
        pagina_modulo_dinamico(mapa[pagina])
    else:
        st.session_state["pagina_atual"] = "inicio"
        pagina_inicio(modulos)


if __name__ == "__main__":
    main()
