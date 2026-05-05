import base64
import html
import importlib.util
import inspect
import json
import logging
import mimetypes
import os
import re
import runpy
import shutil
import subprocess
import sys
import tempfile
import time
from contextlib import suppress
from datetime import datetime
from pathlib import Path

import streamlit as st

# =========================================================
# CONFIGURACAO DA APLICACAO
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_FILE = Path(__file__).name
CONFIG_FILE = BASE_DIR / "modulos_config.json"
HISTORY_FILE = BASE_DIR / "historico_processamentos.json"
LOG_FILE = BASE_DIR / "portal_operacional.log"

ADMIN_USER = "pedro admin"
ADMIN_PASSWORD = "admin pedro"

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Central Operacional de Analises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

# =========================================================
# CORES E IDENTIDADE
# =========================================================
COLORS = {
    "navy": "#061642",
    "navy_2": "#0A1F55",
    "blue": "#0B4EDB",
    "blue_2": "#2F80ED",
    "sky": "#D9E8FF",
    "ice": "#F6FAFF",
    "line": "#DDE7F7",
    "text": "#081B49",
    "muted": "#5D6B8A",
    "success": "#16A34A",
    "warning": "#F59E0B",
    "danger": "#DC2626",
}

# =========================================================
# UTILITARIOS DE ARQUIVO
# =========================================================
def normalizar_nome_arquivo(nome: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", nome.lower())


def encontrar_arquivo(candidatos: list[str]) -> Path | None:
    for nome in candidatos:
        caminho = BASE_DIR / nome
        if caminho.exists():
            return caminho
    try:
        mapa = {normalizar_nome_arquivo(p.name): p for p in BASE_DIR.iterdir() if p.is_file()}
    except Exception:
        return None
    for nome in candidatos:
        chave = normalizar_nome_arquivo(nome)
        if chave in mapa:
            return mapa[chave]
    return None


def detectar_mime(caminho: Path) -> str:
    mime, _ = mimetypes.guess_type(str(caminho))
    return mime or "image/png"


def imagem_base64(caminho: Path | None) -> tuple[str, str]:
    if caminho is None or not caminho.exists():
        return "", "image/png"
    try:
        return base64.b64encode(caminho.read_bytes()).decode("utf-8"), detectar_mime(caminho)
    except Exception as exc:
        logging.exception("Falha ao converter imagem para base64: %s", exc)
        return "", "image/png"


@st.cache_data(show_spinner=False)
def logo_cache(path_str: str | None) -> tuple[str, str]:
    if not path_str:
        return "", "image/png"
    return imagem_base64(Path(path_str))


def encontrar_logo() -> Path | None:
    return encontrar_arquivo([
        "logo_nepomuceno.png.jpeg",
        "logo_nepomuceno.jpeg",
        "logo_nepomuceno.jpg",
        "logo_nepomuceno.png",
        "Expresso Nepomuceno.png",
        "Expresso Nepomuceno.jpeg",
        "Expresso Nepomuceno.jpg",
        "Logo Nepomuceno.png",
        "Logo Nepomuceno.jpeg",
        "Logo Nepomuceno.jpg",
    ])


LOGO_PATH = encontrar_logo()
LOGO_B64, LOGO_MIME = logo_cache(str(LOGO_PATH) if LOGO_PATH else None)


def logo_html(classe: str = "brand-logo") -> str:
    if LOGO_B64:
        return f'<img class="{classe}" src="data:{LOGO_MIME};base64,{LOGO_B64}" alt="Expresso Nepomuceno">'
    return f'<div class="{classe} brand-logo-fallback">E.N</div>'


# =========================================================
# JSON / HISTORICO / LOG
# =========================================================
def ler_json(caminho: Path, padrao):
    if not caminho.exists():
        return padrao
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except Exception as exc:
        logging.exception("Erro lendo JSON %s: %s", caminho, exc)
        return padrao


def salvar_json(caminho: Path, dados) -> None:
    caminho.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def registrar_historico(tipo: str, status: str, detalhes: str, arquivo_saida: str | None = None, duracao: float | None = None) -> None:
    historico = ler_json(HISTORY_FILE, [])
    historico.insert(0, {
        "data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "tipo": tipo,
        "status": status,
        "detalhes": detalhes,
        "arquivo_saida": arquivo_saida or "",
        "duracao_segundos": duracao,
    })
    salvar_json(HISTORY_FILE, historico[:500])
    logging.info("Historico | %s | %s | %s", tipo, status, detalhes)


def limpar_temporarios(*paths: str) -> None:
    for path in paths:
        with suppress(Exception):
            if path and os.path.exists(path):
                os.remove(path)


# =========================================================
# CONFIGURACAO DE MODULOS
# =========================================================
def config_padrao_modulo(nome_arquivo: str) -> dict:
    return {
        "nome": nome_amigavel_script(nome_arquivo),
        "icone": "🧩",
        "descricao": "Modulo operacional adicionado ao portal.",
        "categoria": "Modulo",
        "ordem": 100,
        "ativo": False,
        "modo": "auto",
        "ocultar": False,
        "observacoes": "",
        "criado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }


def carregar_config_modulos() -> dict:
    dados = ler_json(CONFIG_FILE, {})
    if not isinstance(dados, dict):
        return {}
    return dados


def salvar_config_modulos(config: dict) -> None:
    salvar_json(CONFIG_FILE, config)


# =========================================================
# CSS - MODELO LIMPO EXPRESSO NEPOMUCENO
# =========================================================
def aplicar_css() -> None:
    st.markdown(f"""
<style>
:root {{
  --navy: {COLORS['navy']};
  --navy-2: {COLORS['navy_2']};
  --blue: {COLORS['blue']};
  --blue-2: {COLORS['blue_2']};
  --sky: {COLORS['sky']};
  --ice: {COLORS['ice']};
  --line: {COLORS['line']};
  --text: {COLORS['text']};
  --muted: {COLORS['muted']};
}}
html, body, [class*="css"] {{ font-family: Inter, "Segoe UI", Roboto, Arial, sans-serif; }}
.stApp {{ background: linear-gradient(180deg, #FFFFFF 0%, #F2F7FF 100%); color: var(--text); }}
.block-container {{ padding-top: 0.8rem; padding-bottom: 3rem; max-width: 1420px; }}
#MainMenu {{ visibility: hidden; }}
footer {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ visibility: visible !important; background: transparent !important; height: 2.25rem !important; }}
header[data-testid="stHeader"] * {{ color: var(--navy) !important; }}

section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, #07173F 0%, #0B255A 100%); border-right: 1px solid rgba(255,255,255,.10); }}
section[data-testid="stSidebar"] > div {{ padding-top: 1rem; }}
section[data-testid="stSidebar"] * {{ color: #EAF2FF !important; }}
.side-brand {{ display:flex; gap:12px; align-items:center; padding: 10px 6px 16px 6px; }}
.side-mark {{ width:48px; height:48px; border-radius:50%; border:2px solid #5AA2FF; display:flex; align-items:center; justify-content:center; font-weight:900; color:white; box-shadow: 0 0 0 6px rgba(47,128,237,.16); }}
.side-title {{ font-weight:900; font-size:15px; line-height:1.15; }}
.side-subtitle {{ color:#BCD1F2 !important; font-size:12px; margin-top:4px; }}
.side-section {{ margin: 18px 0 8px; font-size:10px; letter-spacing:.16em; color:#9DBDF3 !important; text-transform:uppercase; font-weight:900; }}
.side-bottom-line {{ height:1px; background: rgba(255,255,255,.10); margin: 18px 0; }}
section[data-testid="stSidebar"] .stButton > button {{ width:100%; justify-content:flex-start; background: rgba(255,255,255,.055) !important; border: 1px solid rgba(255,255,255,.10) !important; color:#EAF2FF !important; min-height:46px; border-radius:12px !important; font-weight:800 !important; }}
section[data-testid="stSidebar"] .stButton > button:hover {{ background: rgba(47,128,237,.28) !important; border-color: rgba(154,196,255,.40) !important; transform: translateX(2px); }}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{ background: linear-gradient(135deg, #1D4ED8 0%, #2F80ED 100%) !important; border:0 !important; color:white !important; }}

.topbar {{ background: linear-gradient(135deg, #071A4A 0%, #0A3A8C 100%); border-radius:0 0 0 0; min-height:74px; display:flex; align-items:center; justify-content:space-between; padding: 18px 28px; margin: -0.8rem -1rem 2rem -1rem; box-shadow:0 14px 40px rgba(6,22,66,.20); }}
.topbar-left {{ display:flex; align-items:center; gap:22px; }}
.topbar-logo {{ height:42px; max-width:190px; object-fit:contain; }}
.topbar-divider {{ width:1px; height:42px; background:rgba(255,255,255,.35); }}
.topbar-title {{ color:white; font-size:15px; text-transform:uppercase; letter-spacing:.04em; line-height:1.25; }}
.topbar-actions {{ display:flex; align-items:center; gap:16px; color:white; font-size:13px; }}
.avatar {{ width:34px; height:34px; border-radius:50%; background:#D9E8FF; color:#0A3A8C; display:flex; align-items:center; justify-content:center; font-weight:900; }}

.hero {{ position:relative; background: radial-gradient(circle at 78% 25%, rgba(47,128,237,.22), transparent 28%), linear-gradient(135deg, #071A4A 0%, #0B2B6F 53%, #0A55B8 100%); border-radius:24px; padding:42px 44px 110px; overflow:hidden; color:white; min-height:320px; box-shadow:0 24px 65px rgba(9,31,86,.24); }}
.hero:after {{ content:"EN"; position:absolute; right:52px; top:42px; color:rgba(255,255,255,.10); font-size:112px; font-weight:900; letter-spacing:-8px; }}
.hero:before {{ content:""; position:absolute; right:-150px; bottom:-120px; width:650px; height:300px; background: radial-gradient(ellipse, rgba(151,196,255,.28), transparent 65%); transform: rotate(-10deg); }}
.hero-title {{ position:relative; z-index:1; font-size:38px; line-height:1.08; font-weight:900; max-width:680px; margin:0 0 14px; }}
.hero-text {{ position:relative; z-index:1; color:#D9E8FF; font-size:16px; max-width:640px; margin-bottom:24px; }}
.hero-kpis {{ position:relative; z-index:1; display:flex; gap:12px; flex-wrap:wrap; }}
.hero-kpi {{ min-width:178px; border:1px solid rgba(255,255,255,.18); background:rgba(255,255,255,.08); border-radius:14px; padding:14px 16px; backdrop-filter: blur(8px); }}
.hero-kpi span {{ display:block; color:#BFD7FF; font-size:12px; }}
.hero-kpi strong {{ display:block; color:white; font-size:24px; margin-top:3px; }}

.home-overlap {{ margin-top:-78px; position:relative; z-index:3; }}
.card-grid {{ display:grid; grid-template-columns: repeat(3, minmax(0,1fr)); gap:20px; }}
.module-card {{ background:white; border:1px solid var(--line); border-radius:16px; padding:24px; min-height:210px; box-shadow:0 18px 46px rgba(15,42,92,.10); transition:.16s ease; }}
.module-card:hover {{ transform:translateY(-3px); box-shadow:0 24px 58px rgba(15,42,92,.16); border-color:#B9D4FF; }}
.module-icon {{ width:54px; height:54px; border-radius:16px; display:flex; align-items:center; justify-content:center; background:#EAF3FF; color:var(--blue); font-size:28px; margin-bottom:18px; }}
.module-title {{ color:var(--text); font-size:18px; font-weight:900; margin-bottom:9px; }}
.module-desc {{ color:var(--muted); font-size:13px; line-height:1.55; min-height:58px; }}
.card-stats {{ display:flex; gap:12px; margin-top:16px; }}
.stat-mini {{ flex:1; border:1px solid var(--line); border-radius:10px; padding:8px 10px; }}
.stat-mini strong {{ color:var(--text); font-size:18px; display:block; }}
.stat-mini span {{ color:var(--muted); font-size:11px; }}

.panel {{ background:white; border:1px solid var(--line); border-radius:18px; padding:24px; box-shadow:0 18px 46px rgba(15,42,92,.08); margin-bottom:20px; }}
.panel-title {{ display:flex; align-items:center; gap:12px; color:var(--text); font-weight:900; font-size:22px; margin-bottom:6px; }}
.panel-text {{ color:var(--muted); font-size:14px; margin-bottom:18px; }}
.workflow-title {{ font-size:32px; color:var(--text); font-weight:900; margin:0; }}
.workflow-subtitle {{ color:var(--muted); margin:6px 0 18px; }}
.badge {{ display:inline-flex; align-items:center; gap:6px; border-radius:999px; padding:6px 10px; font-size:12px; font-weight:800; }}
.badge-blue {{ background:#EEF5FF; color:#174EA6; }}
.badge-green {{ background:#DCFCE7; color:#166534; }}
.badge-orange {{ background:#FFF3E2; color:#9A5A00; }}

.step-card {{ background:white; border:1px solid var(--line); border-radius:16px; padding:20px; min-height:265px; box-shadow:0 16px 42px rgba(15,42,92,.08); position:relative; }}
.step-num {{ position:absolute; right:18px; top:16px; width:27px; height:27px; border-radius:50%; display:flex; align-items:center; justify-content:center; background:var(--blue); color:white; font-weight:900; }}
.step-icon {{ width:58px; height:58px; border-radius:50%; background:#EAF3FF; color:var(--blue); display:flex; align-items:center; justify-content:center; font-size:27px; margin: 8px auto 18px; }}
.step-title {{ text-align:center; color:var(--text); font-weight:900; font-size:16px; }}
.step-subtitle {{ text-align:center; color:var(--muted); font-size:12px; margin-bottom:14px; }}
.upload-box {{ border:1.5px dashed #8DB7FF; border-radius:12px; min-height:96px; display:flex; align-items:center; justify-content:center; text-align:center; color:var(--text); background:#FAFCFF; }}

.info-row {{ display:grid; grid-template-columns: 1.4fr 1fr 1fr 1fr 1fr; gap:18px; align-items:center; }}
.info-item strong {{ display:block; color:var(--text); font-size:13px; }}
.info-item span {{ display:block; color:var(--muted); font-size:12px; margin-top:4px; }}

.stButton > button {{ border-radius:11px !important; min-height:40px; border:1px solid #C9DAF5 !important; color:#0647B7 !important; background:white !important; font-weight:800 !important; }}
.stButton > button:hover {{ border-color:#0B4EDB !important; color:#0B4EDB !important; box-shadow:0 6px 18px rgba(11,78,219,.12); }}
.stButton > button[kind="primary"] {{ background:linear-gradient(135deg, #0B4EDB 0%, #0A61F7 100%) !important; color:white !important; border:0 !important; }}
.stDownloadButton > button {{ background:linear-gradient(135deg, #0B4EDB 0%, #0A61F7 100%) !important; color:white !important; border:0 !important; border-radius:11px !important; min-height:42px; font-weight:900 !important; }}
.stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-baseweb="select"] > div {{ color:#081B49 !important; background:#FFFFFF !important; border-color:#C9DAF5 !important; }}
.stCheckbox label span, .stTextInput label, .stTextArea label, .stNumberInput label, .stSelectbox label, .stFileUploader label {{ color:#152A5C !important; font-weight:800 !important; }}
[data-testid="stFileUploader"] section {{ background:#FAFCFF !important; border:1.5px dashed #8DB7FF !important; border-radius:12px !important; }}
[data-testid="stFileUploader"] section * {{ color:#081B49 !important; }}
[data-testid="stFileUploader"] button {{ background:#FFFFFF !important; color:#0B4EDB !important; border:1px solid #C9DAF5 !important; }}
[data-testid="stDataFrame"] {{ border-radius:14px; overflow:hidden; }}
.stProgress > div > div > div > div {{ background:linear-gradient(90deg, #0B4EDB, #2F80ED) !important; }}

@media(max-width: 1050px) {{
  .card-grid {{ grid-template-columns: 1fr; }}
  .hero {{ padding:30px 24px 92px; }}
  .hero-title {{ font-size:30px; }}
  .info-row {{ grid-template-columns:1fr; }}
  .topbar {{ margin-left:-.5rem; margin-right:-.5rem; }}
}}
</style>
""", unsafe_allow_html=True)


# =========================================================
# NAVEGACAO
# =========================================================
PAGES = {
    "inicio": "Início",
    "permanencia": "Análise de Permanência",
    "odometro": "Odômetro V12",
    "editor": "Editor de Módulos",
    "historico": "Histórico",
    "relatorios": "Relatórios",
    "config": "Configurações / Layout",
}

PAGE_ICONS = {
    "inicio": "⌂",
    "permanencia": "◷",
    "odometro": "◴",
    "editor": "⚙",
    "historico": "▤",
    "relatorios": "▥",
    "config": "⚙",
}


def go(page: str) -> None:
    st.session_state.page = page
    st.rerun()


def current_page() -> str:
    if "page" not in st.session_state:
        st.session_state.page = "inicio"
    return st.session_state.page


def sidebar_button(page: str, label: str, icon: str) -> None:
    active = current_page() == page
    if st.sidebar.button(f"{icon}  {label}", key=f"nav_{page}", type="primary" if active else "secondary"):
        go(page)


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown(f"""
        <div class="side-brand">
          <div class="side-mark">E.N</div>
          <div>
            <div class="side-title">Central Operacional</div>
            <div class="side-subtitle">Expresso Nepomuceno</div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        sidebar_button("inicio", "Início", "⌂")
        sidebar_button("permanencia", "Análise de Permanência", "◷")
        sidebar_button("odometro", "Odômetro V12", "◴")
        sidebar_button("editor", "Editor de Módulos", "⚙")
        sidebar_button("historico", "Histórico", "▤")
        sidebar_button("relatorios", "Relatórios", "▥")
        sidebar_button("config", "Configurações / Layout", "⚙")
        st.markdown('<div class="side-bottom-line"></div>', unsafe_allow_html=True)
        if st.button("↪ Sair", key="logout_sidebar"):
            st.session_state.editor_auth = False
            st.toast("Sessão administrativa encerrada.")
        st.caption("Versão 2.5.4")


def render_topbar() -> None:
    st.markdown(f"""
    <div class="topbar">
      <div class="topbar-left">
        {logo_html('topbar-logo')}
        <div class="topbar-divider"></div>
        <div class="topbar-title">Central Operacional<br>de Análises</div>
      </div>
      <div class="topbar-actions">
        <span>🔔</span><span>?</span><div class="avatar">AD</div>
        <div><b>Administrador</b><br><span style="color:#C9D9F5;font-size:12px;">admin@expnepo.com.br</span></div>
        <span>⌄</span>
      </div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# MODULOS E EXECUCAO
# =========================================================
def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".py", "")
    nome = nome.replace("_", " ").replace("-", " ")
    return " ".join(p.capitalize() for p in nome.split()) or nome_arquivo


def script_deve_ser_ignorado(caminho: Path) -> bool:
    nome = caminho.name.lower()
    ignorar = {
        APP_FILE.lower(), "app.py", "app_home_profissional.py", "app_layout_editor.py",
        "app_roadmap_aplicado.py", "app_dark_glass_corrigido.py", "__init__.py",
    }
    if nome in ignorar:
        return True
    return nome.startswith(("streamlit ", "pip ", "python ", "py "))


@st.cache_data(show_spinner=False, ttl=20)
def listar_scripts_python_cached() -> list[str]:
    arquivos = []
    for arq in BASE_DIR.iterdir():
        if arq.is_file() and arq.suffix.lower() == ".py" and not script_deve_ser_ignorado(arq):
            arquivos.append(arq.name)
    return sorted(arquivos, key=str.lower)


def listar_scripts_python() -> list[str]:
    return listar_scripts_python_cached()


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


def inspecionar_script(nome_arquivo: str) -> dict:
    caminho = BASE_DIR / nome_arquivo
    resultado = {"arquivo": nome_arquivo, "existe": caminho.exists(), "main_streamlit": False, "main": False, "erro": "", "funcoes": []}
    if not caminho.exists():
        resultado["erro"] = "Arquivo não encontrado."
        return resultado
    try:
        modulo = carregar_modulo_por_arquivo(caminho)
        resultado["main_streamlit"] = hasattr(modulo, "main_streamlit")
        resultado["main"] = hasattr(modulo, "main")
        resultado["funcoes"] = [n for n, v in inspect.getmembers(modulo, inspect.isfunction) if not n.startswith("_")][:40]
    except Exception as exc:
        resultado["erro"] = str(exc)
    return resultado


def executar_script_configurado(nome_arquivo: str, modo: str = "auto") -> None:
    caminho = BASE_DIR / nome_arquivo
    if not caminho.exists():
        st.error(f"Arquivo não encontrado: {nome_arquivo}")
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
                raise AttributeError("O módulo não possui main_streamlit().")
            modulo.main_streamlit()
        elif modo == "main":
            if not hasattr(modulo, "main"):
                raise AttributeError("O módulo não possui main().")
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


def criar_template_modulo(nome_arquivo: str, titulo: str) -> Path:
    safe = re.sub(r"[^a-zA-Z0-9_]+", "_", Path(nome_arquivo).stem).strip("_").lower() or "modulo"
    path = BASE_DIR / f"{safe}_portal.py"
    if path.exists():
        return path
    conteudo = f'''"""Modulo gerado pelo Editor de Modulos da Central Operacional."""
import streamlit as st

MODULO_CONFIG = {{
    "nome": {titulo!r},
    "icone": "🧩",
    "descricao": "Modulo criado pelo editor integrado.",
    "ativo": True,
}}


def main_streamlit():
    st.title({titulo!r})
    st.info("Substitua este conteudo pela interface do modulo.")
    arquivo = st.file_uploader("Enviar arquivo", type=["xlsx", "xls", "csv"])
    if arquivo:
        st.success(f"Arquivo carregado: {{arquivo.name}}")
'''
    path.write_text(conteudo, encoding="utf-8")
    return path


# =========================================================
# COMPONENTES VISUAIS
# =========================================================
def module_card_html(icon: str, title: str, desc: str, stats: list[tuple[str, str]] | None = None) -> None:
    stats_html = ""
    if stats:
        stats_html = '<div class="card-stats">' + ''.join(
            f'<div class="stat-mini"><strong>{html.escape(v)}</strong><span>{html.escape(k)}</span></div>' for k, v in stats
        ) + '</div>'
    st.markdown(f"""
    <div class="module-card">
      <div class="module-icon">{html.escape(icon)}</div>
      <div class="module-title">{html.escape(title)}</div>
      <div class="module-desc">{html.escape(desc)}</div>
      {stats_html}
    </div>
    """, unsafe_allow_html=True)


def page_header(title: str, subtitle: str, badge: str | None = None) -> None:
    badge_html = f'<span class="badge badge-blue">● {html.escape(badge)}</span>' if badge else ""
    st.markdown(f"""
    <div style="margin:18px 0 22px;">
      <h1 class="workflow-title">{html.escape(title)} {badge_html}</h1>
      <div class="workflow-subtitle">{html.escape(subtitle)}</div>
    </div>
    """, unsafe_allow_html=True)


def panel(title: str, text: str = "") -> None:
    st.markdown(f"""
    <div class="panel-title">{html.escape(title)}</div>
    <div class="panel-text">{html.escape(text)}</div>
    """, unsafe_allow_html=True)


def atualizar_progresso(barra, status, pct: int, texto: str) -> None:
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.12)


def validar_excel_upload(uploaded, nome: str) -> bool:
    if uploaded is None:
        return False
    if not uploaded.name.lower().endswith((".xlsx", ".xls")):
        st.error(f"{nome}: envie um arquivo Excel .xlsx ou .xls")
        return False
    if uploaded.size == 0:
        st.error(f"{nome}: arquivo vazio.")
        return False
    return True


# =========================================================
# PAGINAS
# =========================================================
def pagina_inicio() -> None:
    scripts = listar_scripts_python()
    ativos = [v for v in carregar_config_modulos().values() if isinstance(v, dict) and v.get("ativo") and not v.get("ocultar")]
    historico = ler_json(HISTORY_FILE, [])
    hoje = datetime.now().strftime("%d/%m/%Y")
    concluidos_hoje = sum(1 for h in historico if str(h.get("data", "")).startswith(hoje) and h.get("status") == "sucesso")

    st.markdown(f"""
    <div class="hero">
      <h1 class="hero-title">Bem-vindo à Central Operacional de Análises.</h1>
      <div class="hero-text">Gerencie e acompanhe os processos de análise de forma centralizada, segura e eficiente.</div>
      <div class="hero-kpis">
        <div class="hero-kpi"><span>Processamentos hoje</span><strong>{concluidos_hoje}</strong></div>
        <div class="hero-kpi"><span>Módulos ativos</span><strong>{len(ativos) + 2}</strong></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="home-overlap">', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        module_card_html("◷", "Análise de Permanência", "Visualize e analise eventos de permanência dos veículos nos locais de carga e descarga.", [("Eventos hoje", "56"), ("Pendências", "8")])
        if st.button("Acessar módulo  →", key="home_perm", use_container_width=True):
            go("permanencia")
    with c2:
        module_card_html("◴", "Odômetro V12", "Consulte, valide e processe leituras de odômetro dos veículos da frota V12.", [("Leituras hoje", "342"), ("Pendências", "19")])
        if st.button("Acessar módulo  →", key="home_odo", use_container_width=True):
            go("odometro")
    with c3:
        module_card_html("⚙", "Editor de Módulos", "Crie, edite e gerencie módulos de regras e parâmetros para análises personalizadas.", [("Scripts", str(len(scripts))), ("Ativos", str(len(ativos)))])
        if st.button("Acessar módulo  →", key="home_editor", use_container_width=True):
            go("editor")

    c4, c5 = st.columns([1, 1])
    with c4:
        module_card_html("▤", "Histórico", "Visualize o histórico completo de análises realizadas no sistema.", [("Registros", str(len(historico))), ("Sucesso", "98%")])
        if st.button("Acessar histórico  →", key="home_hist", use_container_width=True):
            go("historico")
    with c5:
        module_card_html("▥", "Relatórios", "Gere relatórios detalhados e personalizados com base nas análises realizadas.", [("Status", "Online"), ("Ambiente", "Prod")])
        if st.button("Acessar relatórios  →", key="home_rel", use_container_width=True):
            go("relatorios")

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="panel">
      <div class="info-row">
        <div><b style="color:#0B4EDB;">ⓘ Informações do Sistema</b></div>
        <div class="info-item"><span>Versão do sistema</span><strong>2.5.4</strong></div>
        <div class="info-item"><span>Ambiente</span><strong>Produção</strong></div>
        <div class="info-item"><span>Último acesso</span><strong>{datetime.now().strftime('%d/%m/%Y %H:%M')}</strong></div>
        <div class="info-item"><span>Usuários online</span><strong>7</strong></div>
      </div>
    </div>
    """, unsafe_allow_html=True)


def carregar_modulo_permanencia():
    caminho = encontrar_arquivo(["Codigo_colado.py", "Código_colado.py", "Codigo colado.py", "Código colado.py"])
    if caminho is None:
        raise FileNotFoundError("Arquivo Codigo_colado.py não encontrado na pasta do app.")
    modulo = carregar_modulo_por_arquivo(caminho)
    faltando = [f for f in ["carregar_dados", "identificar_eventos_carregamento", "montar_ciclos_carregamento", "gerar_resumos", "salvar_saida"] if not hasattr(modulo, f)]
    if faltando:
        raise AttributeError("Funções ausentes: " + ", ".join(faltando))
    return modulo


def pagina_permanencia() -> None:
    page_header("Análise de Permanência", "Processamento da base de permanência com classificação por tempo configurável.", "Ativo")
    tabs = st.tabs(["Visão Geral", "Configurações", "Histórico de Execuções"])
    with tabs[0]:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        panel("Workflow de Processamento", "Faça upload da base de permanência e execute a análise operacional.")
        col1, col2 = st.columns(2)
        with col1:
            tempo_minimo = st.number_input("Tempo mínimo aceitável (minutos)", min_value=0, value=15, step=1)
        with col2:
            tempo_maximo = st.number_input("Tempo máximo aceitável (minutos)", min_value=1, value=55, step=1)
        arquivo = st.file_uploader("Arraste e solte a base de permanência ou clique para selecionar", type=["xlsx", "xls"], key="upload_permanencia")
        if arquivo:
            st.success(f"Arquivo carregado: {arquivo.name}")
        executar = st.button("Executar workflow", type="primary", use_container_width=True, disabled=not arquivo)
        st.markdown('</div>', unsafe_allow_html=True)
        if executar:
            if tempo_maximo <= tempo_minimo:
                st.error("O tempo máximo precisa ser maior que o mínimo.")
                return
            inicio = time.time()
            tmp_path = None
            try:
                permanencia = carregar_modulo_permanencia()
                permanencia.TEMPO_MINIMO = tempo_minimo
                permanencia.TEMPO_MAXIMO = tempo_maximo
                barra = st.progress(0)
                status = st.empty()
                atualizar_progresso(barra, status, 10, "Preparando arquivo")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(arquivo.read())
                    tmp_path = tmp.name
                atualizar_progresso(barra, status, 30, "Lendo base")
                df_base = permanencia.carregar_dados(tmp_path)
                atualizar_progresso(barra, status, 50, "Identificando eventos")
                eventos = permanencia.identificar_eventos_carregamento(df_base)
                atualizar_progresso(barra, status, 70, "Montando ciclos")
                df_resultado, df_alertas = permanencia.montar_ciclos_carregamento(eventos)
                atualizar_progresso(barra, status, 85, "Gerando resumos")
                resumo_geral, resumo_up, resumo_eq, ranking = permanencia.gerar_resumos(df_resultado)
                atualizar_progresso(barra, status, 95, "Gerando Excel final")
                saida = permanencia.salvar_saida(tmp_path, df_base, eventos, df_resultado, df_alertas, resumo_geral, resumo_up, resumo_eq, ranking)
                atualizar_progresso(barra, status, 100, "Finalizado")
                duracao = round(time.time() - inicio, 2)
                registrar_historico("Permanência", "sucesso", arquivo.name, saida, duracao)
                st.success(f"Processo finalizado em {duracao} segundos")
                if not df_resultado.empty:
                    st.dataframe(df_resultado.head(100), use_container_width=True)
                with open(saida, "rb") as f:
                    st.download_button("Baixar Excel Permanência", f, file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except Exception as exc:
                logging.exception("Erro no processamento de permanencia: %s", exc)
                registrar_historico("Permanência", "erro", str(exc), duracao=round(time.time() - inicio, 2))
                st.error(f"Erro ao processar permanência: {exc}")
            finally:
                limpar_temporarios(tmp_path)
    with tabs[1]:
        st.info("Os parâmetros podem ser ajustados antes da execução. O padrão atual é 15 a 55 minutos.")
    with tabs[2]:
        pagina_historico(filtro="Permanência", compact=True)


def pagina_odometro() -> None:
    page_header("Odômetro V12", "Validação e análise de odômetro com integração de bases e geração de relatórios.", "Ativo")
    col_a, col_b = st.columns([2.5, 1])
    with col_b:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown(f'<span class="badge badge-blue">Última execução<br>{datetime.now().strftime("%d/%m/%Y %H:%M")}</span>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with col_a:
        tabs = st.tabs(["Visão Geral", "Configurações", "Histórico de Execuções", "Documentação"])
        with tabs[0]:
            panel("Workflow de Processamento", "Faça upload das bases necessárias para executar o workflow Odômetro V12.")
            c1, c2, c3, c4 = st.columns(4)
            uploads = []
            etapas = [
                ("Base Combustível", "CSV, XLSX ou Parquet", "⛽", "comb"),
                ("Base Km Rodado Maxtrack", "CSV, XLSX ou Parquet", "◴", "maxtrack"),
                ("Base Ativo de Veículos", "CSV, XLSX ou Parquet", "🚚", "ativo"),
                ("Produção Oficial / Cliente", "CSV, XLSX ou Parquet", "👥", "producao"),
            ]
            for idx, (col, item) in enumerate(zip([c1, c2, c3, c4], etapas), start=1):
                nome, subtitulo, icone, key = item
                with col:
                    st.markdown(f"""
                    <div class="step-card">
                      <div class="step-num">{idx}</div>
                      <div class="step-icon">{html.escape(icone)}</div>
                      <div class="step-title">{html.escape(nome)}</div>
                      <div class="step-subtitle">{html.escape(subtitulo)}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    up = st.file_uploader(f"Selecionar {nome}", type=["xlsx", "xls"], key=key, label_visibility="collapsed")
                    uploads.append(up)
                    if up:
                        st.success(f"Enviado: {up.name}")
            qtd = sum(u is not None for u in uploads)
            st.progress(qtd / 4)
            st.caption(f"{qtd} de 4 arquivos carregados")
            executar = st.button("Executar workflow", type="primary", use_container_width=True, disabled=qtd < 4)
            if qtd < 4:
                st.info("Após o envio de todas as bases, clique em Executar workflow para iniciar o processamento.")
            if executar:
                processar_odometro(uploads)
        with tabs[1]:
            st.info("Configurações do Odômetro V12 serão carregadas do script odometro_v12_com_percentual.py.")
        with tabs[2]:
            pagina_historico(filtro="Odômetro", compact=True)
        with tabs[3]:
            st.markdown("O workflow exige quatro bases: Combustível, Km Maxtrack, Ativo de Veículos e Produção Oficial.")


def processar_odometro(uploads) -> None:
    inicio = time.time()
    temp_paths = []
    try:
        nomes = ["Base Combustível", "Base Km Rodado Maxtrack", "Base Ativo de Veículos", "Produção Oficial / Cliente"]
        for uploaded, nome in zip(uploads, nomes):
            if not validar_excel_upload(uploaded, nome):
                return
        barra = st.progress(0)
        status = st.empty()
        atualizar_progresso(barra, status, 10, "Salvando arquivos temporários")
        for uploaded in uploads:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.read())
                temp_paths.append(tmp.name)
        saida = os.path.join(tempfile.gettempdir(), f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        script = encontrar_arquivo(["odometro_v12_com_percentual.py"])
        if script is None:
            st.error("Arquivo odometro_v12_com_percentual.py não encontrado.")
            return
        atualizar_progresso(barra, status, 25, "Executando script do odômetro")
        processo = subprocess.Popen([sys.executable, str(script), *temp_paths, saida], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
        logs = []
        log_box = st.empty()
        progresso = 25
        while True:
            linha = processo.stdout.readline() if processo.stdout else ""
            if linha:
                logs.append(linha.strip())
                log_box.code("\n".join(logs[-15:]))
            if processo.poll() is not None:
                break
            if progresso < 90:
                progresso += 1
                barra.progress(progresso)
                status.info(f"{progresso}% - Processando odômetro V12...")
            time.sleep(0.25)
        if processo.returncode != 0:
            raise RuntimeError("O processamento retornou erro. Verifique o log exibido.")
        if not os.path.exists(saida):
            raise FileNotFoundError("Arquivo final não foi gerado.")
        atualizar_progresso(barra, status, 100, "Finalizado")
        duracao = round(time.time() - inicio, 2)
        registrar_historico("Odômetro", "sucesso", "4 bases processadas", saida, duracao)
        st.success(f"Odômetro finalizado em {duracao} segundos")
        with open(saida, "rb") as f:
            st.download_button("Baixar Excel Odômetro V12", f, file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    except Exception as exc:
        logging.exception("Erro no processamento do odometro: %s", exc)
        registrar_historico("Odômetro", "erro", str(exc), duracao=round(time.time() - inicio, 2))
        st.error(f"Erro ao processar odômetro: {exc}")
    finally:
        limpar_temporarios(*temp_paths)


def editor_login() -> bool:
    if st.session_state.get("editor_auth"):
        return True
    page_header("Área administrativa", "Informe usuário e senha para acessar o Editor de Módulos.")
    with st.form("login_editor"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)
    if entrar:
        if usuario.strip().lower() == ADMIN_USER and senha == ADMIN_PASSWORD:
            st.session_state.editor_auth = True
            st.success("Acesso liberado.")
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")
    return False


def pagina_editor_modulos() -> None:
    if not editor_login():
        return
    page_header("Editor de Módulos", "Configure qualquer arquivo .py para funcionar no portal sem alterar o código original.", "Admin")
    config = carregar_config_modulos()
    scripts = listar_scripts_python()
    if not scripts:
        st.warning("Nenhum arquivo .py adicional encontrado na pasta do app.")
    col_left, col_right = st.columns([1.1, 1])
    with col_left:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        panel("Configurar módulo existente", "Selecione um script, defina como ele aparece no menu e escolha o modo de execução.")
        selecionado = st.selectbox("Arquivo Python", scripts or [""], index=0)
        if selecionado:
            atual = {**config_padrao_modulo(selecionado), **config.get(selecionado, {})}
            info = inspecionar_script(selecionado)
            if info["erro"]:
                st.warning(f"Inspeção com alerta: {info['erro']}")
            else:
                st.success(f"Arquivo válido. main_streamlit: {info['main_streamlit']} | main: {info['main']}")
            with st.form("form_modulo"):
                nome = st.text_input("Nome no menu", value=atual.get("nome", nome_amigavel_script(selecionado)))
                icone = st.text_input("Ícone", value=atual.get("icone", "🧩"))
                descricao = st.text_area("Descrição", value=atual.get("descricao", "Modulo operacional adicionado ao portal."), height=88)
                categoria = st.text_input("Categoria", value=atual.get("categoria", "Modulo"))
                modo = st.selectbox("Modo de execução", ["auto", "main_streamlit", "main", "script"], index=["auto", "main_streamlit", "main", "script"].index(atual.get("modo", "auto")) if atual.get("modo", "auto") in ["auto", "main_streamlit", "main", "script"] else 0)
                ordem = st.number_input("Ordem", min_value=1, max_value=999, value=int(atual.get("ordem", 100)), step=1)
                ativo = st.checkbox("Ativo no menu", value=bool(atual.get("ativo", False)))
                ocultar = st.checkbox("Ocultar temporariamente", value=bool(atual.get("ocultar", False)))
                observacoes = st.text_area("Instruções/observações para este módulo", value=atual.get("observacoes", ""), height=90)
                salvar = st.form_submit_button("Salvar configuração", type="primary", use_container_width=True)
            if salvar:
                config[selecionado] = {"nome": nome, "icone": icone, "descricao": descricao, "categoria": categoria, "modo": modo, "ordem": int(ordem), "ativo": ativo, "ocultar": ocultar, "observacoes": observacoes, "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
                salvar_config_modulos(config)
                listar_scripts_python_cached.clear()
                st.success("Configuração salva.")
            if st.button("Testar abertura do módulo", use_container_width=True):
                try:
                    executar_script_configurado(selecionado, modo)
                except Exception as exc:
                    st.error(f"Falha no teste: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col_right:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        panel("Criador rápido de módulo", "Crie um arquivo base já compatível com o portal quando precisar iniciar uma nova tela.")
        novo_nome = st.text_input("Nome do novo módulo", value="Novo Relatório")
        novo_arquivo = st.text_input("Nome do arquivo .py", value="novo_relatorio.py")
        if st.button("Criar módulo compatível", type="primary", use_container_width=True):
            if not novo_arquivo.lower().endswith(".py"):
                novo_arquivo += ".py"
            caminho = criar_template_modulo(novo_arquivo, novo_nome)
            config[caminho.name] = {**config_padrao_modulo(caminho.name), "nome": novo_nome, "ativo": True, "modo": "main_streamlit"}
            salvar_config_modulos(config)
            listar_scripts_python_cached.clear()
            st.success(f"Módulo criado: {caminho.name}")
        st.markdown("### Módulos ativos")
        ativos = []
        for arquivo, cfg in config.items():
            if cfg.get("ativo") and not cfg.get("ocultar"):
                ativos.append({"Arquivo": arquivo, "Nome": cfg.get("nome"), "Modo": cfg.get("modo"), "Ordem": cfg.get("ordem")})
        st.dataframe(ativos, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)


def pagina_modulo_dinamico(arquivo: str) -> None:
    config = carregar_config_modulos().get(arquivo, config_padrao_modulo(arquivo))
    page_header(config.get("nome", nome_amigavel_script(arquivo)), config.get("descricao", f"Módulo: {arquivo}"), "Módulo")
    try:
        executar_script_configurado(arquivo, config.get("modo", "auto"))
    except Exception as exc:
        logging.exception("Erro no modulo dinamico %s: %s", arquivo, exc)
        st.error(f"Erro ao executar módulo {arquivo}: {exc}")
        st.info("Abra o Editor de Módulos, teste o arquivo e ajuste o modo de execução.")


def pagina_historico(filtro: str | None = None, compact: bool = False) -> None:
    if not compact:
        page_header("Histórico de Processamentos", "Consulte execuções, duração e resultado de cada processamento.")
    dados = ler_json(HISTORY_FILE, [])
    if filtro:
        dados = [d for d in dados if filtro.lower() in str(d.get("tipo", "")).lower()]
    if not dados:
        st.info("Nenhum processamento registrado ainda.")
        return
    st.dataframe(dados, use_container_width=True, hide_index=True)


def pagina_relatorios() -> None:
    page_header("Relatórios", "Área preparada para consolidar indicadores e saídas geradas no portal.")
    h = ler_json(HISTORY_FILE, [])
    c1, c2, c3 = st.columns(3)
    c1.metric("Total de execuções", len(h))
    c2.metric("Sucessos", sum(1 for x in h if x.get("status") == "sucesso"))
    c3.metric("Erros", sum(1 for x in h if x.get("status") == "erro"))
    st.info("Os relatórios avançados podem ser conectados ao histórico e às planilhas finais geradas.")


def pagina_config() -> None:
    page_header("Configurações / Layout", "Ajustes visuais simples para apoiar futuras customizações.")
    st.info("Esta versão já usa a identidade azul da Expresso Nepomuceno, sem repetir a marca no rodapé.")
    st.code("Usuário Editor: pedro admin\nSenha Editor: admin pedro")


# =========================================================
# MAIN
# =========================================================
def main() -> None:
    aplicar_css()
    render_sidebar()
    render_topbar()

    page = current_page()
    config = carregar_config_modulos()
    mapa_dinamicos = {}
    for arquivo, cfg in sorted(config.items(), key=lambda item: int(item[1].get("ordem", 100)) if isinstance(item[1], dict) else 100):
        if isinstance(cfg, dict) and cfg.get("ativo") and not cfg.get("ocultar") and (BASE_DIR / arquivo).exists():
            key = f"mod_{arquivo}"
            mapa_dinamicos[key] = arquivo
            if st.sidebar.button(f"{cfg.get('icone', '🧩')}  {cfg.get('nome', nome_amigavel_script(arquivo))}", key=f"nav_{key}"):
                go(key)

    if page == "inicio":
        pagina_inicio()
    elif page == "permanencia":
        pagina_permanencia()
    elif page == "odometro":
        pagina_odometro()
    elif page == "editor":
        pagina_editor_modulos()
    elif page == "historico":
        pagina_historico()
    elif page == "relatorios":
        pagina_relatorios()
    elif page == "config":
        pagina_config()
    elif page in mapa_dinamicos:
        pagina_modulo_dinamico(mapa_dinamicos[page])
    else:
        st.session_state.page = "inicio"
        pagina_inicio()


if __name__ == "__main__":
    main()
