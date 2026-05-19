import base64
import hashlib
import json
import mimetypes
import os
import re
import sys
import tempfile
import traceback
import types
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

# =========================================================
# CONFIGURAÇÃO
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_NAME = "Central Operacional de Análises"
LOGIN_FILE = BASE_DIR / "Login.xlsx"
HISTORY_FILE = BASE_DIR / "portal_history.json"

# Máximo de tentativas de login antes de bloqueio temporário
MAX_LOGIN_ATTEMPTS = 5

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# IDENTIDADE VISUAL — paleta refinada monocromática azul
# =========================================================
COLORS = {
    "sidebar":    "#07142E",
    "sidebar_2":  "#0C1E42",
    "primary":    "#1340A0",
    "accent":     "#3B72F2",
    "accent_dim": "#2255D4",
    "bg":         "#F5F7FA",
    "surface":    "#FFFFFF",
    "border":     "#E2E8F0",
    "border_2":   "#CBD5E1",
    "text":       "#0F1F3D",
    "muted":      "#64748B",
    "success":    "#16A34A",
    "warning":    "#CA8A04",
    "danger":     "#DC2626",
}


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def find_file(candidates: list[str]) -> Path | None:
    for name in candidates:
        p = BASE_DIR / name
        if p.exists():
            return p
    normalized = {normalize_name(p.name): p for p in BASE_DIR.iterdir() if p.is_file()}
    for name in candidates:
        key = normalize_name(name)
        if key in normalized:
            return normalized[key]
    return None


def get_logo_path() -> Path | None:
    return find_file(
        [
            "logo_nepomuceno.jpeg",
            "logo_nepomuceno.jpg",
            "logo_nepomuceno.png",
            "logo_nepomuceno.png.jpeg",
            "Logo Nepomuceno.jpeg",
            "Logo Nepomuceno.jpg",
            "Logo Nepomuceno.png",
            "Expresso Nepomuceno.jpeg",
            "Expresso Nepomuceno.jpg",
            "Expresso Nepomuceno.png",
        ]
    )


def image_to_base64(path: Path | None) -> str:
    if not path or not path.exists():
        return ""
    try:
        return base64.b64encode(path.read_bytes()).decode("utf-8")
    except Exception:
        return ""


LOGO_PATH = get_logo_path()
LOGO_B64  = image_to_base64(LOGO_PATH)
LOGO_MIME = mimetypes.guess_type(str(LOGO_PATH))[0] if LOGO_PATH else "image/png"
if LOGO_PATH and LOGO_PATH.suffix.lower() in {".jpg", ".jpeg"}:
    LOGO_MIME = "image/jpeg"
LOGO_MIME = LOGO_MIME or "image/png"


# =========================================================
# ESTADO
# =========================================================
PAGES = {
    "inicio":        "Início",
    "odometro":      "Odômetro / Vínculo",
    "tempo":         "Tempo de Carregamento",
    "viagens":       "Viagens em Bloco",
    "historico":     "Histórico",
    "relatorios":    "Relatórios",
    "configuracoes": "Configurações",
}

PAGE_ICONS = {
    "inicio":        "🏠",
    "odometro":      "◴",
    "tempo":         "◔",
    "viagens":       "▦",
    "historico":     "☰",
    "relatorios":    "▥",
    "configuracoes": "⚙",
}


@dataclass
class DiagnosticItem:
    nome:    str
    arquivo: str
    objetivo: str
    status:  str
    detalhe: str


def init_state() -> None:
    defaults = {
        "authenticated":    False,
        "user_name":        "",
        "current_page":     "inicio",
        "portal_history":   load_history(),
        "analysis_report":  [],
        "last_error":       "",
        "login_attempts":   0,
        "login_locked_until": 0.0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# =========================================================
# LOGIN / HISTÓRICO
# =========================================================

def _hash_password(password: str) -> str:
    """Gera hash SHA-256 da senha para comparação segura."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def load_login_table() -> pd.DataFrame:
    """
    Carrega a tabela de usuários do Login.xlsx.
    Suporta senhas em texto plano (legado) e hashes SHA-256.
    Se o arquivo não existir, levanta erro — sem credencial padrão hardcoded.
    """
    if LOGIN_FILE.exists():
        try:
            df = pd.read_excel(LOGIN_FILE)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            pass
    # Sem credencial default hardcoded — retorna DataFrame vazio para forçar
    # configuração correta do arquivo de login.
    return pd.DataFrame(columns=["Chapa", "Senha"])


USERS_DF = load_login_table()


def _is_login_locked() -> bool:
    """Verifica se o login está bloqueado por excesso de tentativas."""
    import time
    locked_until = st.session_state.get("login_locked_until", 0.0)
    if locked_until > time.time():
        remaining = int(locked_until - time.time())
        st.error(f"Muitas tentativas incorretas. Aguarde {remaining}s antes de tentar novamente.")
        return True
    return False


def authenticate(username: str, password: str) -> bool:
    """
    Autentica o usuário comparando hash SHA-256.
    Registra tentativas falhas e bloqueia após MAX_LOGIN_ATTEMPTS.
    """
    import time

    if _is_login_locked():
        return False

    if USERS_DF.empty:
        st.error("Sistema de autenticação não configurado. Contate o administrador.")
        return False

    cols      = {normalize_name(c): c for c in USERS_DF.columns}
    user_col  = cols.get("chapa") or list(USERS_DF.columns)[0]
    pass_col  = cols.get("senha") or (list(USERS_DF.columns)[1] if len(USERS_DF.columns) > 1 else user_col)

    username = str(username).strip()
    password = str(password).strip()
    pwd_hash = _hash_password(password)

    for _, row in USERS_DF.iterrows():
        stored_user = str(row.get(user_col, "")).strip()
        stored_pass = str(row.get(pass_col, "")).strip()

        if stored_user != username:
            continue

        # Aceita hash SHA-256 (64 chars hex) ou texto plano legado
        match = (
            stored_pass == pwd_hash
            or stored_pass == password  # retrocompatibilidade com planilhas legadas
        )
        if match:
            st.session_state["user_name"]      = username
            st.session_state["login_attempts"] = 0
            st.session_state["login_locked_until"] = 0.0
            return True

    # Incrementa contador de falhas
    attempts = st.session_state.get("login_attempts", 0) + 1
    st.session_state["login_attempts"] = attempts
    if attempts >= MAX_LOGIN_ATTEMPTS:
        st.session_state["login_locked_until"] = time.time() + 60  # 60s de bloqueio
        st.session_state["login_attempts"]     = 0
        add_history("Login", "BLOQUEIO", f"Usuário '{username}' bloqueado por excesso de tentativas.")
    else:
        add_history("Login", "FALHA", f"Tentativa inválida para '{username}' ({attempts}/{MAX_LOGIN_ATTEMPTS}).")

    return False


def save_history(records: list[dict[str, Any]]) -> None:
    try:
        HISTORY_FILE.write_text(
            json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def load_history() -> list[dict[str, Any]]:
    if HISTORY_FILE.exists():
        try:
            data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []


def add_history(module_name: str, status: str, detail: str, output_name: str = "") -> None:
    records = st.session_state.get("portal_history", [])
    records.insert(
        0,
        {
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "usuario":   st.session_state.get("user_name", ""),
            "modulo":    module_name,
            "status":    status,
            "detalhe":   detail,
            "saida":     output_name,
        },
    )
    st.session_state["portal_history"] = records[:300]
    save_history(st.session_state["portal_history"])


# =========================================================
# CSS — Design Minimalista Profissional
# =========================================================

def apply_css() -> str:
    logo_html = (
        f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="np-logo" alt="Expresso Nepomuceno">'
        if LOGO_B64
        else '<div class="np-logo-fallback">EN</div>'
    )

    st.markdown(
        f"""
        <style>
        /* ── Variáveis globais ── */
        :root {{
            --c-sidebar:  {COLORS['sidebar']};
            --c-primary:  {COLORS['primary']};
            --c-accent:   {COLORS['accent']};
            --c-accent-d: {COLORS['accent_dim']};
            --c-bg:       {COLORS['bg']};
            --c-surface:  {COLORS['surface']};
            --c-border:   {COLORS['border']};
            --c-border2:  {COLORS['border_2']};
            --c-text:     {COLORS['text']};
            --c-muted:    {COLORS['muted']};
            --c-success:  {COLORS['success']};
            --c-warning:  {COLORS['warning']};
            --c-danger:   {COLORS['danger']};
        }}

        /* ── Base ── */
        html, body, [class*="css"], .stApp {{
            font-family: "Segoe UI", system-ui, sans-serif;
        }}
        .stApp {{
            background: var(--c-bg);
            color: var(--c-text);
        }}
        [data-testid="stHeader"] {{ background: transparent; }}
        #MainMenu, footer {{ visibility: hidden; }}

        .block-container {{
            max-width: 1400px;
            padding-top: 1.25rem;
            padding-bottom: 3rem;
        }}

        /* ── Sidebar ── */
        section[data-testid="stSidebar"] {{
            background: var(--c-sidebar);
            border-right: 1px solid rgba(255,255,255,0.06);
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}

        /* Botões de navegação da sidebar */
        section[data-testid="stSidebar"] .stButton > button {{
            width: 100%;
            justify-content: flex-start;
            min-height: 44px;
            border-radius: 10px !important;
            border: 1px solid transparent !important;
            background: transparent !important;
            color: rgba(255,255,255,0.75) !important;
            font-size: 14px !important;
            font-weight: 500 !important;
            letter-spacing: 0.01em;
            box-shadow: none !important;
            transition: background 0.15s, color 0.15s, border-color 0.15s;
        }}
        section[data-testid="stSidebar"] .stButton > button:hover {{
            background: rgba(255,255,255,0.07) !important;
            color: #FFFFFF !important;
            border-color: rgba(255,255,255,0.10) !important;
        }}
        section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: rgba(59,114,242,0.22) !important;
            border-color: rgba(59,114,242,0.50) !important;
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }}

        /* Brand sidebar */
        .np-sidebar-wrap {{
            padding: 8px 4px 20px 4px;
        }}
        .np-logo {{
            width: 120px;
            height: auto;
            max-height: 56px;
            display: block;
            object-fit: contain;
            margin-bottom: 20px;
        }}
        .np-logo-fallback {{
            width: 42px;
            height: 42px;
            border-radius: 10px;
            background: rgba(59,114,242,0.30);
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            font-weight: 700;
            margin-bottom: 20px;
            letter-spacing: -0.02em;
        }}
        .np-sidebar-title {{
            font-size: 15px;
            font-weight: 700;
            color: #FFFFFF;
            margin: 0;
            letter-spacing: -0.01em;
        }}
        .np-sidebar-subtitle {{
            font-size: 12px;
            color: rgba(255,255,255,0.52);
            margin-top: 3px;
            margin-bottom: 24px;
        }}
        .np-nav-label {{
            font-size: 10px;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            font-weight: 600;
            color: rgba(255,255,255,0.35) !important;
            margin: 0 4px 8px;
        }}
        .np-side-foot {{
            margin-top: 28px;
            padding: 12px 4px 0;
            border-top: 1px solid rgba(255,255,255,0.08);
            font-size: 11px;
            color: rgba(255,255,255,0.38) !important;
            line-height: 1.7;
        }}

        /* ── Topbar ── */
        .np-topbar {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 14px;
            padding: 14px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            margin-bottom: 22px;
        }}
        .np-topbar-left {{ display: flex; align-items: center; gap: 12px; }}
        .np-brand-circle {{
            width: 38px; height: 38px;
            border-radius: 10px;
            background: var(--c-primary);
            color: #FFFFFF;
            display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: 14px; letter-spacing: -0.02em;
        }}
        .np-top-title {{
            font-size: 15px; font-weight: 600;
            color: var(--c-text); letter-spacing: -0.01em;
        }}
        .np-top-subtitle {{
            font-size: 12px; color: var(--c-muted); margin-top: 1px;
        }}
        .np-topbar-right {{ display: flex; align-items: center; gap: 10px; }}
        .np-status-pill {{
            display: inline-flex; align-items: center; gap: 6px;
            padding: 6px 12px; border-radius: 999px;
            background: rgba(22,163,74,0.08);
            border: 1px solid rgba(22,163,74,0.20);
            color: var(--c-success) !important;
            font-size: 12px; font-weight: 600;
        }}
        .np-dot {{
            width: 7px; height: 7px; border-radius: 50%;
            background: var(--c-success);
        }}
        .np-user-chip {{
            width: 36px; height: 36px; border-radius: 999px;
            background: #EEF2FF;
            color: var(--c-primary);
            display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: 12px;
        }}
        .np-user-label {{ font-size: 13px; font-weight: 600; color: var(--c-text); }}
        .np-user-sub   {{ font-size: 11px; color: var(--c-muted); }}

0.04) 100%);
            border-left: 1px solid rgba(255,255,255,0.07);
            pointer-events: none;
        }}
        .np-hero-content {{
            position: relative; z-index: 1;
            max-width: clamp(360px, 52%, 680px);
        }}
        .np-hero-eyebrow {{
            font-size: 12px; font-weight: 600;
            letter-spacing: 0.10em; text-transform: uppercase;
            color: rgba(255,255,255,0.55) !important;
            margin-bottom: 12px;
        }}
        .np-hero-title {{
            color: #FFFFFF !important;
            font-size: clamp(28px, 3.2vw, 44px);
            line-height: 1.10;
            font-weight: 700;
            letter-spacing: -0.02em;
            margin: 0 0 14px 0;
        }}
        .np-hero-text {{
            color: rgba(255,255,255,0.72) !important;
            font-size: 14px;
            line-height: 1.65;
            max-width: 480px;
        }}
        .np-hero-brand {{
            position: absolute; right: 36px; top: 50%;
            transform: translateY(-50%);
            z-index: 2; text-align: center;
            display: flex; flex-direction: column; align-items: center; gap: 10px;
        }}
        .np-hero-brand-ring {{
            width: clamp(160px, 20vw, 260px);
            height: clamp(160px, 20vw, 260px);
            border-radius: 50%;
            border: 1px solid rgba(255,255,255,0.14);
            display: flex; align-items: center; justify-content: center;
            background: rgba(255,255,255,0.06);
        }}
        .np-hero-brand img {{
            width: clamp(130px, 16vw, 210px);
            height: auto;
            max-height: 170px;
            object-fit: contain;
            mix-blend-mode: luminosity;
            filter: brightness(1.8) contrast(0.85);
        }}
        .np-hero-mark {{
            font-size: 10px; font-weight: 700;
            letter-spacing: 0.14em; text-transform: uppercase;
            color: rgba(255,255,255,0.28) !important;
            margin-top: 4px;
        }}
        .np-stat-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            margin-top: 28px;
            max-width: 560px;
        }}
        .np-stat {{
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(255,255,255,0.12);
            border-radius: 10px;
            padding: 14px 16px;
        }}
        .np-stat-value {{
            font-size: 22px; font-weight: 700;
            color: #FFFFFF !important;
            letter-spacing: -0.02em;
        }}
        .np-stat-label {{
            font-size: 11px; margin-top: 2px;
            color: rgba(255,255,255,0.55) !important;
        }}

        /* ── Seção ── */
        .np-section-title {{
            font-size: 16px; font-weight: 700;
            color: var(--c-text);
            letter-spacing: -0.01em;
            margin: 4px 0 2px;
        }}
        .np-section-subtitle {{
            font-size: 13px; color: var(--c-muted);
            margin-bottom: 16px;
        }}
        .np-divider {{
            border: none;
            border-top: 1px solid var(--c-border);
            margin: 6px 0 20px;
        }}

        /* ── Cards ── */
        .np-card {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 14px;
            padding: 20px;
            margin-bottom: 12px;
        }}
        .np-module-card {{
            min-height: 240px;
            display: flex; flex-direction: column;
            justify-content: space-between;
            transition: border-color 0.15s;
        }}
        .np-module-card:hover {{ border-color: var(--c-border2); }}
        .np-icon {{
            width: 40px; height: 40px; border-radius: 10px;
            background: #EEF2FF; color: var(--c-primary);
            display: flex; align-items: center; justify-content: center;
            font-size: 18px; margin-bottom: 14px;
        }}
        .np-card-title {{
            font-size: 15px; font-weight: 700;
            color: var(--c-text); margin-bottom: 6px;
            letter-spacing: -0.01em;
        }}
        .np-card-text {{
            font-size: 13px; line-height: 1.6;
            color: var(--c-muted); min-height: 56px;
        }}
        .np-metrics-box {{
            display: flex; gap: 0;
            border: 1px solid var(--c-border);
            border-radius: 10px;
            overflow: hidden;
            margin-top: 14px;
        }}
        .np-metric-item {{
            flex: 1;
            padding: 10px 14px;
            border-right: 1px solid var(--c-border);
        }}
        .np-metric-item:last-child {{ border-right: none; }}
        .np-metric-value {{
            font-size: 16px; font-weight: 700;
            color: var(--c-text); letter-spacing: -0.01em;
        }}
        .np-metric-label {{
            font-size: 11px; color: var(--c-muted); margin-top: 1px;
        }}

        /* Info/notes card */
        .np-info-card {{
            background: var(--c-bg);
            border: 1px solid var(--c-border);
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 12px;
        }}
        .np-info-title {{
            font-size: 14px; font-weight: 600;
            color: var(--c-text); margin-bottom: 4px;
        }}
        .np-info-text {{
            font-size: 13px; color: var(--c-muted); line-height: 1.6;
        }}
        .np-notes-list {{
            padding-left: 16px; color: var(--c-muted); line-height: 1.75;
            font-size: 13px;
        }}

        /* ── Botões globais ── */
        .stButton > button, .stDownloadButton > button {{
            border-radius: 10px !important;
            min-height: 42px !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            letter-spacing: -0.01em;
            transition: opacity 0.15s !important;
        }}
        .stButton > button[kind="primary"],
        .stDownloadButton > button {{
            background: var(--c-primary) !important;
            color: #FFFFFF !important;
            border: 0 !important;
        }}
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button:hover {{
            background: var(--c-accent-d) !important;
        }}

        /* ── Métricas ── */
        [data-testid="stMetric"] {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 12px;
            padding: 14px 16px;
        }}

        /* ── Upload ── */
        [data-testid="stFileUploader"] section {{
            border: 1px dashed var(--c-border2);
            border-radius: 12px;
            background: #FAFBFF;
        }}

        /* ── DataFrame ── */
        [data-testid="stDataFrame"] {{
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid var(--c-border);
        }}

        /* ── Login ── */
        .np-login-bg {{
            min-height: 100vh;
            background: var(--c-bg);
        }}
        .np-login-wrap {{
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 2rem 1rem;
        }}

        /* Card centralizado via colunas do Streamlit */
        .np-login-card {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 18px;
            padding: 36px 32px 28px;
        }}
        .np-login-logo {{
            width: 120px;
            height: auto;
            max-height: 52px;
            object-fit: contain;
            display: block;
            margin-bottom: 24px;
        }}
        .np-login-logo-fb {{
            width: 40px; height: 40px;
            background: var(--c-primary);
            border-radius: 10px;
            display: flex; align-items: center; justify-content: center;
            color: #fff; font-weight: 700; font-size: 15px;
            margin-bottom: 24px;
        }}
        .np-login-title {{
            font-size: 22px; font-weight: 700;
            color: var(--c-text); letter-spacing: -0.02em;
            margin: 0 0 6px;
        }}
        .np-login-sub {{
            font-size: 13px; color: var(--c-muted); line-height: 1.5;
            margin-bottom: 24px;
        }}
        .np-login-footer {{
            font-size: 11px; color: var(--c-muted);
            text-align: center; margin-top: 16px;
            padding-top: 16px;
            border-top: 1px solid var(--c-border);
        }}

        /* ── Responsivo ── */
        @media (max-width: 1100px) {{
            .np-hero-content  {{ max-width: 100%; }}
            .np-hero-brand    {{
                position: relative; right: auto; top: auto;
                transform: none; margin-top: 20px; text-align: left;
            }}
            .np-hero-brand img {{ max-width: 200px; }}
            .np-stat-grid      {{ grid-template-columns: repeat(2, 1fr); }}
        }}
        @media (max-width: 700px) {{
            .np-hero {{ padding: 28px 22px; }}
            .np-hero-title {{ font-size: 26px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    return logo_html


# =========================================================
# UTILITÁRIOS DE MÓDULO
# =========================================================

def strip_autorun_block(source: str) -> str:
    marker = 'if __name__ == "__main__":'
    idx = source.find(marker)
    if idx != -1:
        return source[:idx]
    return source


class DummyTkWindow:
    def __init__(self, *args, **kwargs): pass
    def title(self, *a, **k): pass
    def geometry(self, *a, **k): pass
    def resizable(self, *a, **k): pass
    def configure(self, *a, **k): pass
    def mainloop(self, *a, **k): pass
    def destroy(self, *a, **k): pass
    def update_idletasks(self, *a, **k): pass


class DummyStyle:
    def theme_use(self, *a, **k): pass
    def configure(self, *a, **k): pass


class DummyWidget:
    def __init__(self, *a, **k): pass
    def pack(self, *a, **k): pass
    def grid(self, *a, **k): pass
    def config(self, *a, **k): pass
    configure = config


def ensure_fake_tk_modules() -> dict[str, Any]:
    backups: dict[str, Any] = {}
    if "tkinter" not in sys.modules:
        tk_module = types.ModuleType("tkinter")
        tk_module.Tk    = DummyTkWindow
        tk_module.BOTH  = tk_module.X = tk_module.Y = tk_module.LEFT = tk_module.RIGHT = 0
        tk_module.END   = "end"
        backups["tkinter"] = None
        sys.modules["tkinter"] = tk_module
    if "tkinter.filedialog" not in sys.modules:
        filedialog = types.ModuleType("tkinter.filedialog")
        filedialog.askopenfilename = lambda *a, **k: ""
        backups["tkinter.filedialog"] = None
        sys.modules["tkinter.filedialog"] = filedialog
    if "tkinter.messagebox" not in sys.modules:
        messagebox = types.ModuleType("tkinter.messagebox")
        messagebox.showinfo  = lambda *a, **k: None
        messagebox.showerror = lambda *a, **k: None
        backups["tkinter.messagebox"] = None
        sys.modules["tkinter.messagebox"] = messagebox
    if "tkinter.ttk" not in sys.modules:
        ttk = types.ModuleType("tkinter.ttk")
        ttk.Style       = DummyStyle
        ttk.Frame       = DummyWidget
        ttk.Label       = DummyWidget
        ttk.Progressbar = DummyWidget
        ttk.Button      = DummyWidget
        backups["tkinter.ttk"] = None
        sys.modules["tkinter.ttk"] = ttk
    return backups


def restore_modules(backups: dict[str, Any]) -> None:
    for name, original in backups.items():
        if original is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = original


def load_module_safely(path: Path, use_fake_tk: bool = False):
    source  = path.read_text(encoding="utf-8", errors="replace")
    source  = strip_autorun_block(source)
    module  = types.ModuleType(f"portal_{path.stem}")
    module.__file__                    = str(path)
    module.__dict__["__name__"]        = module.__name__
    module.__dict__["__package__"]     = None

    backups: dict[str, Any] = {}
    try:
        if use_fake_tk:
            backups = ensure_fake_tk_modules()
        exec(compile(source, str(path), "exec"), module.__dict__)
        return module
    finally:
        if backups:
            restore_modules(backups)


def get_module_diagnostics() -> list[DiagnosticItem]:
    items: list[DiagnosticItem] = []

    od_file = find_file([
        "app_odometro_streamlit_corrigido.py",
        "app_odometro_streamlit_corrigido (1).py",
        "Python_Odometro_Vinculo.py",
    ])
    if od_file and od_file.exists():
        try:
            mod     = load_module_safely(od_file)
            required = ["processar_streamlit", "preparar_bases", "gerar_resultado", "exportar"]
            missing  = [n for n in required if not hasattr(mod, n)]
            if missing:
                items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name,
                    "Cruzar combustível, maxtrack, ativos e produção",
                    "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name,
                    "Cruzar combustível, maxtrack, ativos e produção",
                    "OK", "Módulo integrado com regras V15 preservadas."))
        except Exception as exc:
            items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name,
                "Cruzar combustível, maxtrack, ativos e produção", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Odômetro / Vínculo", "app_odometro_streamlit_corrigido.py",
            "Cruzar combustível, maxtrack, ativos e produção", "Erro", "Arquivo não localizado."))

    tc_file = find_file(["Python_Tempo_Carregamento.py"])
    if tc_file and tc_file.exists():
        try:
            mod      = load_module_safely(tc_file)
            required = ["processar_arquivo", "main_streamlit"]
            missing  = [n for n in required if not hasattr(mod, n)]
            if missing:
                items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name,
                    "Tratar permanência por área e consolidar eventos",
                    "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name,
                    "Tratar permanência por área e consolidar eventos",
                    "OK", "Estrutura principal encontrada."))
        except Exception as exc:
            items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name,
                "Tratar permanência por área e consolidar eventos", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Tempo de Carregamento", "Python_Tempo_Carregamento.py",
            "Tratar permanência por área e consolidar eventos", "Erro", "Arquivo não localizado."))

    vb_file = find_file(["Python_Viagens_Bloco.py"])
    if vb_file and vb_file.exists():
        try:
            mod      = load_module_safely(vb_file, use_fake_tk=True)
            required = ["processar_arquivos", "main_streamlit"]
            missing  = [n for n in required if not hasattr(mod, n)]
            if missing:
                items.append(DiagnosticItem("Viagens em Bloco", vb_file.name,
                    "Integrar Maxtrack, SAP e permanência",
                    "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Viagens em Bloco", vb_file.name,
                    "Integrar Maxtrack, SAP e permanência",
                    "OK", "Estrutura Streamlit segura. Interface tkinter preservada para uso local."))
        except Exception as exc:
            items.append(DiagnosticItem("Viagens em Bloco", vb_file.name,
                "Integrar Maxtrack, SAP e permanência", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Viagens em Bloco", "Python_Viagens_Bloco.py",
            "Integrar Maxtrack, SAP e permanência", "Erro", "Arquivo não localizado."))

    return items


# =========================================================
# COMPONENTES VISUAIS
# =========================================================

def render_topbar() -> None:
    initials   = (st.session_state.get("user_name") or "AD")[:2].upper()
    user_label = st.session_state.get("user_name") or "Administrador"
    st.markdown(
        f"""
        <div class="np-topbar">
            <div class="np-topbar-left">
                <div class="np-brand-circle">EN</div>
                <div>
                    <div class="np-top-title">Central Operacional de Análises</div>
                    <div class="np-top-subtitle">Portal integrado &nbsp;·&nbsp; <span id="np-clock">--/--/---- --:--</span></div>
                </div>
            </div>
            <div class="np-topbar-right">
                <div class="np-status-pill"><span class="np-dot"></span> Online</div>
                <div class="np-user-chip">{initials}</div>
                <div>
                    <div class="np-user-label">{user_label}</div>
                    <div class="np-user-sub">Operador</div>
                </div>
            </div>
        </div>
        <script>
        (function() {{
            function pad(n) {{ return String(n).padStart(2, '0'); }}
            function tick() {{
                var el = document.getElementById('np-clock');
                if (!el) return;
                var now = new Date();
                var d = pad(now.getDate()) + '/' + pad(now.getMonth()+1) + '/' + now.getFullYear();
                var t = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
                el.textContent = d + ' ' + t;
            }}
            tick();
            setInterval(tick, 1000);
        }})();
        </script>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar(logo_html: str) -> str:
    with st.sidebar:
        st.markdown(
            f"""
            <div class="np-sidebar-wrap">
                {logo_html}
                <div class="np-sidebar-title">Central Operacional</div>
                <div class="np-sidebar-subtitle">Expresso Nepomuceno</div>
                <div class="np-nav-label">Navegação</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        for page_key, label in PAGES.items():
            icon = PAGE_ICONS.get(page_key, "")
            kind = "primary" if st.session_state["current_page"] == page_key else "secondary"
            if st.button(f"{icon}  {label}", key=f"nav_{page_key}", type=kind, use_container_width=True):
                st.session_state["current_page"] = page_key
                st.rerun()

        st.markdown(
            """
            <div class="np-side-foot">
                <div>Versão&nbsp;&nbsp;4.1.0</div>
                <div>Ambiente&nbsp;&nbsp;Interno</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    return st.session_state["current_page"]


def render_module_card(
    icon: str,
    title: str,
    desc: str,
    stats: list[tuple[str, str]],
    button_label: str,
    page_key: str,
) -> None:
    stats_html = "".join(
        f'<div class="np-metric-item">'
        f'<div class="np-metric-value">{v}</div>'
        f'<div class="np-metric-label">{l}</div>'
        f'</div>'
        for v, l in stats
    )
    st.markdown(
        f"""
        <div class="np-card np-module-card">
            <div>
                <div class="np-icon">{icon}</div>
                <div class="np-card-title">{title}</div>
                <div class="np-card-text">{desc}</div>
                <div class="np-metrics-box">{stats_html}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(button_label, key=f"card_{page_key}", use_container_width=True):
        st.session_state["current_page"] = page_key
        st.rerun()


# =========================================================
# TELA DE LOGIN — centralizada via colunas nativas
# =========================================================

def render_login_screen() -> None:
    # Espaçador superior
    st.markdown("<div style='height:6vh'></div>", unsafe_allow_html=True)

    # Centralização via colunas nativas do Streamlit (mais confiável que CSS flex)
    _, col, _ = st.columns([1, 1.4, 1])

    with col:
        # Logo
        if LOGO_B64:
            st.markdown(
                f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="np-login-logo" alt="Logo">',
                unsafe_allow_html=True,
            )
        else:
            st.markdown('<div class="np-login-logo-fb">EN</div>', unsafe_allow_html=True)

        # Títulos
        st.markdown('<div class="np-login-title">Acesso ao Portal</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="np-login-sub">Identifique-se com suas credenciais de acesso.</div>',
            unsafe_allow_html=True,
        )

        # Formulário
        with st.form("login_form", clear_on_submit=False):
            username  = st.text_input("Chapa / Usuário", placeholder="Ex: 123456")
            password  = st.text_input("Senha", type="password", placeholder="••••••••")
            submitted = st.form_submit_button(
                "Entrar", use_container_width=True, type="primary"
            )

        if submitted:
            if not username or not password:
                st.warning("Preencha usuário e senha.")
            elif authenticate(username, password):
                st.session_state["authenticated"] = True
                add_history("Login", "OK", "Acesso realizado com sucesso.")
                st.rerun()
            else:
                attempts = st.session_state.get("login_attempts", 0)
                remaining = max(0, MAX_LOGIN_ATTEMPTS - attempts)
                st.error(f"Credenciais inválidas. Tentativas restantes: {remaining}.")

        # Rodapé discreto — sem revelar detalhes de implementação
        st.markdown(
            '<div class="np-login-footer">Acesso restrito. Em caso de dúvidas, contate o administrador.</div>',
            unsafe_allow_html=True,
        )


# =========================================================
# PÁGINAS
# =========================================================

def page_inicio() -> None:
    history     = st.session_state.get("portal_history", [])
    today_str   = datetime.now().strftime("%d/%m/%Y")
    total_today = sum(1 for r in history if r.get("data_hora", "").startswith(today_str))
    total_ok    = sum(1 for r in history if r.get("status") == "OK")

    if LOGO_B64:
        brand_inner = f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" alt="Expresso Nepomuceno">'
    else:
        brand_inner = '<span style="color:rgba(255,255,255,0.7);font-size:28px;font-weight:800;letter-spacing:-0.03em;">EN</span>'

    st.markdown(
        f"""
        <div class="np-hero">
            <div class="np-hero-content">
                <div class="np-hero-eyebrow">Bem-vindo à</div>
                <h1 class="np-hero-title">Central Operacional<br>de Análises.</h1>
                <div class="np-hero-text">
                    Gerencie e acompanhe os processos de análise de forma
                    centralizada e eficiente.
                </div>
                <div class="np-stat-grid">
                    <div class="np-stat">
                        <div class="np-stat-value">{total_today}</div>
                        <div class="np-stat-label">Hoje</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{total_ok}</div>
                        <div class="np-stat-label">Concluídos</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{len(history)}</div>
                        <div class="np-stat-label">Histórico</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{max(len(history) - total_ok, 0)}</div>
                        <div class="np-stat-label">Erros</div>
                    </div>
                </div>
            </div>
            <div class="np-hero-brand">
                <div class="np-hero-brand-ring">
                    {brand_inner}
                </div>
                <div class="np-hero-mark">Expresso Nepomuceno</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="np-section-title">Módulos principais</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Selecione a ferramenta operacional desejada.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        render_module_card("◴", "Odômetro / Vínculo",
            "Cruza combustível, Maxtrack, ativos e produção para gerar o ODOMETRO_MATCH.",
            [("4", "bases"), ("Excel", "saída")], "Acessar →", "odometro")
    with c2:
        render_module_card("◔", "Tempo de Carregamento",
            "Trata permanência por área, ruídos e consolida tempos operacionais.",
            [("1", "base"), ("Resumo", "auto")], "Acessar →", "tempo")
    with c3:
        render_module_card("▦", "Viagens em Bloco",
            "Integra Maxtrack, SAP e permanência para viagens validadas e deduplicadas.",
            [("3", "bases"), ("Score", "dedup")], "Acessar →", "viagens")

    c4, c5 = st.columns(2)
    with c4:
        render_module_card("☰", "Histórico",
            "Consulte o histórico completo de processamentos realizados no portal.",
            [(str(len(history)), "registros"), (datetime.now().strftime("%H:%M"), "agora")],
            "Ver histórico →", "historico")
    with c5:
        render_module_card("▥", "Relatórios",
            "Acompanhe indicadores de uso, módulos e resultados gerados.",
            [(str(total_ok), "sucessos"), (str(max(len(history) - total_ok, 0)), "erros")],
            "Ver relatórios →", "relatorios")


def page_odometro() -> None:
    st.markdown('<div class="np-section-title">Odômetro / Vínculo</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Envie as quatro bases obrigatórias para gerar o arquivo consolidado final.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    module_path = find_file([
        "app_odometro_streamlit_corrigido.py",
        "app_odometro_streamlit_corrigido (1).py",
        "Python_Odometro_Vinculo.py",
    ])
    if not module_path:
        st.error("Arquivo app_odometro_streamlit_corrigido.py não encontrado.")
        return

    col1, col2 = st.columns(2)
    with col1:
        comb   = st.file_uploader("1. Base Combustível",        type=["xlsx", "xls"], key="odom_comb")
        ativo  = st.file_uploader("3. Base Ativo de Veículos",  type=["xlsx", "xls"], key="odom_ativo")
    with col2:
        maxtrack = st.file_uploader("2. Base Km Rodado Maxtrack",      type=["xlsx", "xls"], key="odom_max")
        producao = st.file_uploader("4. Produção Oficial / Cliente",   type=["xlsx", "xls"], key="odom_prod")

    output_name = st.text_input(
        "Nome do arquivo de saída",
        value=f"resultado_odometro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    )
    progress = sum(x is not None for x in [comb, maxtrack, ativo, producao]) / 4
    st.progress(progress)
    st.caption(f"{int(progress * 4)} de 4 arquivos carregados")

    if st.button(
        "Processar Odômetro / Vínculo",
        key="btn_processar_odometro",
        type="primary",
        use_container_width=True,
        disabled=not all([comb, maxtrack, ativo, producao]),
    ):
        try:
            mod = load_module_safely(module_path)
            with st.spinner("Executando processamento do módulo..."):
                excel_bytes, indicadores, resultado_final = mod.processar_streamlit(
                    comb, maxtrack, ativo, producao, output_name
                )

            st.success("Processamento concluído com sucesso.")
            add_history("Odômetro / Vínculo", "OK", "Processamento executado.", output_name)

            if isinstance(indicadores, pd.DataFrame) and not indicadores.empty:
                st.markdown("### Indicadores finais")
                st.dataframe(indicadores, use_container_width=True, hide_index=True)

            if isinstance(resultado_final, pd.DataFrame) and not resultado_final.empty:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Registros", f"{len(resultado_final):,}".replace(",", "."))
                od_match = int(resultado_final["ODOMETRO_MATCH"].notna().sum()) if "ODOMETRO_MATCH" in resultado_final.columns else 0
                c2.metric("ODOMETRO_MATCH", f"{od_match:,}".replace(",", "."))
                salto = int((resultado_final["KM_ENTRE_ABAST"] > 1500).sum()) if "KM_ENTRE_ABAST" in resultado_final.columns else 0
                c3.metric("Saltos > 1500", salto)
                c4.metric("Placas", resultado_final["PLACA"].nunique() if "PLACA" in resultado_final.columns else 0)
                st.markdown("### Prévia do resultado")
                st.dataframe(resultado_final.head(100), use_container_width=True)

            st.download_button(
                "Baixar Excel consolidado",
                data=excel_bytes,
                file_name=output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Erro no módulo de odômetro: {exc}")
            st.code(traceback.format_exc(), language="text")
            add_history("Odômetro / Vínculo", "ERRO", str(exc), output_name)


def page_tempo() -> None:
    st.markdown('<div class="np-section-title">Tempo de Carregamento</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Envie a base de permanência para tratamento e consolidação automática.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    module_path = find_file(["Python_Tempo_Carregamento.py"])
    if not module_path:
        st.error("Arquivo Python_Tempo_Carregamento.py não encontrado.")
        return

    arquivo     = st.file_uploader("Base de permanência / carregamento", type=["xlsx", "xls"], key="tempo_arquivo")
    output_name = st.text_input(
        "Nome do arquivo de saída",
        value=f"tempo_carregamento_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    )

    if st.button(
        "Processar Tempo de Carregamento",
        key="btn_processar_tempo",
        type="primary",
        use_container_width=True,
        disabled=arquivo is None,
    ):
        try:
            mod = load_module_safely(module_path)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_in:
                temp_in.write(arquivo.read())
                path_in = temp_in.name

            path_out = str(
                BASE_DIR / (output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx")
            )

            with st.spinner("Tratando arquivo..."):
                saida, tratado, resumo_placa_local, resumo_local, info = mod.processar_arquivo(
                    path_in, path_out
                )

            st.success("Processamento concluído com sucesso.")
            add_history("Tempo de Carregamento", "OK", "Processamento executado.", Path(saida).name)

            c1, c2, c3 = st.columns(3)
            c1.metric("Linhas tratadas",     len(tratado))
            c2.metric("Resumo placa/local",  len(resumo_placa_local))
            c3.metric("Resumo local",        len(resumo_local))

            st.markdown("### Colunas identificadas")
            st.json(info)
            st.markdown("### Prévia da base tratada")
            st.dataframe(tratado.head(100), use_container_width=True)
            st.markdown("### Resumo por placa e local")
            st.dataframe(resumo_placa_local.head(100), use_container_width=True)
            st.markdown("### Resumo por local")
            st.dataframe(resumo_local.head(100), use_container_width=True)

            with open(saida, "rb") as f:
                st.download_button(
                    "Baixar Excel tratado",
                    f.read(),
                    file_name=Path(saida).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception as exc:
            st.error(f"Erro no módulo de tempo de carregamento: {exc}")
            st.code(traceback.format_exc(), language="text")
            add_history("Tempo de Carregamento", "ERRO", str(exc), output_name)


class StreamlitProgressAdapter:
    def __init__(self):
        self.bar    = st.progress(0)
        self.status = st.empty()
        self.total  = 100

    def set_total(self, total: int):
        self.total = total or 100

    def update(self, value: int, etapa: str = "", detalhe: str = ""):
        pct = max(0, min(int(value), 100))
        self.bar.progress(pct / 100)
        msg = f"{pct}% — {etapa}".strip()
        if detalhe:
            msg += f" | {detalhe}"
        self.status.info(msg)


def page_viagens() -> None:
    st.markdown('<div class="np-section-title">Viagens em Bloco</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Envie os três arquivos para integração, validação SAP e deduplicação.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    module_path = find_file(["Python_Viagens_Bloco.py"])
    if not module_path:
        st.error("Arquivo Python_Viagens_Bloco.py não encontrado.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        arq_max  = st.file_uploader("1. Arquivo Maxtrack",    type=["xlsx", "xls"], key="vb_max")
    with c2:
        arq_sap  = st.file_uploader("2. Arquivo SAP",         type=["xlsx", "xls"], key="vb_sap")
    with c3:
        arq_perm = st.file_uploader("3. Arquivo Permanência", type=["xlsx", "xls"], key="vb_perm")

    if st.button(
        "Processar Viagens em Bloco",
        key="btn_processar_viagens",
        type="primary",
        use_container_width=True,
        disabled=not all([arq_max, arq_sap, arq_perm]),
    ):
        try:
            mod = load_module_safely(module_path, use_fake_tk=True)

            def save_uploaded(uploaded_file) -> str:
                suffix = Path(uploaded_file.name).suffix or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(uploaded_file.read())
                    return tmp.name

            p_max  = save_uploaded(arq_max)
            p_sap  = save_uploaded(arq_sap)
            p_perm = save_uploaded(arq_perm)

            progress = StreamlitProgressAdapter()
            with st.spinner("Processando viagens..."):
                saida, total_viagens, total_placas, total_invalidas, total_duplicadas, total_antes = (
                    mod.processar_arquivos(p_max, p_sap, p_perm, progress)
                )

            st.success("Processamento concluído com sucesso.")
            add_history("Viagens em Bloco", "OK", "Processamento executado.", Path(saida).name)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Antes da dedup",  total_antes)
            c2.metric("Viagens finais",  total_viagens)
            c3.metric("Válidas",         total_viagens - total_invalidas)
            c4.metric("Inválidas",       total_invalidas)
            c5.metric("Placas",          total_placas)
            st.metric("Duplicadas descartadas", total_duplicadas)

            with open(saida, "rb") as f:
                st.download_button(
                    "Baixar arquivo de viagens",
                    f.read(),
                    file_name=Path(saida).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception as exc:
            st.error(f"Erro no módulo de viagens em bloco: {exc}")
            st.code(traceback.format_exc(), language="text")
            add_history("Viagens em Bloco", "ERRO", str(exc), "")


def page_historico() -> None:
    st.markdown('<div class="np-section-title">Histórico</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Registro de processamentos e eventos do portal.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    records = st.session_state.get("portal_history", [])
    if not records:
        st.info("Nenhum registro encontrado até o momento.")
        return

    df = pd.DataFrame(records)
    st.dataframe(df, use_container_width=True, hide_index=True)


def page_relatorios() -> None:
    st.markdown('<div class="np-section-title">Relatórios</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Indicadores de uso do portal e execução dos módulos.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    records = pd.DataFrame(st.session_state.get("portal_history", []))
    if records.empty:
        st.info("Sem dados suficientes para montar relatórios.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de execuções", len(records))
    c2.metric("Sucessos",           int((records["status"] == "OK").sum()))
    c3.metric("Erros",              int((records["status"] == "ERRO").sum()))

    if "modulo" in records.columns:
        resumo = (
            records.groupby(["modulo", "status"], dropna=False)
            .size()
            .reset_index(name="quantidade")
        )
        st.markdown("### Resumo por módulo")
        st.dataframe(resumo, use_container_width=True, hide_index=True)


def page_configuracoes() -> None:
    st.markdown('<div class="np-section-title">Configurações</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Diagnóstico dos módulos e manutenção do portal.</div>', unsafe_allow_html=True)
    st.markdown('<hr class="np-divider">', unsafe_allow_html=True)

    diagnostics = get_module_diagnostics()
    st.session_state["analysis_report"] = diagnostics

    st.markdown("### Diagnóstico dos módulos")
    diag_df = pd.DataFrame([
        {
            "Módulo":   item.nome,
            "Arquivo":  item.arquivo,
            "Objetivo": item.objetivo,
            "Status":   item.status,
            "Detalhe":  item.detalhe,
        }
        for item in diagnostics
    ])
    st.dataframe(diag_df, use_container_width=True, hide_index=True)

    st.markdown("### Observações")
    st.markdown(
        """
        <div class="np-info-card">
            <div class="np-info-title">Pontos identificados</div>
            <div class="np-info-text">
                <ul class="np-notes-list">
                    <li><strong>Viagens em Bloco:</strong> arquivo original usa <code>tkinter</code>. O portal carrega via adaptador seguro sem alterar a lógica de processamento.</li>
                    <li><strong>Odômetro / Vínculo:</strong> módulo com auto-renderização detectada. Carregado de forma controlada — apenas a função de processamento é chamada.</li>
                    <li><strong>Tempo de Carregamento:</strong> integração direta via <code>processar_arquivo</code>.</li>
                </ul>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Recarregar diagnóstico", key="btn_recarregar_diagnostico", use_container_width=True):
            st.session_state["analysis_report"] = get_module_diagnostics()
            st.success("Diagnóstico atualizado.")
    with c2:
        if st.button("Limpar histórico", key="btn_limpar_historico", use_container_width=True):
            st.session_state["portal_history"] = []
            save_history([])
            st.success("Histórico limpo com sucesso.")

    st.markdown("---")
    if st.button("Sair do portal", key="btn_sair_portal", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["current_page"]  = "inicio"
        st.rerun()


# =========================================================
# APP PRINCIPAL
# =========================================================

def main() -> None:
    init_state()
    logo_html = apply_css()

    if not st.session_state.get("authenticated"):
        render_login_screen()
        return

    current_page = render_sidebar(logo_html)
    render_topbar()

    dispatch = {
        "inicio":        page_inicio,
        "odometro":      page_odometro,
        "tempo":         page_tempo,
        "viagens":       page_viagens,
        "historico":     page_historico,
        "relatorios":    page_relatorios,
        "configuracoes": page_configuracoes,
    }

    page_fn = dispatch.get(current_page)
    if page_fn:
        page_fn()
    else:
        st.error("Página não encontrada.")


if __name__ == "__main__":
    main()
