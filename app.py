"""
Central Operacional de Análises — Versão Aprimorada
Design profissional com login otimizado, dashboard visual e navegação intuitiva.
"""
import hashlib
import json
import mimetypes
import os
import re
import sys
import tempfile
import traceback
import types
import base64
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

MAX_LOGIN_ATTEMPTS = 5

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# PALETA DE CORES — Profissional Azul Escuro
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

# =========================================================
# UTILITÁRIOS
# =========================================================

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
    return find_file([
        "logo_nepomuceno.jpeg", "logo_nepomuceno.jpg", "logo_nepomuceno.png",
        "Logo Nepomuceno.jpeg", "Logo Nepomuceno.jpg", "Logo Nepomuceno.png",
        "Expresso Nepomuceno.jpeg", "Expresso Nepomuceno.jpg", "Expresso Nepomuceno.png",
    ])


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
# AUTENTICAÇÃO & HISTÓRICO
# =========================================================

def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def load_login_table() -> pd.DataFrame:
    if LOGIN_FILE.exists():
        try:
            df = pd.read_excel(LOGIN_FILE)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=["Chapa", "Senha"])


USERS_DF = load_login_table()


def _is_login_locked() -> bool:
    import time
    locked_until = st.session_state.get("login_locked_until", 0.0)
    if locked_until > time.time():
        remaining = int(locked_until - time.time())
        return True
    return False


def authenticate(username: str, password: str) -> bool:
    import time

    if _is_login_locked():
        return False

    if USERS_DF.empty:
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

        match = (stored_pass == pwd_hash or stored_pass == password)
        if match:
            st.session_state["user_name"]      = username
            st.session_state["login_attempts"] = 0
            st.session_state["login_locked_until"] = 0.0
            add_history("Login", "SUCESSO", f"Usuário '{username}' autenticado com sucesso.")
            return True

    attempts = st.session_state.get("login_attempts", 0) + 1
    st.session_state["login_attempts"] = attempts
    if attempts >= MAX_LOGIN_ATTEMPTS:
        st.session_state["login_locked_until"] = time.time() + 60
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
    records.insert(0, {
        "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "usuario":   st.session_state.get("user_name", ""),
        "modulo":    module_name,
        "status":    status,
        "detalhe":   detail,
        "saida":     output_name,
    })
    st.session_state["portal_history"] = records[:300]
    save_history(st.session_state["portal_history"])


# =========================================================
# CSS — Design Minimalista Profissional Aprimorado
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
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
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
            transition: all 0.15s ease;
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
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
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
            animation: pulse 2s infinite;
        }}
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.5; }}
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

        /* ── Hero / Dashboard ── */
        .np-hero {{
            background: linear-gradient(135deg, var(--c-sidebar) 0%, var(--c-primary) 100%);
            border-radius: 18px;
            padding: 32px 28px;
            margin-bottom: 28px;
            position: relative;
            overflow: hidden;
            color: #FFFFFF;
            min-height: 200px;
        }}
        .np-hero::before {{
            content: '';
            position: absolute;
            top: -50%;
            right: -10%;
            width: 400px;
            height: 400px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 70%);
            pointer-events: none;
        }}
        .np-hero-content {{
            position: relative; z-index: 1;
            max-width: 60%;
        }}
        .np-hero-eyebrow {{
            font-size: 12px; font-weight: 600;
            letter-spacing: 0.10em; text-transform: uppercase;
            color: rgba(255,255,255,0.55) !important;
            margin-bottom: 12px;
        }}
        .np-hero-title {{
            color: #FFFFFF !important;
            font-size: 32px;
            line-height: 1.15;
            font-weight: 700;
            letter-spacing: -0.02em;
            margin: 0 0 8px 0;
        }}
        .np-hero-text {{
            color: rgba(255,255,255,0.72) !important;
            font-size: 14px;
            line-height: 1.6;
        }}

        /* ── Módulos Grid ── */
        .np-modules-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 16px;
            margin-bottom: 28px;
        }}
        .np-module-card {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 14px;
            padding: 20px;
            min-height: 200px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: all 0.2s ease;
            cursor: pointer;
        }}
        .np-module-card:hover {{
            border-color: var(--c-accent);
            box-shadow: 0 4px 12px rgba(19,64,160,0.08);
            transform: translateY(-2px);
        }}
        .np-module-icon {{
            width: 48px; height: 48px; border-radius: 12px;
            background: linear-gradient(135deg, var(--c-primary) 0%, var(--c-accent) 100%);
            color: #FFFFFF;
            display: flex; align-items: center; justify-content: center;
            font-size: 24px; margin-bottom: 14px;
        }}
        .np-card-title {{
            font-size: 15px; font-weight: 700;
            color: var(--c-text); margin-bottom: 6px;
            letter-spacing: -0.01em;
        }}
        .np-card-text {{
            font-size: 13px; line-height: 1.6;
            color: var(--c-muted);
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

        /* ── Cards genéricos ── */
        .np-card {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 14px;
            padding: 20px;
            margin-bottom: 12px;
        }}

        /* ── Botões ── */
        .stButton > button, .stDownloadButton > button {{
            border-radius: 10px !important;
            min-height: 42px !important;
            font-weight: 600 !important;
            font-size: 14px !important;
            letter-spacing: -0.01em;
            transition: all 0.15s !important;
        }}
        .stButton > button[kind="primary"],
        .stDownloadButton > button {{
            background: var(--c-primary) !important;
            color: #FFFFFF !important;
            border: 0 !important;
            box-shadow: 0 2px 8px rgba(19,64,160,0.2) !important;
        }}
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button:hover {{
            background: var(--c-accent-d) !important;
            box-shadow: 0 4px 12px rgba(19,64,160,0.3) !important;
        }}

        /* ── Login ── */
        .np-login-bg {{
            min-height: 100vh;
            background: linear-gradient(135deg, #07142E 0%, #1340A0 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 2rem 1rem;
        }}
        .np-login-card {{
            background: var(--c-surface);
            border: 1px solid var(--c-border);
            border-radius: 18px;
            padding: 40px 36px 32px;
            max-width: 420px;
            width: 100%;
            box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        }}
        .np-login-logo {{
            width: 120px;
            height: auto;
            max-height: 52px;
            object-fit: contain;
            display: block;
            margin: 0 auto 24px;
        }}
        .np-login-logo-fb {{
            width: 48px; height: 48px;
            background: var(--c-primary);
            border-radius: 12px;
            display: flex; align-items: center; justify-content: center;
            color: #fff; font-weight: 700; font-size: 18px;
            margin: 0 auto 24px;
        }}
        .np-login-title {{
            font-size: 24px; font-weight: 700;
            color: var(--c-text); letter-spacing: -0.02em;
            margin: 0 0 8px;
            text-align: center;
        }}
        .np-login-sub {{
            font-size: 13px; color: var(--c-muted); line-height: 1.5;
            margin-bottom: 28px;
            text-align: center;
        }}
        .np-login-footer {{
            font-size: 11px; color: var(--c-muted);
            text-align: center; margin-top: 20px;
            padding-top: 16px;
            border-top: 1px solid var(--c-border);
        }}

        /* ── Responsivo ── */
        @media (max-width: 700px) {{
            .np-hero {{ padding: 24px 20px; }}
            .np-hero-title {{ font-size: 24px; }}
            .np-hero-content {{ max-width: 100%; }}
            .np-modules-grid {{ grid-template-columns: 1fr; }}
            .np-login-card {{ padding: 32px 24px 24px; }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    return logo_html


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
                    <div class="np-top-subtitle">Portal integrado · <span id="np-clock">--/--/---- --:--</span></div>
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
            if st.button(f"{icon} {label}", key=f"nav_{page_key}", use_container_width=True):
                st.session_state["current_page"] = page_key
                st.rerun()

        st.divider()

        if st.button("🚪 Sair", key="logout_btn", use_container_width=True):
            st.session_state["authenticated"] = False
            st.session_state["user_name"] = ""
            add_history("Logout", "SUCESSO", "Usuário desconectado.")
            st.rerun()


def render_login_page(logo_html: str) -> None:
    st.markdown(
        f"""
        <div class="np-login-bg">
            <div class="np-login-card">
                {logo_html if logo_html else '<div class="np-login-logo-fb">EN</div>'}
                <h1 class="np-login-title">Central Operacional</h1>
                <p class="np-login-sub">Análises & Gestão<br>Insira suas credenciais para continuar</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        pass
    with col2:
        st.markdown("<div style='min-height: 200px;'></div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Acesso ao Sistema")
        chapa = st.text_input("CHAPA", placeholder="Ex: 001", key="login_chapa")
        senha = st.text_input("SENHA", type="password", placeholder="Sua senha", key="login_senha")

        if st.button("🔓 Entrar", use_container_width=True, type="primary"):
            if not chapa or not senha:
                st.error("Preencha todos os campos.")
            elif _is_login_locked():
                remaining = int(st.session_state.get("login_locked_until", 0) - __import__("time").time())
                st.error(f"Muitas tentativas incorretas. Aguarde {remaining}s antes de tentar novamente.")
            elif authenticate(chapa, senha):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Chapa ou senha inválida.")

        st.markdown(
            """
            <div class="np-login-footer">
                Central Operacional de Análises · v1.0<br>
                © 2026 Expresso Nepomuceno
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_home_page() -> None:
    st.markdown(
        f"""
        <div class="np-hero">
            <div class="np-hero-content">
                <div class="np-hero-eyebrow">Bem-vindo</div>
                <h1 class="np-hero-title">Operador</h1>
                <p class="np-hero-text">Acesse os módulos de análise operacional para gerenciar dados de frotas, viagens e tempos de carregamento.</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Módulos")

    modules = [
        ("◴", "Odômetro", "Vínculo de veículos", "odometro"),
        ("◔", "Carregamento", "Tempo de operação", "tempo"),
        ("▦", "Viagens", "Viagens em bloco", "viagens"),
        ("☰", "Histórico", "Registro de ações", "historico"),
        ("▥", "Relatórios", "Análises e dados", "relatorios"),
        ("⚙", "Configurações", "Preferências", "configuracoes"),
    ]

    cols = st.columns(3)
    for idx, (icon, title, desc, page_key) in enumerate(modules):
        with cols[idx % 3]:
            if st.button(
                f"""
                <div class="np-module-card">
                    <div>
                        <div class="np-module-icon">{icon}</div>
                        <div class="np-card-title">{title}</div>
                        <div class="np-card-text">{desc}</div>
                    </div>
                </div>
                """,
                key=f"module_{page_key}",
                use_container_width=True,
            ):
                st.session_state["current_page"] = page_key
                st.rerun()


def render_odometro_page() -> None:
    st.markdown("### Odômetro / Vínculo")
    st.markdown("Cruzar combustível, maxtrack, ativos e produção")

    col1, col2 = st.columns(2)
    with col1:
        st.number_input("Leitura Odômetro", min_value=0, step=1)
    with col2:
        st.text_input("Placa do Veículo")

    if st.button("Registrar"):
        add_history("Odômetro", "SUCESSO", "Leitura de odômetro registrada.")
        st.success("Registrado com sucesso!")


def render_tempo_page() -> None:
    st.markdown("### Tempo de Carregamento")
    st.markdown("Tratar permanência por área e consolidar eventos")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("▶ Iniciar"):
            add_history("Tempo", "INFO", "Cronômetro iniciado.")
    with col2:
        if st.button("⏸ Pausar"):
            add_history("Tempo", "INFO", "Cronômetro pausado.")
    with col3:
        if st.button("⏹ Parar"):
            add_history("Tempo", "SUCESSO", "Tempo registrado.")
            st.success("Tempo registrado!")


def render_viagens_page() -> None:
    st.markdown("### Viagens em Bloco")
    st.markdown("Integrar Maxtrack, SAP e permanência")

    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Origem")
    with col2:
        st.text_input("Destino")

    if st.button("Registrar Viagem"):
        add_history("Viagens", "SUCESSO", "Viagem registrada.")
        st.success("Viagem registrada com sucesso!")


def render_historico_page() -> None:
    st.markdown("### Histórico")
    st.markdown("Registro de ações do sistema")

    history = st.session_state.get("portal_history", [])

    if not history:
        st.info("Nenhum registro no histórico.")
    else:
        for record in history[:50]:
            status_color = {
                "SUCESSO": "🟢",
                "FALHA": "🔴",
                "BLOQUEIO": "🟠",
                "INFO": "🔵",
            }.get(record.get("status", "INFO"), "⚪")

            st.markdown(
                f"""
                **{status_color} {record.get('modulo', 'Sistema')}** — {record.get('data_hora', '')}
                
                {record.get('detalhe', '')}
                """
            )


def render_relatorios_page() -> None:
    st.markdown("### Relatórios")
    st.markdown("Análises e dados operacionais")

    history = st.session_state.get("portal_history", [])

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de Registros", len(history))
    with col2:
        sucesso = len([h for h in history if h.get("status") == "SUCESSO"])
        st.metric("Sucessos", sucesso)
    with col3:
        falhas = len([h for h in history if h.get("status") == "FALHA"])
        st.metric("Falhas", falhas)


def render_configuracoes_page() -> None:
    st.markdown("### Configurações")
    st.markdown("Preferências e informações do sistema")

    st.markdown("**Perfil do Usuário**")
    st.text(f"Usuário: {st.session_state.get('user_name', 'N/A')}")
    st.text(f"Tipo: Operador")

    st.divider()

    st.markdown("**Sobre o Sistema**")
    st.text("Central Operacional de Análises v1.0")
    st.text("© 2026 Expresso Nepomuceno")


# =========================================================
# MAIN
# =========================================================

def main() -> None:
    init_state()

    if not st.session_state["authenticated"]:
        logo_html = apply_css()
        render_login_page(logo_html)
    else:
        logo_html = apply_css()
        render_sidebar(logo_html)
        render_topbar()

        page = st.session_state["current_page"]

        if page == "inicio":
            render_home_page()
        elif page == "odometro":
            render_odometro_page()
        elif page == "tempo":
            render_tempo_page()
        elif page == "viagens":
            render_viagens_page()
        elif page == "historico":
            render_historico_page()
        elif page == "relatorios":
            render_relatorios_page()
        elif page == "configuracoes":
            render_configuracoes_page()


if __name__ == "__main__":
    main()
