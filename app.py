import base64
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

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# IDENTIDADE VISUAL
# =========================================================
COLORS = {
    "sidebar": "#071A4A",
    "sidebar_2": "#0A235D",
    "primary": "#143B8F",
    "primary_2": "#4A40D9",
    "accent": "#4C91FF",
    "bg": "#F4F7FC",
    "card": "#FFFFFF",
    "border": "#DCE5F2",
    "text": "#102554",
    "muted": "#667A99",
    "success": "#16A34A",
    "warning": "#D97706",
    "danger": "#DC2626",
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
LOGO_B64 = image_to_base64(LOGO_PATH)
LOGO_MIME = mimetypes.guess_type(str(LOGO_PATH))[0] if LOGO_PATH else "image/png"
if LOGO_PATH and LOGO_PATH.suffix.lower() in {".jpg", ".jpeg"}:
    LOGO_MIME = "image/jpeg"
LOGO_MIME = LOGO_MIME or "image/png"


# =========================================================
# ESTADO
# =========================================================
PAGES = {
    "inicio": "🏠 Início",
    "odometro": "◴ Odômetro / Vínculo",
    "tempo": "◔ Tempo de Carregamento",
    "viagens": "▦ Viagens em Bloco",
    "historico": "☰ Histórico",
    "relatorios": "▥ Relatórios",
    "configuracoes": "⚙ Configurações",
}


@dataclass
class DiagnosticItem:
    nome: str
    arquivo: str
    objetivo: str
    status: str
    detalhe: str


def init_state() -> None:
    defaults = {
        "authenticated": False,
        "user_name": "",
        "current_page": "inicio",
        "portal_history": load_history(),
        "analysis_report": [],
        "last_error": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# =========================================================
# LOGIN / HISTÓRICO
# =========================================================

def load_login_table() -> pd.DataFrame:
    if LOGIN_FILE.exists():
        try:
            df = pd.read_excel(LOGIN_FILE)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            pass
    return pd.DataFrame({"Chapa": ["232473"], "Senha": ["232473"]})


USERS_DF = load_login_table()


def authenticate(username: str, password: str) -> bool:
    if USERS_DF.empty:
        return False

    cols = {normalize_name(c): c for c in USERS_DF.columns}
    user_col = cols.get("chapa") or list(USERS_DF.columns)[0]
    pass_col = cols.get("senha") or (list(USERS_DF.columns)[1] if len(USERS_DF.columns) > 1 else user_col)

    username = str(username).strip()
    password = str(password).strip()

    for _, row in USERS_DF.iterrows():
        if str(row.get(user_col, "")).strip() == username and str(row.get(pass_col, "")).strip() == password:
            st.session_state["user_name"] = username
            return True
    return False



def save_history(records: list[dict[str, Any]]) -> None:
    try:
        HISTORY_FILE.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
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
            "usuario": st.session_state.get("user_name", ""),
            "modulo": module_name,
            "status": status,
            "detalhe": detail,
            "saida": output_name,
        },
    )
    st.session_state["portal_history"] = records[:300]
    save_history(st.session_state["portal_history"])


# =========================================================
# CSS / LAYOUT
# =========================================================

def apply_css() -> None:
    logo_html = (
        f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" class="np-logo" alt="Expresso Nepomuceno">'
        if LOGO_B64
        else '<div class="np-logo-fallback">EN</div>'
    )
    st.markdown(
        f"""
        <style>
        :root {{
            --np-sidebar: {COLORS['sidebar']};
            --np-sidebar-2: {COLORS['sidebar_2']};
            --np-primary: {COLORS['primary']};
            --np-primary-2: {COLORS['primary_2']};
            --np-accent: {COLORS['accent']};
            --np-bg: {COLORS['bg']};
            --np-card: {COLORS['card']};
            --np-border: {COLORS['border']};
            --np-text: {COLORS['text']};
            --np-muted: {COLORS['muted']};
            --np-success: {COLORS['success']};
            --np-warning: {COLORS['warning']};
            --np-danger: {COLORS['danger']};
        }}

        html, body, [class*="css"], .stApp {{
            font-family: "Segoe UI", Arial, sans-serif;
        }}

        .stApp {{
            background: linear-gradient(180deg, #F7F9FD 0%, #F1F5FB 100%);
            color: var(--np-text);
        }}

        [data-testid="stHeader"] {{
            background: transparent;
        }}
        #MainMenu, footer {{visibility:hidden;}}

        .block-container {{
            max-width: 1380px;
            padding-top: 1rem;
            padding-bottom: 3rem;
        }}

        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, var(--np-sidebar) 0%, #081E56 100%);
            border-right: 1px solid rgba(255,255,255,0.08);
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] .stButton > button {{
            width: 100%;
            justify-content: flex-start;
            min-height: 52px;
            border-radius: 16px !important;
            border: 1px solid rgba(255,255,255,0.14) !important;
            background: rgba(255,255,255,0.08) !important;
            color: #FFFFFF !important;
            font-size: 15px !important;
            font-weight: 700 !important;
            box-shadow: none !important;
        }}
        section[data-testid="stSidebar"] .stButton > button:hover {{
            border-color: rgba(148, 182, 255, .65) !important;
            background: rgba(255,255,255,0.12) !important;
        }}
        section[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, #4C91FF 0%, #5B4FEA 100%) !important;
            border: 1px solid rgba(255,255,255,0.3) !important;
            color: #FFFFFF !important;
        }}

        .np-sidebar-brand {{
            padding: 6px 4px 16px 4px;
        }}
        .np-logo {{
            width: 120px;
            max-width: 100%;
            display: block;
            object-fit: contain;
            margin-bottom: 16px;
        }}
        .np-logo-fallback {{
            width: 84px;
            height: 84px;
            border-radius: 18px;
            background: linear-gradient(135deg, #123B8E 0%, #4B40D8 100%);
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 34px;
            font-weight: 900;
            margin-bottom: 16px;
        }}
        .np-sidebar-title {{
            font-size: 18px;
            font-weight: 800;
            margin: 0;
        }}
        .np-sidebar-subtitle {{
            font-size: 13px;
            opacity: .88;
            margin-top: 4px;
            margin-bottom: 18px;
        }}
        .np-sidebar-section {{
            font-size: 12px;
            letter-spacing: .16em;
            text-transform: uppercase;
            font-weight: 800;
            margin: 10px 6px 10px;
            opacity: .92;
        }}
        .np-side-foot {{
            margin-top: 22px;
            padding: 8px 6px 0;
            font-size: 12px;
            color: rgba(255,255,255,0.86) !important;
        }}

        .np-topbar {{
            background: #FFFFFF;
            border: 1px solid var(--np-border);
            border-radius: 22px;
            padding: 16px 22px;
            box-shadow: 0 16px 40px rgba(15, 23, 42, 0.06);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 18px;
            margin-bottom: 20px;
        }}
        .np-topbar-left {{display:flex; align-items:center; gap:14px;}}
        .np-brand-circle {{
            width: 54px; height:54px; border-radius: 999px;
            background: linear-gradient(135deg, #0A2D6F 0%, #2748C9 100%);
            color: #FFFFFF;
            display:flex; align-items:center; justify-content:center;
            font-weight: 900; font-size: 24px;
        }}
        .np-top-title {{font-size: 18px; font-weight: 800; color: var(--np-text);}}
        .np-top-subtitle {{font-size: 13px; color: var(--np-muted); margin-top: 2px;}}
        .np-topbar-right {{display:flex; align-items:center; gap: 12px;}}
        .np-status-pill {{
            display:inline-flex; align-items:center; gap:8px;
            padding: 10px 14px; border-radius: 999px;
            background: linear-gradient(135deg, #2E2C8E 0%, #4B35C0 100%);
            color:#FFFFFF !important; font-size: 13px; font-weight: 800;
        }}
        .np-dot {{ width:10px; height:10px; border-radius:50%; background:#22C55E; box-shadow: 0 0 0 4px rgba(34,197,94,.18); }}
        .np-user-chip {{
            width: 44px; height:44px; border-radius: 999px; background:#E8F0FF;
            color:#16387E; display:flex; align-items:center; justify-content:center;
            font-weight: 900;
        }}
        .np-user-label {{font-size: 13px; color: var(--np-text); font-weight: 700;}}
        .np-user-email {{font-size: 12px; color: var(--np-muted);}}

        .np-hero {{
            position: relative;
            overflow: hidden;
            background: linear-gradient(135deg, #072B6B 0%, #0D4FAE 42%, #174A91 100%);
            border-radius: 28px;
            min-height: 330px;
            padding: 42px 42px 36px;
            border: 1px solid rgba(255,255,255,0.12);
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.12);
            margin-bottom: 22px;
        }}
        .np-hero::before {{
            content: "";
            position: absolute;
            inset: 0;
            background:
                radial-gradient(circle at 15% 30%, rgba(255,255,255,0.12), transparent 18%),
                linear-gradient(147deg, transparent 0 55%, rgba(255,255,255,0.08) 55% 58%, transparent 58%),
                linear-gradient(160deg, transparent 0 70%, rgba(255,255,255,0.18) 70% 71%, transparent 71% 73%, rgba(255,255,255,0.12) 73% 74%, transparent 74%),
                linear-gradient(170deg, transparent 0 75%, rgba(255,255,255,0.1) 75% 76%, transparent 76% 77%, rgba(255,255,255,0.08) 77% 78%, transparent 78%);
            opacity: .9;
        }}
        .np-hero-content {{position:relative; z-index:1; max-width: 56%;}}
        .np-hero-eyebrow {{color:#CFE1FF !important; font-size: 14px; font-weight: 700; margin-bottom: 14px;}}
        .np-hero-title {{color:#FFFFFF !important; font-size: 58px; line-height: 1.05; font-weight: 900; margin:0 0 16px 0;}}
        .np-hero-text {{color: rgba(255,255,255,0.92) !important; font-size: 18px; max-width: 580px; line-height: 1.6;}}
        .np-hero-brand {{position:absolute; right:56px; top:58px; z-index:1; text-align:right;}}
        .np-hero-brand img {{width: 290px; max-width: 27vw; filter: drop-shadow(0 10px 30px rgba(0,0,0,.18));}}
        .np-hero-mark {{font-size: 34px; font-weight: 900; color: rgba(255,255,255,0.18) !important; margin-top: 12px;}}
        .np-stat-grid {{display:grid; grid-template-columns: repeat(2, minmax(170px, 1fr)); gap: 14px; max-width: 540px; margin-top: 26px;}}
        .np-stat {{background: rgba(11,33,89,0.34); border: 1px solid rgba(255,255,255,0.16); border-radius: 18px; padding: 16px 18px; color:white !important;}}
        .np-stat-value {{font-size: 20px; font-weight: 900;}}
        .np-stat-label {{font-size: 13px; color: rgba(255,255,255,0.88) !important;}}

        .np-section-title {{font-size: 18px; color: var(--np-text); font-weight: 800; margin: 6px 0 4px;}}
        .np-section-subtitle {{font-size: 14px; color: var(--np-muted); margin-bottom: 18px;}}

        .np-card {{
            background: var(--np-card);
            border: 1px solid var(--np-border);
            border-radius: 22px;
            padding: 22px;
            box-shadow: 0 14px 32px rgba(15, 23, 42, 0.06);
            margin-bottom: 14px;
        }}
        .np-module-card {{min-height: 264px; display:flex; flex-direction:column; justify-content:space-between;}}
        .np-icon {{
            width: 64px; height:64px; border-radius: 18px; background:#EEF4FF;
            color:#2F57D3; display:flex; align-items:center; justify-content:center;
            font-size: 28px; font-weight: 700; margin-bottom: 16px;
        }}
        .np-card-title {{font-size: 18px; color: var(--np-text); font-weight: 800; margin-bottom: 10px;}}
        .np-card-text {{font-size: 14px; line-height: 1.6; color: var(--np-muted); min-height: 68px;}}
        .np-metrics-box {{display:flex; gap: 16px; border:1px solid var(--np-border); border-radius:14px; padding: 12px 14px; margin-top: 16px;}}
        .np-metric-item {{flex:1;}}
        .np-metric-value {{font-size: 18px; font-weight: 800; color: var(--np-text);}}
        .np-metric-label {{font-size: 12px; color: var(--np-muted);}}
        .np-notes-list {{padding-left: 16px; color: var(--np-muted); line-height:1.7;}}

        .stButton > button, .stDownloadButton > button {{
            border-radius: 14px !important;
            min-height: 46px !important;
            font-weight: 800 !important;
        }}
        .stButton > button[kind="primary"], .stDownloadButton > button {{
            background: linear-gradient(135deg, #173D91 0%, #4B40D8 100%) !important;
            color: #FFFFFF !important;
            border: 0 !important;
        }}

        [data-testid="stMetric"] {{
            background: #FFFFFF;
            border: 1px solid var(--np-border);
            border-radius: 18px;
            padding: 12px 16px;
            box-shadow: 0 10px 24px rgba(15,23,42,0.05);
        }}

        [data-testid="stFileUploader"] section {{
            border: 1px dashed #9FB1CC;
            border-radius: 16px;
            background: #FBFDFF;
        }}
        [data-testid="stDataFrame"] {{
            border-radius: 16px;
            overflow: hidden;
            border: 1px solid var(--np-border);
        }}

        .np-login-shell {{
            min-height: calc(100vh - 60px);
            display:flex;
            align-items:center;
            justify-content:center;
            padding: 18px;
        }}
        .np-login-card {{
            background: #FFFFFF;
            border: 1px solid var(--np-border);
            border-radius: 28px;
            padding: 34px;
            width: min(460px, 95vw);
            box-shadow: 0 30px 80px rgba(15,23,42,0.12);
        }}
        .np-login-title {{font-size: 30px; font-weight: 900; color: var(--np-text); margin: 6px 0;}}
        .np-login-sub {{font-size: 14px; color: var(--np-muted); margin-bottom: 20px;}}
        .np-info-card {{
            background:#FFFFFF; border:1px solid var(--np-border); border-radius:18px; padding:16px 18px; margin-bottom:12px;
        }}
        .np-info-title {{font-size:15px; font-weight:800; color:var(--np-text); margin-bottom:4px;}}
        .np-info-text {{font-size:13px; color:var(--np-muted); line-height:1.6;}}

        @media (max-width: 1100px) {{
            .np-hero-content {{max-width: 100%;}}
            .np-hero-brand {{position:relative; right:auto; top:auto; margin-top: 24px; text-align:left;}}
            .np-hero-brand img {{max-width: 260px; width: 100%;}}
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
    def __init__(self, *args, **kwargs):
        pass

    def title(self, *args, **kwargs):
        pass

    def geometry(self, *args, **kwargs):
        pass

    def resizable(self, *args, **kwargs):
        pass

    def configure(self, *args, **kwargs):
        pass

    def mainloop(self, *args, **kwargs):
        pass

    def destroy(self, *args, **kwargs):
        pass

    def update_idletasks(self, *args, **kwargs):
        pass


class DummyStyle:
    def theme_use(self, *args, **kwargs):
        pass

    def configure(self, *args, **kwargs):
        pass


class DummyWidget:
    def __init__(self, *args, **kwargs):
        pass

    def pack(self, *args, **kwargs):
        pass

    def grid(self, *args, **kwargs):
        pass

    def config(self, *args, **kwargs):
        pass

    configure = config



def ensure_fake_tk_modules() -> dict[str, Any]:
    backups: dict[str, Any] = {}
    if "tkinter" not in sys.modules:
        tk_module = types.ModuleType("tkinter")
        tk_module.Tk = DummyTkWindow
        tk_module.BOTH = tk_module.X = tk_module.Y = tk_module.LEFT = tk_module.RIGHT = 0
        tk_module.END = "end"
        backups["tkinter"] = None
        sys.modules["tkinter"] = tk_module
    if "tkinter.filedialog" not in sys.modules:
        filedialog = types.ModuleType("tkinter.filedialog")
        filedialog.askopenfilename = lambda *a, **k: ""
        backups["tkinter.filedialog"] = None
        sys.modules["tkinter.filedialog"] = filedialog
    if "tkinter.messagebox" not in sys.modules:
        messagebox = types.ModuleType("tkinter.messagebox")
        messagebox.showinfo = lambda *a, **k: None
        messagebox.showerror = lambda *a, **k: None
        backups["tkinter.messagebox"] = None
        sys.modules["tkinter.messagebox"] = messagebox
    if "tkinter.ttk" not in sys.modules:
        ttk = types.ModuleType("tkinter.ttk")
        ttk.Style = DummyStyle
        ttk.Frame = DummyWidget
        ttk.Label = DummyWidget
        ttk.Progressbar = DummyWidget
        ttk.Button = DummyWidget
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
    source = path.read_text(encoding="utf-8", errors="replace")
    source = strip_autorun_block(source)
    module = types.ModuleType(f"portal_{path.stem}")
    module.__file__ = str(path)
    module.__dict__["__name__"] = module.__name__
    module.__dict__["__package__"] = None

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

    od_file = find_file(["app_odometro_streamlit_corrigido.py", "app_odometro_streamlit_corrigido (1).py", "Python_Odometro_Vinculo.py"])
    if od_file and od_file.exists():
        try:
            mod = load_module_safely(od_file)
            required = ["processar_streamlit", "preparar_bases", "gerar_resultado", "exportar"]
            missing = [name for name in required if not hasattr(mod, name)]
            if missing:
                items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name, "Cruzar combustível, maxtrack, ativos e produção", "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name, "Cruzar combustível, maxtrack, ativos e produção", "OK", "Módulo de odômetro integrado ao portal com as regras V15 preservadas."))
        except Exception as exc:
            items.append(DiagnosticItem("Odômetro / Vínculo", od_file.name, "Cruzar combustível, maxtrack, ativos e produção", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Odômetro / Vínculo", "app_odometro_streamlit_corrigido.py", "Cruzar combustível, maxtrack, ativos e produção", "Erro", "Arquivo não localizado."))

    tc_file = find_file(["Python_Tempo_Carregamento.py"])
    if tc_file and tc_file.exists():
        try:
            mod = load_module_safely(tc_file)
            required = ["processar_arquivo", "main_streamlit"]
            missing = [name for name in required if not hasattr(mod, name)]
            if missing:
                items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name, "Tratar permanência por área e consolidar eventos", "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name, "Tratar permanência por área e consolidar eventos", "OK", "Estrutura principal encontrada."))
        except Exception as exc:
            items.append(DiagnosticItem("Tempo de Carregamento", tc_file.name, "Tratar permanência por área e consolidar eventos", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Tempo de Carregamento", "Python_Tempo_Carregamento.py", "Tratar permanência por área e consolidar eventos", "Erro", "Arquivo não localizado."))

    vb_file = find_file(["Python_Viagens_Bloco.py"])
    if vb_file and vb_file.exists():
        try:
            mod = load_module_safely(vb_file, use_fake_tk=True)
            required = ["processar_arquivos", "main_streamlit"]
            missing = [name for name in required if not hasattr(mod, name)]
            if missing:
                items.append(DiagnosticItem("Viagens em Bloco", vb_file.name, "Integrar Maxtrack, SAP e permanência", "Atenção", f"Funções ausentes: {', '.join(missing)}"))
            else:
                items.append(DiagnosticItem("Viagens em Bloco", vb_file.name, "Integrar Maxtrack, SAP e permanência", "OK", "Estrutura Streamlit segura encontrada. Interface desktop tkinter permanece preservada apenas para uso local."))
        except Exception as exc:
            items.append(DiagnosticItem("Viagens em Bloco", vb_file.name, "Integrar Maxtrack, SAP e permanência", "Erro", str(exc)))
    else:
        items.append(DiagnosticItem("Viagens em Bloco", "Python_Viagens_Bloco.py", "Integrar Maxtrack, SAP e permanência", "Erro", "Arquivo não localizado."))

    return items


# =========================================================
# COMPONENTES VISUAIS
# =========================================================

def render_topbar() -> None:
    initials = (st.session_state.get("user_name") or "AD")[:2].upper()
    user_label = st.session_state.get("user_name") or "Administrador"
    last_access = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.markdown(
        f"""
        <div class="np-topbar">
            <div class="np-topbar-left">
                <div class="np-brand-circle">EN</div>
                <div>
                    <div class="np-top-title">Central Operacional de Análises</div>
                    <div class="np-top-subtitle">Portal operacional integrado • Último acesso {last_access}</div>
                </div>
            </div>
            <div class="np-topbar-right">
                <div class="np-status-pill"><span class="np-dot"></span> Sistema online</div>
                <div class="np-user-chip">{initials}</div>
                <div>
                    <div class="np-user-label">Administrador</div>
                    <div class="np-user-email">{user_label}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_sidebar(logo_html: str) -> str:
    with st.sidebar:
        st.markdown(
            f"""
            <div class="np-sidebar-brand">
                {logo_html}
                <div class="np-sidebar-title">Central Operacional</div>
                <div class="np-sidebar-subtitle">Expresso Nepomuceno</div>
                <div class="np-sidebar-section">Navegação</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        for page_key, label in PAGES.items():
            kind = "primary" if st.session_state["current_page"] == page_key else "secondary"
            if st.button(label, key=f"nav_{page_key}", type=kind, use_container_width=True):
                st.session_state["current_page"] = page_key
                st.rerun()

        st.markdown(
            f"""
            <div class="np-side-foot">
                <div><strong>Versão</strong> 4.0.0</div>
                <div><strong>Ambiente</strong> interno</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    return st.session_state["current_page"]



def render_module_card(icon: str, title: str, desc: str, stats: list[tuple[str, str]], button_label: str, page_key: str) -> None:
    stats_html = "".join(
        [
            f'<div class="np-metric-item"><div class="np-metric-value">{v}</div><div class="np-metric-label">{l}</div></div>'
            for v, l in stats
        ]
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
# TELAS
# =========================================================

def render_login_screen() -> None:
    st.markdown('<div class="np-login-shell">', unsafe_allow_html=True)
    with st.container():
        st.markdown('<div class="np-login-card">', unsafe_allow_html=True)
        if LOGO_B64:
            st.markdown(
                f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" style="width:140px;display:block;margin-bottom:18px;">',
                unsafe_allow_html=True,
            )
        st.markdown('<div class="np-login-title">Acesso ao Portal</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="np-login-sub">Entre com sua chapa e senha cadastradas na planilha de login.</div>',
            unsafe_allow_html=True,
        )

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Chapa / usuário")
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True, type="primary")

        if submitted:
            if authenticate(username, password):
                st.session_state["authenticated"] = True
                add_history("Login", "OK", "Acesso realizado com sucesso.")
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

        st.caption("Acesso validado pela planilha Login.xlsx.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)



def page_inicio() -> None:
    history = st.session_state.get("portal_history", [])
    total_today = sum(1 for row in history if row.get("data_hora", "").startswith(datetime.now().strftime("%d/%m/%Y")))
    total_ok = sum(1 for row in history if row.get("status") == "OK")
    online = "100%"

    brand = f'<img src="data:{LOGO_MIME};base64,{LOGO_B64}" alt="logo">' if LOGO_B64 else '<div class="np-hero-mark">EN</div>'

    st.markdown(
        f"""
        <div class="np-hero">
            <div class="np-hero-content">
                <div class="np-hero-eyebrow">Bem-vindo à</div>
                <h1 class="np-hero-title">Central Operacional de Análises.</h1>
                <div class="np-hero-text">
                    Gerencie e acompanhe os processos de análise de forma centralizada, segura e eficiente.
                </div>
                <div class="np-stat-grid">
                    <div class="np-stat">
                        <div class="np-stat-value">{total_today}</div>
                        <div class="np-stat-label">Processamentos hoje</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{total_ok}</div>
                        <div class="np-stat-label">Concluídos com sucesso</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{len(history)}</div>
                        <div class="np-stat-label">Histórico registrado</div>
                    </div>
                    <div class="np-stat">
                        <div class="np-stat-value">{online}</div>
                        <div class="np-stat-label">Disponibilidade visual</div>
                    </div>
                </div>
            </div>
            <div class="np-hero-brand">{brand}<div class="np-hero-mark">EXPRESSO NEPOMUCENO</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="np-section-title">Módulos principais</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Selecione abaixo a ferramenta operacional desejada.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        render_module_card("◴", "Odômetro / Vínculo", "Cruza combustível, Maxtrack, ativos e produção para gerar o ODOMETRO_MATCH final.", [("4", "bases"), ("Excel", "saída")], "Acessar módulo →", "odometro")
    with c2:
        render_module_card("◔", "Tempo de Carregamento", "Trata permanência por área, ruídos, eventos próximos e consolida tempos operacionais.", [("1", "base"), ("Resumo", "automático")], "Acessar módulo →", "tempo")
    with c3:
        render_module_card("▦", "Viagens em Bloco", "Integra Maxtrack, SAP e permanência para montar viagens validadas e deduplicadas.", [("3", "bases"), ("Score", "deduplicação")], "Acessar módulo →", "viagens")

    c4, c5 = st.columns(2)
    with c4:
        render_module_card("☰", "Histórico", "Consulte o histórico de processamentos realizados no portal.", [(str(len(history)), "registros"), (datetime.now().strftime("%H:%M"), "agora")], "Ver histórico →", "historico")
    with c5:
        render_module_card("▥", "Relatórios", "Acompanhe indicadores simples de uso, módulos e resultados gerados.", [(str(total_ok), "sucessos"), (str(max(len(history)-total_ok,0)), "outros")], "Ver relatórios →", "relatorios")



def page_odometro() -> None:
    st.markdown('<div class="np-section-title">Odômetro / Vínculo</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Envie as quatro bases obrigatórias para gerar o arquivo consolidado final.</div>', unsafe_allow_html=True)

    module_path = find_file(["app_odometro_streamlit_corrigido.py", "app_odometro_streamlit_corrigido (1).py", "Python_Odometro_Vinculo.py"])
    if not module_path:
        st.error("Arquivo app_odometro_streamlit_corrigido.py não encontrado.")
        return

    col1, col2 = st.columns(2)
    with col1:
        comb = st.file_uploader("1. Base Combustível", type=["xlsx", "xls"], key="odom_comb")
        ativo = st.file_uploader("3. Base Ativo de Veículos", type=["xlsx", "xls"], key="odom_ativo")
    with col2:
        maxtrack = st.file_uploader("2. Base Km Rodado Maxtrack", type=["xlsx", "xls"], key="odom_max")
        producao = st.file_uploader("4. Produção Oficial / Cliente", type=["xlsx", "xls"], key="odom_prod")

    output_name = st.text_input("Nome do arquivo de saída", value=f"resultado_odometro_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    progress = sum(x is not None for x in [comb, maxtrack, ativo, producao]) / 4
    st.progress(progress)
    st.caption(f"{int(progress*4)} de 4 arquivos carregados")

    if st.button("Processar Odômetro / Vínculo", key="btn_processar_odometro", type="primary", use_container_width=True, disabled=not all([comb, maxtrack, ativo, producao])):
        try:
            mod = load_module_safely(module_path)
            with st.spinner("Executando processamento do módulo..."):
                excel_bytes, indicadores, resultado_final = mod.processar_streamlit(comb, maxtrack, ativo, producao, output_name)

            st.success("Processamento concluído com sucesso.")
            add_history("Odômetro / Vínculo", "OK", "Processamento executado com sucesso.", output_name)

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

    module_path = find_file(["Python_Tempo_Carregamento.py"])
    if not module_path:
        st.error("Arquivo Python_Tempo_Carregamento.py não encontrado.")
        return

    arquivo = st.file_uploader("Base de permanência / carregamento", type=["xlsx", "xls"], key="tempo_arquivo")
    output_name = st.text_input("Nome do arquivo de saída", value=f"tempo_carregamento_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    if st.button("Processar Tempo de Carregamento", key="btn_processar_tempo", type="primary", use_container_width=True, disabled=arquivo is None):
        try:
            mod = load_module_safely(module_path)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_in:
                temp_in.write(arquivo.read())
                path_in = temp_in.name

            path_out = str(BASE_DIR / (output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx"))

            with st.spinner("Tratando arquivo..."):
                saida, tratado, resumo_placa_local, resumo_local, info = mod.processar_arquivo(path_in, path_out)

            st.success("Processamento concluído com sucesso.")
            add_history("Tempo de Carregamento", "OK", "Processamento executado com sucesso.", Path(saida).name)

            c1, c2, c3 = st.columns(3)
            c1.metric("Linhas tratadas", len(tratado))
            c2.metric("Resumo placa/local", len(resumo_placa_local))
            c3.metric("Resumo local", len(resumo_local))

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
        self.bar = st.progress(0)
        self.status = st.empty()
        self.total = 100

    def set_total(self, total: int):
        self.total = total or 100

    def update(self, value: int, etapa: str = "", detalhe: str = ""):
        pct = max(0, min(int(value), 100))
        self.bar.progress(pct / 100)
        msg = f"{pct}% - {etapa}".strip()
        if detalhe:
            msg += f" | {detalhe}"
        self.status.info(msg)



def page_viagens() -> None:
    st.markdown('<div class="np-section-title">Viagens em Bloco</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Envie os três arquivos para integração, validação SAP e deduplicação.</div>', unsafe_allow_html=True)

    module_path = find_file(["Python_Viagens_Bloco.py"])
    if not module_path:
        st.error("Arquivo Python_Viagens_Bloco.py não encontrado.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        arq_max = st.file_uploader("1. Arquivo Maxtrack", type=["xlsx", "xls"], key="vb_max")
    with c2:
        arq_sap = st.file_uploader("2. Arquivo SAP", type=["xlsx", "xls"], key="vb_sap")
    with c3:
        arq_perm = st.file_uploader("3. Arquivo Permanência", type=["xlsx", "xls"], key="vb_perm")

    if st.button("Processar Viagens em Bloco", key="btn_processar_viagens", type="primary", use_container_width=True, disabled=not all([arq_max, arq_sap, arq_perm])):
        try:
            mod = load_module_safely(module_path, use_fake_tk=True)

            def save_uploaded(uploaded_file) -> str:
                suffix = Path(uploaded_file.name).suffix or ".xlsx"
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(uploaded_file.read())
                    return tmp.name

            p_max = save_uploaded(arq_max)
            p_sap = save_uploaded(arq_sap)
            p_perm = save_uploaded(arq_perm)

            progress = StreamlitProgressAdapter()
            with st.spinner("Processando viagens..."):
                saida, total_viagens, total_placas, total_invalidas, total_duplicadas, total_antes = mod.processar_arquivos(p_max, p_sap, p_perm, progress)

            st.success("Processamento concluído com sucesso.")
            add_history("Viagens em Bloco", "OK", "Processamento executado com sucesso.", Path(saida).name)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Antes da dedup", total_antes)
            c2.metric("Viagens finais", total_viagens)
            c3.metric("Válidas", total_viagens - total_invalidas)
            c4.metric("Inválidas", total_invalidas)
            c5.metric("Placas", total_placas)
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
    st.markdown('<div class="np-section-subtitle">Registro dos processamentos e eventos do portal.</div>', unsafe_allow_html=True)

    records = st.session_state.get("portal_history", [])
    if not records:
        st.info("Nenhum registro encontrado até o momento.")
        return

    df = pd.DataFrame(records)
    st.dataframe(df, use_container_width=True, hide_index=True)



def page_relatorios() -> None:
    st.markdown('<div class="np-section-title">Relatórios</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Indicadores simples de uso do portal e da execução dos módulos.</div>', unsafe_allow_html=True)

    records = pd.DataFrame(st.session_state.get("portal_history", []))
    if records.empty:
        st.info("Sem dados suficientes para montar relatórios.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de execuções", len(records))
    c2.metric("Sucessos", int((records["status"] == "OK").sum()))
    c3.metric("Erros", int((records["status"] == "ERRO").sum()))

    if "modulo" in records.columns:
        resumo = records.groupby(["modulo", "status"], dropna=False).size().reset_index(name="quantidade")
        st.markdown("### Resumo por módulo")
        st.dataframe(resumo, use_container_width=True, hide_index=True)



def page_configuracoes() -> None:
    st.markdown('<div class="np-section-title">Configurações</div>', unsafe_allow_html=True)
    st.markdown('<div class="np-section-subtitle">Diagnóstico dos arquivos carregados, autenticação e manutenção do portal.</div>', unsafe_allow_html=True)

    diagnostics = get_module_diagnostics()
    st.session_state["analysis_report"] = diagnostics

    st.markdown("### Diagnóstico dos blocos do arquivo")
    diag_df = pd.DataFrame([
        {
            "Módulo": item.nome,
            "Arquivo": item.arquivo,
            "Objetivo": item.objetivo,
            "Status": item.status,
            "Detalhe": item.detalhe,
        }
        for item in diagnostics
    ])
    st.dataframe(diag_df, use_container_width=True, hide_index=True)

    st.markdown("### Observações encontradas")
    st.markdown(
        """
        <div class="np-info-card">
            <div class="np-info-title">Principais pontos identificados antes da correção</div>
            <div class="np-info-text">
                <ul class="np-notes-list">
                    <li><strong>Viagens em Bloco:</strong> o arquivo original usa <code>tkinter</code> para interface desktop. Em ambiente Streamlit isso tende a falhar ou travar. O portal agora usa um adaptador seguro sem mexer na lógica de processamento.</li>
                    <li><strong>Odômetro / Vínculo:</strong> o arquivo possui auto-renderização quando detecta Streamlit. Isso pode duplicar interface ao ser importado. O portal agora carrega o módulo de forma controlada e chama apenas a função de processamento.</li>
                    <li><strong>Tempo de Carregamento:</strong> a lógica está adequada para integração. O portal chama diretamente a função <code>processar_arquivo</code>.</li>
                    <li><strong>Layout anterior:</strong> havia conflitos visuais, HTML aparecendo na tela e componentes com risco de duplicidade. O layout foi refeito mantendo a navegação e a funcionalidade operacional.</li>
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

    if st.button("Sair do portal", key="btn_sair_portal", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["current_page"] = "inicio"
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

    if current_page == "inicio":
        page_inicio()
    elif current_page == "odometro":
        page_odometro()
    elif current_page == "tempo":
        page_tempo()
    elif current_page == "viagens":
        page_viagens()
    elif current_page == "historico":
        page_historico()
    elif current_page == "relatorios":
        page_relatorios()
    elif current_page == "configuracoes":
        page_configuracoes()
    else:
        st.error("Página não encontrada.")


if __name__ == "__main__":
    main()
