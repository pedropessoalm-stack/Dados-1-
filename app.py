from __future__ import annotations

import ast
import base64
import html
import json
import mimetypes
import os
import re
import sys
import tempfile
import traceback
import types
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

# ============================================================
# CENTRAL OPERACIONAL DE ANALISES - EXPRESSO NEPOMUCENO
# Versao: 3.0.0
# Objetivo: portal fixo, controlado e sem reconhecimento automatico instavel.
# Modulos integrados esperados na mesma pasta do app.py:
# - Python_Odometro_Vinculo.py
# - Python_Tempo_Carregamento.py
# - Python_Viagens_Bloco.py
# ============================================================

APP_VERSION = "3.0.0"
BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / "historico_processamentos.json"
LOG_FILE = BASE_DIR / "portal_logs.jsonl"

MODULE_FILES = {
    "odometro": "Python_Odometro_Vinculo.py",
    "tempo": "Python_Tempo_Carregamento.py",
    "viagens": "Python_Viagens_Bloco.py",
}

PAGES = {
    "inicio": "Início",
    "odometro": "Odômetro / Vínculo",
    "tempo": "Tempo de Carregamento",
    "viagens": "Viagens em Bloco",
    "historico": "Histórico",
    "relatorios": "Relatórios",
    "configuracoes": "Configurações",
}

st.set_page_config(
    page_title="Central Operacional de Análises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# UTILITARIOS
# ============================================================

def now_str() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def now_file() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def log_event(evento: str, detalhes: dict[str, Any] | None = None) -> None:
    payload = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "evento": evento,
        "detalhes": detalhes or {},
    }
    try:
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except Exception:
        pass


def read_history() -> list[dict[str, Any]]:
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def write_history(items: list[dict[str, Any]]) -> None:
    try:
        HISTORY_FILE.write_text(json.dumps(items[-200:], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        log_event("erro_salvar_historico", {"erro": str(exc)})


def add_history(modulo: str, status: str, arquivo_saida: str = "", detalhes: dict[str, Any] | None = None) -> None:
    items = read_history()
    items.append({
        "data_hora": now_str(),
        "modulo": modulo,
        "status": status,
        "arquivo_saida": arquivo_saida,
        "detalhes": detalhes or {},
    })
    write_history(items)


def safe_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name.strip())
    return name or f"arquivo_{now_file()}.xlsx"


def save_upload(uploaded_file, prefix: str) -> str:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=prefix)
    tmp.write(uploaded_file.getbuffer())
    tmp.flush()
    tmp.close()
    return tmp.name


def cleanup_files(paths: list[str]) -> None:
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.unlink(path)
        except Exception:
            pass


def logo_candidates() -> list[Path]:
    names = [
        "logo_nepomuceno.png.jpeg",
        "logo_nepomuceno.jpeg",
        "logo_nepomuceno.jpg",
        "logo_nepomuceno.png",
        "Expresso Nepomuceno.png",
        "Expresso Nepomuceno.jpg",
        "Expresso Nepomuceno.jpeg",
        "Logo Nepomuceno.png",
        "Logo Nepomuceno.jpg",
        "Logo Nepomuceno.jpeg",
    ]
    found = []
    for name in names:
        p = BASE_DIR / name
        if p.exists():
            found.append(p)
    return found


def image_data_uri(path: Path | None) -> str:
    if path is None or not path.exists():
        return ""
    try:
        mime = mimetypes.guess_type(str(path))[0] or "image/png"
        data = base64.b64encode(path.read_bytes()).decode("utf-8")
        return f"data:{mime};base64,{data}"
    except Exception:
        return ""


def logo_html(compact: bool = False) -> str:
    path = logo_candidates()[0] if logo_candidates() else None
    uri = image_data_uri(path)
    if uri:
        cls = "brand-logo compact" if compact else "brand-logo"
        return f'<img class="{cls}" src="{uri}" alt="Expresso Nepomuceno">'
    return """
    <div class="logo-fallback">
        <span class="plane">▸</span>
        <span><b>EXPRESSO</b><strong>NEPOMUCENO</strong></span>
    </div>
    """


class RemoveAutoRun(ast.NodeTransformer):
    """Remove blocos finais que executam o script ao importar o modulo."""

    def visit_Module(self, node: ast.Module) -> ast.Module:
        new_body = []
        for child in node.body:
            if isinstance(child, ast.If):
                test = ast.unparse(child.test) if hasattr(ast, "unparse") else ""
                if "__name__" in test or "_esta_rodando_em_streamlit" in test:
                    continue
            new_body.append(child)
        node.body = new_body
        return node


@st.cache_resource(show_spinner=False)
def load_module_safely(filename: str) -> types.ModuleType:
    path = BASE_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {filename}")

    source = path.read_text(encoding="utf-8", errors="replace")
    tree = ast.parse(source, filename=str(path))
    tree = RemoveAutoRun().visit(tree)
    ast.fix_missing_locations(tree)

    module_name = f"portal_mod_{path.stem}_{abs(hash(str(path))) % 1000000}"
    module = types.ModuleType(module_name)
    module.__file__ = str(path)
    module.__name__ = module_name
    module.__package__ = ""

    sys.modules[module_name] = module
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = lambda *args, **kwargs: None
        exec(compile(tree, str(path), "exec"), module.__dict__)
    finally:
        st.set_page_config = original_set_page_config
    return module


class StreamlitProgressAdapter:
    def __init__(self):
        self.total = 100
        self.progress = st.progress(0, text="0% - Aguardando início")
        self.status = st.empty()
        self.detail = st.empty()

    def set_total(self, total):
        self.total = max(int(total), 1)

    def update(self, step=None, status=None, detail=None):
        value = 0 if step is None else max(0, min(int(step), self.total))
        pct = int((value / self.total) * 100)
        text = status or "Processando"
        self.progress.progress(pct / 100, text=f"{pct}% - {text}")
        if status:
            self.status.info(status)
        if detail:
            self.detail.caption(str(detail))


# ============================================================
# CSS - VISUAL APROVADO COM MELHOR CONTRASTE
# ============================================================

def apply_css() -> None:
    st.markdown(
        """
<style>
:root {
  --en-navy: #0B1F4D;
  --en-navy-2: #112A5E;
  --en-blue: #2447C7;
  --en-blue-2: #3D63F2;
  --en-light: #F4F8FF;
  --en-card: #FFFFFF;
  --en-border: #D9E2F2;
  --en-text: #071849;
  --en-muted: #62708A;
  --en-success: #17A15F;
}

html, body, [class*="css"] {
  font-family: "Inter", "Segoe UI", Arial, sans-serif !important;
}

.stApp {
  background:
    radial-gradient(circle at top left, rgba(61,99,242,.13), transparent 28%),
    linear-gradient(180deg, #F8FBFF 0%, #EEF4FC 100%) !important;
  color: var(--en-text);
}

.block-container {
  max-width: 1180px;
  padding-top: 1.1rem;
  padding-bottom: 3rem;
}

#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] {
  visibility: visible !important;
  background: transparent !important;
}

section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #102A61 0%, #0B1F4D 100%) !important;
  border-right: 1px solid rgba(255,255,255,.13);
}
section[data-testid="stSidebar"] > div { padding: 1.25rem 1.05rem; }
section[data-testid="stSidebar"] * { color: #F4F8FF !important; }

.sidebar-brand-box {
  margin: 1.3rem 0 1rem 0;
}
.brand-logo {
  width: 150px;
  height: auto;
  object-fit: contain;
  display: block;
}
.brand-logo.compact { width: 112px; }
.logo-fallback {
  color: #FFF;
  line-height: .95;
  letter-spacing: -.02em;
  display: flex;
  align-items: center;
  gap: 6px;
  font-weight: 800;
}
.logo-fallback span:last-child { display: grid; }
.logo-fallback b { font-size: 12px; }
.logo-fallback strong { font-size: 18px; }
.logo-fallback .plane { color: #179DEB; font-size: 20px; transform: scaleX(1.7); }
.sidebar-title { font-size: 16px; font-weight: 900; margin-top: .8rem; }
.sidebar-subtitle { font-size: 12px; opacity: .92; margin-bottom: 1.1rem; }
.sidebar-section { font-size: 11px; letter-spacing: .16em; font-weight: 900; margin: 1rem 0 .55rem; }
.sidebar-version { position: fixed; bottom: 18px; font-size: 11px; opacity: .95; }

section[data-testid="stSidebar"] .stButton > button {
  width: 100% !important;
  min-height: 42px !important;
  justify-content: flex-start !important;
  text-align: left !important;
  background: rgba(255,255,255,.10) !important;
  border: 1px solid rgba(255,255,255,.14) !important;
  border-radius: 13px !important;
  box-shadow: none !important;
  color: #F8FBFF !important;
  font-weight: 750 !important;
  opacity: 1 !important;
}
section[data-testid="stSidebar"] .stButton > button:hover,
section[data-testid="stSidebar"] .stButton > button:focus {
  background: rgba(61,99,242,.72) !important;
  border-color: rgba(255,255,255,.42) !important;
  transform: translateX(2px);
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
  background: linear-gradient(135deg, #36A3FF 0%, #3D63F2 100%) !important;
  border-color: rgba(255,255,255,.65) !important;
  color: #FFFFFF !important;
}

.topbar {
  height: 74px;
  border-radius: 0;
  background: linear-gradient(135deg, #0B1F4D 0%, #152A73 58%, #27349A 100%);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 34px;
  color: white;
  margin: -1.1rem 0 1.3rem 0;
  box-shadow: 0 18px 38px rgba(11,31,77,.16);
  overflow: hidden;
  position: relative;
}
.topbar:after {
  content: "";
  position: absolute;
  right: 105px; top: -60px;
  width: 260px; height: 170px;
  background: rgba(255,255,255,.06);
  transform: rotate(45deg);
}
.topbar-left { display: flex; align-items: center; gap: 28px; z-index: 1; }
.topbar-divider { width: 1px; height: 42px; background: rgba(255,255,255,.38); }
.topbar-title { font-size: 15px; line-height: 1.22; font-weight: 900; letter-spacing: .04em; text-transform: uppercase; }
.topbar-status { z-index: 1; display: flex; align-items: center; gap: 10px; font-size: 12px; font-weight: 800; }
.status-dot { width: 10px; height: 10px; background: #22C55E; border-radius: 999px; box-shadow: 0 0 0 6px rgba(34,197,94,.12); }

.hero {
  background: rgba(255,255,255,.93);
  border: 1px solid var(--en-border);
  border-radius: 20px;
  min-height: 150px;
  padding: 42px 46px;
  box-shadow: 0 20px 48px rgba(11,31,77,.07);
  position: relative;
  overflow: hidden;
  margin-bottom: 24px;
}
.hero:after {
  content: "EN";
  position: absolute;
  right: 32px; top: 4px;
  font-size: 88px;
  line-height: 1;
  font-weight: 950;
  color: rgba(11,31,77,.055);
}
.hero h1, .page-title h1 {
  margin: 0 0 12px 0;
  font-size: 29px;
  color: #071849;
  letter-spacing: -.025em;
  font-weight: 950;
}
.hero p, .page-title p { color: #51617E; margin: 0; font-size: 15px; }
.breadcrumb { font-size: 12px; color: #64748B; margin-bottom: 24px; }

.page-title {
  margin: 20px 0 22px 0;
}
.page-title-row { display: flex; align-items: center; gap: 12px; }
.badge-ok {
  display: inline-flex; align-items: center; gap: 6px;
  background: #DDFBEA; border: 1px solid #8EE6B5; color: #087A41;
  border-radius: 999px; padding: 6px 10px; font-size: 12px; font-weight: 900;
}

.module-card, .wide-card, .info-strip {
  background: rgba(255,255,255,.95);
  border: 1px solid var(--en-border);
  border-radius: 15px;
  box-shadow: 0 18px 35px rgba(11,31,77,.08);
}
.module-card { min-height: 208px; padding: 26px 28px; position: relative; }
.module-icon {
  width: 48px; height: 48px; border-radius: 999px;
  background: #EEF3FF; color: #3156EA;
  display: flex; align-items: center; justify-content: center;
  font-size: 24px; font-weight: 900; margin-bottom: 22px;
}
.module-card h3 { color: #071849; font-size: 18px; font-weight: 950; margin: 0 0 14px 0; }
.module-card p { color: #5A6882; font-size: 13.5px; line-height: 1.55; margin: 0; }
.card-arrow { position: absolute; right: 18px; top: 18px; width: 24px; height: 24px; border-radius: 50%; border: 1px solid #BFD0F4; color: #3156EA; display:flex; align-items:center; justify-content:center; font-weight:900; }
.wide-card { padding: 24px 28px; }
.info-strip { padding: 18px 24px; margin-top: 20px; }
.info-grid { display:grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }
.info-label { font-size: 11px; color:#78849A; margin-bottom: 8px; }
.info-value { color:#071849; font-weight: 900; font-size: 13px; }
.section-title { font-size: 22px; font-weight: 950; color:#071849; margin: 24px 0 10px 0; }
.section-subtitle { color:#61708C; margin-bottom: 20px; }

.stButton > button {
  border-radius: 9px !important;
  border: 1px solid #BFD0F4 !important;
  color: #1642B0 !important;
  background: #FFFFFF !important;
  min-height: 38px !important;
  font-weight: 850 !important;
  box-shadow: none !important;
}
.stButton > button:hover { border-color:#3D63F2 !important; color:#0B1F4D !important; }
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, #2447C7 0%, #3D63F2 100%) !important;
  color: white !important;
  border: 0 !important;
}
.stDownloadButton > button {
  background: linear-gradient(135deg, #2447C7 0%, #3D63F2 100%) !important;
  color: #fff !important;
  border: 0 !important;
  border-radius: 10px !important;
  font-weight: 900 !important;
}

[data-testid="stFileUploader"] section {
  background: #FFFFFF !important;
  border: 1px dashed #97AAD4 !important;
  border-radius: 14px !important;
}
[data-testid="stFileUploader"] * { color: #071849 !important; }
input, textarea, select, div[data-baseweb="select"] > div {
  color: #071849 !important;
  background: #FFFFFF !important;
  border-color: #CAD7EE !important;
}
label, .stMarkdown p, .stCaptionContainer { color: #31405F !important; }
[data-testid="stMetric"] {
  background: #FFFFFF;
  border: 1px solid var(--en-border);
  border-radius: 14px;
  padding: 14px 16px;
  box-shadow: 0 12px 26px rgba(11,31,77,.06);
}
hr { border-color:#DCE5F4; }

@media (max-width: 900px) {
  .topbar { padding: 0 18px; }
  .topbar-title { font-size: 12px; }
  .hero { padding: 30px 24px; }
  .info-grid { grid-template-columns: 1fr 1fr; }
}
</style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# COMPONENTES VISUAIS
# ============================================================

def set_page(page: str) -> None:
    st.session_state["page"] = page
    st.rerun()


def nav_button(label: str, page: str, icon: str, key: str) -> None:
    current = st.session_state.get("page", "inicio")
    btn_type = "primary" if current == page else "secondary"
    if st.sidebar.button(f"{icon} {label}", key=key, type=btn_type, use_container_width=True):
        set_page(page)


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown('<div class="sidebar-brand-box">', unsafe_allow_html=True)
        st.markdown(logo_html(compact=False), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-title">Central Operacional</div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-subtitle">Expresso Nepomuceno</div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-section">Navegação</div>', unsafe_allow_html=True)

        nav_button("Início", "inicio", "⌂", "nav_inicio")
        nav_button("Odômetro / Vínculo", "odometro", "◴", "nav_odometro")
        nav_button("Tempo de Carregamento", "tempo", "◷", "nav_tempo")
        nav_button("Viagens em Bloco", "viagens", "▦", "nav_viagens")
        nav_button("Histórico", "historico", "▤", "nav_historico")
        nav_button("Relatórios", "relatorios", "▥", "nav_relatorios")
        nav_button("Configurações", "configuracoes", "⚙", "nav_configuracoes")

        st.markdown(
            f'<div class="sidebar-version">Versão {APP_VERSION}<br>Ambiente interno</div>',
            unsafe_allow_html=True,
        )


def render_topbar() -> None:
    st.markdown(
        f"""
<div class="topbar">
  <div class="topbar-left">
    {logo_html(compact=True)}
    <div class="topbar-divider"></div>
    <div class="topbar-title">Central Operacional<br>de Análises</div>
  </div>
  <div class="topbar-status"><span class="status-dot"></span> Sistema online</div>
</div>
        """,
        unsafe_allow_html=True,
    )


def page_header(title: str, subtitle: str, active: bool = True) -> None:
    st.markdown(
        f"""
<div class="breadcrumb">Central Operacional de Análises / {html.escape(title)}</div>
<div class="page-title">
  <div class="page-title-row">
    <h1>{html.escape(title)}</h1>
    {'<span class="badge-ok">● Ativo</span>' if active else ''}
  </div>
  <p>{html.escape(subtitle)}</p>
</div>
        """,
        unsafe_allow_html=True,
    )


def render_module_card(icon: str, title: str, desc: str, page: str, key: str) -> None:
    st.markdown(
        f"""
<div class="module-card">
  <div class="card-arrow">›</div>
  <div class="module-icon">{html.escape(icon)}</div>
  <h3>{html.escape(title)}</h3>
  <p>{html.escape(desc)}</p>
</div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Acessar módulo →", key=key, use_container_width=True):
        set_page(page)


def module_missing_box(filename: str) -> bool:
    path = BASE_DIR / filename
    if path.exists():
        return False
    st.error(f"Arquivo obrigatório não encontrado na pasta do app: {filename}")
    st.info("Suba este arquivo junto com o app.py no GitHub/Streamlit Cloud.")
    return True


# ============================================================
# PAGINAS
# ============================================================

def pagina_inicio() -> None:
    st.markdown(
        """
<div class="hero">
  <h1>Bem-vindo à Central Operacional de Análises</h1>
  <p>Acesse as ferramentas de análise, histórico e relatórios do sistema.</p>
</div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Escolha uma ferramenta</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Clique em uma opção abaixo para abrir a página operacional.</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        render_module_card("◴", "Odômetro / Vínculo", "Cruza combustível, Maxtrack, ativos e produção para gerar o ODOMETRO_MATCH final.", "odometro", "home_odometro")
    with c2:
        render_module_card("◷", "Tempo de Carregamento", "Trata permanência por área, ruídos, eventos próximos e resumo por local/placa.", "tempo", "home_tempo")
    with c3:
        render_module_card("▦", "Viagens em Bloco", "Integra Maxtrack, SAP e permanência para montar viagens validadas e deduplicadas.", "viagens", "home_viagens")

    st.write("")
    c4, c5 = st.columns(2)
    with c4:
        render_module_card("▤", "Histórico", "Consulte o histórico de processamentos realizados no portal.", "historico", "home_historico")
    with c5:
        render_module_card("▥", "Relatórios", "Acompanhe indicadores simples de uso, módulos e resultados gerados.", "relatorios", "home_relatorios")

    history = read_history()
    st.markdown(
        f"""
<div class="info-strip">
  <div class="info-grid">
    <div><div class="info-label">Versão do sistema</div><div class="info-value">{APP_VERSION}</div></div>
    <div><div class="info-label">Ambiente</div><div class="info-value">Produção</div></div>
    <div><div class="info-label">Último acesso</div><div class="info-value">{now_str()}</div></div>
    <div><div class="info-label">Processamentos registrados</div><div class="info-value">{len(history)}</div></div>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )


def pagina_odometro() -> None:
    page_header("Odômetro / Vínculo", "Geração do ODOMETRO_MATCH com Maxtrack como fonte principal e Cliente/Produção como complemento.")
    filename = MODULE_FILES["odometro"]
    if module_missing_box(filename):
        return
    mod = load_module_safely(filename)

    st.markdown('<div class="wide-card">', unsafe_allow_html=True)
    st.subheader("1. Envie as quatro bases obrigatórias")
    col1, col2 = st.columns(2)
    with col1:
        combustivel = st.file_uploader("1 - Base Combustível", type=["xlsx", "xls"], key="odo_combustivel")
        ativo = st.file_uploader("3 - Base Ativo de Veículos", type=["xlsx", "xls"], key="odo_ativo")
    with col2:
        maxtrack = st.file_uploader("2 - Base KM Rodado Maxtrack", type=["xlsx", "xls"], key="odo_maxtrack")
        producao = st.file_uploader("4 - Produção Oficial / Cliente", type=["xlsx", "xls"], key="odo_producao")

    nome_saida = st.text_input("Nome do arquivo final", value=f"Resumo_Abastecimento_Odometro_{now_file()}.xlsx", key="odo_saida")
    qtd = sum(x is not None for x in [combustivel, maxtrack, ativo, producao])
    st.progress(qtd / 4, text=f"{qtd} de 4 arquivos carregados")
    st.markdown('</div>', unsafe_allow_html=True)

    if qtd < 4:
        st.warning("Aguardando as quatro bases para liberar o processamento.")
        return

    if st.button("Processar Odômetro / Vínculo", key="btn_processar_odometro", type="primary", use_container_width=True):
        try:
            if not hasattr(mod, "processar_streamlit"):
                raise AttributeError("O módulo não possui a função processar_streamlit().")
            with st.spinner("Processando bases do odômetro..."):
                excel_bytes, indicadores, resultado_final = mod.processar_streamlit(
                    combustivel, maxtrack, ativo, producao, nome_saida
                )
            st.success("Processamento concluído com sucesso.")
            add_history("Odômetro / Vínculo", "Concluído", nome_saida, {"registros": int(len(resultado_final))})

            st.subheader("Indicadores finais")
            st.dataframe(indicadores, use_container_width=True, hide_index=True)
            c1, c2, c3 = st.columns(3)
            c1.metric("Registros finais", f"{len(resultado_final):,}".replace(",", "."))
            c2.metric("ODOMETRO_MATCH preenchido", f"{int(resultado_final['ODOMETRO_MATCH'].notna().sum()):,}".replace(",", ".") if "ODOMETRO_MATCH" in resultado_final else "-")
            c3.metric("Arquivo", safe_filename(nome_saida))
            st.download_button(
                "Baixar Excel final",
                data=excel_bytes,
                file_name=safe_filename(nome_saida if nome_saida.lower().endswith(".xlsx") else f"{nome_saida}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="download_odometro",
            )
        except Exception as exc:
            add_history("Odômetro / Vínculo", "Erro", detalhes={"erro": str(exc)})
            log_event("erro_odometro", {"erro": str(exc), "trace": traceback.format_exc()})
            st.error("Erro ao processar Odômetro / Vínculo. Verifique abas, colunas e formatos dos arquivos.")
            st.exception(exc)


def pagina_tempo_carregamento() -> None:
    page_header("Tempo de Carregamento", "Tratamento de permanência por área com junção de eventos próximos e geração de resumos em Excel.")
    filename = MODULE_FILES["tempo"]
    if module_missing_box(filename):
        return
    mod = load_module_safely(filename)

    st.markdown('<div class="wide-card">', unsafe_allow_html=True)
    arquivo = st.file_uploader("Selecione o arquivo Excel de permanência", type=["xlsx", "xls", "xlsm"], key="tempo_upload")
    st.markdown('</div>', unsafe_allow_html=True)

    if not arquivo:
        st.warning("Aguardando upload do arquivo.")
        return

    if st.button("Processar Tempo de Carregamento", key="btn_tempo_processar", type="primary", use_container_width=True):
        temp_in = ""
        temp_out = ""
        try:
            if not hasattr(mod, "processar_arquivo"):
                raise AttributeError("O módulo não possui a função processar_arquivo().")
            temp_in = save_upload(arquivo, "tempo_carregamento_")
            temp_out = str(Path(tempfile.gettempdir()) / f"{Path(arquivo.name).stem}_Base_Tratada_{now_file()}.xlsx")
            progress = st.progress(0, text="0% - Iniciando")
            status = st.empty()
            status.info("Processando arquivo...")
            progress.progress(0.35, text="35% - Processando dados")
            caminho_saida, tratado, resumo_placa_local, resumo_local, info = mod.processar_arquivo(temp_in, temp_out)
            progress.progress(1.0, text="100% - Finalizado")
            status.success("Processamento concluído")

            add_history("Tempo de Carregamento", "Concluído", Path(caminho_saida).name, {"linhas_validas": int(info.get("linhas_validas", 0))})
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Linha cabeçalho", info.get("linha_cabecalho", "-"))
            c2.metric("Linhas originais", info.get("linhas_original", "-"))
            c3.metric("Linhas tratadas", info.get("linhas_tratadas", "-"))
            c4.metric("Linhas válidas", info.get("linhas_validas", "-"))

            with st.expander("Colunas identificadas", expanded=False):
                st.json({"Placa": info.get("col_placa"), "Área": info.get("col_area"), "Entrada": info.get("col_entrada"), "Saída": info.get("col_saida")})

            st.subheader("Prévia da base tratada")
            st.dataframe(tratado.head(100), use_container_width=True)
            st.subheader("Resumo por placa/local")
            st.dataframe(resumo_placa_local.head(100), use_container_width=True)
            st.subheader("Resumo por local")
            st.dataframe(resumo_local.head(100), use_container_width=True)

            with open(caminho_saida, "rb") as f:
                st.download_button(
                    "Baixar Excel tratado",
                    f.read(),
                    file_name=Path(caminho_saida).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="download_tempo",
                )
        except Exception as exc:
            add_history("Tempo de Carregamento", "Erro", detalhes={"erro": str(exc)})
            log_event("erro_tempo", {"erro": str(exc), "trace": traceback.format_exc()})
            st.error("Erro ao processar Tempo de Carregamento. Verifique abas, colunas e formatos do arquivo.")
            st.exception(exc)
        finally:
            cleanup_files([temp_in])


def pagina_viagens_bloco() -> None:
    page_header("Viagens em Bloco", "Processamento integrado de Maxtrack, SAP e Permanência em Área com validação, score e deduplicação.")
    filename = MODULE_FILES["viagens"]
    if module_missing_box(filename):
        return
    mod = load_module_safely(filename)

    st.markdown('<div class="wide-card">', unsafe_allow_html=True)
    st.subheader("1. Envie as três bases")
    c1, c2, c3 = st.columns(3)
    with c1:
        arq_max = st.file_uploader("Maxtrack", type=["xlsx", "xls"], key="viagens_max")
    with c2:
        arq_sap = st.file_uploader("SAP", type=["xlsx", "xls"], key="viagens_sap")
    with c3:
        arq_perm = st.file_uploader("Permanência em Área", type=["xlsx", "xls"], key="viagens_perm")
    qtd = sum(x is not None for x in [arq_max, arq_sap, arq_perm])
    st.progress(qtd / 3, text=f"{qtd} de 3 arquivos carregados")
    st.markdown('</div>', unsafe_allow_html=True)

    if qtd < 3:
        st.warning("Aguardando as três bases para liberar o processamento.")
        return

    if st.button("Processar Viagens em Bloco", key="btn_viagens_processar", type="primary", use_container_width=True):
        paths: list[str] = []
        try:
            if not hasattr(mod, "processar_arquivos"):
                raise AttributeError("O módulo não possui a função processar_arquivos().")
            p_max = save_upload(arq_max, "viagens_maxtrack_")
            p_sap = save_upload(arq_sap, "viagens_sap_")
            p_perm = save_upload(arq_perm, "viagens_perm_")
            paths = [p_max, p_sap, p_perm]
            progress = StreamlitProgressAdapter()
            with st.spinner("Processando viagens em bloco..."):
                saida, total_viagens, total_placas, total_invalidas, total_duplicadas, total_antes = mod.processar_arquivos(p_max, p_sap, p_perm, progress)

            st.success("Processamento concluído com sucesso.")
            add_history("Viagens em Bloco", "Concluído", Path(saida).name, {"viagens": int(total_viagens), "placas": int(total_placas)})
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Antes dedup", total_antes)
            c2.metric("Viagens finais", total_viagens)
            c3.metric("Placas", total_placas)
            c4.metric("Inválidas", total_invalidas)
            c5.metric("Duplicadas", total_duplicadas)

            with open(saida, "rb") as f:
                st.download_button(
                    "Baixar Excel de viagens",
                    f.read(),
                    file_name=Path(saida).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="download_viagens",
                )
        except Exception as exc:
            add_history("Viagens em Bloco", "Erro", detalhes={"erro": str(exc)})
            log_event("erro_viagens", {"erro": str(exc), "trace": traceback.format_exc()})
            st.error("Erro ao processar Viagens em Bloco. Verifique abas, colunas e formatos dos arquivos.")
            st.exception(exc)
        finally:
            cleanup_files(paths)


def pagina_historico() -> None:
    page_header("Histórico", "Registro dos processamentos executados no portal.")
    history = read_history()
    if not history:
        st.info("Nenhum processamento registrado ainda.")
        return
    df = pd.DataFrame(history)
    st.dataframe(df.iloc[::-1], use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar histórico JSON",
        json.dumps(history, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name=f"historico_processamentos_{now_file()}.json",
        mime="application/json",
        key="download_historico",
    )


def pagina_relatorios() -> None:
    page_header("Relatórios", "Indicadores simples do portal operacional.")
    history = read_history()
    df = pd.DataFrame(history)
    c1, c2, c3 = st.columns(3)
    c1.metric("Processamentos registrados", len(history))
    c2.metric("Módulos ativos", 3)
    c3.metric("Último acesso", now_str())

    if not df.empty:
        st.subheader("Processamentos por módulo")
        resumo = df.groupby(["modulo", "status"], dropna=False).size().reset_index(name="quantidade")
        st.dataframe(resumo, use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados suficientes para relatório.")


def pagina_configuracoes() -> None:
    page_header("Configurações", "Status técnico dos arquivos integrados ao portal.")
    rows = []
    for key, filename in MODULE_FILES.items():
        path = BASE_DIR / filename
        rows.append({
            "Módulo": PAGES[key],
            "Arquivo": filename,
            "Encontrado": "Sim" if path.exists() else "Não",
            "Tamanho KB": round(path.stat().st_size / 1024, 1) if path.exists() else 0,
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.info("O portal usa módulos fixos e controlados. O reconhecimento automático instável foi removido para evitar travamentos e execução indevida de scripts.")


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    apply_css()
    if "page" not in st.session_state:
        st.session_state["page"] = "inicio"

    render_sidebar()
    render_topbar()

    page = st.session_state.get("page", "inicio")
    if page == "inicio":
        pagina_inicio()
    elif page == "odometro":
        pagina_odometro()
    elif page == "tempo":
        pagina_tempo_carregamento()
    elif page == "viagens":
        pagina_viagens_bloco()
    elif page == "historico":
        pagina_historico()
    elif page == "relatorios":
        pagina_relatorios()
    elif page == "configuracoes":
        pagina_configuracoes()
    else:
        st.session_state["page"] = "inicio"
        st.rerun()


if __name__ == "__main__":
    main()
