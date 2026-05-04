import ast
import base64
import html
import importlib.util
import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st


# =========================================================
# CONFIGURACAO DA APLICACAO
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_FILE = Path(__file__).name

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Central Operacional de Analises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# IDENTIDADE VISUAL
# =========================================================
CORES = {
    "azul_900": "#071A3A",
    "azul_800": "#0B1F4D",
    "azul_700": "#123B7A",
    "azul_500": "#1D8ACB",
    "azul_400": "#2F80ED",
    "fundo": "#F4F7FB",
    "card": "#FFFFFF",
    "borda": "#DDE6F2",
    "texto": "#0F172A",
    "texto_suave": "#53657D",
    "sucesso": "#147A3E",
    "alerta": "#A15C04",
    "erro": "#B42318",
}


# =========================================================
# FUNCOES DE ARQUIVO / LOGO
# =========================================================
def normalizar_nome_arquivo(nome: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", nome.lower())


def encontrar_arquivo(candidatos: list[str]) -> Path | None:
    """Procura um arquivo na pasta do app aceitando variacoes de nome."""
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
        'class="co-logo" alt="Expresso Nepomuceno">'
    )
else:
    LOGO_HTML = '<div class="co-logo-fallback">EN</div>'


# =========================================================
# CSS CORPORATIVO - prefixado para reduzir conflitos
# =========================================================
def aplicar_css_base() -> None:
    st.markdown(
        f"""
<style>
:root {{
    --co-azul-900: {CORES['azul_900']};
    --co-azul-800: {CORES['azul_800']};
    --co-azul-700: {CORES['azul_700']};
    --co-azul-500: {CORES['azul_500']};
    --co-azul-400: {CORES['azul_400']};
    --co-fundo: {CORES['fundo']};
    --co-card: {CORES['card']};
    --co-borda: {CORES['borda']};
    --co-texto: {CORES['texto']};
    --co-texto-suave: {CORES['texto_suave']};
}}

html, body, [class*="css"], .stApp {{
    font-family: "Segoe UI", Roboto, Arial, sans-serif;
}}

.stApp {{
    background:
        radial-gradient(circle at top left, rgba(47, 128, 237, 0.10), transparent 34%),
        linear-gradient(180deg, #F8FBFF 0%, #EDF3FA 100%);
}}

.block-container {{
    padding-top: 1.1rem;
    padding-bottom: 4.5rem;
    max-width: 1480px;
}}

#MainMenu, footer, header {{
    visibility: hidden;
}}

/* Correcao definitiva de contraste em paginas internas */
main .stMarkdown, main .stMarkdown p, main label, main span, main div[data-testid="stWidgetLabel"],
main div[data-testid="stMarkdownContainer"], main div[data-testid="stMetricLabel"],
main div[data-testid="stMetricValue"], main div[data-testid="stCaptionContainer"] {{
    color: var(--co-texto) !important;
}}

main div[data-testid="stMarkdownContainer"] p,
main .stCaptionContainer,
main small {{
    color: var(--co-texto-suave) !important;
}}

/* Sidebar profissional sem aparencia de radio button */
section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #071A3A 0%, #0B1328 100%);
    border-right: 1px solid rgba(255, 255, 255, 0.09);
}}

section[data-testid="stSidebar"] > div {{
    padding-top: 1.2rem;
}}

section[data-testid="stSidebar"] * {{
    color: #E8EEF8 !important;
}}

section[data-testid="stSidebar"] .stButton > button {{
    width: 100%;
    min-height: 48px;
    justify-content: flex-start;
    text-align: left;
    border-radius: 15px !important;
    border: 1px solid rgba(255, 255, 255, 0.10) !important;
    background: rgba(255, 255, 255, 0.055) !important;
    color: #F8FAFC !important;
    font-weight: 760 !important;
    padding: 0.72rem 0.9rem !important;
    box-shadow: none !important;
}}

section[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(29, 138, 203, 0.20) !important;
    border-color: rgba(147, 197, 253, 0.46) !important;
    transform: translateX(2px);
}}

section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] > div:has(button[kind="primary"]) button {{
    background: linear-gradient(135deg, #1D8ACB 0%, #2F80ED 100%) !important;
    border-color: rgba(255,255,255,0.22) !important;
    color: #FFFFFF !important;
}}

.co-sidebar-brand {{
    padding: 10px 4px 16px 4px;
}}

.co-sidebar-logo-wrap {{
    display: flex;
    align-items: center;
    gap: 11px;
}}

.co-sidebar-logo {{
    width: 42px;
    height: 42px;
    border-radius: 13px;
    background: rgba(255, 255, 255, 0.13);
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 900;
    color: #FFFFFF;
    letter-spacing: 0.04em;
}}

.co-sidebar-title {{
    font-size: 15px;
    font-weight: 900;
    line-height: 1.15;
    color: #FFFFFF;
}}

.co-sidebar-subtitle {{
    margin-top: 6px;
    font-size: 12px;
    line-height: 1.45;
    color: #A8B3C7 !important;
}}

.co-sidebar-section {{
    font-size: 11px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    font-weight: 900;
    color: #93C5FD !important;
    margin-top: 18px;
    margin-bottom: 8px;
}}

/* Headers */
.co-hero {{
    background: linear-gradient(135deg, #0B1F4D 0%, #123B7A 56%, #1D8ACB 100%);
    border-radius: 26px;
    padding: 34px 38px;
    border: 1px solid rgba(255, 255, 255, 0.16);
    box-shadow: 0 22px 54px rgba(15, 23, 42, 0.18);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 28px;
    margin-bottom: 26px;
}}

.co-header-compact {{
    background: linear-gradient(135deg, #0B1F4D 0%, #123B7A 58%, #1D8ACB 100%);
    border-radius: 20px;
    padding: 18px 22px;
    border: 1px solid rgba(255, 255, 255, 0.14);
    box-shadow: 0 16px 38px rgba(15, 23, 42, 0.14);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    margin-bottom: 18px;
}}

.co-header-left {{
    display: flex;
    align-items: center;
    gap: 20px;
}}

.co-logo {{
    width: 230px;
    max-width: 26vw;
    height: auto;
    border-radius: 16px;
    object-fit: contain;
    box-shadow: 0 10px 28px rgba(0, 0, 0, 0.22);
}}

.co-header-compact .co-logo {{
    width: 132px;
    max-width: 18vw;
}}

.co-logo-fallback {{
    width: 96px;
    height: 64px;
    border-radius: 16px;
    background: rgba(255, 255, 255, 0.15);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    font-weight: 900;
    color: white !important;
}}

.co-header-compact .co-logo-fallback {{
    width: 58px;
    height: 44px;
    font-size: 19px;
}}

.co-eyebrow {{
    font-size: 12px;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    font-weight: 900;
    color: #BDE7FF !important;
    margin-bottom: 8px;
}}

.co-hero-title {{
    color: #FFFFFF !important;
    font-size: 42px;
    line-height: 1.04;
    font-weight: 950;
    margin: 0;
}}

.co-page-title {{
    color: #FFFFFF !important;
    font-size: 27px;
    line-height: 1.08;
    font-weight: 950;
    margin: 0;
}}

.co-header-subtitle {{
    color: #DCEAFE !important;
    font-size: 15px;
    margin-top: 8px;
    max-width: 860px;
}}

.co-header-badge {{
    min-width: 156px;
    text-align: center;
    padding: 12px 16px;
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.12);
    border: 1px solid rgba(255, 255, 255, 0.24);
    color: #FFFFFF !important;
    font-weight: 900;
    font-size: 13px;
}}

/* Cards */
.co-section-title {{
    color: var(--co-texto) !important;
    font-size: 26px;
    font-weight: 950;
    margin: 10px 0 8px 0;
}}

.co-section-subtitle {{
    color: var(--co-texto-suave) !important;
    font-size: 15px;
    line-height: 1.6;
    margin-bottom: 18px;
}}

.co-card {{
    background: rgba(255, 255, 255, 0.98);
    border: 1px solid var(--co-borda);
    border-radius: 22px;
    padding: 22px;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.08);
    margin-bottom: 16px;
}}

.co-module-card {{
    min-height: 235px;
    border-top: 5px solid #1D8ACB;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
}}

.co-icon {{
    width: 56px;
    height: 56px;
    border-radius: 18px;
    background: #EAF5FF;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    margin-bottom: 18px;
}}

.co-card-title {{
    color: var(--co-texto) !important;
    font-size: 21px;
    font-weight: 940;
    margin: 0 0 9px 0;
}}

.co-card-text {{
    color: var(--co-texto-suave) !important;
    font-size: 15px;
    line-height: 1.62;
}}

.co-card-footer {{
    color: #0B3A7A !important;
    font-weight: 900;
    font-size: 14px;
    margin-top: 18px;
}}

.co-panel-title {{
    color: var(--co-texto) !important;
    font-size: 20px;
    font-weight: 930;
    margin-bottom: 4px;
}}

.co-panel-text {{
    color: var(--co-texto-suave) !important;
    font-size: 14px;
    line-height: 1.56;
}}

.co-upload-card {{
    background: #FFFFFF;
    border: 1px solid #DDE6F2;
    border-radius: 20px;
    padding: 18px 18px 14px 18px;
    box-shadow: 0 12px 28px rgba(15, 23, 42, 0.065);
    margin-bottom: 14px;
}}

.co-upload-head {{
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 14px;
    margin-bottom: 12px;
}}

.co-upload-step {{
    width: 34px;
    height: 34px;
    border-radius: 12px;
    background: #EAF5FF;
    color: #0B3A7A !important;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 950;
}}

.co-upload-title {{
    color: var(--co-texto) !important;
    font-size: 16px;
    font-weight: 920;
    margin: 0;
}}

.co-upload-desc {{
    color: var(--co-texto-suave) !important;
    font-size: 13px;
    margin-top: 3px;
}}

.co-pill {{
    display: inline-flex;
    align-items: center;
    gap: 7px;
    border-radius: 999px;
    padding: 7px 12px;
    font-size: 12px;
    font-weight: 900;
    white-space: nowrap;
}}

.co-pill-ok {{
    color: #166534 !important;
    background: #ECFDF5;
    border: 1px solid #B7F4CF;
}}

.co-pill-wait {{
    color: #92400E !important;
    background: #FFF7ED;
    border: 1px solid #FED7AA;
}}

.co-pill-info {{
    color: #1D4ED8 !important;
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
}}

.co-message-success,
.co-message-warning,
.co-message-info,
.co-message-error {{
    border-radius: 17px;
    padding: 15px 17px;
    margin-bottom: 16px;
    font-weight: 760;
}}

.co-message-success {{
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    color: #166534 !important;
}}

.co-message-warning {{
    background: #FFF7ED;
    border: 1px solid #FED7AA;
    color: #9A3412 !important;
}}

.co-message-info {{
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    color: #1D4ED8 !important;
}}

.co-message-error {{
    background: #FEF3F2;
    border: 1px solid #FECDCA;
    color: #B42318 !important;
}}

.co-status-card {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    background: #FFFFFF;
    border: 1px solid var(--co-borda);
    border-radius: 22px;
    padding: 22px 26px;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.07);
    margin-top: 10px;
}}

.co-muted {{
    color: var(--co-texto-suave) !important;
}}

/* Componentes nativos */
.stButton > button {{
    border-radius: 14px !important;
    border: 1px solid #CBD5E1 !important;
    font-weight: 850 !important;
    min-height: 44px;
}}

.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, #123B7A 0%, #1D8ACB 100%) !important;
    color: #FFFFFF !important;
    border: 0 !important;
}}

.stDownloadButton > button {{
    border-radius: 14px !important;
    background: linear-gradient(135deg, #123B7A 0%, #1D8ACB 100%) !important;
    color: #FFFFFF !important;
    border: 0 !important;
    font-weight: 900 !important;
    min-height: 45px;
}}

.stProgress > div > div > div > div {{
    background: linear-gradient(90deg, #123B7A 0%, #1D8ACB 100%);
}}

[data-testid="stFileUploader"] section {{
    border-radius: 15px;
    border: 1px dashed #8EA8C9;
    background: #F8FBFF;
}}

[data-testid="stFileUploader"] button {{
    border-radius: 12px !important;
}}

[data-testid="stDataFrame"] {{
    border-radius: 16px;
    overflow: hidden;
}}

[data-testid="stMetric"] {{
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 17px;
    padding: 15px 16px;
    box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
}}

hr {{
    margin: 1.15rem 0;
}}

@media (max-width: 900px) {{
    .co-hero, .co-header-compact, .co-status-card {{
        display: block;
    }}

    .co-header-left {{
        display: block;
    }}

    .co-logo {{
        max-width: 100%;
        width: 100%;
        margin-bottom: 18px;
    }}

    .co-header-compact .co-logo {{
        max-width: 100%;
        width: 120px;
    }}

    .co-hero-title {{
        font-size: 31px;
    }}

    .co-page-title {{
        font-size: 24px;
    }}

    .co-header-badge {{
        margin-top: 16px;
        width: fit-content;
    }}
}}
</style>
""",
        unsafe_allow_html=True,
    )


aplicar_css_base()


# =========================================================
# MODELO DE MODULO DINAMICO SEGURO
# =========================================================
@dataclass
class ModuloDinamico:
    arquivo: str
    caminho: Path
    nome: str
    icone: str
    descricao: str
    categoria: str
    ordem: int
    slug: str
    ativo: bool
    status: str
    erro: str = ""


IGNORAR_EXATOS = {
    APP_FILE.lower(),
    "app.py",
    "app_executivo.py",
    "app_executivo_final.py",
    "app_home_profissional.py",
    "app_home_limpa_click.py",
    "app_profissional_final.py",
    "__init__.py",
    "codigo_colado.py",
    "código_colado.py",
    "codigo colado.py",
    "código colado.py",
    "odometro_v12_com_percentual.py",
}

PREFIXOS_INVALIDOS = (
    "streamlit ",
    "pip ",
    "py ",
    "python ",
    "teste",
    "test_",
    "_",
)

SLUGS_RESERVADOS = {
    "inicio",
    "permanencia",
    "analisepermanencia",
    "odometro",
    "odometrov12",
    "appodometrostreamlit",
    "permanenciaarea",
    "codigocolado",
}


def slugify(texto: str) -> str:
    texto = texto.lower().strip()
    texto = re.sub(r"[^a-z0-9]+", "", texto)
    return texto or "modulo"


def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".py", "")
    nome = nome.replace("_", " ").replace("-", " ")
    return " ".join(parte.capitalize() for parte in nome.split())


def script_deve_ser_ignorado(caminho: Path) -> bool:
    nome_lower = caminho.name.lower()
    if nome_lower in IGNORAR_EXATOS:
        return True

    if nome_lower.startswith(PREFIXOS_INVALIDOS):
        return True

    return False


def ler_config_modulo_por_ast(caminho: Path) -> tuple[dict[str, Any] | None, bool, str]:
    """Valida sem importar: evita executar codigo de modulo quebrado na montagem do menu."""
    try:
        arvore = ast.parse(caminho.read_text(encoding="utf-8"), filename=str(caminho))
    except UnicodeDecodeError:
        return None, False, "Arquivo sem codificacao UTF-8."
    except SyntaxError as exc:
        return None, False, f"Erro de sintaxe: {exc.msg} na linha {exc.lineno}."
    except Exception as exc:
        return None, False, f"Nao foi possivel ler o arquivo: {exc}"

    config = None
    possui_main_streamlit = False

    for node in arvore.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "main_streamlit":
            possui_main_streamlit = True

        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "MODULO_CONFIG":
                    try:
                        config = ast.literal_eval(node.value)
                    except Exception:
                        return None, possui_main_streamlit, "MODULO_CONFIG precisa ser um dicionario literal."

    if not isinstance(config, dict):
        return None, possui_main_streamlit, "MODULO_CONFIG nao encontrado."

    if not possui_main_streamlit:
        return config, False, "Funcao main_streamlit() nao encontrada."

    return config, True, ""


def validar_modulo_dinamico(caminho: Path, slugs_usados: set[str]) -> ModuloDinamico:
    config, possui_main, erro = ler_config_modulo_por_ast(caminho)

    if config is None:
        return ModuloDinamico(
            arquivo=caminho.name,
            caminho=caminho,
            nome=nome_amigavel_script(caminho.name),
            icone="🧩",
            descricao="Modulo ignorado por nao seguir o padrao exigido.",
            categoria="Ignorado",
            ordem=9999,
            slug=slugify(caminho.stem),
            ativo=False,
            status="invalido",
            erro=erro,
        )

    nome = str(config.get("nome") or nome_amigavel_script(caminho.name)).strip()
    icone = str(config.get("icone") or "🧩").strip()
    descricao = str(config.get("descricao") or "Modulo operacional adicionado ao portal.").strip()
    categoria = str(config.get("categoria") or "Modulos adicionais").strip()
    ativo = bool(config.get("ativo", True))

    try:
        ordem = int(config.get("ordem", 100))
    except Exception:
        ordem = 100

    slug = slugify(str(config.get("slug") or nome))

    if not possui_main:
        return ModuloDinamico(caminho.name, caminho, nome, icone, descricao, categoria, ordem, slug, False, "invalido", erro)

    if not ativo:
        return ModuloDinamico(caminho.name, caminho, nome, icone, descricao, categoria, ordem, slug, False, "inativo", "Modulo marcado como ativo=False.")

    if slug in SLUGS_RESERVADOS or slug in slugs_usados:
        return ModuloDinamico(caminho.name, caminho, nome, icone, descricao, categoria, ordem, slug, False, "duplicado", "Nome/slug duplicado ou reservado.")

    slugs_usados.add(slug)
    return ModuloDinamico(caminho.name, caminho, nome, icone, descricao, categoria, ordem, slug, True, "ok")


def descobrir_modulos_dinamicos() -> tuple[list[ModuloDinamico], list[ModuloDinamico]]:
    validos: list[ModuloDinamico] = []
    problemas: list[ModuloDinamico] = []
    slugs_usados: set[str] = set()

    for caminho in sorted(BASE_DIR.glob("*.py"), key=lambda p: p.name.lower()):
        if script_deve_ser_ignorado(caminho):
            continue

        modulo = validar_modulo_dinamico(caminho, slugs_usados)
        if modulo.status == "ok" and modulo.ativo:
            validos.append(modulo)
        else:
            problemas.append(modulo)

    validos.sort(key=lambda m: (m.ordem, m.nome.lower()))
    problemas.sort(key=lambda m: m.arquivo.lower())
    return validos, problemas


MODULOS_DINAMICOS, MODULOS_COM_PROBLEMA = descobrir_modulos_dinamicos()


# =========================================================
# FUNCOES GERAIS DE EXECUCAO
# =========================================================
def atualizar_progresso(barra, status, pct: int, texto: str) -> None:
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.18)


def nome_modulo_seguro(caminho_script: Path) -> str:
    base = re.sub(r"\W+", "_", caminho_script.stem, flags=re.UNICODE).strip("_")
    if not base:
        base = "modulo"
    return f"modulo_{base}_{abs(hash(str(caminho_script))) % 10_000_000}"


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


def remover_style_tags(conteudo: str) -> str:
    return re.sub(r"<style[\s\S]*?</style>", "", conteudo, flags=re.IGNORECASE)


def executar_modulo_streamlit(modulo) -> None:
    """Executa modulo dinamico com protecao basica contra set_page_config e CSS global."""
    original_set_page_config = st.set_page_config
    original_markdown = st.markdown

    def markdown_sem_css_global(body, *args, **kwargs):
        if isinstance(body, str) and "<style" in body.lower():
            body_limpo = remover_style_tags(body)
            if body_limpo.strip():
                return original_markdown(body_limpo, *args, **kwargs)
            st.warning("CSS global do modulo foi bloqueado para preservar o layout do portal.")
            return None
        return original_markdown(body, *args, **kwargs)

    try:
        st.set_page_config = set_page_config_noop
        st.markdown = markdown_sem_css_global

        if not hasattr(modulo, "main_streamlit"):
            st.error("Este modulo nao possui main_streamlit().")
            return

        modulo.main_streamlit()
    finally:
        st.set_page_config = original_set_page_config
        st.markdown = original_markdown


def validar_funcoes_modulo(modulo, funcoes_obrigatorias: list[str]) -> list[str]:
    return [funcao for funcao in funcoes_obrigatorias if not hasattr(modulo, funcao)]


def ir_para(pagina: str) -> None:
    st.session_state.pagina_atual = pagina
    st.rerun()


# =========================================================
# COMPONENTES VISUAIS
# =========================================================
def render_header_home() -> None:
    data_atual = datetime.now().strftime("%d/%m/%Y")
    st.markdown(
        f"""
        <div class="co-hero">
            <div class="co-header-left">
                {LOGO_HTML}
                <div>
                    <div class="co-eyebrow">Inteligencia operacional</div>
                    <h1 class="co-hero-title">Central Operacional de Analises</h1>
                    <div class="co-header-subtitle">
                        Portal executivo para processamento de bases, controle de rotinas e geracao de arquivos tratados.
                    </div>
                </div>
            </div>
            <div class="co-header-badge">Ambiente interno<br>{data_atual}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_header_compacto(titulo: str, subtitulo: str, icone: str) -> None:
    data_atual = datetime.now().strftime("%d/%m/%Y")
    st.markdown(
        f"""
        <div class="co-header-compact">
            <div class="co-header-left">
                {LOGO_HTML}
                <div>
                    <div class="co-eyebrow">{html.escape(icone)} Modulo operacional</div>
                    <h1 class="co-page-title">{html.escape(titulo)}</h1>
                    <div class="co-header-subtitle">{html.escape(subtitulo)}</div>
                </div>
            </div>
            <div class="co-header-badge">Online<br>{data_atual}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_card(titulo: str, texto: str, icone: str = "•", rodape: str = "") -> None:
    st.markdown(
        f"""
        <div class="co-card co-module-card">
            <div>
                <div class="co-icon">{html.escape(icone)}</div>
                <h3 class="co-card-title">{html.escape(titulo)}</h3>
                <div class="co-card-text">{html.escape(texto)}</div>
            </div>
            <div class="co-card-footer">{html.escape(rodape)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_message(tipo: str, texto: str) -> None:
    classe = {
        "success": "co-message-success",
        "warning": "co-message-warning",
        "info": "co-message-info",
        "error": "co-message-error",
    }.get(tipo, "co-message-info")

    st.markdown(f'<div class="{classe}">{texto}</div>', unsafe_allow_html=True)


def render_status_card() -> None:
    st.markdown(
        """
        <div class="co-status-card">
            <div>
                <div class="co-panel-title">Central pronta para operacao</div>
                <div class="co-panel-text">
                    Selecione um modulo no menu lateral ou acesse diretamente pelos cards principais.
                </div>
            </div>
            <div class="co-pill co-pill-ok">● Sistema online</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_card(numero: int, titulo: str, descricao: str, arquivo) -> None:
    status_html = (
        '<span class="co-pill co-pill-ok">● Carregado</span>'
        if arquivo is not None
        else '<span class="co-pill co-pill-wait">● Pendente</span>'
    )
    nome_arquivo = html.escape(arquivo.name) if arquivo is not None else "Aguardando arquivo Excel"

    st.markdown(
        f"""
        <div class="co-upload-card">
            <div class="co-upload-head">
                <div style="display:flex; gap:12px; align-items:flex-start;">
                    <div class="co-upload-step">{numero}</div>
                    <div>
                        <div class="co-upload-title">{html.escape(titulo)}</div>
                        <div class="co-upload-desc">{html.escape(descricao)}</div>
                        <div class="co-upload-desc"><b>Arquivo:</b> {nome_arquivo}</div>
                    </div>
                </div>
                {status_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> str:
    paginas_fixas = [
        ("inicio", "🏠 Início"),
        ("permanencia", "📊 Análise de Permanência"),
        ("odometro", "🚛 Odômetro V12"),
    ]

    paginas_dinamicas = [(f"dyn:{m.slug}", f"{m.icone} {m.nome}") for m in MODULOS_DINAMICOS]
    todas_paginas = paginas_fixas + paginas_dinamicas

    if "pagina_atual" not in st.session_state:
        st.session_state.pagina_atual = "inicio"

    ids_validos = {p[0] for p in todas_paginas}
    if st.session_state.pagina_atual not in ids_validos:
        st.session_state.pagina_atual = "inicio"

    with st.sidebar:
        st.markdown(
            """
            <div class="co-sidebar-brand">
                <div class="co-sidebar-logo-wrap">
                    <div class="co-sidebar-logo">EN</div>
                    <div>
                        <div class="co-sidebar-title">Central de Analises</div>
                        <div class="co-sidebar-subtitle">Processamento de bases e relatorios internos.</div>
                    </div>
                </div>
            </div>
            <div class="co-sidebar-section">Navegacao</div>
            """,
            unsafe_allow_html=True,
        )

        for pagina_id, rotulo in paginas_fixas:
            tipo = "primary" if st.session_state.pagina_atual == pagina_id else "secondary"
            if st.button(rotulo, key=f"nav_{pagina_id}", type=tipo, use_container_width=True):
                st.session_state.pagina_atual = pagina_id
                st.rerun()

        if MODULOS_DINAMICOS:
            st.markdown('<div class="co-sidebar-section">Modulos adicionais</div>', unsafe_allow_html=True)
            for modulo in MODULOS_DINAMICOS:
                pagina_id = f"dyn:{modulo.slug}"
                tipo = "primary" if st.session_state.pagina_atual == pagina_id else "secondary"
                if st.button(f"{modulo.icone} {modulo.nome}", key=f"nav_{pagina_id}", type=tipo, use_container_width=True):
                    st.session_state.pagina_atual = pagina_id
                    st.rerun()

        if MODULOS_COM_PROBLEMA:
            with st.expander("Diagnostico de modulos"):
                st.caption("Arquivos ignorados por seguranca ou falta de padrao.")
                for modulo in MODULOS_COM_PROBLEMA:
                    st.caption(f"{modulo.arquivo}: {modulo.erro}")

    return st.session_state.pagina_atual


# =========================================================
# PAGINA INICIAL
# =========================================================
def pagina_inicio() -> None:
    render_header_home()

    st.markdown('<div class="co-section-title">Módulos principais</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="co-section-subtitle">Acesse rapidamente os processos operacionais homologados da central.</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        render_card(
            "Análise de Permanência",
            "Processa a base de permanência, identifica eventos, classifica tempos e gera Excel final para acompanhamento operacional.",
            "📊",
            "Base Excel",
        )
        if st.button("Abrir Análise de Permanência", key="home_permanencia", use_container_width=True, type="primary"):
            ir_para("permanencia")

    with col2:
        render_card(
            "Odômetro V12",
            "Consolida combustível, Maxtrack, ativos e produção oficial para gerar o relatório final com cruzamento das quatro bases.",
            "🚛",
            "4 bases obrigatórias",
        )
        if st.button("Abrir Odômetro V12", key="home_odometro", use_container_width=True, type="primary"):
            ir_para("odometro")

    with col3:
        render_card(
            "Módulos Adicionais",
            "Novos módulos entram no portal somente quando seguem o padrão seguro com MODULO_CONFIG e main_streamlit().",
            "🧩",
            "Plugins controlados",
        )
        if MODULOS_DINAMICOS:
            primeiro = MODULOS_DINAMICOS[0]
            if st.button("Abrir Módulo Adicional", key="home_dinamico", use_container_width=True):
                ir_para(f"dyn:{primeiro.slug}")
        else:
            st.button("Nenhum módulo adicional ativo", key="home_dinamico_disabled", use_container_width=True, disabled=True)

    render_status_card()


# =========================================================
# PAGINA FIXA 1 - ANALISE DE PERMANENCIA
# =========================================================
def pagina_permanencia() -> None:
    render_header_compacto(
        "Análise de Permanência",
        "Processamento da base de permanência com classificação por tempo configurável.",
        "📊",
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
        render_message("error", "Arquivo <b>Codigo_colado.py</b> não encontrado na pasta do app.")
        st.stop()

    try:
        permanencia = carregar_modulo_por_arquivo(caminho_permanencia)
    except Exception as e:
        render_message("error", f"Erro ao importar <b>{html.escape(caminho_permanencia.name)}</b>: {html.escape(str(e))}")
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
        render_message(
            "error",
            "O arquivo de permanência foi encontrado, mas não possui as funções esperadas: "
            + ", ".join(html.escape(f) for f in funcoes_faltando),
        )
        st.stop()

    st.markdown('<div class="co-card">', unsafe_allow_html=True)
    st.markdown('<div class="co-panel-title">Parâmetros de tratamento</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="co-panel-text">Defina a faixa aceitável para classificação dos tempos antes de processar a base.</div>',
        unsafe_allow_html=True,
    )

    col_param1, col_param2 = st.columns(2)
    with col_param1:
        tempo_minimo = st.number_input("Tempo mínimo aceitável (minutos)", min_value=0, value=15, step=1)
    with col_param2:
        tempo_maximo = st.number_input("Tempo máximo aceitável (minutos)", min_value=1, value=55, step=1)
    st.markdown('</div>', unsafe_allow_html=True)

    if tempo_maximo <= tempo_minimo:
        render_message("error", "O tempo máximo precisa ser maior que o tempo mínimo.")

    st.markdown('<div class="co-card">', unsafe_allow_html=True)
    st.markdown('<div class="co-panel-title">Importação da base</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="co-panel-text">Envie o arquivo Excel de permanência para iniciar o processamento.</div>',
        unsafe_allow_html=True,
    )
    arquivo = st.file_uploader("Selecione o Excel de permanência", type=["xlsx", "xls"], key="upload_permanencia")
    st.markdown('</div>', unsafe_allow_html=True)

    if not arquivo:
        render_message("warning", "Aguardando upload da base de permanência para iniciar o processamento.")
        return

    render_message("success", f"Arquivo carregado: <b>{html.escape(arquivo.name)}</b>")

    if st.button("Processar Permanência", use_container_width=True, type="primary"):
        if tempo_maximo <= tempo_minimo:
            render_message("error", "Corrija os tempos antes de processar.")
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
                    "Baixar Excel Permanência",
                    f,
                    file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            render_message("error", f"Erro ao processar permanência: {html.escape(str(e))}")


# =========================================================
# PAGINA FIXA 2 - ODOMETRO
# =========================================================
def pagina_odometro() -> None:
    render_header_compacto(
        "Odômetro V12",
        "Consolidação do ODOMETRO_MATCH com as quatro bases obrigatórias.",
        "🚛",
    )

    render_message("info", "Envie as quatro bases obrigatórias para gerar o Excel final do odômetro.")

    col1, col2 = st.columns(2)

    with col1:
        render_upload_card(1, "Base Combustível", "Arquivo de abastecimentos e dados de combustível.", st.session_state.get("comb"))
        comb = st.file_uploader("Upload da Base Combustível", type=["xlsx", "xls"], key="comb", label_visibility="collapsed")

        render_upload_card(3, "Base Ativo de Veículos", "Cadastro de veículos, placas e ativos operacionais.", st.session_state.get("ativo"))
        ativo = st.file_uploader("Upload da Base Ativo de Veículos", type=["xlsx", "xls"], key="ativo", label_visibility="collapsed")

    with col2:
        render_upload_card(2, "Base Km Rodado Maxtrack", "Base de quilometragem rodada extraída do Maxtrack.", st.session_state.get("maxtrack"))
        maxtrack = st.file_uploader("Upload da Base Km Rodado Maxtrack", type=["xlsx", "xls"], key="maxtrack", label_visibility="collapsed")

        render_upload_card(4, "Produção Oficial / Cliente", "Arquivo oficial de produção utilizado no cruzamento final.", st.session_state.get("producao"))
        producao = st.file_uploader("Upload da Produção Oficial / Cliente", type=["xlsx", "xls"], key="producao", label_visibility="collapsed")

    qtd_carregados = sum(arquivo is not None for arquivo in [comb, maxtrack, ativo, producao])
    st.progress(qtd_carregados / 4)
    st.caption(f"{qtd_carregados} de 4 arquivos carregados")

    if not (comb and maxtrack and ativo and producao):
        render_message("warning", "Aguardando upload das quatro bases para liberar o processamento.")
        return

    render_message("success", "Todas as bases foram carregadas. O processamento já pode ser iniciado.")

    if st.button("Processar Odômetro V12", use_container_width=True, type="primary"):
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

            saida = os.path.join(tempfile.gettempdir(), f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            script = encontrar_arquivo(["odometro_v12_com_percentual.py"])

            if script is None:
                render_message("error", "Arquivo <b>odometro_v12_com_percentual.py</b> não encontrado.")
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
                render_message("error", "O processamento retornou erro. Verifique os logs exibidos acima.")
                st.stop()

            if not os.path.exists(saida):
                render_message("error", "Arquivo final não foi gerado.")
                st.stop()

            atualizar_progresso(barra, status, 100, "Finalizado")
            tempo_total = round(time.time() - inicio, 2)
            st.success(f"Odômetro finalizado em {tempo_total} segundos")

            with open(saida, "rb") as f:
                st.download_button(
                    "Baixar Excel Odômetro V12",
                    f,
                    file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            render_message("error", f"Erro ao processar odômetro: {html.escape(str(e))}")


# =========================================================
# MODULOS DINAMICOS
# =========================================================
def pagina_modulo_dinamico(modulo_info: ModuloDinamico) -> None:
    render_header_compacto(modulo_info.nome, modulo_info.descricao, modulo_info.icone)

    if not modulo_info.caminho.exists():
        render_message("error", f"Arquivo não encontrado: <b>{html.escape(modulo_info.arquivo)}</b>")
        st.stop()

    try:
        modulo = carregar_modulo_por_arquivo(modulo_info.caminho)
    except Exception as e:
        render_message(
            "error",
            f"Erro ao carregar o módulo <b>{html.escape(modulo_info.nome)}</b>: {html.escape(str(e))}",
        )
        st.stop()

    if not hasattr(modulo, "main_streamlit"):
        render_message("error", "Módulo inválido: função <code>main_streamlit()</code> não encontrada.")
        st.stop()

    render_message("success", "Módulo validado e carregado com segurança.")
    executar_modulo_streamlit(modulo)


# =========================================================
# APLICACAO PRINCIPAL
# =========================================================
def main() -> None:
    pagina = render_sidebar()

    if pagina == "inicio":
        pagina_inicio()
    elif pagina == "permanencia":
        pagina_permanencia()
    elif pagina == "odometro":
        pagina_odometro()
    elif pagina.startswith("dyn:"):
        slug = pagina.replace("dyn:", "", 1)
        modulo = next((m for m in MODULOS_DINAMICOS if m.slug == slug), None)
        if modulo is None:
            render_header_compacto("Página não encontrada", "O módulo solicitado não está ativo ou foi removido.", "⚠️")
            render_message("error", "Página não encontrada.")
        else:
            pagina_modulo_dinamico(modulo)
    else:
        render_header_compacto("Página não encontrada", "Selecione uma opção válida no menu lateral.", "⚠️")
        render_message("error", "Página não encontrada.")


if __name__ == "__main__":
    main()
