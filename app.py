import os
import sys
import time
import base64
import tempfile
import subprocess
import importlib.util
from pathlib import Path
from datetime import datetime

import streamlit as st


# =========================================================
# CONFIGURACAO BASE
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
APP_FILE = Path(__file__).name

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="Sistema Operacional de Analises",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# IDENTIDADE VISUAL
# =========================================================
CORES = {
    "azul_escuro": "#111D4A",
    "azul_medio": "#153E75",
    "azul_claro": "#1D8ACB",
    "cinza_fundo": "#F4F7FB",
    "cinza_borda": "#E2E8F0",
    "cinza_texto": "#334155",
    "cinza_suave": "#64748B",
    "branco": "#FFFFFF",
}


def encontrar_logo() -> Path | None:
    candidatos = [
        "logo_nepomuceno.jpeg",
        "logo_nepomuceno.jpg",
        "logo_nepomuceno.png",
        "Logo Nepomuceno.jpeg",
        "Logo Nepomuceno.jpg",
        "Logo Nepomuceno.png",
        "WhatsApp Image 2025-08-12 at 15.22.05.jpeg",
    ]

    for nome in candidatos:
        caminho = BASE_DIR / nome
        if caminho.exists():
            return caminho

    return None


def imagem_base64(caminho: Path | None) -> str:
    if caminho is None or not caminho.exists():
        return ""

    try:
        dados = caminho.read_bytes()
        return base64.b64encode(dados).decode("utf-8")
    except Exception:
        return ""


LOGO_PATH = encontrar_logo()
LOGO_B64 = imagem_base64(LOGO_PATH)
LOGO_HTML = (
    f'<img src="data:image/jpeg;base64,{LOGO_B64}" class="company-logo" alt="Expresso Nepomuceno">'
    if LOGO_B64
    else '<div class="company-logo-fallback">EN</div>'
)


st.markdown(
    f"""
<style>
:root {{
    --azul-escuro: {CORES['azul_escuro']};
    --azul-medio: {CORES['azul_medio']};
    --azul-claro: {CORES['azul_claro']};
    --cinza-fundo: {CORES['cinza_fundo']};
    --cinza-borda: {CORES['cinza_borda']};
    --cinza-texto: {CORES['cinza_texto']};
    --cinza-suave: {CORES['cinza_suave']};
    --branco: {CORES['branco']};
}}

html, body, [class*="css"] {{
    font-family: "Segoe UI", Roboto, Arial, sans-serif;
}}

.stApp {{
    background: linear-gradient(180deg, #F8FAFC 0%, #EEF3F8 100%);
}}

.block-container {{
    padding-top: 1.25rem;
    padding-bottom: 3rem;
    max-width: 1420px;
}}

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #0F1B3D 0%, #10182E 100%);
    border-right: 1px solid rgba(255, 255, 255, 0.08);
}}

[data-testid="stSidebar"] * {{
    color: #E5E7EB;
}}

[data-testid="stSidebar"] .stRadio label {{
    color: #E5E7EB !important;
}}

[data-testid="stSidebar"] [role="radiogroup"] label {{
    background: rgba(255, 255, 255, 0.06);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 12px;
    padding: 0.65rem 0.75rem;
    margin-bottom: 0.45rem;
}}

[data-testid="stSidebar"] [role="radiogroup"] label:hover {{
    background: rgba(29, 138, 203, 0.22);
    border-color: rgba(29, 138, 203, 0.55);
}}

[data-testid="stSidebar"] hr {{
    border-color: rgba(255, 255, 255, 0.12);
}}

.main-header {{
    background: linear-gradient(135deg, #101A3C 0%, #142B5C 58%, #1A5F95 100%);
    border-radius: 22px;
    padding: 26px 30px;
    border: 1px solid rgba(255, 255, 255, 0.12);
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 22px;
    margin-bottom: 22px;
}}

.header-left {{
    display: flex;
    align-items: center;
    gap: 22px;
}}

.company-logo {{
    width: 280px;
    max-width: 32vw;
    height: auto;
    border-radius: 14px;
    object-fit: contain;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.20);
}}

.company-logo-fallback {{
    width: 96px;
    height: 64px;
    border-radius: 14px;
    background: rgba(255, 255, 255, 0.12);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    font-weight: 900;
    color: white;
}}

.header-eyebrow {{
    font-size: 12px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-weight: 800;
    color: #9FD3F7;
    margin-bottom: 7px;
}}

.header-title {{
    color: white;
    font-size: 36px;
    line-height: 1.05;
    font-weight: 850;
    margin: 0;
}}

.header-subtitle {{
    color: #CBD5E1;
    font-size: 15px;
    margin-top: 9px;
    max-width: 780px;
}}

.header-badge {{
    min-width: 150px;
    text-align: center;
    padding: 12px 15px;
    border-radius: 999px;
    background: rgba(255, 255, 255, 0.10);
    border: 1px solid rgba(255, 255, 255, 0.18);
    color: white;
    font-weight: 750;
}}

.page-title-card {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid var(--cinza-borda);
    border-radius: 18px;
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
    width: 46px;
    height: 46px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #EAF5FF;
    color: var(--azul-medio);
    font-size: 25px;
    font-weight: 800;
}}

.page-title {{
    color: #1E293B;
    font-size: 31px;
    font-weight: 850;
    margin: 0;
}}

.page-subtitle {{
    color: var(--cinza-suave);
    margin: 4px 0 0 0;
    font-size: 15px;
}}

.executive-card {{
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid var(--cinza-borda);
    border-radius: 18px;
    padding: 20px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06);
    margin-bottom: 16px;
}}

.card-title {{
    font-size: 17px;
    color: #1E293B;
    font-weight: 800;
    margin-bottom: 6px;
}}

.card-text {{
    color: var(--cinza-suave);
    font-size: 14px;
    line-height: 1.55;
}}

.status-pill {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #ECFDF5;
    border: 1px solid #BBF7D0;
    color: #166534;
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 750;
}}

.info-pill {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    color: #1D4ED8;
    padding: 7px 12px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 750;
}}

.warning-card {{
    background: #FFF7ED;
    border: 1px solid #FED7AA;
    color: #9A3412;
    border-radius: 16px;
    padding: 16px 18px;
    margin-bottom: 16px;
    font-weight: 650;
}}

.success-card {{
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    color: #166534;
    border-radius: 16px;
    padding: 16px 18px;
    margin-bottom: 16px;
    font-weight: 650;
}}

.stButton > button {{
    border-radius: 12px !important;
    border: 1px solid #CBD5E1 !important;
    font-weight: 750 !important;
    min-height: 42px;
}}

.stDownloadButton > button {{
    border-radius: 12px !important;
    background: linear-gradient(135deg, #153E75 0%, #1D8ACB 100%) !important;
    color: white !important;
    border: 0 !important;
    font-weight: 800 !important;
    min-height: 45px;
}}

.stProgress > div > div > div > div {{
    background: linear-gradient(90deg, #153E75 0%, #1D8ACB 100%);
}}

[data-testid="stMetric"] {{
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    padding: 15px 16px;
    box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
}}

[data-testid="stFileUploader"] section {{
    border-radius: 16px;
    border: 1px dashed #94A3B8;
    background: #F8FAFC;
}}

hr {{
    margin: 1.3rem 0;
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
# FUNCOES GERAIS
# =========================================================
def atualizar_progresso(barra, status, pct: int, texto: str) -> None:
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.20)


def listar_scripts_python() -> list[str]:
    ignorar = {
        APP_FILE,
        "app.py",
        "app_executivo.py",
        "__init__.py",
    }

    arquivos: list[str] = []

    for arq in BASE_DIR.iterdir():
        if arq.is_file() and arq.suffix.lower() == ".py" and arq.name not in ignorar:
            arquivos.append(arq.name)

    return sorted(arquivos)


def nome_amigavel_script(nome_arquivo: str) -> str:
    nome = nome_arquivo.replace(".py", "")
    nome = nome.replace("_", " ").replace("-", " ")
    return " ".join(parte.capitalize() for parte in nome.split())


def set_page_config_noop(*args, **kwargs) -> None:
    """Evita erro quando um modulo interno tambem chama st.set_page_config()."""
    return None


def carregar_modulo_por_arquivo(caminho_script: Path):
    nome_modulo = caminho_script.stem.replace(" ", "_").replace("-", "_")
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


def executar_modulo_streamlit(modulo) -> None:
    original_set_page_config = st.set_page_config
    try:
        st.set_page_config = set_page_config_noop

        if hasattr(modulo, "main_streamlit"):
            modulo.main_streamlit()
        elif hasattr(modulo, "main"):
            modulo.main()
        else:
            st.warning("Este modulo nao possui funcao main_streamlit() nem main().")
    finally:
        st.set_page_config = original_set_page_config


def render_header() -> None:
    data_atual = datetime.now().strftime("%d/%m/%Y")
    st.markdown(
        f"""
        <div class="main-header">
            <div class="header-left">
                {LOGO_HTML}
                <div>
                    <div class="header-eyebrow">Inteligencia operacional</div>
                    <h1 class="header-title">Sistema Operacional de Analises</h1>
                    <div class="header-subtitle">
                        Plataforma executiva para processamento de bases, indicadores operacionais e geracao de arquivos tratados.
                    </div>
                </div>
            </div>
            <div class="header-badge">Atualizado em<br>{data_atual}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_page_title(icone: str, titulo: str, subtitulo: str) -> None:
    st.markdown(
        f"""
        <div class="page-title-card">
            <div class="page-title-row">
                <div class="page-icon">{icone}</div>
                <div>
                    <h2 class="page-title">{titulo}</h2>
                    <p class="page-subtitle">{subtitulo}</p>
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
            <div class="card-title">{titulo}</div>
            <div class="card-text">{texto}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar(menu: list[str]) -> str:
    with st.sidebar:
        if LOGO_B64:
            st.image(str(LOGO_PATH), use_container_width=True)
        else:
            st.markdown("### Expresso Nepomuceno")

        st.markdown("---")
        st.markdown("**Navegacao**")

        if "pagina_atual" not in st.session_state:
            st.session_state.pagina_atual = menu[0]

        if st.session_state.pagina_atual not in menu:
            st.session_state.pagina_atual = menu[0]

        pagina = st.radio(
            "Selecione uma area",
            menu,
            index=menu.index(st.session_state.pagina_atual),
            label_visibility="collapsed",
            key="menu_principal",
        )

        st.session_state.pagina_atual = pagina

        st.markdown("---")
        st.caption("Ambiente interno de analises")
        st.caption(f"Pasta: {BASE_DIR.name}")

    return pagina


# =========================================================
# PAGINAS
# =========================================================
def pagina_inicio(scripts_detectados: list[str]) -> None:
    render_page_title(
        "🏠",
        "Visao geral",
        "Central de modulos para processamento de dados operacionais.",
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("Modulos ativos", len(scripts_detectados) + 2)
    col2.metric("Padrao de layout", "Executivo")
    col3.metric("Arquivos Python", len(scripts_detectados))

    st.markdown("### Modulos disponiveis")

    c1, c2, c3 = st.columns(3)
    with c1:
        render_card(
            "📊 Analise de Permanencia",
            "Processa base de permanencia, identifica eventos, classifica tempos e gera Excel final para acompanhamento operacional.",
        )
    with c2:
        render_card(
            "🚛 Odometro V12",
            "Cruza bases de combustivel, Maxtrack, ativos e producao oficial para consolidacao do relatorio final.",
        )
    with c3:
        render_card(
            "🧩 Modulos dinamicos",
            "Scripts Python detectados automaticamente na pasta do sistema e executados quando possuem main_streamlit() ou main().",
        )

    st.markdown(
        """
        <div class="success-card">
            ✅ Para adicionar novos modulos, coloque o arquivo .py na mesma pasta do app.py e inclua uma funcao <code>main_streamlit()</code> ou <code>main()</code>.
        </div>
        """,
        unsafe_allow_html=True,
    )


def pagina_permanencia() -> None:
    render_page_title(
        "📊",
        "Analise de Permanencia",
        "Processamento da base de permanencia com classificacao por tempo configuravel.",
    )

    caminho_permanencia = BASE_DIR / "Codigo_colado.py"

    if not caminho_permanencia.exists():
        st.error("Arquivo Codigo_colado.py nao encontrado na pasta do app.")
        st.stop()

    try:
        permanencia = carregar_modulo_por_arquivo(caminho_permanencia)
    except Exception as e:
        st.error(f"Erro ao importar Codigo_colado.py: {e}")
        st.stop()

    st.markdown("<span class='info-pill'>⚙️ Parametros de processamento</span>", unsafe_allow_html=True)
    st.write("")

    col_param1, col_param2 = st.columns(2)

    with col_param1:
        tempo_minimo = st.number_input(
            "Tempo minimo aceitavel (minutos)",
            min_value=0,
            value=15,
            step=1,
        )

    with col_param2:
        tempo_maximo = st.number_input(
            "Tempo maximo aceitavel (minutos)",
            min_value=1,
            value=55,
            step=1,
        )

    if tempo_maximo <= tempo_minimo:
        st.error("O tempo maximo precisa ser maior que o minimo.")

    arquivo = st.file_uploader(
        "Selecione o Excel de permanencia",
        type=["xlsx", "xls"],
        key="upload_permanencia",
    )

    if not arquivo:
        st.markdown(
            """
            <div class="warning-card">
                Aguardando upload da base de permanencia para iniciar o processamento.
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f"""
        <div class="success-card">
            ✅ Arquivo carregado: <b>{arquivo.name}</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🚀 Processar Permanencia", use_container_width=True):
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

            st.markdown("### Previa dos resultados")
            if not df_resultado.empty:
                st.dataframe(df_resultado.head(100), use_container_width=True)
            else:
                st.warning("Nenhum resultado gerado.")

            with open(saida, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel Permanencia",
                    f,
                    file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Erro ao processar permanencia: {e}")


def pagina_odometro() -> None:
    render_page_title(
        "🚛",
        "Odometro V12",
        "Consolidacao do ODOMETRO_MATCH com as quatro bases obrigatorias.",
    )

    st.info("Envie as 4 bases obrigatorias para gerar o Excel final do odometro.")

    col1, col2 = st.columns(2)

    with col1:
        comb = st.file_uploader("1 - Base Combustivel", type=["xlsx", "xls"], key="comb")
        ativo = st.file_uploader("3 - Base Ativo de Veiculos", type=["xlsx", "xls"], key="ativo")

    with col2:
        maxtrack = st.file_uploader("2 - Base Km Rodado Maxtrack", type=["xlsx", "xls"], key="maxtrack")
        producao = st.file_uploader("4 - Producao Oficial / Cliente", type=["xlsx", "xls"], key="producao")

    qtd_carregados = sum(arquivo is not None for arquivo in [comb, maxtrack, ativo, producao])
    st.progress(qtd_carregados / 4)
    st.caption(f"{qtd_carregados} de 4 arquivos carregados")

    if not (comb and maxtrack and ativo and producao):
        st.markdown(
            """
            <div class="warning-card">
                Aguardando upload das quatro bases para liberar o processamento.
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        """
        <div class="success-card">
            ✅ Todas as bases foram carregadas. O processamento ja pode ser iniciado.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🚀 Processar Odometro V12", use_container_width=True):
        inicio = time.time()
        barra = st.progress(0)
        status = st.empty()
        log_box = st.empty()

        try:
            atualizar_progresso(barra, status, 10, "Salvando arquivos temporarios")

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

            script = BASE_DIR / "odometro_v12_com_percentual.py"

            if not script.exists():
                st.error("Arquivo odometro_v12_com_percentual.py nao encontrado.")
                st.stop()

            atualizar_progresso(barra, status, 25, "Executando script do odometro")

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
                    status.info(f"{progresso}% - Processando odometro V12...")

                time.sleep(0.4)

            if processo.returncode != 0:
                st.error("O processamento retornou erro.")
                st.stop()

            if not os.path.exists(saida):
                st.error("Arquivo final nao foi gerado.")
                st.stop()

            atualizar_progresso(barra, status, 100, "Finalizado")

            tempo_total = round(time.time() - inicio, 2)
            st.success(f"Odometro finalizado em {tempo_total} segundos")

            with open(saida, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel Odometro V12",
                    f,
                    file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.error(f"Erro ao processar odometro: {e}")


def pagina_modulo_dinamico(script: str) -> None:
    nome = nome_amigavel_script(script)
    caminho_script = BASE_DIR / script

    render_page_title(
        "🧩",
        nome,
        f"Modulo detectado automaticamente: {script}",
    )

    if not caminho_script.exists():
        st.error(f"Arquivo nao encontrado: {script}")
        st.stop()

    try:
        modulo = carregar_modulo_por_arquivo(caminho_script)
    except Exception as e:
        st.error(f"Erro ao carregar modulo {script}: {e}")
        st.stop()

    if hasattr(modulo, "main_streamlit") or hasattr(modulo, "main"):
        st.markdown(
            """
            <div class="success-card">
                ✅ Modulo compativel com o site. A tela abaixo pertence ao script selecionado.
            </div>
            """,
            unsafe_allow_html=True,
        )
        executar_modulo_streamlit(modulo)
    else:
        st.markdown(
            """
            <div class="warning-card">
                Para este modulo rodar integrado ao site, o arquivo precisa ter uma funcao <code>main_streamlit()</code> ou <code>main()</code>.
            </div>
            """,
            unsafe_allow_html=True,
        )


# =========================================================
# APLICACAO
# =========================================================
def main() -> None:
    scripts_detectados = listar_scripts_python()

    modulos_fixos = [
        "🏠 Inicio",
        "📊 Analise de Permanencia",
        "🚛 Odometro V12",
    ]

    modulos_dinamicos = [
        f"🧩 {nome_amigavel_script(s)}"
        for s in scripts_detectados
        if s not in ["Codigo_colado.py", "odometro_v12_com_percentual.py"]
    ]

    menu = modulos_fixos + modulos_dinamicos

    pagina = render_sidebar(menu)
    render_header()

    mapa_dinamico = {
        f"🧩 {nome_amigavel_script(s)}": s
        for s in scripts_detectados
        if s not in ["Codigo_colado.py", "odometro_v12_com_percentual.py"]
    }

    if pagina == "🏠 Inicio":
        pagina_inicio(scripts_detectados)
    elif pagina == "📊 Analise de Permanencia":
        pagina_permanencia()
    elif pagina == "🚛 Odometro V12":
        pagina_odometro()
    elif pagina in mapa_dinamico:
        pagina_modulo_dinamico(mapa_dinamico[pagina])
    else:
        st.error("Pagina nao encontrada.")


if __name__ == "__main__":
    main()
