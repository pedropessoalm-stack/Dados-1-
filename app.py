import os
import sys
import time
import tempfile
import subprocess
from datetime import datetime

import streamlit as st
import Codigo_colado as permanencia


st.set_page_config(
    page_title="Sistema de Análises",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================
# MENU SUPERIOR
# =========================
st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
}
.menu-card {
    background-color: #111827;
    padding: 18px;
    border-radius: 14px;
    border: 1px solid #374151;
    margin-bottom: 25px;
}
.titulo-sistema {
    font-size: 34px;
    font-weight: 800;
    color: white;
}
.subtitulo {
    color: #CBD5E1;
    font-size: 15px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="menu-card">
    <div class="titulo-sistema">🚀 Sistema Operacional de Análises</div>
    <div class="subtitulo">Escolha o módulo abaixo para iniciar o processamento.</div>
</div>
""", unsafe_allow_html=True)

pagina = st.radio(
    "Menu",
    [
        "📊 Análise de Permanência",
        "🚛 Odômetro V12"
    ],
    horizontal=True,
    label_visibility="collapsed"
)

st.divider()


def atualizar_progresso(barra, status, pct, texto):
    barra.progress(pct)
    status.info(f"{pct}% - {texto}")
    time.sleep(0.3)


# =========================================================
# PÁGINA 1 - PERMANÊNCIA
# =========================================================
if pagina == "📊 Análise de Permanência":

    st.title("📊 Análise de Permanência")

    col_param1, col_param2 = st.columns(2)

    with col_param1:
        tempo_minimo = st.number_input(
            "Tempo mínimo aceitável (minutos)",
            min_value=0,
            value=15,
            step=1
        )

    with col_param2:
        tempo_maximo = st.number_input(
            "Tempo máximo aceitável (minutos)",
            min_value=1,
            value=55,
            step=1
        )

    if tempo_maximo <= tempo_minimo:
        st.error("O tempo máximo precisa ser maior que o mínimo.")

    arquivo = st.file_uploader(
        "Selecione o Excel de permanência",
        type=["xlsx", "xls"]
    )

    if arquivo:
        st.success(f"Arquivo carregado: {arquivo.name}")

        if st.button("🚀 Processar Permanência", use_container_width=True):

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
                    ranking_improcedentes=ranking
                )

                atualizar_progresso(barra, status, 100, "Finalizado")

                tempo_total = round(time.time() - inicio, 2)

                st.success(f"✅ Processo finalizado em {tempo_total} segundos")

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
                        "⬇️ Baixar Excel Permanência",
                        f,
                        file_name=f"resultado_permanencia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Erro ao processar: {e}")


# =========================================================
# PÁGINA 2 - ODÔMETRO
# =========================================================
if pagina == "🚛 Odômetro V12":

    st.title("🚛 Odômetro V12 - ODOMETRO_MATCH")

    st.info("Envie as 4 bases obrigatórias para gerar o Excel final do odômetro.")

    col1, col2 = st.columns(2)

    with col1:
        comb = st.file_uploader("1️⃣ Base Combustível", type=["xlsx", "xls"])
        ativo = st.file_uploader("3️⃣ Base Ativo de Veículos", type=["xlsx", "xls"])

    with col2:
        maxtrack = st.file_uploader("2️⃣ Base Km Rodado Maxtrack", type=["xlsx", "xls"])
        producao = st.file_uploader("4️⃣ Produção Oficial / Cliente", type=["xlsx", "xls"])

    if comb and maxtrack and ativo and producao:

        st.success("✅ Todas as bases foram carregadas.")

        if st.button("🚀 Processar Odômetro V12", use_container_width=True):

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

                saida = os.path.join(
                    tempfile.gettempdir(),
                    f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )

                script = os.path.join(os.getcwd(), "odometro_v12_com_percentual.py")

                if not os.path.exists(script):
                    st.error("Arquivo odometro_v12_com_percentual.py não encontrado na pasta atual.")
                    st.stop()

                atualizar_progresso(barra, status, 25, "Executando script do odômetro")

                processo = subprocess.Popen(
                    [sys.executable, script, p1, p2, p3, p4, saida],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace"
                )

                logs = []
                progresso = 25

                while True:
                    linha = processo.stdout.readline()

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
                    st.error("❌ O processamento retornou erro.")
                    st.stop()

                if not os.path.exists(saida):
                    st.error("Arquivo final não foi gerado.")
                    st.stop()

                atualizar_progresso(barra, status, 100, "Finalizado")

                tempo_total = round(time.time() - inicio, 2)

                st.success(f"✅ Odômetro finalizado em {tempo_total} segundos")

                with open(saida, "rb") as f:
                    st.download_button(
                        "⬇️ Baixar Excel Odômetro V12",
                        f,
                        file_name=f"resultado_odometro_v12_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Erro ao processar odômetro: {e}")

    else:
        st.warning("Aguardando upload das 4 bases.")