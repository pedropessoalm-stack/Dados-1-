import tempfile
import unicodedata
from pathlib import Path
from datetime import time, datetime

import pandas as pd
import streamlit as st


TEMPO_MINIMO_SEGUNDOS = 60
JUNTAR_INTERVALO_SEGUNDOS = 300
TEMPO_RUIDO_SEGUNDOS = 60


def normalizar(txt):
    txt = str(txt).strip()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return txt.upper()


def segundos_excel(seg):
    if pd.isna(seg) or seg < 0:
        return 0
    return seg / 86400


def detectar_cabecalho(df):
    palavras = ["PLACA", "AREA", "ENTRADA", "SAIDA"]
    melhor = 0
    score_max = -1

    for i, row in df.iterrows():
        texto = " ".join(normalizar(x) for x in row.values)
        score = sum(p in texto for p in palavras)

        if score > score_max:
            score_max = score
            melhor = i

    return melhor


def localizar_coluna(df, nomes):
    cols = {normalizar(c): c for c in df.columns}

    for nome in nomes:
        nome_norm = normalizar(nome)
        for c_norm, c in cols.items():
            if c_norm == nome_norm:
                return c

    for nome in nomes:
        nome_norm = normalizar(nome)
        for c_norm, c in cols.items():
            if "MOTORISTA" in c_norm:
                continue
            if nome_norm in c_norm:
                return c

    return None


def travar_saida_no_dia(entrada, saida):
    if pd.isna(entrada) or pd.isna(saida):
        return saida

    fim_do_dia = pd.Timestamp.combine(entrada.date(), time(23, 59, 59))

    if saida.date() > entrada.date():
        return fim_do_dia

    return saida


def processar_arquivo(caminho_entrada, caminho_saida):
    caminho = Path(caminho_entrada)

    df_raw = pd.read_excel(caminho, header=None)
    df_raw = df_raw.dropna(how="all").dropna(axis=1, how="all")

    linha = detectar_cabecalho(df_raw)

    df_original = pd.read_excel(caminho, header=linha)
    df_original = df_original.dropna(how="all").dropna(axis=1, how="all")
    df_original.columns = [str(c).strip() for c in df_original.columns]

    df = df_original.copy()

    col_placa = localizar_coluna(df, ["Placa"])
    col_area = localizar_coluna(df, ["Nome da Área", "Nome da Area", "Área"])
    col_entrada = localizar_coluna(df, ["Data de entrada"])
    col_saida = localizar_coluna(df, ["Data de saída", "Data de saida"])

    if not all([col_placa, col_area, col_entrada, col_saida]):
        raise ValueError(
            f"Erro ao localizar colunas obrigatórias. "
            f"Placa={col_placa}, Área={col_area}, Entrada={col_entrada}, Saída={col_saida}"
        )

    df[col_entrada] = pd.to_datetime(df[col_entrada], errors="coerce", dayfirst=True)
    df[col_saida] = pd.to_datetime(df[col_saida], errors="coerce", dayfirst=True)

    df = df[df[col_entrada].notna()].copy()
    df[col_saida] = df[col_saida].fillna(df[col_entrada])
    df = df[df[col_saida] >= df[col_entrada]].copy()

    df["DATA_SAIDA_ORIGINAL"] = df[col_saida]

    df[col_saida] = df.apply(
        lambda x: travar_saida_no_dia(x[col_entrada], x[col_saida]),
        axis=1
    )

    df = df.sort_values([col_placa, col_entrada]).reset_index(drop=True)

    df["TEMPO_SEG_ORIGINAL"] = (
        df[col_saida] - df[col_entrada]
    ).dt.total_seconds()

    df["FLAG_RUIDO"] = df["TEMPO_SEG_ORIGINAL"] < TEMPO_RUIDO_SEGUNDOS
    df["ID_ORIGINAL"] = df.index + 1

    resultado = []

    for placa, grupo in df.groupby(col_placa, dropna=False):

        grupo = grupo.sort_values(col_entrada).reset_index(drop=True)
        atual = None

        for i in range(len(grupo)):
            row = grupo.loc[i]

            if atual is None:
                atual = row.copy()
                atual["QTD_EVENTOS_ORIGINAIS"] = 1
                atual["REGRA_JUNCAO"] = "EVENTO_INICIAL"
                continue

            mesma_area = normalizar(atual[col_area]) == normalizar(row[col_area])
            intervalo = (row[col_entrada] - atual[col_saida]).total_seconds()

            proxima_area = None
            if i + 1 < len(grupo):
                proxima_area = grupo.loc[i + 1, col_area]

            ruido_confirmado_aba = (
                row["FLAG_RUIDO"]
                and proxima_area is not None
                and normalizar(atual[col_area]) == normalizar(proxima_area)
                and normalizar(row[col_area]) != normalizar(atual[col_area])
            )

            if mesma_area and 0 <= intervalo <= JUNTAR_INTERVALO_SEGUNDOS:
                atual[col_saida] = max(atual[col_saida], row[col_saida])
                atual["DATA_SAIDA_ORIGINAL"] = max(
                    atual["DATA_SAIDA_ORIGINAL"],
                    row["DATA_SAIDA_ORIGINAL"]
                )
                atual["QTD_EVENTOS_ORIGINAIS"] += 1
                atual["REGRA_JUNCAO"] = "MESMA_AREA_GAP_ATE_5_MIN"

            elif ruido_confirmado_aba:
                atual[col_saida] = max(atual[col_saida], row[col_saida])
                atual["DATA_SAIDA_ORIGINAL"] = max(
                    atual["DATA_SAIDA_ORIGINAL"],
                    row["DATA_SAIDA_ORIGINAL"]
                )
                atual["QTD_EVENTOS_ORIGINAIS"] += 1
                atual["REGRA_JUNCAO"] = "RUIDO_ENTRE_MESMA_AREA_A_B_A"

            else:
                resultado.append(atual)
                atual = row.copy()
                atual["QTD_EVENTOS_ORIGINAIS"] = 1
                atual["REGRA_JUNCAO"] = "NOVO_EVENTO"

        if atual is not None:
            resultado.append(atual)

    tratado = pd.DataFrame(resultado)

    tratado["TEMPO_SEG"] = (
        tratado[col_saida] - tratado[col_entrada]
    ).dt.total_seconds()

    tratado["STATUS_TRATAMENTO"] = "VALIDO"

    tratado.loc[
        tratado["TEMPO_SEG"] < TEMPO_MINIMO_SEGUNDOS,
        "STATUS_TRATAMENTO"
    ] = "EXPURGADO_TEMPO_PEQUENO"

    tratado["TEMPO"] = tratado["TEMPO_SEG"].apply(segundos_excel)
    tratado["DATA_ENTRADA"] = tratado[col_entrada].dt.date
    tratado["HORA_ENTRADA"] = tratado[col_entrada].dt.hour

    base_valida = tratado[
        tratado["STATUS_TRATAMENTO"] != "EXPURGADO_TEMPO_PEQUENO"
    ].copy()

    resumo_placa_local = (
        base_valida
        .groupby([col_placa, col_area], dropna=False)
        .agg(
            QUANTIDADE_PERMANENCIAS=("TEMPO_SEG", "count"),
            TEMPO_TOTAL_SEG=("TEMPO_SEG", "sum"),
            TEMPO_MEDIO_SEG=("TEMPO_SEG", "mean"),
            PRIMEIRA_ENTRADA=(col_entrada, "min"),
            ULTIMA_SAIDA=(col_saida, "max")
        )
        .reset_index()
    )

    resumo_placa_local["TEMPO_TOTAL"] = resumo_placa_local["TEMPO_TOTAL_SEG"].apply(segundos_excel)
    resumo_placa_local["TEMPO_MEDIO"] = resumo_placa_local["TEMPO_MEDIO_SEG"].apply(segundos_excel)

    resumo_placa_local = resumo_placa_local[
        [
            col_placa,
            col_area,
            "QUANTIDADE_PERMANENCIAS",
            "TEMPO_TOTAL",
            "TEMPO_MEDIO",
            "PRIMEIRA_ENTRADA",
            "ULTIMA_SAIDA"
        ]
    ]

    resumo_local = (
        base_valida
        .groupby(col_area, dropna=False)
        .agg(
            QUANTIDADE_PLACA=(col_placa, "nunique"),
            TEMPO_TOTAL_SEG=("TEMPO_SEG", "sum"),
            PRIMEIRA_ENTRADA=(col_entrada, "min"),
            ULTIMA_SAIDA=(col_saida, "max")
        )
        .reset_index()
    )

    resumo_local["TEMPO_MEDIO_POR_PLACA_SEG"] = (
        resumo_local["TEMPO_TOTAL_SEG"] / resumo_local["QUANTIDADE_PLACA"]
    )

    resumo_local["TEMPO_TOTAL"] = resumo_local["TEMPO_TOTAL_SEG"].apply(segundos_excel)
    resumo_local["TEMPO_MEDIO_POR_PLACA"] = resumo_local["TEMPO_MEDIO_POR_PLACA_SEG"].apply(segundos_excel)

    resumo_local = resumo_local[
        [
            col_area,
            "QUANTIDADE_PLACA",
            "TEMPO_TOTAL",
            "TEMPO_MEDIO_POR_PLACA",
            "PRIMEIRA_ENTRADA",
            "ULTIMA_SAIDA"
        ]
    ]

    base_proj = base_valida[
        base_valida[col_area].astype(str).apply(lambda x: "PROJ" in normalizar(x))
    ].copy()

    if not base_proj.empty:
        base_proj["DATA"] = base_proj[col_entrada].dt.date
        base_proj["HORA"] = base_proj[col_entrada].dt.hour

        locais_proj = sorted(base_proj[col_area].dropna().unique())

        dados_proj_longo = (
            base_proj
            .groupby([col_area, "DATA", "HORA"], dropna=False)
            .agg(
                QUANTIDADE_VEICULOS=(col_placa, "nunique"),
                QUANTIDADE_CHEGADAS=(col_placa, "count"),
                TEMPO_TOTAL_SEG=("TEMPO_SEG", "sum"),
                TEMPO_MEDIO_SEG=("TEMPO_SEG", "mean")
            )
            .reset_index()
        )
    else:
        locais_proj = []
        dados_proj_longo = pd.DataFrame(columns=[
            col_area,
            "DATA",
            "HORA",
            "QUANTIDADE_VEICULOS",
            "QUANTIDADE_CHEGADAS",
            "TEMPO_TOTAL_SEG",
            "TEMPO_MEDIO_SEG"
        ])

    colunas_prioritarias = [
        "ID_ORIGINAL",
        col_placa,
        col_area,
        col_entrada,
        col_saida,
        "DATA_SAIDA_ORIGINAL",
        "DATA_ENTRADA",
        "HORA_ENTRADA",
        "TEMPO",
        "TEMPO_SEG",
        "QTD_EVENTOS_ORIGINAIS",
        "REGRA_JUNCAO",
        "STATUS_TRATAMENTO",
        "FLAG_RUIDO"
    ]

    outras_colunas = [c for c in tratado.columns if c not in colunas_prioritarias]
    tratado = tratado[colunas_prioritarias + outras_colunas]

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:

        abas = {
            "BASE_ORIGINAL": df_original,
            "BASE_TRATADA": tratado,
            "RESUMO_PLACA_LOCAL": resumo_placa_local,
            "RESUMO_LOCAL": resumo_local
        }

        for nome_aba, dataframe in abas.items():
            dataframe.to_excel(writer, sheet_name=nome_aba, index=False)

        wb = writer.book

        fmt_header = wb.add_format({
            "bold": True,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_header_amarelo = wb.add_format({
            "bold": True,
            "bg_color": "#FFE699",
            "font_color": "black",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_titulo = wb.add_format({
            "bold": True,
            "font_size": 13,
            "bg_color": "#D9EAF7",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_tempo = wb.add_format({
            "num_format": "[h]:mm:ss",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_tempo_vertical = wb.add_format({
            "num_format": "[h]:mm",
            "align": "center",
            "valign": "vcenter",
            "rotation": 90,
            "border": 1
        })

        fmt_data = wb.add_format({
            "num_format": "dd/mm/yyyy hh:mm:ss",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_data_curta = wb.add_format({
            "num_format": "dd/mm/yyyy",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_numero = wb.add_format({
            "num_format": "0",
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        fmt_media = wb.add_format({
            "num_format": "0.0",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#FCE4D6"
        })

        fmt_texto = wb.add_format({
            "valign": "vcenter",
            "border": 1
        })

        fmt_zero = wb.add_format({
            "num_format": "0",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#FCE4D6"
        })

        fmt_baixo = wb.add_format({
            "num_format": "0",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#F8CBAD"
        })

        fmt_alto = wb.add_format({
            "num_format": "0",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "bg_color": "#FF0000",
            "font_color": "white",
            "bold": True
        })

        for nome_aba, dataframe in abas.items():
            ws = writer.sheets[nome_aba]

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(dataframe), len(dataframe.columns) - 1)
            ws.hide_gridlines(2)

            for col_idx, col_nome in enumerate(dataframe.columns):
                ws.write(0, col_idx, col_nome, fmt_header)

                nome_norm = normalizar(col_nome)

                try:
                    largura = max(
                        len(str(col_nome)) + 3,
                        dataframe[col_nome].astype(str).str.len().max() + 2
                        if not dataframe.empty else 12
                    )
                except Exception:
                    largura = len(str(col_nome)) + 3

                largura = min(max(largura, 12), 45)

                if col_nome in [
                    "TEMPO",
                    "TEMPO_TOTAL",
                    "TEMPO_MEDIO",
                    "TEMPO_MEDIO_POR_PLACA"
                ]:
                    ws.set_column(col_idx, col_idx, 20, fmt_tempo)

                elif col_nome in [
                    "PRIMEIRA_ENTRADA",
                    "ULTIMA_SAIDA",
                    col_entrada,
                    col_saida,
                    "DATA_SAIDA_ORIGINAL"
                ]:
                    ws.set_column(col_idx, col_idx, 22, fmt_data)

                elif col_nome == "DATA_ENTRADA":
                    ws.set_column(col_idx, col_idx, 14, fmt_data_curta)

                elif (
                    "QUANTIDADE" in nome_norm
                    or "QTD" in nome_norm
                    or "ID" in nome_norm
                    or "HORA" in nome_norm
                    or "TEMPO_SEG" in nome_norm
                ):
                    ws.set_column(col_idx, col_idx, largura, fmt_numero)

                else:
                    ws.set_column(col_idx, col_idx, largura, fmt_texto)

        ws_proj = wb.add_worksheet("RESUMO_PROJ_HORA")
        writer.sheets["RESUMO_PROJ_HORA"] = ws_proj
        ws_proj.hide_gridlines(2)

        linha_atual = 0

        if len(locais_proj) == 0:
            ws_proj.write(linha_atual, 0, "Nenhum local PROJ encontrado.", fmt_header)
        else:
            for local in locais_proj:
                base_local = dados_proj_longo[dados_proj_longo[col_area] == local].copy()
                datas = sorted(base_local["DATA"].dropna().unique())

                if len(datas) == 0:
                    continue

                ws_proj.merge_range(linha_atual, 0, linha_atual, 25, local, fmt_titulo)
                linha_atual += 1

                ws_proj.write(linha_atual, 0, "DATA", fmt_header)
                ws_proj.merge_range(linha_atual, 1, linha_atual, 24, "JANELA", fmt_header_amarelo)
                ws_proj.write(linha_atual, 25, "TOTAL", fmt_header)
                linha_atual += 1

                ws_proj.write(linha_atual, 0, "", fmt_header)
                for h in range(24):
                    ws_proj.write(linha_atual, h + 1, h, fmt_header)
                ws_proj.write(linha_atual, 25, "TOTAL", fmt_header)
                linha_atual += 1

                inicio_dados = linha_atual

                for data in datas:
                    ws_proj.write_datetime(
                        linha_atual,
                        0,
                        pd.Timestamp(data).to_pydatetime(),
                        fmt_data_curta
                    )

                    total_dia = 0

                    for h in range(24):
                        filtro = base_local[
                            (base_local["DATA"] == data)
                            & (base_local["HORA"] == h)
                        ]

                        qtd = int(filtro["QUANTIDADE_VEICULOS"].sum()) if not filtro.empty else 0
                        total_dia += qtd

                        if qtd >= 5:
                            formato = fmt_alto
                        elif qtd > 0:
                            formato = fmt_baixo
                        else:
                            formato = fmt_zero

                        ws_proj.write(linha_atual, h + 1, qtd, formato)

                    ws_proj.write(linha_atual, 25, total_dia, fmt_numero)
                    linha_atual += 1

                fim_dados = linha_atual - 1

                linha_atual += 1

                ws_proj.write(linha_atual, 0, "MEDIA QTD", fmt_header_amarelo)

                for h in range(24):
                    col_excel = h + 1
                    formula = f"=AVERAGE({chr(65 + col_excel)}{inicio_dados + 1}:{chr(65 + col_excel)}{fim_dados + 1})"
                    ws_proj.write_formula(linha_atual, col_excel, formula, fmt_media)

                ws_proj.write(linha_atual, 25, "", fmt_header_amarelo)
                linha_atual += 2

                ws_proj.write(linha_atual, 0, "TEMPO TOTAL", fmt_header_amarelo)

                total_tempo_geral = 0

                for h in range(24):
                    tempo_total_hora = base_local[base_local["HORA"] == h]["TEMPO_TOTAL_SEG"].sum()
                    total_tempo_geral += tempo_total_hora
                    ws_proj.write(linha_atual, h + 1, segundos_excel(tempo_total_hora), fmt_tempo_vertical)

                ws_proj.write(linha_atual, 25, segundos_excel(total_tempo_geral), fmt_tempo)
                linha_atual += 1

                ws_proj.write(linha_atual, 0, "TEMPO MÉDIO", fmt_header_amarelo)

                for h in range(24):
                    filtro_hora = base_local[base_local["HORA"] == h]

                    if not filtro_hora.empty:
                        tempo_medio_hora = (
                            filtro_hora["TEMPO_TOTAL_SEG"].sum()
                            / filtro_hora["QUANTIDADE_CHEGADAS"].sum()
                        )
                    else:
                        tempo_medio_hora = 0

                    ws_proj.write(linha_atual, h + 1, segundos_excel(tempo_medio_hora), fmt_tempo_vertical)

                total_chegadas = base_local["QUANTIDADE_CHEGADAS"].sum()
                total_tempo = base_local["TEMPO_TOTAL_SEG"].sum()

                tempo_medio_geral = total_tempo / total_chegadas if total_chegadas > 0 else 0
                ws_proj.write(linha_atual, 25, segundos_excel(tempo_medio_geral), fmt_tempo)
                linha_atual += 3

            ws_proj.set_column(0, 0, 16)
            ws_proj.set_column(1, 24, 4)
            ws_proj.set_column(25, 25, 13)
            ws_proj.freeze_panes(2, 1)

    info = {
        "linha_cabecalho": linha + 1,
        "col_placa": col_placa,
        "col_area": col_area,
        "col_entrada": col_entrada,
        "col_saida": col_saida,
        "linhas_original": len(df_original),
        "linhas_tratadas": len(tratado),
        "linhas_validas": len(base_valida),
        "locais_proj": len(locais_proj),
    }

    return caminho_saida, tratado, resumo_placa_local, resumo_local, info


def main_streamlit():
    st.set_page_config(page_title="Permanencia Area", page_icon="📍", layout="wide")

    st.title("📍 Resumo Permanência por Área")

    st.info("Envie o Excel de permanência. O sistema vai tratar ruídos, juntar eventos próximos e gerar o arquivo final.")

    arquivo = st.file_uploader(
        "Selecione o arquivo Excel",
        type=["xlsx", "xls", "xlsm"],
        key="upload_resumo_permanencia_area"
    )

    if not arquivo:
        st.warning("Aguardando upload do arquivo.")
        return

    st.success(f"Arquivo carregado: {arquivo.name}")

    if st.button("🚀 Processar Resumo Permanência por Área", use_container_width=True):

        inicio = datetime.now()

        barra = st.progress(0)
        status = st.empty()

        try:
            status.info("10% - Salvando arquivo temporário...")
            barra.progress(10)

            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(arquivo.name).suffix) as tmp:
                tmp.write(arquivo.read())
                entrada = tmp.name

            saida = str(
                Path(tempfile.gettempdir()) /
                f"{Path(arquivo.name).stem}_Base_Tratada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )

            status.info("35% - Processando dados...")
            barra.progress(35)

            caminho_saida, tratado, resumo_placa_local, resumo_local, info = processar_arquivo(
                entrada,
                saida
            )

            status.info("100% - Finalizado.")
            barra.progress(100)

            tempo_total = (datetime.now() - inicio).total_seconds()

            st.success(f"✅ Arquivo processado em {tempo_total:.1f} segundos.")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Linha cabeçalho", info["linha_cabecalho"])
            c2.metric("Linhas originais", info["linhas_original"])
            c3.metric("Linhas tratadas", info["linhas_tratadas"])
            c4.metric("Linhas válidas", info["linhas_validas"])

            st.markdown("### Colunas identificadas")
            st.write({
                "Placa": info["col_placa"],
                "Área": info["col_area"],
                "Entrada": info["col_entrada"],
                "Saída": info["col_saida"],
            })

            st.markdown("### Prévia da base tratada")
            st.dataframe(tratado.head(100), use_container_width=True)

            st.markdown("### Resumo por placa/local")
            st.dataframe(resumo_placa_local.head(100), use_container_width=True)

            st.markdown("### Resumo por local")
            st.dataframe(resumo_local.head(100), use_container_width=True)

            with open(caminho_saida, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel tratado",
                    f,
                    file_name=Path(caminho_saida).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Erro ao processar: {e}")




def main():
    """Entrada padrao para o carregador do site e para o Streamlit."""
    return main_streamlit()


def app():
    """Alias opcional para plataformas que procuram funcao app()."""
    return main_streamlit()


if __name__ == "__main__":
    main()
