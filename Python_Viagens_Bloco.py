import sys
import time
import threading
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


# =========================================================
# CONFIGURACAO GERAL
# =========================================================
EMPRESA = "Expresso Nepomuceno"
MAXTRACK_SHEET = "RL - Viagens"
SAP_SHEET = 0
PERMANENCIA_HEADER_ROW = 3

START = "17 POSTO DE COMBUSTIVEL FABRICA ARACRUZ"

FILIALS = [
    "ARA - ENTRADA DA FILIAL 20 KM/H",
    "FILIAL-ARACRUZ (ARA)",
]

FIELD_TOKENS = [
    "ARA - OFFROAD",
    "OFFROAD 40 KM/H",
]

RETURN_PRIMARY = [
    "ARA - ENTRADA FABRICA NORTE 20 KM/H",
]

RETURN_CONFIRM = [
    "ARA - DESCARGA DA FABRICA 20 KM/H",
    "DESCARREGAMENTO ARACRUZ",
    "PATIO DE DESCARGA SUZANO ARACRUZ",
    "PÁTIO DE DESCARGA SUZANO ARACRUZ",
    "DESAMARRACAO - ARACRUZ",
    "DESAMARRAÇÃO - ARACRUZ",
    "08 DESAMARRACAO ARACRUZ",
    "08 DESAMARRAÇÃO ARACRUZ",
    "PONTO DE VARRICAO PATIO ARACRUZ",
    "PONTO DE VARRIÇÃO PÁTIO ARACRUZ",
    "FABRICA-BALANCA DE ENTRADA (ARA)",
    "FABRICA-BALANÇA DE ENTRADA (ARA)",
    "24 FILA BALANCA DE ENTRADA ARACRUZ",
    "24 FILA BALANÇA DE ENTRADA ARACRUZ",
    "FABRICA-DESCARGA (ARA)",
    "FABRICA-ARACRUZ (ARA)",
]

FACTORY_INTERNAL = [
    "01 FABRICA SUZANO ARACRUZ",
    "FABRICA ARACRUZ 1",
    "FÁBRICA ARACRUZ 1",
    "FABRICA ARACRUZ",
    "FÁBRICA ARACRUZ",
    "FABRICA AR - 141796",
    "FABRICA AR - 141797",
    "FABRICA AR - 141801",
    "FABRICA AR - 141805",
    "FABRICA FIBRIA - 141783",
    "FABRICA FIBRIA - 141754",
    "FABRICA FIBRIA - 141761",
    "FABRICA FIBRIA - 141752",
    "TPF",
    "ABA-FABRICA (ARA)",
    "FABRICA-ARACRUZ (ARA)",
    "FABRICA-BALANCA DE ENTRADA (ARA)",
    "FABRICA-BALANÇA DE ENTRADA (ARA)",
]

FACTORY_EXIT = [
    "25 FILA BALANCA SAIDA ARA",
    "25 FILA BALANÇA SAÍDA ARA",
    "05 BALANCA DE SAIDA ARACRUZ",
    "05 BALANÇA DE SAÍDA ARACRUZ",
    "FILA BALANCA SAIDA",
    "FILA BALANÇA SAÍDA",
]

DISTANCIA_MINIMA_NOVO_INICIO = 5.0
TOLERANCIA_HORAS_SAP = 24
TOLERANCIA_HORAS_PERMANENCIA = 12


# =========================================================
# INTERFACE DE PROGRESSO
# =========================================================
class ProgressUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{EMPRESA} - Processamento de Viagens")
        self.root.geometry("760x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#F4F7FA")

        self.start_time = time.time()
        self.total_steps = 100
        self.current_step = 0
        self.done = False

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "EN.Horizontal.TProgressbar",
            troughcolor="#E5EAF0",
            background="#00A651",
            bordercolor="#E5EAF0",
            lightcolor="#00A651",
            darkcolor="#008A44",
        )

        self.header = tk.Frame(self.root, bg="#003B5C", height=58)
        self.header.pack(fill="x")

        self.label_empresa = tk.Label(
            self.header,
            text=EMPRESA,
            bg="#003B5C",
            fg="white",
            font=("Segoe UI", 16, "bold")
        )
        self.label_empresa.pack(side="left", padx=22, pady=14)

        self.label_modulo = tk.Label(
            self.header,
            text="Análise Maxtrack + SAP + Permanência em Área",
            bg="#003B5C",
            fg="#D6EAF8",
            font=("Segoe UI", 10)
        )
        self.label_modulo.pack(side="right", padx=22, pady=18)

        self.body = tk.Frame(self.root, bg="#F4F7FA")
        self.body.pack(fill="both", expand=True, padx=26, pady=18)

        self.label_title = tk.Label(
            self.body,
            text="Preparando processamento...",
            bg="#F4F7FA",
            fg="#1F2D3D",
            font=("Segoe UI", 14, "bold")
        )
        self.label_title.pack(pady=(4, 8))

        self.label_status = tk.Label(
            self.body,
            text="",
            bg="#F4F7FA",
            fg="#34495E",
            font=("Segoe UI", 10),
            wraplength=690,
            justify="center"
        )
        self.label_status.pack(pady=(0, 10))

        self.progress = ttk.Progressbar(
            self.body,
            orient="horizontal",
            length=680,
            mode="determinate",
            style="EN.Horizontal.TProgressbar"
        )
        self.progress.pack(pady=8)

        self.label_percent = tk.Label(
            self.body,
            text="0%",
            bg="#F4F7FA",
            fg="#00A651",
            font=("Segoe UI", 13, "bold")
        )
        self.label_percent.pack()

        self.label_time = tk.Label(
            self.body,
            text="Decorrido: 00:00 | Restante: calculando...",
            bg="#F4F7FA",
            fg="#2C3E50",
            font=("Segoe UI", 10, "bold")
        )
        self.label_time.pack(pady=(10, 4))

        self.label_detail = tk.Label(
            self.body,
            text="",
            bg="#F4F7FA",
            fg="#7F8C8D",
            font=("Segoe UI", 9)
        )
        self.label_detail.pack()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        if self.done:
            self.root.destroy()

    def set_total(self, total):
        self.total_steps = max(int(total), 1)
        self.progress["maximum"] = self.total_steps

    def update(self, step=None, status=None, detail=None):
        if step is not None:
            self.current_step = max(0, min(int(step), self.total_steps))

        if status:
            self.label_title.config(text=status)

        if detail:
            self.label_status.config(text=detail)
            self.label_detail.config(text=detail)

        self.progress["value"] = self.current_step

        pct = int((self.current_step / self.total_steps) * 100)
        self.label_percent.config(text=f"{pct}%")

        elapsed = time.time() - self.start_time
        remaining = None
        if self.current_step > 0:
            remaining = (elapsed / self.current_step) * (self.total_steps - self.current_step)

        self.label_time.config(
            text=f"Decorrido: {format_seconds(elapsed)} | Restante: {format_seconds(remaining) if remaining is not None else 'calculando...'}"
        )

        self.root.update_idletasks()

    def finish(self, message):
        self.done = True
        self.update(step=self.total_steps, status="Processamento concluído", detail=message)
        messagebox.showinfo("Concluído", message)
        self.root.destroy()

    def fail(self, message):
        self.done = True
        self.update(status="Erro no processamento", detail=message)
        messagebox.showerror("Erro", message)
        self.root.destroy()


def format_seconds(seconds):
    if seconds is None:
        return "calculando..."
    seconds = int(max(0, seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


# =========================================================
# FUNCOES BASE
# =========================================================
def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).upper().strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto


def obter_distancia(valor) -> float:
    if pd.isna(valor):
        return 0.0
    try:
        txt = str(valor).strip().replace(",", ".")
        if txt == "":
            return 0.0
        return float(txt)
    except Exception:
        return 0.0


def combinar_data_hora(col_data, col_hora):
    data = pd.to_datetime(col_data, errors="coerce", dayfirst=True)
    hora_txt = col_hora.astype(str).str.strip().replace({"nan": None, "NaT": None, "": None})
    hora = pd.to_timedelta(hora_txt, errors="coerce")
    return data + hora


def escolher_arquivo(titulo: str):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    arquivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
    )
    root.destroy()
    return arquivo


# =========================================================
# MAXTRACK
# =========================================================
def validar_colunas_maxtrack(df: pd.DataFrame):
    obrigatorias = [
        "Identificador/Placa",
        "Frota",
        "Início",
        "Fim",
        "Área(s)",
        "Distância (Km)",
    ]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(f"Maxtrack sem colunas obrigatórias: {', '.join(faltando)}")


def marcar_area_maxtrack(area: str) -> dict:
    area_norm = normalizar_texto(area)

    flags = {
        "area_norm": area_norm,
        "start": normalizar_texto(START) in area_norm,
        "filial": any(normalizar_texto(tok) in area_norm for tok in FILIALS),
        "manut": "MANUTEN" in area_norm,
        "field": any(normalizar_texto(tok) in area_norm for tok in FIELD_TOKENS),
        "ret_primary": any(normalizar_texto(tok) in area_norm for tok in RETURN_PRIMARY),
        "ret_confirm": any(normalizar_texto(tok) in area_norm for tok in RETURN_CONFIRM),
        "factory_internal": any(normalizar_texto(tok) in area_norm for tok in FACTORY_INTERNAL),
        "factory_exit": any(normalizar_texto(tok) in area_norm for tok in FACTORY_EXIT),
    }

    flags["factory_related"] = any(
        flags[k] for k in ["start", "ret_primary", "ret_confirm", "factory_internal", "factory_exit"]
    )
    return flags


def preparar_maxtrack(caminho_maxtrack: str, progress=None):
    if progress:
        progress.update(5, "Lendo Maxtrack", Path(caminho_maxtrack).name)

    df = pd.read_excel(caminho_maxtrack, sheet_name=MAXTRACK_SHEET)
    validar_colunas_maxtrack(df)

    if progress:
        progress.update(10, "Tratando datas do Maxtrack", "Convertendo início, fim e distância")

    df["Início"] = pd.to_datetime(df["Início"], errors="coerce", dayfirst=True)
    df["Fim"] = pd.to_datetime(df["Fim"], errors="coerce", dayfirst=True)
    df["DIST_KM_NUM"] = df["Distância (Km)"].apply(obter_distancia)

    df = df.dropna(subset=["Identificador/Placa", "Início", "Fim"]).copy()
    df = df.sort_values(["Identificador/Placa", "Início", "Fim"]).reset_index(drop=True)

    if progress:
        progress.update(15, "Mapeando áreas do Maxtrack", "Classificando fábrica, campo, filial, manutenção e retorno")

    marc = df["Área(s)"].apply(marcar_area_maxtrack).apply(pd.Series)
    df = pd.concat([df, marc], axis=1)
    return df


# =========================================================
# SAP
# =========================================================
def preparar_sap(caminho_sap: str, progress=None):
    if progress:
        progress.update(20, "Lendo SAP", Path(caminho_sap).name)

    sap = pd.read_excel(caminho_sap, sheet_name=SAP_SHEET)

    obrigatorias = ["Placa", "DtSaídaFáb", "HrSaídaFab"]
    faltando = [c for c in obrigatorias if c not in sap.columns]
    if faltando:
        raise ValueError(f"SAP sem colunas obrigatórias: {', '.join(faltando)}")

    if progress:
        progress.update(25, "Tratando SAP", "Montando início oficial: DtSaídaFáb + HrSaídaFab")

    sap = sap.copy()
    sap["SAP_PLACA"] = sap["Placa"].astype(str).str.strip()
    sap["SAP_DT_INICIO_REAL"] = combinar_data_hora(sap["DtSaídaFáb"], sap["HrSaídaFab"])

    if "Saída" in sap.columns and "Hora Saída" in sap.columns:
        sap["SAP_DT_FIM_REAL"] = combinar_data_hora(sap["Saída"], sap["Hora Saída"])
    else:
        sap["SAP_DT_FIM_REAL"] = pd.NaT

    if "DtFimDesFb" in sap.columns and "HrFimDesFb" in sap.columns:
        sap["SAP_DT_FIM_DESCARGA"] = combinar_data_hora(sap["DtFimDesFb"], sap["HrFimDesFb"])
    else:
        sap["SAP_DT_FIM_DESCARGA"] = pd.NaT

    if "Data Entr" in sap.columns and "Hora Entr" in sap.columns:
        sap["SAP_DT_ENTRADA_FAB"] = combinar_data_hora(sap["Data Entr"], sap["Hora Entr"])
    else:
        sap["SAP_DT_ENTRADA_FAB"] = pd.NaT

    campos_extras = ["Movimento", "Nº Doc", "CMM/CEM", "CTRC/ACT", "UP", "Distância", "Município", "Material"]
    for c in campos_extras:
        if c not in sap.columns:
            sap[c] = ""

    sap = sap.dropna(subset=["SAP_DT_INICIO_REAL"]).copy()
    return sap


def existe_inicio_sap_proximo(sap, placa, dt_ref, horas=TOLERANCIA_HORAS_SAP):
    placa = str(placa).strip()
    janela_ini = dt_ref - pd.Timedelta(hours=horas)
    janela_fim = dt_ref + pd.Timedelta(hours=horas)

    cand = sap[
        (sap["SAP_PLACA"] == placa) &
        (sap["SAP_DT_INICIO_REAL"] >= janela_ini) &
        (sap["SAP_DT_INICIO_REAL"] <= janela_fim)
    ].copy()

    if cand.empty:
        return False, pd.NaT, ""

    cand["DIF"] = (cand["SAP_DT_INICIO_REAL"] - dt_ref).abs()
    melhor = cand.sort_values("DIF").iloc[0]
    desc = f"Mov:{melhor.get('Movimento', '')} Doc:{melhor.get('Nº Doc', '')} UP:{melhor.get('UP', '')}"
    return True, melhor["SAP_DT_INICIO_REAL"], desc


def existe_fim_sap_proximo(sap, placa, dt_ref, horas=TOLERANCIA_HORAS_SAP):
    placa = str(placa).strip()
    janela_ini = dt_ref - pd.Timedelta(hours=horas)
    janela_fim = dt_ref + pd.Timedelta(hours=horas)

    cand = sap[sap["SAP_PLACA"] == placa].copy()
    eventos = []

    for _, r in cand.iterrows():
        if pd.notna(r.get("SAP_DT_FIM_REAL")):
            eventos.append(("SAIDA_HORA_SAIDA", r["SAP_DT_FIM_REAL"]))
        if pd.notna(r.get("SAP_DT_FIM_DESCARGA")):
            eventos.append(("FIM_DESCARGA", r["SAP_DT_FIM_DESCARGA"]))
        if pd.notna(r.get("SAP_DT_ENTRADA_FAB")):
            eventos.append(("ENTRADA_FAB", r["SAP_DT_ENTRADA_FAB"]))

    if not eventos:
        return False, pd.NaT, ""

    base = pd.DataFrame(eventos, columns=["TIPO", "DATA"])
    base = base[(base["DATA"] >= janela_ini) & (base["DATA"] <= janela_fim)].copy()

    if base.empty:
        return False, pd.NaT, ""

    base["DIF"] = (base["DATA"] - dt_ref).abs()
    melhor = base.sort_values("DIF").iloc[0]
    return True, melhor["DATA"], melhor["TIPO"]


# =========================================================
# PERMANENCIA EM AREA
# =========================================================
def preparar_permanencia(caminho_perm: str, progress=None):
    if progress:
        progress.update(30, "Lendo Permanência em Área", Path(caminho_perm).name)

    perm = pd.read_excel(caminho_perm, sheet_name=0, header=PERMANENCIA_HEADER_ROW)

    obrigatorias = ["Placa", "Frota", "Serial", "Nome da Área", "Data de entrada", "Data de saída", "Tempo dentro da área"]
    faltando = [c for c in obrigatorias if c not in perm.columns]
    if faltando:
        raise ValueError(f"Permanência sem colunas obrigatórias: {', '.join(faltando)}")

    if progress:
        progress.update(35, "Tratando Permanência em Área", "Mapeando BCA, TT, ADV, PROJ e FABRICA-ARACRUZ")

    perm = perm.copy()
    perm["PERM_PLACA"] = perm["Placa"].astype(str).str.strip()
    perm["PERM_FROTA"] = perm["Frota"].astype(str).str.strip()
    perm["PERM_SERIAL"] = perm["Serial"].astype(str).str.strip()
    perm["PERM_AREA"] = perm["Nome da Área"].astype(str)
    perm["PERM_AREA_NORM"] = perm["PERM_AREA"].apply(normalizar_texto)
    perm["PERM_INICIO"] = pd.to_datetime(perm["Data de entrada"], errors="coerce", dayfirst=True)
    perm["PERM_FIM"] = pd.to_datetime(perm["Data de saída"], errors="coerce", dayfirst=True)

    perm["PERM_BCA"] = perm["PERM_AREA_NORM"].str.contains("BCA", na=False)
    perm["PERM_TT"] = perm["PERM_AREA_NORM"].str.contains("TT", na=False)
    perm["PERM_ADV"] = perm["PERM_AREA_NORM"].str.contains("ADV", na=False)
    perm["PERM_PROJ"] = perm["PERM_AREA_NORM"].str.contains("PROJ", na=False)
    perm["PERM_FABRICA_ARACRUZ"] = perm["PERM_AREA_NORM"].str.contains("FABRICA-ARACRUZ", na=False)
    perm["PERM_DESCARGA"] = perm["PERM_AREA_NORM"].str.contains("DESCARGA", na=False)
    perm["PERM_OFFROAD"] = perm["PERM_AREA_NORM"].str.contains("OFFROAD|BCA|PROJ|ADV|TT", regex=True, na=False)

    # Familias operacionais para score e deduplicacao
    perm["PERM_AREA_FABRICA"] = perm["PERM_AREA_NORM"].str.contains(
        "FABRICA|DESCARGA|BALANCA DE ENTRADA|ABA-FABRICA", regex=True, na=False
    )
    perm["PERM_AREA_CAMPO_REAL"] = perm["PERM_AREA_NORM"].str.contains(
        "OFFROAD|BCA|PROJ", regex=True, na=False
    )
    perm["PERM_AREA_SUPORTE"] = perm["PERM_AREA_NORM"].str.contains(
        "TT|ADV|PS|REF", regex=True, na=False
    )
    perm["PERM_AREA_MANUTENCAO_FILIAL"] = perm["PERM_AREA_NORM"].str.contains(
        "DM-|FILIAL", regex=True, na=False
    )

    perm = perm.dropna(subset=["PERM_INICIO", "PERM_FIM"]).copy()
    perm = perm.sort_values(["PERM_PLACA", "PERM_INICIO", "PERM_FIM"]).reset_index(drop=True)
    return perm


def obter_contexto_permanencia_viagem(perm, placa, inicio, fim):
    placa = str(placa).strip()

    if pd.isna(inicio) or pd.isna(fim):
        return {}

    janela_ini = inicio - pd.Timedelta(hours=TOLERANCIA_HORAS_PERMANENCIA)
    janela_fim = fim + pd.Timedelta(hours=TOLERANCIA_HORAS_PERMANENCIA)

    trecho = perm[
        (perm["PERM_PLACA"] == placa) &
        (perm["PERM_INICIO"] <= janela_fim) &
        (perm["PERM_FIM"] >= janela_ini)
    ].copy()

    dentro = trecho[(trecho["PERM_INICIO"] <= fim) & (trecho["PERM_FIM"] >= inicio)].copy()

    def any_col(col):
        return bool(dentro[col].any()) if not dentro.empty else False

    tem_bca = any_col("PERM_BCA")
    tem_tt = any_col("PERM_TT")
    tem_adv = any_col("PERM_ADV")
    tem_proj = any_col("PERM_PROJ")
    tem_fabrica = any_col("PERM_FABRICA_ARACRUZ")
    tem_descarga = any_col("PERM_DESCARGA")

    if tem_proj and tem_adv:
        class_campo = "CAMPO_CARREGAMENTO_COM_ADVERSIDADE"
    elif tem_proj:
        class_campo = "CAMPO_CARREGAMENTO"
    elif tem_adv:
        class_campo = "CAMPO_ADVERSIDADE"
    elif tem_tt:
        class_campo = "CAMPO_TROCA_TURNO"
    elif tem_bca:
        class_campo = "CAMPO_BCA"
    else:
        class_campo = "CAMPO_SEM_DETALHE"

    perm_campo = dentro[dentro["PERM_OFFROAD"]].copy()
    perm_bca = dentro[dentro["PERM_BCA"]].copy()
    perm_fabrica = dentro[dentro["PERM_FABRICA_ARACRUZ"] | dentro["PERM_DESCARGA"]].copy()

    areas = ""
    if not dentro.empty:
        areas = " | ".join(dentro["PERM_AREA"].dropna().astype(str).drop_duplicates().head(20).tolist())

    return {
        "PERM_TEM_BCA": tem_bca,
        "PERM_TEM_TT": tem_tt,
        "PERM_TEM_ADV": tem_adv,
        "PERM_TEM_PROJ": tem_proj,
        "PERM_TEM_FABRICA_ARACRUZ": tem_fabrica,
        "PERM_TEM_DESCARGA": tem_descarga,
        "CLASSIFICACAO_CAMPO": class_campo,
        "PERM_CAMPO_INICIO": perm_campo["PERM_INICIO"].min() if not perm_campo.empty else pd.NaT,
        "PERM_CAMPO_FIM": perm_campo["PERM_FIM"].max() if not perm_campo.empty else pd.NaT,
        "PERM_BCA_INICIO": perm_bca["PERM_INICIO"].min() if not perm_bca.empty else pd.NaT,
        "PERM_BCA_FIM": perm_bca["PERM_FIM"].max() if not perm_bca.empty else pd.NaT,
        "PERM_FABRICA_INICIO": perm_fabrica["PERM_INICIO"].min() if not perm_fabrica.empty else pd.NaT,
        "PERM_AREAS_RESUMO": areas,
    }


# =========================================================
# TRAVA DE SEQUENCIA CRONOLOGICA
# =========================================================
def sequencia_valida_valores(inicio, saida_fabrica, chegada_filial, saida_filial, chegada_campo, saida_campo, chegada_fabrica):
    datas = [inicio, saida_fabrica]

    if pd.notna(chegada_filial):
        datas.append(chegada_filial)

    if pd.notna(saida_filial):
        datas.append(saida_filial)

    datas.extend([chegada_campo, saida_campo, chegada_fabrica])
    datas = [d for d in datas if pd.notna(d)]

    if len(datas) < 2:
        return False

    return all(datas[i] <= datas[i + 1] for i in range(len(datas) - 1))


def validar_sequencia_dataframe(df):
    status = []
    motivo = []

    for _, row in df.iterrows():
        ok = sequencia_valida_valores(
            row.get("INICIO VIAGEM"),
            row.get("SAIDA FABRICA"),
            row.get("CHEGADA FILIAL / MANUTENCAO"),
            row.get("SAIDA FILIAL / MANUTENCAO"),
            row.get("CHEGADA CAMPO"),
            row.get("SAIDA CAMPO"),
            row.get("CHEGADA FABRICA FIM DE VIAGEM"),
        )

        ok_refinado = sequencia_valida_valores(
            row.get("INICIO VIAGEM"),
            row.get("SAIDA FABRICA"),
            row.get("CHEGADA FILIAL / MANUTENCAO"),
            row.get("SAIDA FILIAL / MANUTENCAO"),
            row.get("CHEGADA CAMPO REFINADA"),
            row.get("SAIDA CAMPO REFINADA"),
            row.get("CHEGADA FABRICA REFINADA"),
        )

        if ok and ok_refinado:
            status.append("VALIDA")
            motivo.append("")
        elif ok and not ok_refinado:
            status.append("VALIDA_ORIGINAL_REFINO_INVALIDO")
            motivo.append("Refino por permanencia gerou inversao cronologica; usar datas originais.")
        else:
            status.append("INVALIDA")
            motivo.append("Sequencia de datas invertida.")

    df["STATUS_SEQUENCIA"] = status
    df["MOTIVO_SEQUENCIA"] = motivo
    return df


# =========================================================
# DETECCAO DE VIAGENS
# =========================================================
def detectar_fechamento_factory_block(g, start_idx, data_minima):
    ultimo_fim_campo = pd.NaT
    n = start_idx

    while n < len(g):
        row = g.loc[n]

        if row["field"] and row["Fim"] >= data_minima:
            ultimo_fim_campo = row["Fim"]

        if row["factory_related"] and row["Início"] >= data_minima:
            bloco_inicio = row["Início"]
            bloco_has_primary = bool(row["ret_primary"])
            bloco_has_confirm = bool(row["ret_confirm"])

            n2 = n
            while n2 + 1 < len(g) and bool(g.loc[n2 + 1, "factory_related"]):
                n2 += 1
                prox = g.loc[n2]
                bloco_has_primary = bloco_has_primary or bool(prox["ret_primary"])
                bloco_has_confirm = bloco_has_confirm or bool(prox["ret_confirm"])

            if bloco_has_primary or bloco_has_confirm:
                return {
                    "saida_campo": ultimo_fim_campo,
                    "chegada_fabrica": bloco_inicio,
                    "motivo_fechamento": "ENTRADA_FABRICA_BLOCO" if bloco_has_primary else "DESCARGA_BLOCO",
                    "fim_bloco_idx": n2,
                }

            n = n2

        n += 1

    return None


def determinar_inicio_por_transicao(g, sap, placa, idx_fechamento):
    if idx_fechamento + 1 >= len(g):
        return None

    trecho = g.iloc[idx_fechamento + 1:].copy()
    teve_saida_fabrica = False
    distancia_acumulada = 0.0

    for idx in trecho.index:
        row = g.loc[idx]

        if row["factory_exit"]:
            teve_saida_fabrica = True

        distancia_acumulada += obter_distancia(row["Distância (Km)"])

        if row["start"]:
            return {"inicio_idx": idx, "inicio_dt": row["Início"], "motivo_inicio": "TRANSICAO_POSTO"}

        ok_sap, dt_sap, _ = existe_inicio_sap_proximo(sap, placa, row["Início"], horas=12)
        if ok_sap:
            return {"inicio_idx": idx, "inicio_dt": dt_sap, "motivo_inicio": "TRANSICAO_SAP_SAIDA_FAB"}

        if teve_saida_fabrica and distancia_acumulada >= DISTANCIA_MINIMA_NOVO_INICIO:
            return {"inicio_idx": idx, "inicio_dt": row["Início"], "motivo_inicio": "TRANSICAO_SAIDA_FAB_DESLOCAMENTO"}

    return None


def localizar_bloco_saida_fabrica(g, s):
    ultima_saida_fabrica = g.at[s, "Fim"]
    k = s

    while k + 1 < len(g):
        row = g.loc[k + 1]

        if row["filial"] or row["manut"] or row["field"]:
            break

        if row["factory_related"] or row["area_norm"] in ("", "-"):
            ultima_saida_fabrica = max(ultima_saida_fabrica, row["Fim"])
            k += 1
            continue

        break

    return ultima_saida_fabrica, k


def localizar_bloco_filial_manut(g, idx_inicio_proc, data_minima):
    chegada_filial = pd.NaT
    saida_filial = pd.NaT
    tipo_filial = ""
    achou = False
    ultimo_fim = pd.NaT

    m = idx_inicio_proc
    while m < len(g):
        row = g.loc[m]

        if row["field"] and row["Início"] >= data_minima:
            break

        if (row["filial"] or row["manut"]) and row["Início"] >= data_minima:
            if not achou:
                chegada_filial = row["Início"]
                tipo_filial = "FILIAL" if row["filial"] else "MANUTENCAO"
                achou = True

            if row["filial"]:
                tipo_filial = "FILIAL"

            ultimo_fim = row["Fim"]

        m += 1

    if achou:
        saida_filial = ultimo_fim

    return {"chegada_filial": chegada_filial, "saida_filial": saida_filial, "tipo_filial": tipo_filial, "idx_final": m, "achou": achou}


def localizar_chegada_campo(g, idx_inicio_proc, data_minima):
    p = idx_inicio_proc
    while p < len(g):
        row = g.loc[p]
        if row["field"] and row["Início"] >= data_minima:
            return {"idx": p, "dt": row["Início"]}
        p += 1
    return None


def extrair_viagens_com_transicao(g, sap):
    viagens = []
    g = g.sort_values(["Início", "Fim"]).copy().reset_index(drop=True)

    placa = str(g.iloc[0]["Identificador/Placa"]).strip()
    frota = g.iloc[0]["Frota"]

    fila_inicios = []
    for s in [i for i in g.index if bool(g.at[i, "start"])]:
        fila_inicios.append({"inicio_idx": s, "inicio_dt": g.at[s, "Início"], "motivo_inicio": "POSTO_DIRETO"})

    fila_inicios = sorted(fila_inicios, key=lambda x: x["inicio_dt"])
    usados = set()
    i = 0

    while i < len(fila_inicios):
        inicio_info = fila_inicios[i]
        s = inicio_info["inicio_idx"]

        if s in usados:
            i += 1
            continue

        usados.add(s)
        inicio_viagem = inicio_info["inicio_dt"]
        motivo_inicio = inicio_info["motivo_inicio"]

        saida_fabrica, k = localizar_bloco_saida_fabrica(g, s)
        idx_apos_saida = k + 1

        bloco_filial = localizar_bloco_filial_manut(g, idx_apos_saida, saida_fabrica)

        if bloco_filial["achou"]:
            idx_busca_campo = bloco_filial["idx_final"]
            data_minima_campo = bloco_filial["saida_filial"]
        else:
            idx_busca_campo = idx_apos_saida
            data_minima_campo = saida_fabrica

        campo = localizar_chegada_campo(g, idx_busca_campo, data_minima_campo)
        if campo is None:
            i += 1
            continue

        fechamento = detectar_fechamento_factory_block(g, campo["idx"], campo["dt"])
        if fechamento is None:
            i += 1
            continue

        if not sequencia_valida_valores(
            inicio_viagem,
            saida_fabrica,
            bloco_filial["chegada_filial"],
            bloco_filial["saida_filial"],
            campo["dt"],
            fechamento["saida_campo"],
            fechamento["chegada_fabrica"],
        ):
            i += 1
            continue

        viagens.append({
            "PLACA": placa,
            "FROTA": frota,
            "INICIO VIAGEM": inicio_viagem,
            "SAIDA FABRICA": saida_fabrica,
            "CHEGADA FILIAL / MANUTENCAO": bloco_filial["chegada_filial"],
            "SAIDA FILIAL / MANUTENCAO": bloco_filial["saida_filial"],
            "CHEGADA CAMPO": campo["dt"],
            "SAIDA CAMPO": fechamento["saida_campo"],
            "CHEGADA FABRICA FIM DE VIAGEM": fechamento["chegada_fabrica"],
            "TIPO FILIAL": bloco_filial["tipo_filial"],
            "MOTIVO FECHAMENTO MAXTRACK": fechamento["motivo_fechamento"],
            "MOTIVO INICIO MAXTRACK": motivo_inicio,
        })

        prox = determinar_inicio_por_transicao(g, sap, placa, fechamento["fim_bloco_idx"])
        if prox is not None:
            ja_existe = any(abs((prox["inicio_dt"] - x["inicio_dt"]).total_seconds()) < 60 for x in fila_inicios)
            if not ja_existe:
                fila_inicios.append(prox)
                fila_inicios = sorted(fila_inicios, key=lambda x: x["inicio_dt"])

        i += 1

    return viagens


def aplicar_validacao_sap(resumo, sap):
    saida = resumo.copy()

    ini_ok, ini_data, ini_desc = [], [], []
    fim_ok, fim_data, fim_tipo = [], [], []
    sap_movimento = []
    sap_doc = []
    sap_up = []
    decisao = []

    for _, row in saida.iterrows():
        ok_i, dt_i, desc_i = existe_inicio_sap_proximo(sap, row["PLACA"], row["INICIO VIAGEM"])
        ok_f, dt_f, tipo_f = existe_fim_sap_proximo(sap, row["PLACA"], row["CHEGADA FABRICA FIM DE VIAGEM"])

        mov = ""
        doc = ""
        up = ""
        if ok_i and pd.notna(dt_i):
            cand_sap = sap[
                (sap["SAP_PLACA"].astype(str).str.strip() == str(row["PLACA"]).strip()) &
                (sap["SAP_DT_INICIO_REAL"] == dt_i)
            ]
            if not cand_sap.empty:
                mov = str(cand_sap.iloc[0].get("Movimento", ""))
                doc = str(cand_sap.iloc[0].get("Nº Doc", ""))
                up = str(cand_sap.iloc[0].get("UP", ""))

        ini_ok.append(ok_i)
        ini_data.append(dt_i)
        ini_desc.append(desc_i)
        fim_ok.append(ok_f)
        fim_data.append(dt_f)
        fim_tipo.append(tipo_f)
        sap_movimento.append(mov)
        sap_doc.append(doc)
        sap_up.append(up)

        if ok_i and ok_f:
            decisao.append("VALIDADO_MAXTRACK_SAP")
        elif ok_i:
            decisao.append("INICIO_VALIDADO_SAP_FIM_PENDENTE")
        elif ok_f:
            decisao.append("FIM_VALIDADO_SAP_INICIO_PENDENTE")
        else:
            decisao.append("PENDENTE_SAP")

    saida["SAP_VALIDA_INICIO"] = ini_ok
    saida["SAP_DATA_INICIO_VALIDADA"] = ini_data
    saida["SAP_INFO_INICIO"] = ini_desc
    saida["SAP_MOVIMENTO"] = sap_movimento
    saida["SAP_DOC"] = sap_doc
    saida["SAP_UP"] = sap_up
    saida["SAP_VALIDA_FIM"] = fim_ok
    saida["SAP_DATA_FIM_VALIDADA"] = fim_data
    saida["SAP_TIPO_FIM_VALIDADO"] = fim_tipo
    saida["DECISAO_FINAL"] = decisao
    return saida


def aplicar_permanencia(resumo, perm):
    contextos = []

    for _, row in resumo.iterrows():
        contextos.append(obter_contexto_permanencia_viagem(
            perm,
            row["PLACA"],
            row["INICIO VIAGEM"],
            row["CHEGADA FABRICA FIM DE VIAGEM"]
        ))

    ctx_df = pd.DataFrame(contextos)
    saida = pd.concat([resumo.reset_index(drop=True), ctx_df.reset_index(drop=True)], axis=1)

    saida["CHEGADA CAMPO REFINADA"] = saida["CHEGADA CAMPO"]
    saida["SAIDA CAMPO REFINADA"] = saida["SAIDA CAMPO"]
    saida["CHEGADA FABRICA REFINADA"] = saida["CHEGADA FABRICA FIM DE VIAGEM"]

    mask_campo_ini = saida["PERM_CAMPO_INICIO"].notna()
    mask_campo_fim = saida["PERM_CAMPO_FIM"].notna()
    mask_fabrica = saida["PERM_FABRICA_INICIO"].notna()

    saida.loc[mask_campo_ini, "CHEGADA CAMPO REFINADA"] = saida.loc[mask_campo_ini, "PERM_CAMPO_INICIO"]
    saida.loc[mask_campo_fim, "SAIDA CAMPO REFINADA"] = saida.loc[mask_campo_fim, "PERM_CAMPO_FIM"]
    saida.loc[mask_fabrica, "CHEGADA FABRICA REFINADA"] = saida.loc[mask_fabrica, "PERM_FABRICA_INICIO"]

    saida = validar_sequencia_dataframe(saida)

    # Se o refino gerou inversao, restaura datas originais no resumo visual.
    mask_refino_ruim = saida["STATUS_SEQUENCIA"].eq("VALIDA_ORIGINAL_REFINO_INVALIDO")
    saida.loc[mask_refino_ruim, "CHEGADA CAMPO REFINADA"] = saida.loc[mask_refino_ruim, "CHEGADA CAMPO"]
    saida.loc[mask_refino_ruim, "SAIDA CAMPO REFINADA"] = saida.loc[mask_refino_ruim, "SAIDA CAMPO"]
    saida.loc[mask_refino_ruim, "CHEGADA FABRICA REFINADA"] = saida.loc[mask_refino_ruim, "CHEGADA FABRICA FIM DE VIAGEM"]
    saida = validar_sequencia_dataframe(saida)

    return saida


# =========================================================
# SCORE E DEDUPLICACAO INTELIGENTE
# =========================================================
def calcular_score_viagem(row):
    """
    Score não muda a lógica de detecção.
    Ele só decide qual viagem manter quando houver conflito/duplicidade.
    """
    score = 0

    if bool(row.get("SAP_VALIDA_INICIO", False)):
        score += 5
    if bool(row.get("SAP_VALIDA_FIM", False)):
        score += 4
    if pd.notna(row.get("CHEGADA CAMPO")):
        score += 4
    if bool(row.get("PERM_TEM_PROJ", False)):
        score += 3
    if bool(row.get("PERM_TEM_DESCARGA", False)) or bool(row.get("PERM_TEM_FABRICA_ARACRUZ", False)):
        score += 3
    if bool(row.get("PERM_TEM_BCA", False)):
        score += 2
    if bool(row.get("PERM_TEM_ADV", False)):
        score += 1
    if bool(row.get("PERM_TEM_TT", False)):
        score += 1

    motivo_inicio = str(row.get("MOTIVO INICIO MAXTRACK", ""))
    if motivo_inicio == "POSTO_DIRETO":
        score += 1
    if "TRANSICAO_SAP" in motivo_inicio:
        score += 3
    if "TRANSICAO_SAIDA_FAB_DESLOCAMENTO" in motivo_inicio:
        score += 2

    # penalidades leves para casos mais fracos
    if not bool(row.get("SAP_VALIDA_INICIO", False)) and motivo_inicio == "POSTO_DIRETO":
        score -= 2
    if str(row.get("CLASSIFICACAO_CAMPO", "")) == "CAMPO_SEM_DETALHE":
        score -= 1
    if str(row.get("STATUS_SEQUENCIA", "")) != "VALIDA":
        score -= 10

    return score


def chave_conflito_viagem(row):
    """
    Chave de conflito:
    - se houver SAP_MOVIMENTO, ele é a fonte forte;
    - se não houver, agrupa por placa + janela de início em 6h.
    """
    placa = str(row.get("PLACA", "")).strip()
    mov = str(row.get("SAP_MOVIMENTO", "")).strip()
    doc = str(row.get("SAP_DOC", "")).strip()

    if mov and mov.lower() not in ("nan", "none", "0"):
        return f"{placa}|SAP_MOV|{mov}"

    if doc and doc.lower() not in ("nan", "none", "0"):
        return f"{placa}|SAP_DOC|{doc}"

    dt = row.get("INICIO VIAGEM")
    if pd.isna(dt):
        return f"{placa}|SEM_DATA"

    dt = pd.Timestamp(dt)
    janela_6h = dt.floor("6h")
    return f"{placa}|JANELA|{janela_6h:%Y-%m-%d %H:%M}"


def deduplicar_viagens(resumo):
    """
    Mantém a lógica existente e apenas aplica uma camada posterior:
    quando há viagens concorrentes/duplicadas, mantém a de melhor score.
    As demais ficam na aba DUPLICADAS_DESCARTADAS.
    """
    if resumo.empty:
        resumo["SCORE_VIAGEM"] = []
        resumo["CHAVE_CONFLITO"] = []
        resumo["STATUS_DEDUP"] = []
        return resumo, resumo.copy()

    base = resumo.copy()
    base["SCORE_VIAGEM"] = base.apply(calcular_score_viagem, axis=1)
    base["CHAVE_CONFLITO"] = base.apply(chave_conflito_viagem, axis=1)
    base["STATUS_DEDUP"] = "MANTIDA"

    mantidas_idx = []
    descartadas_idx = []

    for chave, grupo in base.groupby("CHAVE_CONFLITO", dropna=False):
        if len(grupo) == 1:
            mantidas_idx.append(grupo.index[0])
            continue

        grupo_ord = grupo.sort_values(
            by=[
                "SCORE_VIAGEM",
                "SAP_VALIDA_INICIO",
                "SAP_VALIDA_FIM",
                "INICIO VIAGEM",
            ],
            ascending=[False, False, False, True],
        )

        manter = grupo_ord.index[0]
        mantidas_idx.append(manter)

        for idx in grupo_ord.index[1:]:
            descartadas_idx.append(idx)

    mantidas = base.loc[mantidas_idx].copy().sort_values(["PLACA", "INICIO VIAGEM"]).reset_index(drop=True)
    descartadas = base.loc[descartadas_idx].copy().sort_values(["PLACA", "INICIO VIAGEM"]).reset_index(drop=True)

    if not descartadas.empty:
        descartadas["STATUS_DEDUP"] = "DUPLICADA_DESCARTADA"
        descartadas["MOTIVO_DEDUP"] = "Conflito com outra viagem da mesma placa/SAP/janela; mantida a de maior SCORE_VIAGEM."

    mantidas["ID VIAGEM"] = [f"V{i:05d}" for i in range(1, len(mantidas) + 1)]
    return mantidas, descartadas


def montar_resumo_padrao(writer, resumo):
    workbook = writer.book
    ws = workbook.add_worksheet("RESUMO_PADRAO")

    fmt_titulo = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "align": "left", "valign": "vcenter"})
    fmt_sub = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "align": "center", "valign": "vcenter", "border": 1})
    fmt_dt = workbook.add_format({"num_format": "dd/mm/yyyy hh:mm:ss", "align": "center", "valign": "vcenter", "border": 1})
    fmt_td = workbook.add_format({"num_format": "[h]:mm:ss", "align": "center", "valign": "vcenter", "border": 1})
    fmt_blank = workbook.add_format({"border": 1})

    linha = 0
    resumo_visual = resumo[resumo["STATUS_SEQUENCIA"].eq("VALIDA")].copy()

    for _, row in resumo_visual.iterrows():
        ws.merge_range(linha, 0, linha, 2, f"{row['ID VIAGEM']} | {row['PLACA']} | Frota {row['FROTA']}", fmt_titulo)
        linha += 1

        tempo_campo_visual = row["SAIDA CAMPO REFINADA"] - row["CHEGADA CAMPO REFINADA"] if pd.notna(row["SAIDA CAMPO REFINADA"]) and pd.notna(row["CHEGADA CAMPO REFINADA"]) else row["TEMPO CAMPO"]
        tempo_retorno_visual = row["CHEGADA FABRICA REFINADA"] - row["SAIDA CAMPO REFINADA"] if pd.notna(row["CHEGADA FABRICA REFINADA"]) and pd.notna(row["SAIDA CAMPO REFINADA"]) else row["TEMPO SAIDA CAMPO -> CHEGADA FABRICA"]
        ciclo_visual = row["CHEGADA FABRICA REFINADA"] - row["INICIO VIAGEM"] if pd.notna(row["CHEGADA FABRICA REFINADA"]) else row["CICLO TOTAL"]

        blocos = [
            ("INICIO VIAGEM", row["INICIO VIAGEM"], "SAIDA FABRICA", row["SAIDA FABRICA"], row["TEMPO INICIO -> SAIDA FABRICA"]),
            ("SAIDA FABRICA", row["SAIDA FABRICA"], "CHEGADA FILIAL / MANUTENCAO", row["CHEGADA FILIAL / MANUTENCAO"], row["TEMPO SAIDA FABRICA -> CHEGADA FILIAL/MANUT"]),
            ("CHEGADA FILIAL / MANUTENCAO", row["CHEGADA FILIAL / MANUTENCAO"], "SAIDA FILIAL / MANUTENCAO", row["SAIDA FILIAL / MANUTENCAO"], row["TEMPO FILIAL / MANUTENCAO"]),
            ("SAIDA FILIAL / MANUTENCAO", row["SAIDA FILIAL / MANUTENCAO"], "CHEGADA CAMPO", row["CHEGADA CAMPO REFINADA"], row["TEMPO SAIDA FILIAL/MANUT -> CHEGADA CAMPO"]),
            ("CHEGADA CAMPO", row["CHEGADA CAMPO REFINADA"], "SAIDA CAMPO", row["SAIDA CAMPO REFINADA"], tempo_campo_visual),
            ("SAIDA CAMPO", row["SAIDA CAMPO REFINADA"], "CHEGADA FABRICA FIM DE VIAGEM", row["CHEGADA FABRICA REFINADA"], tempo_retorno_visual),
        ]

        for label_a, val_a, label_b, val_b, tempo in blocos:
            ws.write_row(linha, 0, [label_a, label_b, "TEMPO"], fmt_sub)
            linha += 1

            if pd.notna(val_a):
                ws.write_datetime(linha, 0, pd.Timestamp(val_a).to_pydatetime(), fmt_dt)
            else:
                ws.write_blank(linha, 0, None, fmt_blank)

            if pd.notna(val_b):
                ws.write_datetime(linha, 1, pd.Timestamp(val_b).to_pydatetime(), fmt_dt)
            else:
                ws.write_blank(linha, 1, None, fmt_blank)

            if pd.notna(tempo):
                ws.write_number(linha, 2, pd.Timedelta(tempo).total_seconds() / 86400, fmt_td)
            else:
                ws.write_blank(linha, 2, None, fmt_blank)

            linha += 2

        ws.write(linha, 0, "CICLO TOTAL", fmt_sub)
        ws.write_blank(linha, 1, None, fmt_blank)
        if pd.notna(ciclo_visual):
            ws.write_number(linha, 2, pd.Timedelta(ciclo_visual).total_seconds() / 86400, fmt_td)
        else:
            ws.write_blank(linha, 2, None, fmt_blank)

        linha += 3

    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 34)
    ws.set_column(2, 2, 18)
    ws.freeze_panes(1, 0)


def processar_arquivos(caminho_maxtrack, caminho_sap, caminho_perm, progress=None):
    if progress:
        progress.set_total(100)
        progress.update(1, "Iniciando processamento", "Preparando leitura dos arquivos")

    df_max = preparar_maxtrack(caminho_maxtrack, progress)
    sap = preparar_sap(caminho_sap, progress)
    perm = preparar_permanencia(caminho_perm, progress)

    placas = list(df_max["Identificador/Placa"].dropna().unique())
    total_placas = len(placas)

    if progress:
        progress.update(40, "Extraindo viagens", f"Processando {total_placas} placas")

    viagens = []
    base_step = 40
    end_step = 70
    step_range = end_step - base_step

    for idx, (placa, grupo) in enumerate(df_max.groupby("Identificador/Placa", sort=False), start=1):
        viagens.extend(extrair_viagens_com_transicao(grupo, sap))

        if progress:
            step = base_step + int((idx / max(total_placas, 1)) * step_range)
            progress.update(step, "Extraindo viagens", f"Placa {idx}/{total_placas}: {placa}")

    resumo = pd.DataFrame(viagens)
    if resumo.empty:
        raise ValueError("Nenhuma viagem foi identificada com a lógica atual.")

    if progress:
        progress.update(72, "Calculando tempos", "Montando resumo por viagem e validando sequência")

    resumo = resumo.sort_values(["PLACA", "INICIO VIAGEM"]).reset_index(drop=True)
    resumo.insert(0, "ID VIAGEM", [f"V{i:05d}" for i in range(1, len(resumo) + 1)])

    pares_tempo = [
        ("INICIO VIAGEM", "SAIDA FABRICA", "TEMPO INICIO -> SAIDA FABRICA"),
        ("SAIDA FABRICA", "CHEGADA FILIAL / MANUTENCAO", "TEMPO SAIDA FABRICA -> CHEGADA FILIAL/MANUT"),
        ("CHEGADA FILIAL / MANUTENCAO", "SAIDA FILIAL / MANUTENCAO", "TEMPO FILIAL / MANUTENCAO"),
        ("SAIDA FILIAL / MANUTENCAO", "CHEGADA CAMPO", "TEMPO SAIDA FILIAL/MANUT -> CHEGADA CAMPO"),
        ("CHEGADA CAMPO", "SAIDA CAMPO", "TEMPO CAMPO"),
        ("SAIDA CAMPO", "CHEGADA FABRICA FIM DE VIAGEM", "TEMPO SAIDA CAMPO -> CHEGADA FABRICA"),
        ("INICIO VIAGEM", "CHEGADA FABRICA FIM DE VIAGEM", "CICLO TOTAL"),
    ]

    for a, b, c in pares_tempo:
        resumo[c] = resumo[b] - resumo[a]

    resumo = validar_sequencia_dataframe(resumo)

    if progress:
        progress.update(76, "Validando com SAP", "Cruzando início e fim por placa e horário")

    resumo = aplicar_validacao_sap(resumo, sap)

    if progress:
        progress.update(82, "Cruzando permanência em área", "Aplicando BCA, TT, ADV, PROJ e FABRICA-ARACRUZ")

    resumo = aplicar_permanencia(resumo, perm)

    resumo["TEMPO CAMPO REFINADO"] = resumo["SAIDA CAMPO REFINADA"] - resumo["CHEGADA CAMPO REFINADA"]
    resumo["TEMPO RETORNO REFINADO"] = resumo["CHEGADA FABRICA REFINADA"] - resumo["SAIDA CAMPO REFINADA"]
    resumo["CICLO TOTAL REFINADO"] = resumo["CHEGADA FABRICA REFINADA"] - resumo["INICIO VIAGEM"]
    resumo = validar_sequencia_dataframe(resumo)

    # Camada de melhoria: score + deduplicação inteligente.
    # Não altera a detecção original; apenas escolhe a melhor viagem quando há duplicidade.
    resumo_antes_dedup = resumo.copy()
    resumo, duplicadas_descartadas = deduplicar_viagens(resumo)

    invalidas = resumo[~resumo["STATUS_SEQUENCIA"].eq("VALIDA")].copy()
    resumo_validas = resumo[resumo["STATUS_SEQUENCIA"].eq("VALIDA")].copy()

    parametros = pd.DataFrame({
        "BASE": [
            "MAXTRACK", "SAP", "PERMANENCIA", "PERMANENCIA", "PERMANENCIA", "PERMANENCIA", "PERMANENCIA", "TRAVA"
        ],
        "USO": [
            "Monta ciclo principal e sequência de eventos",
            "Valida início/fim oficial da viagem",
            "BCA = entrada/saída de offroad",
            "TT = troca de turno",
            "ADV = adversidade no offroad",
            "PROJ = carregamento no offroad",
            "FABRICA-ARACRUZ (ARA) = descarga/fábrica",
            "Impede resumo com datas fora de ordem cronológica",
            "Deduplicação por SAP/placa/janela usando SCORE_VIAGEM",
        ],
    })

    saida = Path(caminho_maxtrack).with_name(
        f"{Path(caminho_maxtrack).stem}_3_BASES_SCORE_DEDUP.xlsx"
    )

    if progress:
        progress.update(88, "Gerando Excel", "Criando abas, cores e formatação do resumo padrão")

    with pd.ExcelWriter(saida, engine="xlsxwriter", datetime_format="dd/mm/yyyy hh:mm:ss") as writer:
        resumo.to_excel(writer, sheet_name="RESUMO_VALIDADO", index=False)
        resumo_validas.to_excel(writer, sheet_name="RESUMO_VALIDAS", index=False)
        invalidas.to_excel(writer, sheet_name="VIAGENS_INVALIDAS", index=False)
        duplicadas_descartadas.to_excel(writer, sheet_name="DUPLICADAS_DESCARTADAS", index=False)
        resumo_antes_dedup.to_excel(writer, sheet_name="ANTES_DEDUP_AUDITORIA", index=False)
        resumo.to_excel(writer, sheet_name="BLOCOS_AUXILIARES", index=False)
        sap.to_excel(writer, sheet_name="SAP_TRATADO", index=False)
        perm.to_excel(writer, sheet_name="PERMANENCIA_TRATADA", index=False)
        parametros.to_excel(writer, sheet_name="PARAMETROS", index=False)

        montar_resumo_padrao(writer, resumo)

        workbook = writer.book
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "align": "center"})
        fmt_dt = workbook.add_format({"num_format": "dd/mm/yyyy hh:mm:ss"})
        fmt_td = workbook.add_format({"num_format": "[h]:mm:ss"})

        abas = {
            "RESUMO_VALIDADO": resumo,
            "RESUMO_VALIDAS": resumo_validas,
            "VIAGENS_INVALIDAS": invalidas,
            "DUPLICADAS_DESCARTADAS": duplicadas_descartadas,
            "ANTES_DEDUP_AUDITORIA": resumo_antes_dedup,
            "BLOCOS_AUXILIARES": resumo,
            "SAP_TRATADO": sap,
            "PERMANENCIA_TRATADA": perm,
            "PARAMETROS": parametros,
        }

        dt_cols = [
            "INICIO VIAGEM", "SAIDA FABRICA",
            "CHEGADA FILIAL / MANUTENCAO", "SAIDA FILIAL / MANUTENCAO",
            "CHEGADA CAMPO", "SAIDA CAMPO", "CHEGADA FABRICA FIM DE VIAGEM",
            "SAP_DATA_INICIO_VALIDADA", "SAP_DATA_FIM_VALIDADA",
            "SAP_DT_INICIO_REAL", "SAP_DT_FIM_REAL", "SAP_DT_FIM_DESCARGA", "SAP_DT_ENTRADA_FAB",
            "PERM_INICIO", "PERM_FIM",
            "PERM_CAMPO_INICIO", "PERM_CAMPO_FIM", "PERM_BCA_INICIO", "PERM_BCA_FIM",
            "PERM_FABRICA_INICIO",
            "CHEGADA CAMPO REFINADA", "SAIDA CAMPO REFINADA", "CHEGADA FABRICA REFINADA",
        ]

        td_cols = [c for _, _, c in pares_tempo] + [
            "TEMPO CAMPO REFINADO", "TEMPO RETORNO REFINADO", "CICLO TOTAL REFINADO"
        ]

        for aba, frame in abas.items():
            ws = writer.sheets[aba]
            for col_idx, col in enumerate(frame.columns):
                ws.write(0, col_idx, col, fmt_header)
                width = min(max(len(str(col)) + 2, 14), 42)
                ws.set_column(col_idx, col_idx, width)

                if col in dt_cols:
                    ws.set_column(col_idx, col_idx, 20, fmt_dt)
                elif col in td_cols:
                    ws.set_column(col_idx, col_idx, 18, fmt_td)

            ws.freeze_panes(1, 0)
            if len(frame.columns) > 0:
                ws.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))

    if progress:
        progress.update(100, "Finalizando", "Arquivo gerado com sucesso")

    return saida, len(resumo), resumo["PLACA"].nunique(), len(invalidas), len(duplicadas_descartadas), len(resumo_antes_dedup)


def run_processing(progress, arquivo_max, arquivo_sap, arquivo_perm):
    try:
        saida, total_viagens, total_placas, total_invalidas, total_duplicadas, total_antes_dedup = processar_arquivos(
            arquivo_max, arquivo_sap, arquivo_perm, progress
        )

        msg = (
            f"Processamento concluído.\n\n"
            f"Arquivo gerado:\n{saida}\n\n"
            f"Viagens antes da deduplicação: {total_antes_dedup}\n"
            f"Viagens finais mantidas: {total_viagens}\n"
            f"Viagens válidas no resumo padrão: {total_viagens - total_invalidas}\n"
            f"Viagens inválidas separadas: {total_invalidas}\n"
            f"Duplicadas descartadas: {total_duplicadas}\n"
            f"Placas: {total_placas}"
        )
        progress.finish(msg)

    except Exception as e:
        progress.fail(f"Erro ao processar arquivos:\n{e}")


def main():
    arquivo_max = escolher_arquivo("Escolha o arquivo MAXTRACK")
    if not arquivo_max:
        return

    arquivo_sap = escolher_arquivo("Escolha o arquivo SAP")
    if not arquivo_sap:
        return

    arquivo_perm = escolher_arquivo("Escolha o arquivo PERMANÊNCIA EM ÁREA")
    if not arquivo_perm:
        return

    progress = ProgressUI()

    thread = threading.Thread(
        target=run_processing,
        args=(progress, arquivo_max, arquivo_sap, arquivo_perm),
        daemon=True
    )
    thread.start()

    progress.root.mainloop()


if __name__ == "__main__":
    main()
