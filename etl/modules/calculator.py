# =============================================================================
# modules/calculator.py — Cálculo de KPIs logísticos
# =============================================================================
# REGRAS (conforme prompt):
#
# ENTREGA (agrupamento):
#   Mesma Data Coleta + Mesmo Cliente + Mesma Transportadora + Mesma Cidade
#   = 1 Entrega
#
# STATUS DE PRAZO:
#   • Data Efetiva da Entrega > Prev. Entrega Cliente  → ATRASO
#   • Data Efetiva da Entrega ≤ Prev. Entrega Cliente  → NO PRAZO
#   • Data Efetiva da Entrega vazia (NaT)              → NO PRAZO
#     (será reclassificado automaticamente quando a data for preenchida)
#
# COLUNAS CALCULADAS PELO EXCEL SÃO IGNORADAS.
# =============================================================================

import pandas as pd
import numpy as np
import logging

logger = logging.getLogger(__name__)

# Colunas que definem 1 entrega única
CHAVE_ENTREGA = ["data_coleta", "cliente", "transportadora_norm", "municipio"]


def calcular_status_prazo(row: pd.Series) -> str:
    """
    Retorna 'NO PRAZO' ou 'ATRASO' para cada linha.

    Regra:
      - Se Data Efetiva vazia  → 'NO PRAZO'
      - Se Data Efetiva > Prev. Entrega Cliente → 'ATRASO'
      - Caso contrário → 'NO PRAZO'
    """
    data_efetiva  = row.get("data_efetiva_entrega")
    prev_cliente  = row.get("prev_entrega_cliente")

    # Campo vazio = ainda não entregue = conta como NO PRAZO
    if pd.isna(data_efetiva):
        return "NO PRAZO"

    # Prev. Entrega Cliente ausente: não dá para avaliar → NO PRAZO
    if pd.isna(prev_cliente):
        return "NO PRAZO"

    if data_efetiva > prev_cliente:
        return "ATRASO"

    return "NO PRAZO"


def calcular_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adiciona ao DataFrame as colunas calculadas pelo Python:

    Colunas adicionadas
    -------------------
    status_prazo        : 'NO PRAZO' | 'ATRASO'
    entrega_no_prazo    : 1 se NO PRAZO, 0 se ATRASO
    entrega_em_atraso   : 1 se ATRASO, 0 se NO PRAZO
    id_entrega          : hash da chave de agrupamento (Data+Cliente+Transp+Cidade)
    total_impostos_calc : ICMS + Difal + FCP  (recalculado independente do Excel)
    custo_total_envio   : Frete + Custo Embalagem
    """
    df = df.copy()

    # ------------------------------------------------------------------
    # 1. Status de prazo por linha
    # ------------------------------------------------------------------
    df["status_prazo"]     = df.apply(calcular_status_prazo, axis=1)
    df["entrega_no_prazo"] = (df["status_prazo"] == "NO PRAZO").astype(int)
    df["entrega_em_atraso"]= (df["status_prazo"] == "ATRASO").astype(int)

    logger.info(
        f"Status prazo calculado — "
        f"NO PRAZO: {df['entrega_no_prazo'].sum()} | "
        f"ATRASO: {df['entrega_em_atraso'].sum()}"
    )

    # ------------------------------------------------------------------
    # 2. ID de entrega — chave composta para agrupamento
    #    Cada combinação única = 1 entrega física
    # ------------------------------------------------------------------
    # Garante que data_coleta seja string para concatenação segura
    df["_data_str"] = df["data_coleta"].dt.strftime("%Y-%m-%d").fillna("9999-01-01")

    # Garantir que colunas duplicadas retornem Series (não DataFrame)
    def _col(df, c):
        s = df[c]
        return s.iloc[:, 0] if hasattr(s, "iloc") and s.ndim == 2 else s

    df["id_entrega"] = (
        _col(df, "_data_str").str.strip()
        + "|" + _col(df, "cliente").str.strip().str.upper()
        + "|" + _col(df, "transportadora_norm").str.strip().str.upper()
        + "|" + _col(df, "municipio").str.strip().str.upper()
    )

    df.drop(columns=["_data_str"], inplace=True)

    # ------------------------------------------------------------------
    # 3. Indicador booleano: é a primeira linha desse id_entrega?
    #    Usado para contar entregas sem duplicar (soma apenas o primeiro)
    # ------------------------------------------------------------------
    df["is_primeira_linha_entrega"] = (
        ~df.duplicated(subset=["id_entrega"], keep="first")
    ).astype(int)

    total_entregas_unicas = df["is_primeira_linha_entrega"].sum()
    logger.info(f"Entregas únicas identificadas: {total_entregas_unicas}")

    # ------------------------------------------------------------------
    # 4. Impostos — recálculo independente do Excel
    # ------------------------------------------------------------------
    for col in ["icms", "difal", "fcp"]:
        if col not in df.columns:
            df[col] = 0.0

    df["total_impostos_calc"] = df["icms"] + df["difal"] + df["fcp"]

    # ------------------------------------------------------------------
    # 5. Custo total de envio por linha
    # ------------------------------------------------------------------
    for col in ["frete", "custo_embalagem"]:
        if col not in df.columns:
            df[col] = 0.0

    df["custo_total_envio"] = df["frete"] + df["custo_embalagem"]

    # ------------------------------------------------------------------
    # 6. Ano e Mês extraídos da data real de coleta (não do Excel)
    # ------------------------------------------------------------------
    df["ano"]  = df["data_coleta"].dt.year
    df["mes"]  = df["data_coleta"].dt.month

    logger.info("Cálculo de KPIs concluído.")
    return df


def resumo_entregas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna um DataFrame com 1 linha por entrega única, consolidando:
    - Status de prazo (voto majoritário das linhas do grupo)
    - Totais financeiros do grupo
    - Volumes e peso total

    Útil para análises agregadas e exports resumidos.
    """
    # Agrega por id_entrega
    agg = df.groupby("id_entrega", as_index=False).agg(
        data_coleta           = ("data_coleta",           "first"),
        empresa               = ("empresa",                "first"),
        cliente               = ("cliente",                "first"),
        municipio             = ("municipio",              "first"),
        uf                    = ("uf",                     "first"),
        regiao                = ("regiao",                 "first"),
        tipo_cliente          = ("tipo_cliente",           "first"),
        natureza_operacao     = ("natureza_operacao",      "first"),
        transportadora        = ("transportadora_norm",    "first"),
        modalidade            = ("modalidade",             "first"),
        prev_entrega_cliente  = ("prev_entrega_cliente",   "first"),
        data_efetiva_entrega  = ("data_efetiva_entrega",   "first"),
        status_prazo          = ("status_prazo",           "first"),
        entrega_no_prazo      = ("entrega_no_prazo",       "first"),
        entrega_em_atraso     = ("entrega_em_atraso",      "first"),
        nr_notas              = ("nf",                     "count"),
        vr_nf_total           = ("vr_nf",                  "sum"),
        frete_total           = ("frete",                  "sum"),
        peso_total_kg         = ("peso_kg",                "sum"),
        volumes_total         = ("volumes",                "sum"),
        icms_total            = ("icms",                   "sum"),
        difal_total           = ("difal",                  "sum"),
        fcp_total             = ("fcp",                    "sum"),
        total_impostos_calc   = ("total_impostos_calc",    "sum"),
        custo_embalagem_total = ("custo_embalagem",        "sum"),
        custo_total_envio     = ("custo_total_envio",      "sum"),
        ano                   = ("ano",                    "first"),
        mes                   = ("mes",                    "first"),
    )

    logger.info(f"Resumo de entregas gerado: {len(agg)} linhas únicas.")
    return agg
