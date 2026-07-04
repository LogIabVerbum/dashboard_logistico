# =============================================================================
# modules/exporter.py — Exportação dos dados processados
# =============================================================================
# Gera dois formatos de saída consumidos pelo dashboard:
#   1. Parquet — dados completos (eficiente para grandes volumes)
#   2. JSON    — dados completos (compatível com JavaScript/HTML puro)
#
# Além disso, gera um JSON de métricas agregadas pré-calculadas
# para acelerar a carga inicial do dashboard.
# =============================================================================

import os
import json
import logging
from datetime import datetime, date

import pandas as pd
import numpy as np

from config import OUTPUT_PARQUET, OUTPUT_JSON, PROCESSED_DIR

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Serialização segura para JSON (lida com datetime, NaT, NaN, numpy types)
# ---------------------------------------------------------------------------

class _SafeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (pd.Timestamp, datetime, date)):
            return obj.isoformat() if not pd.isna(obj) else None
        if isinstance(obj, float) and (obj != obj):   # NaN
            return None
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        if isinstance(obj, (np.bool_,)):
            return bool(obj)
        return super().default(obj)


def _df_to_json_records(df: pd.DataFrame) -> list:
    """Converte DataFrame para lista de dicts JSON-safe."""
    # Converte Timestamp/NaT para string ISO antes do dump
    df_copy = df.copy()
    for col in df_copy.select_dtypes(include=["datetime64[ns]", "datetimetz"]):
        df_copy[col] = df_copy[col].apply(
            lambda x: x.isoformat() if not pd.isna(x) else None
        )
    return df_copy.where(pd.notnull(df_copy), None).to_dict(orient="records")


# ---------------------------------------------------------------------------
# Exportações principais
# ---------------------------------------------------------------------------

def exportar_parquet(df: pd.DataFrame) -> None:
    """Salva o DataFrame completo em Parquet."""
    df.to_parquet(OUTPUT_PARQUET, index=False)
    tamanho_kb = os.path.getsize(OUTPUT_PARQUET) / 1024
    logger.info(f"Parquet salvo: {OUTPUT_PARQUET} ({tamanho_kb:.1f} KB)")


def exportar_json(df: pd.DataFrame) -> None:
    """Salva o DataFrame completo em JSON (array de objetos)."""
    registros = _df_to_json_records(df)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(registros, f, cls=_SafeEncoder, ensure_ascii=False)
    tamanho_kb = os.path.getsize(OUTPUT_JSON) / 1024
    logger.info(f"JSON salvo: {OUTPUT_JSON} ({tamanho_kb:.1f} KB)")


def exportar_metricas_agregadas(df: pd.DataFrame) -> None:
    """
    Pré-calcula e salva métricas agregadas para carregamento rápido
    no dashboard. Arquivo: data/processed/metricas.json
    """
    metricas_path = os.path.join(PROCESSED_DIR, "metricas.json")
    agora = datetime.now().isoformat()

    # Filtra apenas primeiras linhas de cada entrega para contagens
    df_entregas = df[df["is_primeira_linha_entrega"] == 1].copy()

    def _agg_empresa(df_emp: pd.DataFrame) -> dict:
        """Agrega métricas para um subconjunto de empresa."""
        if df_emp.empty:
            return {}

        total_entregas   = int(df_emp["is_primeira_linha_entrega"].sum())
        no_prazo         = int(df_emp["entrega_no_prazo"].sum())
        em_atraso        = int(df_emp["entrega_em_atraso"].sum())
        pct_prazo        = round(no_prazo / total_entregas * 100) if total_entregas else 0

        # Usar df_emp para todos os financeiros (escopo correto)
        df_fin = df_emp  # mesmo subconjunto filtrado
        return {
            "total_entregas":        total_entregas,
            "entregas_no_prazo":     no_prazo,
            "entregas_em_atraso":    em_atraso,
            "pct_prazo":             pct_prazo,
            "faturamento_total":     round(float(df_fin["vr_nf"].sum()), 2),
            "frete_total":           round(float(df_fin["frete"].sum()), 2),
            "peso_total_kg":         round(float(df_fin["peso_kg"].sum()), 3),
            "volumes_total":         int(df_fin["volumes"].sum()),
            "total_impostos":        round(float(df_fin["total_impostos_calc"].sum()), 2),
            "custo_embalagem_total": round(float(df_fin["custo_embalagem"].sum()), 2),
            "custo_total_envio":     round(float(df_fin["custo_total_envio"].sum()), 2),
        }

    # --- Global ---
    metricas_global = _agg_empresa(df_entregas)

    # --- Por empresa ---
    metricas_iab    = _agg_empresa(
        df_entregas[df_entregas["empresa"] == "IAB"]
    )
    metricas_verbum = _agg_empresa(
        df_entregas[df_entregas["empresa"] == "VERBUM"]
    )

    # --- Por mês (global) ---
    por_mes = (
        df_entregas.groupby(["ano_ref", "ano", "mes"], as_index=False)
        .agg(
            total_entregas   = ("is_primeira_linha_entrega", "sum"),
            no_prazo         = ("entrega_no_prazo",          "sum"),
            em_atraso        = ("entrega_em_atraso",         "sum"),
            faturamento      = ("vr_nf",                     "sum"),
            frete            = ("frete",                     "sum"),
            peso_kg          = ("peso_kg",                   "sum"),
            volumes          = ("volumes",                   "sum"),
            total_impostos   = ("total_impostos_calc",       "sum"),
            custo_envio      = ("custo_total_envio",         "sum"),
        )
        .sort_values(["ano", "mes"])
    )
    por_mes["pct_prazo"] = (
        por_mes["no_prazo"] / por_mes["total_entregas"].replace(0, np.nan) * 100
    ).round().fillna(0).astype(int)

    # --- Por transportadora ---
    por_transp = (
        df_entregas.groupby("transportadora_norm", as_index=False)
        .agg(
            total_entregas = ("is_primeira_linha_entrega", "sum"),
            no_prazo       = ("entrega_no_prazo",          "sum"),
            em_atraso      = ("entrega_em_atraso",         "sum"),
            frete_total    = ("frete",                     "sum"),
        )
        .sort_values("total_entregas", ascending=False)
    )
    por_transp["pct_prazo"] = (
        por_transp["no_prazo"] / por_transp["total_entregas"].replace(0, np.nan) * 100
    ).round().fillna(0).astype(int)

    # --- Por UF ---
    por_uf = (
        df_entregas.groupby("uf", as_index=False)
        .agg(
            total_entregas = ("is_primeira_linha_entrega", "sum"),
            no_prazo       = ("entrega_no_prazo",          "sum"),
            em_atraso      = ("entrega_em_atraso",         "sum"),
            faturamento    = ("vr_nf",                     "sum"),
        )
        .sort_values("total_entregas", ascending=False)
    )

    # --- Por Ano Ref ---
    por_ano_ref = {}
    for ano in df_entregas["ano_ref"].dropna().unique():
        df_ar = df_entregas[df_entregas["ano_ref"] == str(ano)]
        por_ano_ref[str(ano)] = {
            "iab":    _agg_empresa(df_ar[df_ar["empresa"] == "IAB"]),
            "verbum": _agg_empresa(df_ar[df_ar["empresa"] == "VERBUM"]),
            "global": _agg_empresa(df_ar),
        }

    metricas = {
        "gerado_em":       agora,
        "total_linhas_df": len(df),
        "global":          metricas_global,
        "iab":             metricas_iab,
        "verbum":          metricas_verbum,
        "por_mes":         _df_to_json_records(por_mes),
        "por_transportadora": _df_to_json_records(por_transp),
        "por_uf":          _df_to_json_records(por_uf),
        "por_ano_ref":     por_ano_ref,
    }

    with open(metricas_path, "w", encoding="utf-8") as f:
        json.dump(metricas, f, cls=_SafeEncoder, ensure_ascii=False, indent=2)

    tamanho_kb = os.path.getsize(metricas_path) / 1024
    logger.info(f"Métricas agregadas salvas: {metricas_path} ({tamanho_kb:.1f} KB)")
