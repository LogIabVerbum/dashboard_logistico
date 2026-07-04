# =============================================================================
# modules/reader.py — Leitura integral do Quadro de Envios
# =============================================================================
# Lê TODAS as linhas a cada execução (não incremental).
# O operador pode ter alterado qualquer campo, em qualquer posição.
# =============================================================================

import pandas as pd
import logging
from config import QUADRO_ENVIOS_PATH, ABA_LANCAMENTOS, COLUNAS_ENTRADA

logger = logging.getLogger(__name__)


def ler_quadro_envios(caminho: str = QUADRO_ENVIOS_PATH) -> pd.DataFrame:
    """
    Lê a aba 'Lançamentos' completa, retornando apenas as colunas
    relevantes para o ETL. Fórmulas do Excel são ignoradas (data_only=True
    via openpyxl por baixo do pandas).

    Returns
    -------
    pd.DataFrame
        DataFrame bruto, sem nenhuma transformação aplicada.
    """
    logger.info(f"Lendo arquivo: {caminho}")

    try:
        df = pd.read_excel(
            caminho,
            sheet_name=ABA_LANCAMENTOS,
            engine="openpyxl",
            dtype=str,          # tudo como string primeiro → evita coerção errada
        )
    except FileNotFoundError:
        logger.error(f"Arquivo não encontrado: {caminho}")
        raise
    except Exception as e:
        logger.error(f"Erro ao ler o arquivo: {e}")
        raise

    logger.info(f"Linhas lidas (raw): {len(df)}")

    # Seleciona apenas colunas conhecidas que existam no arquivo
    colunas_disponiveis = [c for c in COLUNAS_ENTRADA if c in df.columns]
    colunas_ausentes    = [c for c in COLUNAS_ENTRADA if c not in df.columns]

    if colunas_ausentes:
        logger.warning(f"Colunas esperadas não encontradas no arquivo: {colunas_ausentes}")

    df = df[colunas_disponiveis].copy()

    # Remove linhas completamente vazias
    df.dropna(how="all", inplace=True)

    logger.info(f"Linhas após remoção de vazias: {len(df)}")
    return df
