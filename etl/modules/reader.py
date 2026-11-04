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


# Colunas que são FÓRMULAS no Quadro. Se vierem vazias, o arquivo foi salvo
# sem os valores calculados (ex.: pelo importar_base.py) e precisa ser recalculado.
COLUNAS_FORMULA_CHECAGEM = ["Prev. Entrega Cliente", "Cliente"]
MIN_PREENCHIDAS = 0.5   # abaixo de 50% preenchido = fórmulas sem valor


def _formulas_ok(df: pd.DataFrame) -> bool:
    base = df[df["NF"].notna()] if "NF" in df.columns else df
    if base.empty:
        return True
    for col in COLUNAS_FORMULA_CHECAGEM:
        if col in base.columns and base[col].notna().mean() < MIN_PREENCHIDAS:
            logger.warning(f"Coluna de fórmula '{col}' veio vazia em "
                           f"{(1 - base[col].notna().mean()):.0%} das linhas.")
            return False
    return True


def ler_quadro_envios(caminho: str = QUADRO_ENVIOS_PATH) -> pd.DataFrame:
    """Lê o Quadro; se as fórmulas vierem sem valor, recalcula pelo Excel e lê de novo.
    Se mesmo assim continuar vazio, PARA o ETL (melhor do que gerar dashboard zerado)."""
    df = _ler_quadro_envios(caminho)
    if _formulas_ok(df):
        return df

    from modules.recalc_excel import recalcular_excel
    logger.warning("Quadro salvo sem os valores das fórmulas — tentando recalcular pelo Excel...")
    if recalcular_excel(caminho):
        df = _ler_quadro_envios(caminho)
        if _formulas_ok(df):
            logger.info("Quadro recalculado com sucesso — seguindo com o ETL.")
            return df

    raise RuntimeError(
        "O Quadro de Envios está com as fórmulas sem valor (cliente/prev. entrega vazios). "
        "Abra o Quadro no Excel, clique em Salvar, feche e rode o ATUALIZAR DASHBOARD de novo. "
        "O dashboard NÃO foi atualizado para não sair zerado."
    )


def _ler_quadro_envios(caminho: str = QUADRO_ENVIOS_PATH) -> pd.DataFrame:
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
