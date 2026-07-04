# =============================================================================
# modules/snapshot.py — Detecção de mudanças e preservação de histórico
# =============================================================================
# Lógica de comparação:
#   • Chave: Data Coleta + CNPJ padronizado + Nº Pedido + Empresa
#   • A cada execução, lê o snapshot anterior (se existir)
#   • Compara linha a linha:
#       - Linha nova (chave não existia) → INSERT
#       - Linha alterada (chave existe, algum campo mudou) → UPDATE
#       - Linha removida (chave estava no snapshot, sumiu do Excel) → DELETE
#   • Grava log detalhado de cada mudança
#   • Salva o novo snapshot como "latest"
#   • Mantém cópia datada para histórico permanente
# =============================================================================

import os
import hashlib
import json
import logging
from datetime import datetime

import pandas as pd

from config import (
    SNAPSHOT_LATEST,
    SNAPSHOTS_DIR,
    CHANGE_LOG,
    CHAVE_SNAPSHOT,
)

logger = logging.getLogger(__name__)

# Colunas que, se mudarem, registramos no log de alterações
COLUNAS_MONITORADAS = [
    "natureza_operacao",
    "transportadora_norm",
    "vr_nf",
    "frete",
    "peso_kg",
    "volumes",
    "icms",
    "difal",
    "fcp",
    "total_impostos_calc",
    "custo_embalagem",
    "custo_total_envio",
    "prev_entrega_cliente",
    "data_efetiva_entrega",
    "status_prazo",
    "entrega_concluida",
    "modalidade",
    "municipio",
    "tipo_cliente",
    "cliente",
]


def _hash_linha(row: pd.Series) -> str:
    """Gera hash MD5 dos campos monitorados de uma linha."""
    campos = {c: str(row.get(c, "")) for c in COLUNAS_MONITORADAS}
    conteudo = json.dumps(campos, sort_keys=True, ensure_ascii=False)
    return hashlib.md5(conteudo.encode()).hexdigest()


def _chave_str(row: pd.Series) -> str:
    """Concatena os campos da chave composta em uma string única."""
    partes = []
    for campo in CHAVE_SNAPSHOT:
        v = row.get(campo, "")
        if hasattr(v, "strftime"):          # datetime
            try:
                v = v.strftime("%Y-%m-%d")
            except Exception:
                v = "9999-01-01"            # NaT → data sentinela
        partes.append(str(v).strip())
    return "||".join(partes)


def _preparar_indice(df: pd.DataFrame) -> dict:
    """
    Retorna dict {chave_str: (hash_linha, row_dict)} para comparação rápida.
    """
    indice = {}
    for _, row in df.iterrows():
        chave = _chave_str(row)
        h     = _hash_linha(row)
        indice[chave] = (h, row.to_dict())
    return indice


def comparar_e_registrar(df_novo: pd.DataFrame) -> dict:
    """
    Compara o DataFrame atual com o snapshot anterior.

    Retorna
    -------
    dict com chaves:
        'inseridos'  : int
        'alterados'  : int
        'removidos'  : int
        'sem_mudanca': int
        'log_path'   : str (caminho do CSV de log)
    """
    agora = datetime.now()
    timestamp = agora.strftime("%Y-%m-%d %H:%M:%S")

    # ------------------------------------------------------------------
    # 1. Lê snapshot anterior (se existir)
    # ------------------------------------------------------------------
    if os.path.exists(SNAPSHOT_LATEST):
        df_anterior = pd.read_parquet(SNAPSHOT_LATEST)
        indice_anterior = _preparar_indice(df_anterior)
        logger.info(f"Snapshot anterior carregado: {len(df_anterior)} linhas")
    else:
        indice_anterior = {}
        logger.info("Nenhum snapshot anterior encontrado — primeira execução.")

    # ------------------------------------------------------------------
    # 2. Indexa o DataFrame novo
    # ------------------------------------------------------------------
    indice_novo = _preparar_indice(df_novo)

    # ------------------------------------------------------------------
    # 3. Detecta inserções e alterações
    # ------------------------------------------------------------------
    registros_log = []
    inseridos = alterados = sem_mudanca = 0

    for chave, (hash_novo, row_novo) in indice_novo.items():
        if chave not in indice_anterior:
            # Linha nova
            inseridos += 1
            registros_log.append({
                "timestamp":    timestamp,
                "tipo":         "INSERT",
                "chave":        chave,
                "campo":        None,
                "valor_antes":  None,
                "valor_depois": None,
            })
        else:
            hash_ant, row_ant = indice_anterior[chave]
            if hash_novo != hash_ant:
                # Linha alterada — descobre quais campos mudaram
                alterados += 1
                for col in COLUNAS_MONITORADAS:
                    v_ant = str(row_ant.get(col, ""))
                    v_nov = str(row_novo.get(col, ""))
                    if v_ant != v_nov:
                        registros_log.append({
                            "timestamp":    timestamp,
                            "tipo":         "UPDATE",
                            "chave":        chave,
                            "campo":        col,
                            "valor_antes":  v_ant,
                            "valor_depois": v_nov,
                        })
            else:
                sem_mudanca += 1

    # ------------------------------------------------------------------
    # 4. Detecta remoções (chave existia antes, não existe mais)
    # ------------------------------------------------------------------
    removidos = 0
    for chave in indice_anterior:
        if chave not in indice_novo:
            removidos += 1
            registros_log.append({
                "timestamp":    timestamp,
                "tipo":         "DELETE",
                "chave":        chave,
                "campo":        None,
                "valor_antes":  None,
                "valor_depois": None,
            })

    # ------------------------------------------------------------------
    # 5. Grava log de mudanças
    # ------------------------------------------------------------------
    if registros_log:
        df_log = pd.DataFrame(registros_log)
        if os.path.exists(CHANGE_LOG):
            df_log_existente = pd.read_csv(CHANGE_LOG)
            df_log = pd.concat([df_log_existente, df_log], ignore_index=True)
        df_log.to_csv(CHANGE_LOG, index=False, encoding="utf-8-sig")
        logger.info(f"Log de mudanças atualizado: {len(registros_log)} registros → {CHANGE_LOG}")
    else:
        logger.info("Nenhuma mudança detectada em relação ao snapshot anterior.")

    # ------------------------------------------------------------------
    # 6. Salva novo snapshot (latest + cópia datada)
    # ------------------------------------------------------------------
    df_novo.to_parquet(SNAPSHOT_LATEST, index=False)

    snapshot_datado = os.path.join(
        SNAPSHOTS_DIR,
        f"snapshot_{agora.strftime('%Y%m%d_%H%M%S')}.parquet"
    )
    df_novo.to_parquet(snapshot_datado, index=False)
    logger.info(f"Snapshot salvo: {snapshot_datado}")

    resumo = {
        "inseridos":   inseridos,
        "alterados":   alterados,
        "removidos":   removidos,
        "sem_mudanca": sem_mudanca,
        "log_path":    CHANGE_LOG,
    }

    logger.info(
        f"Comparação concluída — "
        f"Inseridos: {inseridos} | "
        f"Alterados: {alterados} | "
        f"Removidos: {removidos} | "
        f"Sem mudança: {sem_mudanca}"
    )

    return resumo
