# =============================================================================
# etl_runner.py — Orquestrador principal do ETL Logístico IAB & Verbum
# =============================================================================
# COMO USAR:
#   Execute manualmente sempre que quiser atualizar o dashboard:
#
#   python etl_runner.py
#
# O script vai:
#   1. Ler TODAS as linhas do Quadro de Envios (releitura completa)
#   2. Limpar e normalizar todos os campos
#   3. Recalcular KPIs (entregas, prazo, atraso, impostos, custos)
#   4. Comparar com o estado anterior e registrar mudanças no log
#   5. Salvar os dados processados (Parquet + JSON) para o dashboard
#   6. Exibir resumo da execução no terminal
# =============================================================================

import sys
import os
import logging
from datetime import datetime

# Garante que os módulos locais sejam encontrados
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import (
    QUADRO_ENVIOS_PATH,
    PROCESSED_DIR, SNAPSHOTS_DIR, LOGS_DIR,
)
from modules.reader     import ler_quadro_envios
from modules.cleaner    import limpar_e_normalizar
from modules.calculator import calcular_kpis
from modules.snapshot   import comparar_e_registrar
from modules.exporter   import exportar_parquet, exportar_json, exportar_metricas_agregadas

# ---------------------------------------------------------------------------
# Configuração de logging — exibe no terminal E grava em arquivo
# ---------------------------------------------------------------------------
os.makedirs(LOGS_DIR, exist_ok=True)

log_file = os.path.join(
    LOGS_DIR,
    f"etl_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(log_file, encoding="utf-8"),
    ],
)
logger = logging.getLogger("etl_runner")


# ---------------------------------------------------------------------------
# Garante que as pastas de saída existam
# ---------------------------------------------------------------------------
for pasta in [PROCESSED_DIR, SNAPSHOTS_DIR, LOGS_DIR]:
    os.makedirs(pasta, exist_ok=True)


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------
def executar():
    inicio = datetime.now()
    logger.info("=" * 65)
    logger.info("  INÍCIO DO ETL — Dashboard Logístico IAB & Verbum")
    logger.info(f"  Arquivo fonte: {QUADRO_ENVIOS_PATH}")
    logger.info("=" * 65)

    try:
        # ------------------------------------------------------------------
        # ETAPA 1 — Leitura integral do Quadro de Envios
        # ------------------------------------------------------------------
        logger.info("[ ETAPA 1 ] Leitura do arquivo...")
        df_raw = ler_quadro_envios()

        # ------------------------------------------------------------------
        # ETAPA 2 — Limpeza e normalização
        # ------------------------------------------------------------------
        logger.info("[ ETAPA 2 ] Limpeza e normalização...")
        df_limpo = limpar_e_normalizar(df_raw)

        # ------------------------------------------------------------------
        # ETAPA 3 — Cálculo de KPIs
        # ------------------------------------------------------------------
        logger.info("[ ETAPA 3 ] Cálculo de KPIs...")
        df_final = calcular_kpis(df_limpo)

        # ------------------------------------------------------------------
        # ETAPA 4 — Detecção de mudanças e snapshot
        # ------------------------------------------------------------------
        logger.info("[ ETAPA 4 ] Comparação com snapshot anterior...")
        resumo_mudancas = comparar_e_registrar(df_final)

        # ------------------------------------------------------------------
        # ETAPA 5 — Exportação
        # ------------------------------------------------------------------
        logger.info("[ ETAPA 5 ] Exportando dados processados...")
        exportar_parquet(df_final)
        exportar_json(df_final)
        exportar_metricas_agregadas(df_final)

    except Exception as e:
        logger.error(f"ERRO CRÍTICO: {e}", exc_info=True)
        logger.error("ETL INTERROMPIDO — verifique o erro acima.")
        sys.exit(1)

    # ------------------------------------------------------------------
    # Resumo final no terminal
    # ------------------------------------------------------------------
    duracao = (datetime.now() - inicio).total_seconds()

    logger.info("=" * 65)
    logger.info("  ETL CONCLUÍDO COM SUCESSO")
    logger.info(f"  Duração total:      {duracao:.1f}s")
    logger.info(f"  Linhas processadas: {len(df_final)}")
    logger.info(f"  Inserções:          {resumo_mudancas['inseridos']}")
    logger.info(f"  Atualizações:       {resumo_mudancas['alterados']}")
    logger.info(f"  Remoções:           {resumo_mudancas['removidos']}")
    logger.info(f"  Sem mudança:        {resumo_mudancas['sem_mudanca']}")
    logger.info(f"  Log de mudanças:    {resumo_mudancas['log_path']}")
    logger.info(f"  Log de execução:    {log_file}")
    logger.info("=" * 65)

    # Imprime no terminal de forma destacada
    print("\n" + "─" * 55)
    print("  ✅  ETL finalizado!")
    print(f"  📦  Linhas válidas:   {len(df_final)}")
    empresas = df_final["empresa"].value_counts().to_dict()
    for emp, qtd in sorted(empresas.items()):
        print(f"       {emp:<10}: {qtd} linhas")
    print(f"  🔄  Mudanças:  +{resumo_mudancas['inseridos']} ins "
          f"/ ~{resumo_mudancas['alterados']} alt "
          f"/ -{resumo_mudancas['removidos']} rem")
    print(f"  ⏱  Tempo:     {duracao:.1f}s")
    print("─" * 55 + "\n")


if __name__ == "__main__":
    executar()
