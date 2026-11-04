# =============================================================================
# modules/recalc_excel.py — Recalcula as fórmulas de uma planilha pelo Excel
# =============================================================================
# Por que existe: o importar_base.py grava o Quadro de Envios com openpyxl, que
# APAGA os valores já calculados das fórmulas (cliente, UF, região, prev. entrega...).
# O ETL lê só esses valores, então o dashboard saía zerado.
#
# O que faz: abre a planilha no Excel (em segundo plano, invisível), manda
# recalcular tudo, salva e fecha. NÃO altera nenhum dado digitado — é o mesmo
# que abrir no Excel e apertar Salvar.
#
# Requisito: Microsoft Excel instalado no computador (Windows).
# =============================================================================

import os
import subprocess
import logging

logger = logging.getLogger(__name__)

_PS_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$p = $env:RECALC_ARQUIVO
$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    # Open(arquivo, UpdateLinks=0, ReadOnly=False)
    $wb = $xl.Workbooks.Open($p, 0, $false)
    if ($wb.ReadOnly) { Write-Output 'SOMENTE_LEITURA'; exit 3 }
    $xl.CalculateFull()
    $wb.Save()
    Write-Output 'OK'
    exit 0
} catch {
    Write-Output ('ERRO: ' + $_.Exception.Message)
    exit 2
} finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($xl) { $xl.Quit() | Out-Null; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
}
"""


def recalcular_excel(caminho: str, timeout: int = 300) -> bool:
    """Abre `caminho` no Excel, recalcula todas as fórmulas, salva e fecha.

    Retorna True se deu certo. Nunca levanta exceção: em caso de falha só
    registra o motivo no log e retorna False.
    """
    if os.name != "nt":
        logger.warning("Recálculo pelo Excel só funciona no Windows — pulado.")
        return False
    if not os.path.isfile(caminho):
        logger.warning(f"Recálculo: arquivo não encontrado: {caminho}")
        return False

    logger.info("Recalculando fórmulas no Excel (segundo plano)...")
    env = dict(os.environ, RECALC_ARQUIVO=os.path.abspath(caminho))
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", _PS_SCRIPT],
            env=env, capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        logger.warning("Recálculo: o Excel demorou demais e foi interrompido.")
        return False
    except Exception as e:
        logger.warning(f"Recálculo: não foi possível chamar o Excel ({e}).")
        return False

    saida = (r.stdout or "").strip()
    if r.returncode == 0 and "OK" in saida:
        logger.info("Fórmulas recalculadas e Quadro salvo pelo Excel.")
        return True
    if r.returncode == 3 or "SOMENTE_LEITURA" in saida:
        logger.warning("Recálculo: o Quadro está aberto por alguém (somente leitura). "
                       "Feche o arquivo em todos os computadores e rode de novo.")
        return False
    logger.warning(f"Recálculo falhou: {saida or r.stderr.strip()}")
    return False
