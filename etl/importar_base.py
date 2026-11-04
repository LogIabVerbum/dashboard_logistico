# importar_base.py
# Importa o relatório de faturamento do Sankhya (.xlsx) e lança novas NFs no Quadro de Envios.
# Versão: atualizada em 05/08/2026 — ajustado para novo formato do Relatório e novo Quadro de Envios (aba "Lançamentos", NF em col O)

import os
import re
import sys
import logging
import warnings
import datetime
import shutil
from copy import copy

import pandas as pd
import openpyxl
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=UserWarning)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Localização da pasta raiz (multi-máquina)
# ---------------------------------------------------------------------------

def _encontrar_pasta() -> str:
    """
    Retorna o caminho para a pasta que contém 'Relatorio de faturamento.xlsx'.

    Estratégia de busca (ordem de prioridade):
    1. Pasta do próprio script e até 3 níveis acima — funciona para qualquer estrutura
       de diretórios, incluindo a pasta etl/ do Miqueias.
    2. Variáveis de ambiente OneDrive + subpastas conhecidas.
    3. Caminho absoluto legado do Miqueias.
    4. Fallback: pasta do script (com aviso).
    """
    ARQUIVO = "Relatorio de faturamento.xlsx"

    def _achou(pasta):
        return os.path.exists(os.path.join(pasta, ARQUIVO))

    # ── 1. Busca relativa ao script (funciona independentemente de onde está instalado) ──
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidatos_relativos = [script_dir]
    # Sobe até 3 níveis (etl → dashboard_logistico → 10 - DASHIBOARDs → ...)
    pasta = script_dir
    for _ in range(3):
        pasta = os.path.dirname(pasta)
        candidatos_relativos.append(pasta)
        # Também verifica subpastas de nomes conhecidos dentro de cada nível
        for sub in (
            "Arquivos de Claude Relatórios - Relatórios Logística",
            "Arquivos de Claude Relatorios - Relatorios Logistica",
        ):
            candidatos_relativos.append(os.path.join(pasta, sub))

    for c in candidatos_relativos:
        if _achou(c):
            log.info(f"Pasta encontrada (relativa ao script): {c}")
            return c

    # ── 2. Variáveis de ambiente OneDrive ────────────────────────────────────────────────
    bases = []
    for var in ("OneDriveCommercial", "OneDrive"):
        v = os.environ.get(var, "")
        if v:
            bases.append(v)
    perfil = os.environ.get("USERPROFILE", "")
    if perfil:
        bases.append(os.path.join(perfil, "OneDrive - Instituto Alfa e Beto"))

    sufixos = [
        r"Logística-IAB\10 - DASHIBOARDs\Arquivos de Claude Relatórios - Relatórios Logística",
        r"Logistica-IAB\10 - DASHIBOARDs\Arquivos de Claude Relatorios - Relatorios Logistica",
        r"Arquivos de Anderson Pissaia - Logística-IAB\10 - DASHIBOARDs\Arquivos de Claude Relatórios - Relatórios Logística",
    ]

    for base in bases:
        for suf in sufixos:
            caminho = os.path.join(base, suf)
            if _achou(caminho):
                log.info(f"Pasta encontrada (OneDrive env): {caminho}")
                return caminho

    # ── 3. Caminhos absolutos conhecidos por máquina ─────────────────────────────────────
    caminhos_conhecidos = [
        # Miqueias
        r"C:\Users\mique\Instituto Alfa e Beto\Claude Relatórios - Relatórios Logística",
        r"C:\Users\mique\OneDrive - Instituto Alfa e Beto\Logistica-IAB\10 - DASHIBOARDs\Arquivos de Claude Relatorios - Relatorios Logistica",
        # Augusto
        r"C:\Users\AugustoCesar\Instituto Alfa e Beto\Claude Relatórios - Relatórios Logística",
    ]
    for caminho in caminhos_conhecidos:
        if _achou(caminho):
            log.info(f"Pasta encontrada (caminho fixo): {caminho}")
            return caminho

    # ── 4. Fallback: pasta do script (avisa mas não trava) ───────────────────────────────
    log.warning(
        f"'Relatorio de faturamento.xlsx' não encontrado em nenhum caminho conhecido.\n"
        f"Usando pasta do script como fallback: {script_dir}\n"
        f"Verifique se o arquivo existe e se o OneDrive está sincronizado."
    )
    return script_dir


# ---------------------------------------------------------------------------
# Caminhos
# ---------------------------------------------------------------------------
PASTA_DASHBOARDS = _encontrar_pasta()

# Relatório de faturamento fica DENTRO da pasta retornada
ARQUIVO_BASE = os.path.join(PASTA_DASHBOARDS, "Relatorio de faturamento.xlsx")

# Quadro de Envios e backups
# - Caminho normal: um nível acima de PASTA_DASHBOARDS (dentro de 10 - DASHIBOARDs)
# - Caminhos Miqueias (SharePoint direto, duas variantes conhecidas)
_quadros_mique = [
    # Miqueias
    r"C:\Users\mique\Instituto Alfa e Beto\Logistica-IAB\10 - DASHIBOARDs\Quadro de Envios.xlsx",
    r"C:\Users\mique\OneDrive - Instituto Alfa e Beto\Logistica-IAB\10 - DASHIBOARDs\Quadro de Envios.xlsx",
    # Augusto
    r"C:\Users\AugustoCesar\OneDrive - Instituto Alfa e Beto\Arquivos de Anderson Pissaia - Logística-IAB\10 - DASHIBOARDs\Quadro de Envios.xlsx",
]
_quadro_encontrado = next((p for p in _quadros_mique if os.path.exists(p)), None)
QUADRO_ENVIOS = (
    _quadro_encontrado
    if _quadro_encontrado
    else os.path.join(os.path.dirname(PASTA_DASHBOARDS), "Quadro de Envios.xlsx")
)
PASTA_BACKUP = os.path.join(os.path.dirname(QUADRO_ENVIOS), "backups")   # destino antigo (usado só como reserva)

# Backups do Quadro de Envios vão para: Logística-IAB\1 - QUADRO DE ENVIOS\<AAAA>\<MMAAAA>
# (ex.: ...\1 - QUADRO DE ENVIOS\2026\112026). Pastas do ano/mês são criadas automaticamente.
PASTA_QUADRO_ENVIOS_RAIZ = os.path.join(
    os.path.dirname(os.path.dirname(QUADRO_ENVIOS)),   # 10 - DASHIBOARDs -> Logística-IAB
    "1 - QUADRO DE ENVIOS",
)

# ---------------------------------------------------------------------------
# Mapeamento cabeçalho do Relatório de Faturamento → nome interno
# (O arquivo exportado pelo Sankhya já vem com nomes em maiúsculo/underscore)
# ---------------------------------------------------------------------------
RENAME_MAP = {
    # Relatório novo (snake_case exportado pelo Sankhya)
    "CNPJ_REMETENTE":           "CNPJ_REMETENTE",
    "NRO_NFE":                  "NRO_NFE",
    "DT_EMISSAO_NF":            "DT_EMISSAO_NF",
    "CPF_CNPJ_DESTINATARIO":    "CPF_CNPJ_DEST",
    "CLIENTE_DESTINATARIO":     "NOME_PARCEIRO",
    "CEP":                      "CEP_DESTINO",
    "NATUREZA_OPERACAO":        "NATUREZA_OPERACAO",
    "TIPO_OPERACAO":            "NATUREZA_OPERACAO",
    "VALOR_TOTAL_NOTA":         "VALOR_TOTAL_NOTA",
    "VALOR_BASE_CALCULO":       "VALOR_BASE_CALCUL",
    "VALOR_ICMS":               "VALOR_ICMS",
    "VALOR_DIFAL":              "VALOR_DIFAL",
    "VALOR_FCP":                "VALOR_FCP",
    "PESO":                     "PESO",
    "QTD_VOLUME":               "QTD_VOLUME",
    "CIDADE_DESTINO":           "MUNICIPIO_DESTINO",
    "UF_DESTINO":               "UF_DESTINO",
    "TRANSPORTADORA":           "TRANSPORTADORA",
    # Nomes antigos (compatibilidade retroativa)
    "Número da NF-e":           "NRO_NFE",
    "Dt. Emissão NF":           "DT_EMISSAO_NF",
    "Nome Parceiro":            "NOME_PARCEIRO",
    "CPF / CNPJ Dest.":         "CPF_CNPJ_DEST",
    "Município Destino":        "MUNICIPIO_DESTINO",
    "UF Destino":               "UF_DESTINO",
    "Valor Total da Nota":      "VALOR_TOTAL_NOTA",
    "Valor Base Cálculo":       "VALOR_BASE_CALCUL",
    "Natureza Operação":        "NATUREZA_OPERACAO",
    "Peso Líquido (kg)":        "PESO",
    "Peso Bruto (kg)":          "PESO",
}

# ---------------------------------------------------------------------------
# Colunas do Quadro de Envios  (aba "Lançamentos")
# Estrutura detectada em 05/08/2026:
#   A=Data Coleta  D=CNPJ/CPF  E=Cliente  G=Municipio  H=UF  J=Cep
#   N=Empresa      O=NF        P=Natureza  R=Vr NF      T=Peso (kg)
#   U=Base Calc.   V=ICMS      W=Difal     X=FCP        Z=Volumes
#   \=Transportadora
# ---------------------------------------------------------------------------
QE_ABA = "Lançamentos"   # nome da aba no Quadro de Envios

QE_COLUNAS = {
    "DT_EMISSAO_NF":     1,   # A  — Data Coleta
    "ANO":               2,   # B  — Ano  (derivado da data)
    "MES":               3,   # C  — Mês  (derivado da data)
    "CPF_CNPJ_DEST":     4,   # D  — CNPJ / CPF
    "NOME_PARCEIRO":     5,   # E  — Cliente
    "PAIS":              6,   # F  — País  (sempre BRASIL)
    "MUNICIPIO_DESTINO": 7,   # G  — Municipio
    "UF_DESTINO":        8,   # H  — UF
    "REGIAO":            9,   # I  — Região  (calculada por UF)
    "CEP_DESTINO":       10,  # J  — Cep
    "TIPO_CLIENTE":      11,  # K  — Tipo de Cliente  (histórico ou derivado)
    "EMPRESA":           14,  # N  — Empresa  (sempre "IAB")
    "NRO_NFE":           15,  # O  — NF
    "NATUREZA_OPERACAO": 16,  # P  — NATUREZA DA OPERAÇÃO
    "VALOR_TOTAL_NOTA":  18,  # R  — Vr NF
    "PESO":              20,  # T  — Peso (kg)
    "VALOR_BASE_CALCUL": 21,  # U  — Base Cálculo ICM
    "VALOR_ICMS":        22,  # V  — ICMS
    "VALOR_DIFAL":       23,  # W  — Difal
    "VALOR_FCP":         24,  # X  — FCP
    "QTD_VOLUME":        26,  # Z  — Volumes
    "TRANSPORTADORA":    28,  # \  — Transportadora
    "CONS_ANALISE":      13,  # M  — Cons Analise  (sempre "Não" para linhas novas)
    "ANO_REF":           43,  # AQ — Ano Ref  (calendário escolar out-set)
    "DISTANCIA":         45,  # AS — Distância (km)  (fórmula copiada da linha anterior)
}

# Colunas com fórmulas que devem ser copiadas da linha anterior (adaptando o nº de linha)
# Distância (45) é tratada separadamente; aqui ficam as demais.
QE_COLUNAS_FORMULA_COPIAR = [
    30,   # AD — Prev. Envio
    31,   # AE — Prev. Entrega Cliente
    32,   # AF — Prev. Entrega Transportadora
    37,   # AK — Entrega No Prazo
    38,   # AL — Entrega em Atraso
    39,   # AM — Total Entregas
]

# ---------------------------------------------------------------------------
# Tabela de regiões por UF
# ---------------------------------------------------------------------------
_UF_REGIAO = {
    "AC": "NORTE",      "AM": "NORTE",      "AP": "NORTE",
    "PA": "NORTE",      "RO": "NORTE",      "RR": "NORTE",      "TO": "NORTE",
    "AL": "NORDESTE",   "BA": "NORDESTE",   "CE": "NORDESTE",
    "MA": "NORDESTE",   "PB": "NORDESTE",   "PE": "NORDESTE",
    "PI": "NORDESTE",   "RN": "NORDESTE",   "SE": "NORDESTE",
    "DF": "CENTRO-OESTE", "GO": "CENTRO-OESTE", "MS": "CENTRO-OESTE", "MT": "CENTRO-OESTE",
    "ES": "SUDESTE",    "MG": "SUDESTE",    "RJ": "SUDESTE",    "SP": "SUDESTE",
    "PR": "SUL",        "RS": "SUL",        "SC": "SUL",
}

# Palavras-chave no nome do cliente que indicam ESCOLA
_KW_ESCOLA = (
    "COLEGIO", "COLÉGIO", "ESCOLA", "EDUCACIONAL", "LICEU",
    "CENTRO DE ENSINO", "CENTRO EDUCACIONAL", "ENSINO FUNDAMENTAL",
    "ENSINO MEDIO", "ENSINO MÉDIO", "CRECHE", "JARDIM DE INFANCIA",
    "JARDIM DE INFÂNCIA",
)

# Palavras-chave no nome que indicam MUNICÍPIO
_KW_MUNICIPIO = (
    "MUNICIPIO", "MUNICÍPIO", "PREFEITURA", "SECRETARIA MUNICIPAL",
    "SECRETARIA DE EDUCACAO", "SECRETARIA DE EDUCAÇÃO",
    "FUNDO MUNICIPAL",
)

# CNPJs conhecidos de tipo E-BOOK
_CNPJ_EBOOK = {"19004863000165"}  # ARVORE DE LIVROS — remove pontuação para comparar

# Mapeamento CNPJ remetente → Empresa
# Chave: apenas dígitos do CNPJ base (8 primeiros dígitos sem pontuação)
_CNPJ_BASE_EMPRESA = {
    "57638518": "VERBUM",   # 57.638.518/0001-72 e 57.638.518/0002-53
    "08458084": "IAB",      # 08.458.084/0001-13
}

def _empresa_por_cnpj(cnpj_remetente) -> str:
    """Identifica a empresa (IAB ou VERBUM) pelo CNPJ do remetente."""
    if not cnpj_remetente:
        return "IAB"
    digitos = str(cnpj_remetente).replace(".", "").replace("/", "").replace("-", "").strip()
    base = digitos[:8]  # primeiros 8 dígitos identificam a empresa
    return _CNPJ_BASE_EMPRESA.get(base, "IAB")

QE_LINHA_INICIO = 2   # linha 1 = cabeçalho
QE_COL_NF       = 15  # coluna O = número da NF


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_data(val):
    """Converte data serial do Excel (float) ou string para datetime.date."""
    if val is None:
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val if isinstance(val, datetime.date) else val.date()
    if isinstance(val, float):
        # Serial date do Excel: dias desde 1899-12-30
        try:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val)))
        except Exception:
            return None
    if isinstance(val, str):
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _fmt_cep(val) -> str:
    """Formata CEP como string de 8 dígitos com zero à esquerda."""
    if val is None:
        return ""
    s = str(val).strip().replace("-", "").replace(".", "")
    # remove ".0" de floats
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(8) if s.isdigit() else s


def _copiar_estilo_linha(ws, linha_ref: int, linha_dest: int, n_cols: int = 45):
    """Copia fonte, fill, borda, alinhamento e formato de número de toda a linha de referência."""
    for col in range(1, n_cols + 1):
        src = ws.cell(row=linha_ref,  column=col)
        dst = ws.cell(row=linha_dest, column=col)
        if src.has_style:
            dst.font         = copy(src.font)
            dst.fill         = copy(src.fill)
            dst.border       = copy(src.border)
            dst.alignment    = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection   = copy(src.protection)


def _regiao_por_uf(uf) -> str:
    """Retorna a região brasileira para uma UF."""
    if not uf:
        return ""
    return _UF_REGIAO.get(str(uf).strip().upper(), "")


def _ano_ref(dt) -> int:
    """
    Retorna o Ano de Referência conforme o calendário escolar IAB.
    O ano escolar vai de 01/10 a 30/09 do ano seguinte.
    Exemplo: data em out-dez/2025 → Ano Ref 2026; data em jan-set/2026 → Ano Ref 2026.
    """
    if dt is None:
        return None
    return dt.year + 1 if dt.month >= 10 else dt.year


def _tipo_cliente(cpf_cnpj, nome, historico: dict = None) -> str:
    """
    Determina o Tipo de Cliente.
    1. Consulta o histórico do Quadro de Envios (prioridade máxima)
    2. CPF (sem '/')         → PESSOA FÍSICA
    3. CNPJ de e-book        → E-BOOK
    4. CNPJ com nome escola  → ESCOLA
    5. CNPJ com nome municip → MUNICÍPIO
    6. Demais CNPJs          → PESSOA JURÍDICA
    """
    doc = str(cpf_cnpj or "").strip()
    doc_digits = doc.replace(".", "").replace("-", "").replace("/", "")
    nome_upper = str(nome or "").upper()

    # 1. Histórico: mesmo documento já classificado no Quadro
    if historico:
        tipo_hist = historico.get(doc_digits)
        if tipo_hist:
            return tipo_hist

    # 2. CPF: não contém '/' e tem até 11 dígitos
    if "/" not in doc and len(doc_digits) <= 11:
        return "PESSOA FÍSICA"

    # 3. E-BOOK: CNPJ específico
    if doc_digits in _CNPJ_EBOOK:
        return "E-BOOK"

    # 4. Escola
    if any(kw in nome_upper for kw in _KW_ESCOLA):
        return "ESCOLA"

    # 5. Município / Prefeitura
    if any(kw in nome_upper for kw in _KW_MUNICIPIO):
        return "MUNICÍPIO"

    # Default
    return "PESSOA JURÍDICA"


# ---------------------------------------------------------------------------
# Leitura do arquivo base (Relatorio de faturamento.xlsx)
# ---------------------------------------------------------------------------

def ler_arquivo_base() -> pd.DataFrame:
    log.info(f"Lendo: {ARQUIVO_BASE}")
    if not os.path.exists(ARQUIVO_BASE):
        raise FileNotFoundError(
            f"Arquivo base não encontrado: {ARQUIVO_BASE}\n"
            "Verifique se o OneDrive está sincronizado e se o arquivo foi salvo corretamente."
        )

    # Ler com openpyxl (engine padrão para .xlsx)
    df = pd.read_excel(
        ARQUIVO_BASE,
        engine="openpyxl",
        header=0,
        dtype=str,
    )
    df.dropna(how="all", inplace=True)

    # TIPO_OPERACAO tem prioridade sobre NATUREZA_OPERACAO — resolve antes do rename
    if "TIPO_OPERACAO" in df.columns:
        df["NATUREZA_OPERACAO"] = df["TIPO_OPERACAO"]
        df.drop(columns=["TIPO_OPERACAO"], inplace=True)

    # Renomear colunas
    df.rename(columns=RENAME_MAP, inplace=True)

    # Remove eventuais duplicatas residuais
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    # Garantir que todas as colunas internas existam (preenche com None se ausente)
    colunas_internas = set(RENAME_MAP.values())
    for col in colunas_internas:
        if col not in df.columns:
            df[col] = None

    # Conversões de tipo
    df["DT_EMISSAO_NF"]    = df["DT_EMISSAO_NF"].apply(_parse_data)
    df["NRO_NFE"]          = pd.to_numeric(df["NRO_NFE"], errors="coerce")
    df["VALOR_TOTAL_NOTA"] = pd.to_numeric(df["VALOR_TOTAL_NOTA"], errors="coerce")
    df["VALOR_BASE_CALCUL"]= pd.to_numeric(df["VALOR_BASE_CALCUL"], errors="coerce")
    df["VALOR_ICMS"]       = pd.to_numeric(df["VALOR_ICMS"],        errors="coerce")
    df["VALOR_DIFAL"]      = pd.to_numeric(df["VALOR_DIFAL"],       errors="coerce")
    df["VALOR_FCP"]        = pd.to_numeric(df["VALOR_FCP"],         errors="coerce")
    df["PESO"]             = pd.to_numeric(df["PESO"],              errors="coerce")
    df["QTD_VOLUME"]       = pd.to_numeric(df["QTD_VOLUME"],        errors="coerce")
    df["CEP_DESTINO"]      = df["CEP_DESTINO"].apply(_fmt_cep)

    # Normalizar UF
    df["UF_DESTINO"] = df["UF_DESTINO"].str.strip().str.upper()

    # Remover linhas sem NF válida
    df.dropna(subset=["NRO_NFE"], inplace=True)
    df["NRO_NFE"] = df["NRO_NFE"].astype(int)

    log.info(f"Arquivo base: {len(df)} linhas, {df['NRO_NFE'].nunique()} NFs únicas")
    return df


# ---------------------------------------------------------------------------
# Leitura das NFs já presentes no Quadro de Envios
# ---------------------------------------------------------------------------

def ler_nfs_quadro() -> set:
    log.info(f"Lendo NFs existentes em: {QUADRO_ENVIOS}")
    if not os.path.exists(QUADRO_ENVIOS):
        raise FileNotFoundError(
            f"Quadro de Envios não encontrado: {QUADRO_ENVIOS}\n"
            "Verifique se o OneDrive está sincronizado."
        )

    wb = load_workbook(QUADRO_ENVIOS, read_only=True, data_only=True)
    ws = wb[QE_ABA] if QE_ABA in wb.sheetnames else wb.active

    nfs = set()
    for row in ws.iter_rows(min_row=QE_LINHA_INICIO, values_only=True):
        val = row[QE_COL_NF - 1]  # índice 0-based
        if val is not None:
            try:
                nfs.add(int(float(str(val))))
            except (ValueError, TypeError):
                pass
    wb.close()
    log.info(f"NFs já no Quadro: {len(nfs)}")
    return nfs


# ---------------------------------------------------------------------------
# Leitura do histórico de Tipo de Cliente do Quadro de Envios
# ---------------------------------------------------------------------------

def ler_historico_tipo_cliente() -> dict:
    """
    Lê o Quadro de Envios e retorna um dicionário {doc_digits: tipo_cliente}
    com a classificação já registrada para cada CPF/CNPJ.
    Usa apenas dígitos como chave para ignorar formatação.
    """
    wb = load_workbook(QUADRO_ENVIOS, read_only=True, data_only=True)
    ws = wb[QE_ABA] if QE_ABA in wb.sheetnames else wb.active

    historico = {}
    for row in ws.iter_rows(min_row=QE_LINHA_INICIO, values_only=True):
        doc  = row[QE_COLUNAS["CPF_CNPJ_DEST"]  - 1]   # col D
        tipo = row[QE_COLUNAS["TIPO_CLIENTE"]    - 1]   # col K
        if doc and tipo and str(tipo).strip() not in ("None", ""):
            chave = str(doc).replace(".", "").replace("-", "").replace("/", "").strip()
            tipo_norm = str(tipo).strip().upper()
            # Normalizar "Pessoa Física" → "PESSOA FÍSICA"
            if tipo_norm == "PESSOA FISICA":
                tipo_norm = "PESSOA FÍSICA"
            elif tipo_norm == "PESSOA JURIDICA":
                tipo_norm = "PESSOA JURÍDICA"
            elif tipo_norm == "MUNICIPIO":
                tipo_norm = "MUNICÍPIO"
            historico[chave] = tipo_norm

    wb.close()
    log.info(f"Histórico de Tipo de Cliente: {len(historico)} CPF/CNPJs únicos")
    return historico


# ---------------------------------------------------------------------------
# Backup do Quadro de Envios
# ---------------------------------------------------------------------------

def fazer_backup():
    agora = datetime.datetime.now()
    if os.path.isdir(PASTA_QUADRO_ENVIOS_RAIZ):
        # ...\1 - QUADRO DE ENVIOS\2026\112026
        pasta_dest = os.path.join(PASTA_QUADRO_ENVIOS_RAIZ, agora.strftime("%Y"), agora.strftime("%m%Y"))
    else:
        log.warning(f"Pasta '{PASTA_QUADRO_ENVIOS_RAIZ}' não encontrada — backup salvo na pasta antiga.")
        pasta_dest = PASTA_BACKUP
    os.makedirs(pasta_dest, exist_ok=True)
    ts = agora.strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(pasta_dest, f"Quadro_Envios_backup_{ts}.xlsx")
    shutil.copy2(QUADRO_ENVIOS, dest)
    log.info(f"Backup salvo: {dest}")


# ---------------------------------------------------------------------------
# Lançamento das NFs novas no Quadro de Envios
# ---------------------------------------------------------------------------

def lancar_nfs_novas(df_base: pd.DataFrame, nfs_existentes: set, historico: dict = None) -> int:
    # Agrupa por NF (pega a primeira linha — dados de cabeçalho da NF)
    df_nf = (
        df_base
        .sort_values("NRO_NFE")
        .drop_duplicates(subset=["NRO_NFE"], keep="first")
    )

    novas = df_nf[~df_nf["NRO_NFE"].isin(nfs_existentes)].copy()
    if novas.empty:
        log.info("Nenhuma NF nova para lançar.")
        return 0

    log.info(f"NFs novas a lançar: {len(novas)}")

    wb = load_workbook(QUADRO_ENVIOS)
    ws = wb[QE_ABA] if QE_ABA in wb.sheetnames else wb.active

    # Encontra a ÚLTIMA linha com NF preenchida e insere logo após.
    # (não usa "break no primeiro None" pois pode haver gaps no meio da planilha)
    ultima_nf_linha = QE_LINHA_INICIO - 1
    for i, row in enumerate(
        ws.iter_rows(min_row=QE_LINHA_INICIO, min_col=QE_COL_NF, max_col=QE_COL_NF, values_only=True),
        start=QE_LINHA_INICIO,
    ):
        if row[0] is not None:
            ultima_nf_linha = i
    ultima_linha = ultima_nf_linha + 1
    log.info(f"Última linha com NF: {ultima_nf_linha}  →  inserindo a partir da linha {ultima_linha}")

    # Lê todas as fórmulas da última linha preenchida para replicar nas novas
    # (Distância + colunas de prevísão e entregas)
    _linha_ref = ultima_nf_linha
    _formulas_ref = {}   # {col: formula_string}
    for col in QE_COLUNAS_FORMULA_COPIAR + [QE_COLUNAS["DISTANCIA"]]:
        for offset in range(0, 10):
            v = ws.cell(row=_linha_ref - offset, column=col).value
            if v and str(v).startswith("="):
                _formulas_ref[col] = str(v)
                break
    log.info(f"Fórmulas de referência carregadas: {sorted(_formulas_ref.keys())}")

    def _adaptar_formula(col: int, nova_linha: int) -> str | None:
        """
        Adapta a fórmula da linha de referência para nova_linha.

        Regra: substitui SOMENTE referências de linha relativas, ou seja,
        sequências de dígitos precedidas por letras de coluna (e opcionalmente
        por um $ de coluna-absoluta), MAS não precedidas por $ de linha-absoluta.
          $F500   → $F{nova}   (coluna abs, linha relativa → atualiza)
          AK500   → AK{nova}   (tudo relativo → atualiza)
          $F$2    → $F$2       (linha absoluta → preserva)
          $F$1000 → $F$1000    (linha absoluta → preserva)
          0, 2    (constantes sem letra de coluna antes) → preserva
        """
        tmpl = _formulas_ref.get(col)
        if not tmpl:
            return None
        return re.sub(
            r'([A-Za-z]\$?)(\d+)',
            lambda m: m.group(0) if m.group(1).endswith('$') else m.group(1) + str(nova_linha),
            tmpl,
        )

    def _str(val):
        return str(val).strip() if val is not None and str(val) not in ("None", "nan") else ""

    def _natureza(val):
        """Remove o prefixo de código numérico do Sankhya (ex: '1010100 - ') e retorna só a descrição."""
        s = _str(val)
        # Padrão: dígitos seguidos de ' - ' no início → remove
        return re.sub(r'^\d+\s*-\s*', '', s).strip()

    def _num(val):
        try:
            v = float(val) if val is not None else None
            return abs(v) if v is not None else None
        except (ValueError, TypeError):
            return None

    for _, nf in novas.iterrows():
        linha = ultima_linha
        ultima_linha += 1

        # 1. Copiar formatação (fonte, fill, borda, alinhamento, número) da linha anterior
        _copiar_estilo_linha(ws, _linha_ref, linha)

        dt  = nf["DT_EMISSAO_NF"]
        uf  = _str(nf["UF_DESTINO"])
        doc = _str(nf["CPF_CNPJ_DEST"])
        nom = _str(nf["NOME_PARCEIRO"])

        # 2. Gravar valores
        ws.cell(row=linha, column=QE_COLUNAS["DT_EMISSAO_NF"],     value=dt)
        ws.cell(row=linha, column=QE_COLUNAS["ANO"],                value=dt.year  if dt else None)
        ws.cell(row=linha, column=QE_COLUNAS["MES"],                value=dt.month if dt else None)
        ws.cell(row=linha, column=QE_COLUNAS["CPF_CNPJ_DEST"],      value=doc)
        ws.cell(row=linha, column=QE_COLUNAS["NOME_PARCEIRO"],       value=nom)
        ws.cell(row=linha, column=QE_COLUNAS["PAIS"],                value="BRASIL")
        ws.cell(row=linha, column=QE_COLUNAS["MUNICIPIO_DESTINO"],   value=_str(nf["MUNICIPIO_DESTINO"]))
        ws.cell(row=linha, column=QE_COLUNAS["UF_DESTINO"],          value=uf)
        ws.cell(row=linha, column=QE_COLUNAS["REGIAO"],              value=_regiao_por_uf(uf))
        ws.cell(row=linha, column=QE_COLUNAS["CEP_DESTINO"],         value=nf["CEP_DESTINO"])
        ws.cell(row=linha, column=QE_COLUNAS["TIPO_CLIENTE"],        value=_tipo_cliente(doc, nom, historico))
        ws.cell(row=linha, column=QE_COLUNAS["CONS_ANALISE"],        value="Não")
        ws.cell(row=linha, column=QE_COLUNAS["EMPRESA"],             value=_empresa_por_cnpj(nf.get("CNPJ_REMETENTE")))
        ws.cell(row=linha, column=QE_COLUNAS["NRO_NFE"],             value=int(nf["NRO_NFE"]))
        ws.cell(row=linha, column=QE_COLUNAS["NATUREZA_OPERACAO"],   value=_natureza(nf["NATUREZA_OPERACAO"]))
        ws.cell(row=linha, column=QE_COLUNAS["VALOR_TOTAL_NOTA"],    value=_num(nf["VALOR_TOTAL_NOTA"]))
        ws.cell(row=linha, column=QE_COLUNAS["PESO"],                value=_num(nf["PESO"]))
        ws.cell(row=linha, column=QE_COLUNAS["VALOR_BASE_CALCUL"],   value=_num(nf["VALOR_BASE_CALCUL"]))
        ws.cell(row=linha, column=QE_COLUNAS["VALOR_ICMS"],          value=_num(nf["VALOR_ICMS"]))
        ws.cell(row=linha, column=QE_COLUNAS["VALOR_DIFAL"],         value=_num(nf["VALOR_DIFAL"]))
        ws.cell(row=linha, column=QE_COLUNAS["VALOR_FCP"],           value=_num(nf["VALOR_FCP"]))
        ws.cell(row=linha, column=QE_COLUNAS["QTD_VOLUME"],          value=_num(nf["QTD_VOLUME"]))
        ws.cell(row=linha, column=QE_COLUNAS["TRANSPORTADORA"],      value=_str(nf["TRANSPORTADORA"]))
        ws.cell(row=linha, column=QE_COLUNAS["ANO_REF"],             value=_ano_ref(dt))

        # 3. Gravar fórmulas adaptadas (Prev. Envio, Entregas, Distância)
        for col in QE_COLUNAS_FORMULA_COPIAR + [QE_COLUNAS["DISTANCIA"]]:
            f = _adaptar_formula(col, linha)
            if f:
                ws.cell(row=linha, column=col, value=f)

        # Usar esta linha como nova referência de estilo/fórmula para a próxima inserção
        _linha_ref = linha

    wb.save(QUADRO_ENVIOS)
    wb.close()
    log.info(f"{len(novas)} NF(s) lançada(s) no Quadro de Envios.")

    # O openpyxl apaga os valores calculados das fórmulas ao salvar.
    # Recalcula pelo Excel (não altera nenhum dado) para o dashboard não sair zerado.
    try:
        from modules.recalc_excel import recalcular_excel
        if not recalcular_excel(QUADRO_ENVIOS):
            log.warning("ATENÇÃO: abra o Quadro de Envios no Excel e clique em Salvar antes de "
                        "rodar o ATUALIZAR DASHBOARD.")
    except Exception as e:
        log.warning(f"Recálculo automático indisponível ({e}). Abra o Quadro no Excel e salve.")
    return len(novas)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    log.info("=" * 60)
    log.info("IMPORTAR BASE — Quadro de Envios")
    log.info("=" * 60)

    try:
        df_base        = ler_arquivo_base()
        nfs_existentes = ler_nfs_quadro()
        historico      = ler_historico_tipo_cliente()
        fazer_backup()
        total          = lancar_nfs_novas(df_base, nfs_existentes, historico)

        log.info("=" * 60)
        if total > 0:
            log.info(f"CONCLUÍDO — {total} NF(s) nova(s) lançada(s).")
            log.info("Abra o Quadro de Envios, complete frete, prazo e demais campos.")
        else:
            log.info("CONCLUÍDO — Nenhuma NF nova. Quadro já está atualizado.")
        log.info("=" * 60)

    except FileNotFoundError as e:
        log.error(str(e))
        sys.exit(1)
    except Exception as e:
        log.exception(f"Erro inesperado: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
