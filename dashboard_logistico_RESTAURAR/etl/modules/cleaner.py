# =============================================================================
# modules/cleaner.py — Limpeza, normalização e correlações
# =============================================================================

import re
import pandas as pd
import re as _re

# Mapa CNPJ → Nome padronizado (tabela_correspondencia_nome_clientes.xlsx)
_MAPA_NOMES_PADR = {"16582288000190": "COL. MONTE ALTO DA TIJUCA", "60613552000105": "ESC. AGNUS", "43932923000194": "COL. PORTO REAL", "57516228000156": "COL. ARCA EDUCA", "10548157000100": "COL. BOSQUE DOS MANACIAIS - UNIDADE CURITIBA", "50964171000109": "COL. VILA BELA", "62719230000162": "COL. BOSQUE DOS MANACIAIS - UNIDADE RECIFE", "54736220000107": "ASSOCIACAO PETROPOLITANA DE EDUCACAO E CULTURA - APEC", "23819759000104": "COLEGIO CAMINHOS E COLINAS", "3455397600103": "COL. BATISTA DE CAMAÇARI", "16582288000351": "COL. MONTE ALTO BOTAFOGO", "35936521000130": "ASSOCIAÇÃO CIDADE VIVA", "00580852000103": "COL. NAÇÕES UNIDAS", "33466264000194": "ESC. SÃO FRANCISCO DE SALES", "52405166000191": "COL. MAR ALTO", "56228516000142": "COL. BARCO A VELA", "36195301000165": "ASSOCIAÇÃO EDUCACIONAL CANADENSE DE SÃO PAULO", "50703960000196": "COL. CAMINHO NOVO", "29761234000133": "ESC. CRISTÃ MERCÊS", "33544370001110": "ASSOCIAÇÃO NOBREGA DE EDUCAÇÃO ASSISTÊ", "12947194000116": "COL. PAMPEANO", "26462477000182": "ESC. BONS VENTOS", "08335714000162": "SEMEAR SCHOOL", "60454342000103": "COL. PIO X", "04455445000117": "COL. PIO X", "59968697000131": "COL. LUMINE PORTUS", "24412991000188": "AÇÃO COMUNITARIA DE TAMANDARÉ", "03690751000175": "BAZAR RIO", "56237499000100": "COL. SETE MARES", "05209619000124": "CENTRO DE ENSINO HYARTE", "12803829000101": "CENTRO EDUCACIONAL ATOBA", "62211580000113": "CENTRO EDUCACIONAL CAMINHO", "47942708000125": "COL. MARISTA - UBERLÂNDIA", "54348190000161": "COLEGIO ARAUCARIAS", "27929089000202": "COLEGIO ARQUIDIOCESANO FAMILIA DE NAZARE", "34553976000103": "COLEGIO BATISTA DE CAMACARI", "00497704000111": "COLEGIO CATAMARA", "10847762001315": "COLEGIO DAS DAMAS DA INSTRUCAO CRISTA DO RECIFE", "10548157000290": "COL. BOSQUE DOS MANACIAIS - UNIDADE CASCAVEL", "55145110000133": "COLEGIO ILLUMINARE LTDA", "56349104000160": "COLEGIO LUZ DO ALTO", "45898691000201": "COLEGIO LUZEIROS", "28398228000109": "COLEGIO MARISTAUDI CHAMPAGNAT LTDA", "52579682000132": "COLEGIO MONTE CARMELO", "27557042000184": "COLEGIO NOSSA SENHORA DAS GRACAS", "40301427000144": "COLEGIO SACRAMENTO", "10548157000371": "COL. BOSQUE DOS MANACIAIS - UNIDADE FRANCISCO BELTÃO", "10548157000452": "COL. BOSQUE DOS MANACIAIS - UNIDADE TOLEDO", "11369387000174": "COL. CRISTÃO AGGREGARE", "55241543000192": "COL. DIVINO SABER", "45359586000113": "COL. MONTE ALTO DE NITERÓI - (VILA REAL)", "11233342000178": "COL. MONTEIRO LOBATO", "14213556000170": "COL. NAVEGANTES", "55952395000114": "COL. PÍNCARIO SOCIEDADE", "18926708000133": "COL. SERRA E MAR", "20231133000186": "ESC. WINNER", "02231133000186": "ESC. WINNER", "03760938000106": "ESC. ELENA GUERRA", "47008618000161": "ESC. BILINGUE", "51910155000104": "ESC. CARAVELAS", "43905888000114": "ESC. CRISTA SHEMAH", "07535734000114": "ESC. MOPPE LTDA", "51006232000198": "ESC. PARADISO", "50303252000168": "ESC. SER", "30767703000100": "ESC. TRINITAS", "56875117000173": "COL. MAR ABERTO", "46008618000161": "ESC. BILIGUE FOX", "04071106000137": "FUNDACAO UNIVERSIDADE FEDERAL DO ACRE", "23103586000115": "INSTITUTO DE APRENDIZAGEM NOSSA SENHORA DO BOM SUCESSO LTDA", "54392338000165": "INSTITUTO DE ENSINO SÃO JOSE - UNAI", "20872280000135": "INSTITUTO EDUCACIONAL ALETEA", "50114307000191": "COL. MONTES CLAROS", "220706438000144": "INSTITUTO PRESBETERIANO GAMMON", "63019772000438": "INSTITUTO DE ENSINO SÃO JOSE - SJC", "60849128000156": "JARDIM DE MARIA EDUCAÇÃO AFETIVA LTDA", "20736476000100": "PRÓ SECULO", "61805204000194": "SAFARI KIDS", "50303252000148": "ESC. SER", "55903254000101": "COL. FORTE DO RIO BRANCO", "57370596000139": "COL. TRINITAS", "49870101000102": "COL. VERITATIS", "07112308000178": "COL. CAMÕES"}

def _nome_padronizado(cnpj_cpf, nome_original):
    """Retorna nome padronizado se existir na tabela, senão retorna o nome original."""
    if not cnpj_cpf: return nome_original
    dig = _re.sub(r"\D","", str(cnpj_cpf))
    return _MAPA_NOMES_PADR.get(dig, nome_original)

import logging
from config import (
    CONS_ANALISE_SIM, CONS_ANALISE_NAO,
    EMPRESA_MAP, EMPRESA_DESCARTAR,
    TRANSPORTADORA_MAP,
)

# ---------------------------------------------------------------------------
# Mapa de normalização de Natureza de Operação
# Consolida variações textuais na categoria canônica do dashboard
# ---------------------------------------------------------------------------
# Mapeamento oficial baseado em correlação_tipodeoperação.xlsx
NATUREZA_MAP = {
    # VENDA
    "VENDA":                                              "VENDA",
    "Venda":                                              "VENDA",
    "NOTA DE VENDA DE PRODUTO PRODUZIDO":                 "VENDA",
    "VENDA ENTREGA FUTURA - SIMPLES FAT":                 "VENDA",
    "VENDAS DE PRODUTOS - MI":                            "VENDA",
    "1010100 - VENDAS DE PRODUTOS - MI":                  "VENDA",

    # BONIFICAÇÃO
    "BONIFICAÇÃO":                                        "BONIFICAÇÃO",
    "BONIFICACAO":                                        "BONIFICAÇÃO",
    "REMESSA DE BONIFICAÇÃO - SAIDA":                     "BONIFICAÇÃO",
    "2040100 - BONIFICACAO":                              "BONIFICAÇÃO",

    # DEVOLUÇÃO
    "DEVOLUÇÃO":                                          "DEVOLUÇÃO",
    "DEVOLUÇÃO DE VENDA - NF PRÓPRIA":                    "DEVOLUÇÃO",
    "DEVOLUÇÃO DE VENDA - NF TERCEIROS":                  "DEVOLUÇÃO",
    "DEVOLUÇÃO DE REMESSA DE BONIFICAÇÃO":                "DEVOLUÇÃO",
    "DEVOLUÇÃO DE REMESSA - ENTREGA FUTURA":              "DEVOLUÇÃO",
    "ESTORNO VDA - ENTREGA FUTURA SIMPLES FAT":           "DEVOLUÇÃO",

    # SIMPLES REMESSA
    "SIMPLES REMESSA":                                    "SIMPLES REMESSA",
    "SIMPLES REMESSA (OUTRAS)":                           "SIMPLES REMESSA",
    "OUTRAS SAÍDAS":                                      "SIMPLES REMESSA",

    # CONSIGNAÇÃO
    "CONSIGNAÇÃO":                                        "CONSIGNAÇÃO",
    "REMESSA EM CONSIGNAÇÃO - P/ VENDA":                  "CONSIGNAÇÃO",
    "VENDA CONSIGNADA":                                   "CONSIGNAÇÃO",
    "ROTORNO MERC CONSIG COMERCIALIZADA":                 "ACERTO CONSIGNAÇÃO",

    # TRANSFERÊNCIA
    "TRANSFERENCIA":                                      "TRANSFERENCIA",
    "TRANSFERENCIA ENTRE FILIAIS - SAIDA":                "TRANSFERENCIA",
    "TRANSFERENCIA ENTRE FILIAIS - SAIDA LIVROS PROPRIOS":"TRANSFERENCIA",

    # REMESSA VENDA ENTREGA FUTURA
    "REMESSA VENDA ENTREGA FUTURA":                       "REMESSA VENDA ENTREGA FUTURA",

    # OUTRAS ENTRADAS
    "OUTRAS ENTRADAS":                                    "OUTRAS ENTRADAS",
    "RETORNO SIMPLES REMESSA (OUTRAS) PROPRIO":           "OUTRAS ENTRADAS",

    # AJUSTE DE ESTOQUE
    "AJUSTE DE ESTOQUE - SAIDA/DESCARTE":                 "AJUSTE DE ESTOQUE - SAIDA/DESCARTE",
    "AJUSTE DE ESTOQUE - SAÍDA DE ESTOQUE COM TERCEIROS": "AJUSTE DE ESTOQUE - SAÍDA DE ESTOQUE COM TERCEIROS",

    # NOTA DE VENDA DE SERVIÇOS
    "NOTA DE VENDA DE SERVIÇOS COALA":                    "NOTA DE VENDA DE SERVIÇOS COALA",

    # USO FORA DO ESTABELECIMENTO
    "USO FORA DO ESTABELECIMENTO":                        "OUTRAS SAÍDAS",
    "OUTRAS SAIDAS":                                      "OUTRAS SAÍDAS",
}

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def _normalizar_cnpj_cpf(valor: str) -> str:
    """
    Remove toda pontuação e retorna apenas dígitos.
    CPF  → 11 dígitos  →  formata como  000.000.000-00
    CNPJ → 14 dígitos  →  formata como  00.000.000/0000-00
    Qualquer outro tamanho → retorna os dígitos limpos (sem formatação).
    """
    if pd.isna(valor) or str(valor).strip() == "":
        return ""

    digitos = re.sub(r"\D", "", str(valor))

    if len(digitos) == 11:                       # CPF
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    elif len(digitos) == 14:                     # CNPJ
        return (
            f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}"
            f"/{digitos[8:12]}-{digitos[12:]}"
        )
    else:
        return digitos                           # mantém dígitos sem formatação


def _normalizar_cons_analise(valor: str) -> str:
    """Padroniza Cons Analise para 'Sim' ou 'Não'."""
    v = str(valor).strip()
    if v in CONS_ANALISE_SIM:
        return "Sim"
    if v in CONS_ANALISE_NAO:
        return "Não"
    return "Não"                                 # seguro: desconhecido → descarta


def _normalizar_empresa(valor: str) -> str | None:
    """
    Retorna nome padronizado da empresa ou None se deve ser descartada.
    """
    v = str(valor).strip()
    if v in EMPRESA_DESCARTAR:
        return None
    return EMPRESA_MAP.get(v, v.upper())         # fallback: upper do original


def _normalizar_transportadora(valor: str) -> str:
    """
    1. Remove código numérico do início (ex: '232 - ').
    2. Normaliza espaços e converte para upper.
    3. Faz lookup na tabela de correlação.
    4. Se não encontrar, loga aviso e retorna o nome limpo.
    """
    if pd.isna(valor) or str(valor).strip() == "":
        return "NÃO INFORMADO"

    # Remove código numérico e separador no início: "232 - Jadlog" → "Jadlog"
    nome_limpo = re.sub(r"^\d+\s*-\s*", "", str(valor).strip())
    nome_upper = nome_limpo.upper().strip()

    if nome_upper in TRANSPORTADORA_MAP:
        return TRANSPORTADORA_MAP[nome_upper]

    # Tenta match case-insensitive direto no mapa
    for chave, destino in TRANSPORTADORA_MAP.items():
        if chave.upper() == nome_upper:
            return destino

    logger.warning(
        f"Transportadora sem mapeamento: '{valor}' (limpo: '{nome_upper}'). "
        "Adicione ao config.py → TRANSPORTADORA_MAP."
    )
    return nome_limpo.title()                    # retorna formatado como título


def _para_data(serie: pd.Series) -> pd.Series:
    """Converte série para datetime, coercionando erros para NaT."""
    return pd.to_datetime(serie, errors="coerce", dayfirst=False)


def _para_float(serie: pd.Series) -> pd.Series:
    """Converte série de strings para float, coercionando erros para 0."""
    return pd.to_numeric(serie.str.replace(",", "."), errors="coerce").fillna(0.0)


def _para_int(serie: pd.Series) -> pd.Series:
    """Converte série para int, coercionando erros para 0."""
    return pd.to_numeric(serie, errors="coerce").fillna(0).astype(int)


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def limpar_e_normalizar(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Recebe o DataFrame bruto do reader e retorna o DataFrame limpo,
    pronto para o calculador de KPIs.

    Etapas
    ------
    1. Filtra Cons Analise ≠ "Sim"
    2. Normaliza e filtra Empresa (descarta SERENA, KOALA)
    3. Normaliza CNPJ/CPF
    4. Normaliza Transportadora via tabela de correlação
    5. Converte tipos de dados (datas, floats, inteiros)
    6. Renomeia colunas para snake_case limpo
    """
    df = df_raw.copy()

    # ------------------------------------------------------------------
    # 1. Cons Analise — filtro principal
    # ------------------------------------------------------------------
    df["_cons"] = df["Cons Analise"].apply(_normalizar_cons_analise)
    antes = len(df)
    df = df[df["_cons"] == "Sim"].copy()
    df.drop(columns=["_cons"], inplace=True)
    logger.info(f"Filtro Cons Analise: {antes - len(df)} linhas removidas | "
                f"Restam: {len(df)}")

    # ------------------------------------------------------------------
    # 2. Empresa — normaliza e descarta SERENA / KOALA
    # ------------------------------------------------------------------
    df["_empresa_norm"] = df["Empresa"].apply(_normalizar_empresa)
    antes = len(df)
    df = df[df["_empresa_norm"].notna()].copy()
    logger.info(f"Filtro Empresa (Serena/Koala): {antes - len(df)} linhas removidas | "
                f"Restam: {len(df)}")
    df["Empresa"] = df["_empresa_norm"]
    df.drop(columns=["_empresa_norm"], inplace=True)

    # ------------------------------------------------------------------
    # 3. CNPJ / CPF — padronização de formato
    # ------------------------------------------------------------------
    df["cnpj_cpf_norm"] = df["CNPJ / CPF"].apply(_normalizar_cnpj_cpf)

    # ------------------------------------------------------------------
    # 4. Transportadora — correlação
    # ------------------------------------------------------------------
    df["transportadora_norm"] = df["Transportadora"].apply(_normalizar_transportadora)
    # NAO INFORMADO → PRÓPRIO
    df["transportadora_norm"] = df["transportadora_norm"].replace(
        {"NÃO INFORMADO": "PRÓPRIO", "NAO INFORMADO": "PRÓPRIO", "Não Informado": "PRÓPRIO"}
    )
    # Normalizar nome do cliente para o dashboard (via tabela de correspondência)
    if "CNPJ / CPF" in df.columns:
        df["cliente"] = df.apply(
            lambda r: _nome_padronizado(r.get("CNPJ / CPF",""), r.get("cliente", r.get("Cliente",""))),
            axis=1
        )
    # NAO INFORMADO → PRÓPRIO (vem do arquivo_base sem transportadora definida)
    df["transportadora_norm"] = df["transportadora_norm"].replace(
        {"NÃO INFORMADO": "PRÓPRIO", "NAO INFORMADO": "PRÓPRIO", "Não Informado": "PRÓPRIO"}
    )

    # ------------------------------------------------------------------
    # 5. Conversão de tipos
    # ------------------------------------------------------------------
    # Datas
    for col_orig, col_dest in [
        ("Data Coleta",              "data_coleta"),
        ("Prev. Envio",              "prev_envio"),
        ("Prev. Entrega Cliente",    "prev_entrega_cliente"),
        ("Prev. Entrega Transporadora", "prev_entrega_transportadora"),
        ("Data Efetiva da Entrega",  "data_efetiva_entrega"),
    ]:
        if col_orig in df.columns:
            df[col_dest] = _para_data(df[col_orig])

    # Financeiros
    for col_orig, col_dest in [
        ("Vr NF",            "vr_nf"),
        ("Frete",            "frete"),
        ("Base Cálculo ICM", "base_calc_icm"),
        ("ICMS",             "icms"),
        ("Difal",            "difal"),
        ("FCP",              "fcp"),
        ("Total Impostos",   "total_impostos"),
        ("Custo Embalagem",  "custo_embalagem"),
        ("Peso (kg)",        "peso_kg"),
    ]:
        if col_orig in df.columns:
            df[col_dest] = _para_float(df[col_orig])

    # Inteiros
    for col_orig, col_dest in [
        ("Volumes",    "volumes"),
        ("Nº Pedido",  "nr_pedido"),
        ("NF",         "nf"),
        ("Cfop",       "cfop"),
    ]:
        if col_orig in df.columns:
            df[col_dest] = _para_int(df[col_orig])

    # ------------------------------------------------------------------
    # 6. Renomear / padronizar colunas restantes (strings)
    # ------------------------------------------------------------------
    rename_map = {
        "Cliente":             "cliente",
        "Municipio":           "municipio",
        "UF":                  "uf",
        "Região":              "regiao",
        "Tipo de Cliente":     "tipo_cliente",
        "Empresa":             "empresa",
        "NATUREZA DA OPERAÇÃO":"natureza_operacao",
        "Modalidade":          "modalidade",
        "Entrega Concluída?":  "entrega_concluida",
        "Origem da Venda":     "origem_venda",
        "Cód Rastreio":        "cod_rastreio",
        "Obs.":                "obs",
        "Ano Ref":             "ano_ref",
        "Origem_Escola":       "origem_escola",
        "Cons Analise":        "cons_analise",
        "CNPJ / CPF":          "cnpj_cpf_raw",
        "Transportadora":      "transportadora_raw",
    }
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns},
              inplace=True)

    # ------------------------------------------------------------------
    # 7. Natureza de Operação — normalização de variações textuais
    # ------------------------------------------------------------------
    if "natureza_operacao" in df.columns:
        df["natureza_operacao"] = df["natureza_operacao"].apply(
            lambda v: NATUREZA_MAP.get(str(v).strip(), str(v).strip())
        )

    # Garante que colunas de string não tenham espaços extras
    str_cols = df.select_dtypes(include="object").columns
    # Eliminar duplicatas de colunas antes de aplicar strip
    str_cols_uniq = list(dict.fromkeys(str_cols))
    for col in str_cols_uniq:
        try:
            df[col] = df[col].str.strip()
        except Exception:
            pass

    # Remover colunas duplicadas — manter apenas a primeira ocorrência
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    logger.info(f"Normalização concluída. Shape final: {df.shape}")
    return df
