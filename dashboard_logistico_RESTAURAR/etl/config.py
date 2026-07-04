# =============================================================================
# config.py — Configurações centrais do ETL Logístico IAB & Verbum
# =============================================================================

import os

# -----------------------------------------------------------------------------
# CAMINHOS
# -----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ajuste este caminho para o local real do arquivo no seu computador
def _encontrar_onedrive_dashboards(nome_arquivo=None):
    """Localiza pasta 10-DASHIBOARDs no OneDrive de qualquer usuario."""
    import os
    candidatos = [
        os.environ.get("OneDriveCommercial", ""),
        os.environ.get("OneDrive", ""),
        os.path.join(os.environ.get("USERPROFILE",""), "OneDrive - Instituto Alfa e Beto"),
        os.path.join(os.environ.get("USERPROFILE",""), "OneDrive"),
    ]
    subfixos = [
        os.path.join("Logística-IAB", "10 - DASHIBOARDs"),
        os.path.join("Logistica-IAB", "10 - DASHIBOARDs"),
        "10 - DASHIBOARDs",
    ]
    for base in candidatos:
        if not base: continue
        for sub in subfixos:
            pasta = os.path.join(base, sub)
            if nome_arquivo:
                arq = os.path.join(pasta, nome_arquivo)
                if os.path.isfile(arq): return arq
            else:
                if os.path.isdir(pasta): return pasta
    # Fallback: relativo ao script (etl/ -> dashboard_logistico/ -> 10-DASHIBOARDs/)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    fallback = os.path.normpath(os.path.join(script_dir, "..", ".."))
    return os.path.join(fallback, nome_arquivo) if nome_arquivo else fallback

QUADRO_ENVIOS_PATH = _encontrar_onedrive_dashboards("Quadro de Envios.xlsx")

DATA_DIR          = os.path.join(BASE_DIR, "data")
PROCESSED_DIR     = os.path.join(DATA_DIR, "processed")
SNAPSHOTS_DIR     = os.path.join(DATA_DIR, "snapshots")
LOGS_DIR          = os.path.join(DATA_DIR, "logs")

# Saída principal consumida pelo dashboard
OUTPUT_PARQUET    = os.path.join(PROCESSED_DIR, "lancamentos.parquet")
OUTPUT_JSON       = os.path.join(PROCESSED_DIR, "lancamentos.json")
SNAPSHOT_LATEST   = os.path.join(SNAPSHOTS_DIR, "snapshot_latest.parquet")
CHANGE_LOG        = os.path.join(LOGS_DIR,      "change_log.csv")

# Aba de dados operacionais dentro do Quadro de Envios
ABA_LANCAMENTOS   = "Lançamentos"

# -----------------------------------------------------------------------------
# NORMALIZAÇÃO — Cons Analise
# Qualquer variação de "sim" → "Sim" | qualquer variação de "não" → "Não"
# -----------------------------------------------------------------------------
CONS_ANALISE_SIM = {"sim", "SIM", "Sim"}
CONS_ANALISE_NAO = {"nao", "NAO", "não", "NÃO", "Não", "nÃO"}

# -----------------------------------------------------------------------------
# NORMALIZAÇÃO — Empresa
# Empresas válidas; qualquer outro valor é descartado do dashboard
# -----------------------------------------------------------------------------
EMPRESA_MAP = {
    "iab":    "IAB",
    "IAB":    "IAB",
    "Iab":    "IAB",
    "verbum": "VERBUM",
    "VERBUM": "VERBUM",
    "Verbum": "VERBUM",
}
# Empresas a descartar completamente
EMPRESA_DESCARTAR = {"SERENA", "Serena", "serena", "KOALA", "Koala", "koala"}

# -----------------------------------------------------------------------------
# NORMALIZAÇÃO — Transportadoras
# Chave: nome limpo (sem código numérico, upper). Valor: nome no dashboard.
# -----------------------------------------------------------------------------
TRANSPORTADORA_MAP = {
    # --- Jadlog ---
    "JADLOG LOGISTICA S.A":                              "JADLOG",
    "JADLOG LOGISTICA S.A.":                             "JADLOG",

    # --- Braspress ---
    "BRASPRESS":                                         "BRASPRESS",
    "BRASPRESS TRANSPORTES URGENTES LTDA":               "BRASPRESS",
    "BRAPRESS":                                          "BRASPRESS",
    "BTU - BRASPRESS - UDI":                             "BRASPRESS",
    "1689 - BRASPRESS TRANSPORTES URGENTES LTDA":        "BRASPRESS",

    # --- Correios ---
    "CORREIOS":                                          "CORREIOS",
    "EMPRESA BRASILEIRA DE CORREIOS E TELEGRAFOS":       "CORREIOS",
    "FERNANDES E GARCIA AGENCIA DE SERVICOS POSTAIS LTDA -": "CORREIOS",
    "FERNANDES E GARCIA AGENCIA DE SERVI":               "CORREIOS",

    # --- TNT ---
    "TNT":                                               "TNT",
    "TNT MERCURIO CARGAS E ENCOMENDAS EXPRESSAS S/A":    "TNT",
    "TNT MERCURIO":                                      "TNT",
    "TNT MERCURIO CARGAS E ENCOMENDAS":                  "TNT",
    "TNT MERCURIO CARGAS E":                             "TNT",
    "FEDEX EXPRESS CORPORATION":                         "TNT",

    # --- Uberlog ---
    "UBERLOG":                                           "UBERLOG",
    "UBERLOG TRANSPORTES E SERVICOS LTDA":               "UBERLOG",
    "UBERLOG TRANSPORTES E SERVICOS LTDA - ME":          "UBERLOG",
    "UBER FRANQUIA":                                     "UBERLOG",
    "UBERFRANQUIA":                                      "UBERLOG",
    "UBER FRAQUIA RODOVIARIA":                           "UBERLOG",

    # --- Gollog / VRG ---
    "GOLLOG":                                            "GOLLOG",
    "VRG LINHAS AEREAS S.A.":                            "GOLLOG",
    "VRG LINHAS AEREAS":                                 "GOLLOG",
    "VRG LINHAS AERES":                                  "GOLLOG",
    "GOL LINHAS AEREAS S.A":                             "GOL CARGO",

    # --- Azul Cargo ---
    "AZUL LINHAS AEREAS BRASILEIRAS S.A.":               "AZUL CARGO",

    # --- Ativa ---
    "ATIVA DISTRIBUICAO E LOGISTICA LTDA":               "ATIVA",

    # --- Clara ---
    "CLARA TRANSPORTES E OPERACOES LOGISTICA LTDA ME":   "CLARA",
    "CLARA TRANSPORTES E OPERACOES LOGISTICA LTDA. UDI": "CLARA",
    "CLARA TRANSPORTES E OPERACOES LOGIS":               "CLARA",

    # --- AMS ---
    "AMS TRANSPORTES":                                   "AMS",

    # --- CW3 ---
    "CW3 TRANSPORTES E LOGISTICA LTDA":                  "CW3",
    "CW3":                                               "CW3",

    # --- TAM Cargo ---
    "TAM CARGO":                                         "TAM CARGO",
    "TAM LINHAS AEREAS":                                 "TAM CARGO",

    # --- Atlas ---
    "EMPRESA DE TRANSPORTES ATLAS LTDA":                 "ATLAS",
    "ATLAS LOG.":                                        "ATLAS",

    # --- Panservice ---
    "PANSERVICE":                                        "PANSERVICE",
    "PANSERVICE TRANSPORTES":                            "PANSERVICE",
    "PANSERVICE TRANSPORTES E LOCACAO LTDA":             "PANSERVICE",

    # --- Global ---
    "GLOBAL LOG":                                        "GLOBAL",
    "GLOBAL":                                            "GLOBAL",
    "GLOBAL TRANSPORTES":                                "GLOBAL",

    # --- Patrus ---
    "PATRUS TRANSPORTES URGENTES":                       "PATRUS",
    "PATRUS":                                            "PATRUS",

    # --- TG ---
    "TG TRANSPORTES GERAIS E DISTRIBUICAO LTDA":         "TG",
    "TG TRANSPORTES GERAIS E DIST LTDA":                 "TG",

    # --- Jamef ---
    "JAMEF TRANSPORTES LTDA":                            "JAMEF",
    "JAMEF TRANSPORTES LIMITADA":                        "JAMEF",
    "JAMEF ENCOMENDAS URGENTES":                         "JAMEF",
    "JAMEF":                                             "JAMEF",

    # --- Rodonaves ---
    "RODONAVES TRANSPORTES E ENCOMENDAS LTDA":           "RODONAVES",
    "RODONAVES TRANSPORTES E":                           "RODONAVES",

    # --- Antares ---
    "ANTARES TRANSPORTES LTDA - ME":                     "ANTARES",
    "ANTARES TRANSPORTES":                               "ANTARES",

    # --- Passaro ---
    "PASSARO TRANSPORTES LOGISTICA E ARMAZENAGEM LTDA":  "PASSARO",
    "PASSARO TRANSPORTES":                               "PASSARO",

    # --- Transcajuru ---
    "TRANSCAJURU":                                       "TRANSCAJURU",

    # --- Tab Log ---
    "TAB LOG TRANSPORTES LTDA":                          "TAB LOG",

    # --- Alli ---
    "ALLI LOGISTICA INTEGRADA LTDA":                     "ALLI",
    "ALLI":                                              "ALLI",

    # --- DHL ---
    "DHL EXPRESS":                                       "DHL",
    "DHL EXPRESS (BRAZIL) LTDA":                         "DHL",

    # --- Cantelle ---
    "CANTELLI":                                          "CANTELLE",
    "CANTELLE VIAGENS E TURISMO LTDA":                   "CANTELLE",

    # --- Velog ---
    "VELOG TRANSPORTES & LOGISTICA":                     "VELOG",
    "VELOG TRANSPORTES E LOGISTICA LTDA.":               "VELOG",
    "VASLOG TRANSPORTES E LOGISTICA":                    "VELOG",

    # --- Machado ---
    "MACHADO DIAS TRANSPORTES LTDA - EMITENTE":          "MACHADO",

    # --- Peixoto ---
    "PEIXOTO COMERCIO INDUSTRIA":                        "PEIXOTO",

    # --- L&C ---
    "L&C LOGISTICA E TRANSPORTES":                       "L&C",

    # --- Próprio (veículo da empresa) ---
    "PROPRIO":                                           "PRÓPRIO",
    "PRÓPRIO":                                           "PRÓPRIO",
    "LIDER ENTREGAS INTELIGENTES":                       "PRÓPRIO",
    "LEVARE TRANSPORTES":                                "PRÓPRIO",

    # --- Sem informação ---
    "NAO INFORMADO":                                     "NÃO INFORMADO",
    "NÃO INFORMADO":                                     "NÃO INFORMADO",
}

# -----------------------------------------------------------------------------
# COLUNAS USADAS NO ETL
# (apenas as que o Python realmente lê — ignora colunas calculadas do Excel)
# -----------------------------------------------------------------------------
COLUNAS_ENTRADA = [
    "Data Coleta",
    "CNPJ / CPF",
    "Cliente",
    "Municipio",
    "UF",
    "Região",
    "Tipo de Cliente",
    "Nº Pedido",
    "Cons Analise",
    "Empresa",
    "NF",
    "NATUREZA DA OPERAÇÃO",
    "Cfop",
    "Vr NF",
    "Frete",
    "Peso (kg)",
    "Base Cálculo ICM",
    "ICMS",
    "Difal",
    "FCP",
    "Total Impostos",
    "Volumes",
    "Custo Embalagem",
    "Transportadora",
    "Modalidade",
    "Prev. Envio",
    "Prev. Entrega Cliente",
    "Prev. Entrega Transporadora",
    "Entrega Concluída?",
    "Data Efetiva da Entrega",
    "Origem da Venda",
    "Cód Rastreio",
    "Obs.",
    "Ano Ref",
    "Origem_Escola",
]

# Chave composta para detecção de mudanças entre execuções
# (Data Coleta + CNPJ padronizado + Nº Pedido + Empresa)
CHAVE_SNAPSHOT = ["data_coleta", "cnpj_cpf_norm", "nr_pedido", "empresa", "ano_ref"]
