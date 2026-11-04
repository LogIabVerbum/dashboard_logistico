# -*- coding: utf-8 -*-
"""
Análise de Estoque - gerador do HTML
--------------------------------------------
Lê a planilha de movimentação do ERP e o cadastro de itens/programas e gera
"Análise de Estoque.html" na mesma pasta.
Também lê o Quadro de Envios (aba Lançamentos) e atualiza o cadastro de clientes
(cadastro_clientes.csv), marcando os clientes novos desde a última atualização.

Uso:  python gerar_analise.py
      (ou dê dois cliques, se o Python estiver associado a arquivos .py)

Requisitos: Python 3.9+ com pandas e openpyxl  ->  pip install pandas openpyxl
"""
from pathlib import Path
from datetime import datetime
import json, sys, warnings
warnings.filterwarnings("ignore")  # esconde avisos do pandas/openpyxl que não afetam o resultado
import pandas as pd
import numpy as np

# ============ CONFIGURAÇÃO ============
PASTA = Path(__file__).resolve().parent            # pasta onde estão os arquivos (dashboard_logistico\estoque)
PASTA_DASHBOARDS = PASTA.parent.parent              # ...\10 - DASHIBOARDs (funciona em qualquer computador)
# Arquivo base: procura primeiro na pasta de relatórios; se não achar, usa a pasta do script.
PASTA_BASE = PASTA_DASHBOARDS / "Arquivos de Claude Relatórios - Relatórios Logística"
NOME_BASE = "Movimentação de Estoque.xlsx"
ARQ_MOVIMENTACAO = PASTA_BASE / NOME_BASE if (PASTA_BASE / NOME_BASE).exists() else PASTA / NOME_BASE
ARQ_CADASTRO = PASTA / "cadastro_itens_programas.csv"
ARQ_KITS = PASTA / "kits_programa.csv"          # composição dos kits de cada programa (Aluno, Professor, Kit Coordenador...)
ARQ_PAPEIS = PASTA / "papeis_programa.csv"      # por item e programa: Aluno ou Professor (professor/coordenação/turma/escola)
# Quadro de Envios (aba Lançamentos): fonte do cadastro de clientes, lido a cada atualização.
PASTA_QUADRO = PASTA_DASHBOARDS
NOME_QUADRO = "Quadro de Envios.xlsx"
ARQ_QUADRO = PASTA_QUADRO / NOME_QUADRO if (PASTA_QUADRO / NOME_QUADRO).exists() else PASTA / NOME_QUADRO
ABA_QUADRO = "Lançamentos"
ARQ_CLIENTES = PASTA / "cadastro_clientes.csv"   # gerado automaticamente (não editar)
# Relatório de faturamento (ERP): valor total de cada nota. Usado na aba Margem (venda x custo).
NOME_FATURAMENTO = "Relatorio de faturamento.xlsx"
ARQ_FATURAMENTO = PASTA_BASE / NOME_FATURAMENTO if (PASTA_BASE / NOME_FATURAMENTO).exists() else PASTA / NOME_FATURAMENTO
CNPJ_EMPRESAS = {"08458084000113": "IAB", "57638518000172": "Verbum (Matriz)", "57638518000253": "Verbum (Filial)"}
RAIZES_GRUPO = {"08458084", "57638518"}   # notas entre IAB e Verbum são internas: ficam fora da margem
# Tipos de operação do faturamento que são RECEITA e os que são DEVOLUÇÃO de venda
OPS_RECEITA = {"NOTA DE VENDA DE PRODUTO PRODUZIDO", "VENDA ENTREGA FUTURA - SIMPLES FAT", "VENDA CONSIGNADA", "NF DEFINIÇÃO DE PREÇO"}
OPS_DEVOLUCAO_VENDA = {"DEVOLUÇÃO DE VENDA - NF PRÓPRIA", "DEVOLUÇÃO DE VENDA - NF TERCEIROS", "ESTORNO VDA - ENTREGA FUTURA SIMPLES FAT",
                       "RETORNO MERC CONSIG COMERCIALIZADA"}
# Naturezas da movimentação que formam o CUSTO da venda (saídas) e as devoluções que voltam ao estoque (entradas)
NAT_SAIDA_VENDA = {"NOTA DE VENDA DE PRODUTO PRODUZIDO", "REMESSA VENDA ENTREGA FUTURA", "VENDA CONSIGNADA"}
NAT_DEVOLUCAO_VENDA = {"DEVOLUÇÃO DE VENDA - NF PRÓPRIA", "DEVOLUÇÃO DE VENDA - NF TERCEIROS", "DEVOLUÇÃO DE REMESSA - ENTREGA FUTURA"}
ARQ_SAIDA = PASTA / "Análise de Estoque.html"
ARQ_MODELO = PASTA / "analise_modelo.html"
VERSAO_MODELO = "2026-10-09c"   # tem que ser igual à do analise_modelo.html
NATUREZA_IMPLANTACAO = "IMPLANTAÇÃO DE SALDOS DE ESTOQUES"
# Consignação IAB -> Verbum: o ERP só traz a entrada na Verbum; o script cria a saída no IAB.
NATUREZA_CONSIG_VERBUM = r"COMPRA PARA REVENDA POR CONSIGNA"
NATUREZA_CONSIG_IAB = "REMESSA EM CONSIGNAÇÃO PARA A VERBUM (saída criada pelo script)"
# Colunas usadas da planilha (letra -> nome interno). As demais são ignoradas.
COLUNAS = {"A": "cnpj", "C": "codigo", "F": "custo_unit", "G": "data", "H": "descricao",
           "I": "empresa", "J": "cidade", "K": "natureza", "L": "parceiro", "M": "nf",
           "N": "qtd", "Q": "tipo", "R": "uf_estado"}
EMPRESAS = {"INSTITUTO ALFA E BETO": "IAB", "VERBUM EDUCACAO S.A FILIAL": "Verbum (Filial)",
            "VERBUM EDUCACAO S.A": "Verbum (Matriz)"}
# Empresas ignoradas (erro de lançamento). Basta um trecho do nome da coluna I.
EMPRESAS_IGNORADAS = ["LINCKY CONSULTORIA"]
# E-books (descrição começando com "EB - ") são digitais, sem estoque físico: ficam FORA da análise.
EBOOK_PREFIXO = "EB - "
# TEMPORÁRIO: brindes da Coala estão lançados no estoque do IAB (a Coala ainda não tem CNPJ no ERP).
# Enquanto for True, os itens do grupo "Brinde Coala" são mostrados como empresa "Coala".
# Quando a Coala tiver CNPJ próprio no ERP, mude para False (ou apague este bloco).
COALA_BRINDES_COMO_EMPRESA = True
# Classes por descrição (coluna H). Brindes não são reclassificados.
CLASSES_POR_DESCRICAO = {
    "MIDIA": r"\bCDS?\b|\bDVDS?\b|M[ÍI]DIA|PEN ?DRIVE|BLU.?RAY|\bVHS\b",
    "ARQUIVOS": r"ARQUIVO|DOCUMENTO|PRESTA[CÇ][AÃ]O DE CONTAS|PROVAS? GR[ÁA]FICAS?|AMOSTRAS DE LIVROS ANTIGOS|PROP(OSTAS)? ANTIGAS",
    "EMBALAGEM": r"(^|\| )CAIXAS? (N[ºO°.]|IAB|VERBUM|[ÔO]MEGA|ITENS|- MODELO|MODELO|PADR)|(^|\| )CAIXA\b(?! DE HIST)(?! PORTA)|FILME|STRE[TS]C?H|FITAS? ADESIVAS?|EMBALAGEM",
    "UNIFORMES": r"\bCAMIS(A|AS|ETA|ETAS)\b|\bCAL[CÇ]AS?\b|POLO PIQUET|CAMISETA POLO|UNIFORME",
    "MAQUINAS E EQUIPAMENTOS": r"NOTEBOOK|COMPUTADOR|\bCPU|TABLET|TECLADO|\bMOUSE\b|MONITOR|IMPRESSORA|BALAN[CÇ]A|MICRO.?ONDAS|FONES? DE OUVIDO|HEADSET|MICROFONE|CH MAGN|ESTABILIZADOR|NOBREAK|PROJETOR|CELULAR",
}
# Itens que passam a ser BRINDE pela descrição (entram no grupo "Brinde IAB").
BRINDES_POR_DESCRICAO = r"CRACH[ÁA]|MEDALHA|DISPLAY"
# Itens IGNORADOS (não entram na análise), pela descrição.
IGNORAR_POR_DESCRICAO = r"PAPEL COUCH"
# Programas que NÃO ganham aba própria (consultar na aba Estoque Geral)
PROGRAMAS_SEM_ABA = {"Brindes", "Outros itens", "MIDIA", "ARQUIVOS", "UNIFORMES", "EMBALAGEM", "MAQUINAS E EQUIPAMENTOS"}
# Programas do 1º ano: seus itens saem da aba "Alfabetização 2º ano" (a tabela repete o material do 1º ano no 2º)
PROGRAMAS_1ANO = ["Alfabetização 1º ano", "Alfabetização 1º ano + Jogos Digitais"]
# Itens que também saem da aba "Alfabetização 2º ano" (pedido do Anderson): Caderno de Revisão 1, Testes (1-4), Agenda do Professor
TIRAR_DA_ALFA_2ANO = r"CADERNO DE REVIS[ÃA]O 1\b|^TESTES \(1-4\)|TESTES DE ALFABETIZA|AGENDA DO PROFESSOR"
# Programas EXCLUÍDOS do projeto (somem de todos os itens)
PROGRAMAS_EXCLUIDOS = {"Alfabetização 3º ao 5º ano", "Revisão dos Anos Iniciais I", "Revisão dos Anos Iniciais II"}
# Programas de revisão que ficam só com os itens que não repetem no programa regular do mesmo ano
REVISAO_SO_EXCLUSIVOS = {"Revisão 6º ano", "Revisão 7º ano", "Revisão 8º ano", "Revisão 9º ano"}
# Programas que aparecem dentro da aba de outro programa
PROGRAMA_ALIAS = {"Coleção Leituras – material de apoio": "Coleção Leituras"}
# Coleções da Verbum (definem que o programa é da Verbum no menu)
COLECOES_VERBUM = ("Germinar", "Enraizar", "Plenitude", "Movimente", "Consciência")


def alias_programa(p):
    """'Acessório X' não é programa: vira material do professor da coleção da Verbum daquele ano."""
    p = PROGRAMA_ALIAS.get(p, p)
    if p.startswith("Acessório "):
        ano = p[len("Acessório "):]
        return ("Germinar " if ano.startswith("Pré") else "Enraizar ") + ano
    return p


def dono_programa(p):
    return "Verbum" if p.startswith(COLECOES_VERBUM) else "IAB"
# Sempre na tabela do professor/coordenação/turma (materiais de uso da turma)
MATERIAIS_DA_TURMA = r"ACESS[ÓO]RIO|BONECOS?\b(?! QUE)|CARTAZ|CARTELA|LETRAS DO ALFABETO|LETRINHA|\bEVA\b|MATERIAL DOURADO"
PALAVRAS_PROFESSOR = r"PROFESSOR|MANUAL|GUIA|ORIENTA[CÇ]|COORDENA|GESTOR|SECRETARIA|REGISTRO|AGENDA|ALFABETIZADOR|CARTAZ|CARTELA|BONECO|GIGANTE"
# Serviços não são produtos: ficam FORA da análise.
# Identificados pela descrição (coluna H) ou pelo código. Acrescente aqui se aparecer outro.
SERVICOS_PALAVRAS = r"SERVI[CÇ]O|PROCESSAMENTO DE DADOS|PROC\. DADOS|INSTRU[CÇ][AÃ]O, TREINAMENTO|INSTR TREIN|PROPAGANDA E PUBLICIDADE|LOCA[CÇ][AÃ]O\s*,\s*MONTAGEM|TRANSFER[EÊ]NCIA ESTOQUE"
SERVICOS_CODIGOS = {2029, 2050, 2114, 10016, 10029, 10032, 10047, 10050, 10208, 10209}
# ======================================


def carregar_movimentacao():
    df = pd.read_excel(ARQ_MOVIMENTACAO, usecols=",".join(COLUNAS))
    letras = list(COLUNAS)  # read_excel mantém a ordem das colunas do arquivo
    df.columns = [COLUNAS[l] for l in letras]
    df["cnpj"] = df["cnpj"].fillna("").astype(str).str.strip()
    df["natureza"] = df["natureza"].fillna("").astype(str).str.strip()
    df["tipo"] = df["tipo"].fillna("").astype(str).str.strip().str.upper()
    df["empresa"] = df["empresa"].fillna("").astype(str).str.strip()
    ign = df["empresa"].str.upper().apply(lambda e: any(x in e for x in EMPRESAS_IGNORADAS))
    if ign.any(): print(f"  Empresas ignoradas: {int(ign.sum())} movimentos removidos")
    df = df[~ign].copy()
    df["empresa_nome"] = df["empresa"].map(EMPRESAS).fillna(df["empresa"].str.title())
    df["qtd"] = pd.to_numeric(df["qtd"], errors="coerce").fillna(0)
    df["codigo"] = pd.to_numeric(df["codigo"], errors="coerce")
    df = df.dropna(subset=["codigo"])
    df["codigo"] = df["codigo"].astype(int)
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    uf = df["uf_estado"].fillna("").astype(str).str.split(" - ", n=1, expand=True)
    df["uf"] = uf[0].str.strip()
    df["estado"] = uf[1].str.strip().str.upper() if uf.shape[1] > 1 else ""
    servico = df["codigo"].isin(SERVICOS_CODIGOS) | df["descricao"].fillna("").astype(str).str.upper().str.contains(SERVICOS_PALAVRAS, regex=True)
    ign_d = df["descricao"].fillna("").astype(str).str.upper().str.contains(IGNORAR_POR_DESCRICAO, regex=True)
    print(f"  Itens ignorados pela descrição: {df.loc[ign_d, 'codigo'].nunique()} códigos ({int(ign_d.sum())} movimentos)")
    df = df[~ign_d].copy(); servico = servico[~ign_d]
    ebook = df["descricao"].fillna("").astype(str).str.strip().str.upper().str.startswith(EBOOK_PREFIXO)
    print(f"  E-books desconsiderados: {df.loc[ebook, 'codigo'].nunique()} códigos ({int(ebook.sum())} movimentos)")
    df = df[~ebook].copy()
    servico = servico[~ebook]
    df.attrs["servicos"] = sorted(set(df.loc[servico, "codigo"]))
    df.attrs["linhas_servico"] = int(servico.sum())
    df = df[~servico].copy()
    df = espelhar_consignacao(df)
    df = anular_duplicidade_terceiros(df)
    df["mes"] = df["data"].dt.month
    df["ano"] = df["data"].dt.year
    return df


def espelhar_consignacao(df):
    """A remessa em consignação do IAB para a Verbum NÃO vem no relatório do IAB (o ERP trata como estoque
    em poder de terceiros); só aparece a entrada na Verbum ("COMPRA PARA REVENDA POR CONSIGNAÇÃO").
    Para o saldo do IAB ser só o estoque próprio, cria a SAÍDA correspondente no IAB (mesma NF, data e qtd).
    Se um dia o ERP passar a trazer essa saída, a NF já existente no IAB não é duplicada."""
    vb = df["empresa_nome"].str.startswith("Verbum") & df["natureza"].str.upper().str.contains(NATUREZA_CONSIG_VERBUM)
    if not vb.any():
        return df
    iab = df[df["empresa_nome"].eq("IAB") & df["tipo"].str.startswith("SA")]
    ja = set(zip(iab["nf"].astype(str), iab["codigo"], iab["qtd"]))
    novo = df[vb].copy()
    novo = novo[[(str(n), c, q) not in ja for n, c, q in zip(novo["nf"], novo["codigo"], novo["qtd"])]]
    if novo.empty:
        return df
    pv = df.loc[df["empresa_nome"].eq("IAB") & df["parceiro"].fillna("").astype(str).str.upper().str.contains("VERBUM EDUCA"), ["cnpj", "parceiro"]]
    cnpj_vb, nome_vb = (pv.iloc[0]["cnpj"], pv.iloc[0]["parceiro"]) if len(pv) else ("", "VERBUM EDUCACAO S.A FILIAL")
    iab_cnpj = novo["cnpj"].iloc[0]
    novo = novo.assign(empresa="INSTITUTO ALFA E BETO", empresa_nome="IAB", tipo="SAÍDA",
                       natureza=NATUREZA_CONSIG_IAB, parceiro=nome_vb, cnpj=cnpj_vb)
    print(f"  Consignação IAB → Verbum: {int(novo['qtd'].sum()):,} un. ({novo['codigo'].nunique()} itens) "
          f"saem do saldo do IAB (estavam em poder da Verbum)".replace(",", "."))
    return pd.concat([df, novo], ignore_index=True)


def anular_duplicidade_terceiros(df):
    """Envio por "Simples remessa (outras)": o ERP tratava como estoque EM PODER DE TERCEIROS e depois baixou tudo com
    "AJUSTE DE ESTOQUE - SAÍDA DE ESTOQUE COM TERCEIROS" (ex.: Pilar/AL, 21/05/2026). A saída física foi UMA só
    (combinado em 22/10/2026, vale para TODAS as análises): quando a Simples Remessa do mesmo cliente/item também baixou
    o estoque antes do ajuste, o ajuste é reduzido nessa quantidade, para ficar só uma saída. Vale para qualquer
    cliente e empresa com o mesmo caso."""
    nat = df["natureza"].astype(str).str.upper()
    sai = df["tipo"].str.startswith("SA")
    aj = sai & nat.str.contains("SAÍDA DE ESTOQUE COM TERCEIROS")
    sr = sai & nat.eq("SIMPLES REMESSA (OUTRAS)")
    if not aj.any() or not sr.any():
        return df
    df = df.copy()
    chave = (df["empresa_nome"].astype(str) + "|" + df["parceiro"].fillna("").astype(str).str.upper().str.strip()
             + "|" + df["cidade"].fillna("").astype(str).str.upper())
    anulado, casos = 0.0, set()
    for (c, cod), rows in df[aj].groupby([chave[aj], df.loc[aj, "codigo"]]):
        rem = float(df.loc[sr & chave.eq(c) & df["codigo"].eq(cod) & (df["data"] <= rows["data"].max()), "qtd"].sum())
        for idx in rows.sort_values("data", ascending=False).index:    # tira do ajuste o que a remessa já baixou
            if rem <= 0:
                break
            tira = min(rem, float(df.at[idx, "qtd"]))
            df.at[idx, "qtd"] = float(df.at[idx, "qtd"]) - tira
            rem -= tira; anulado += tira; casos.add(c.split("|", 2)[1] + " / " + c.split("|", 2)[2].title())
    if anulado:
        print(f"  Em poder de terceiros: {int(anulado):,} un. baixadas em dobro (simples remessa + ajuste) anuladas - "
              f"ficou só uma saída ({'; '.join(sorted(casos))})".replace(",", "."))
    return df


def carregar_cadastro():
    if not ARQ_CADASTRO.exists():
        print(f"Aviso: {ARQ_CADASTRO.name} não encontrado - programas ficarão em branco.")
        return pd.DataFrame(columns=["Codigo", "Grupo", "Empresa_item", "Programas"])
    cad = pd.read_csv(ARQ_CADASTRO, sep=";", encoding="utf-8-sig")
    cad["Codigo"] = pd.to_numeric(cad["Codigo"], errors="coerce").astype("Int64")
    cad["Programas"] = cad["Programas"].fillna("").map(filtrar_programas)
    return cad.dropna(subset=["Codigo"])


def filtrar_programas(txt):
    """Tira os programas excluídos e, nas Revisões 6º-9º, mantém só itens que não repetem em outros programas."""
    ps = [p.strip() for p in str(txt).split("|") if p.strip() and p.strip() not in PROGRAMAS_EXCLUIDOS]
    # Revisão Nº ano: sai o item que já está no programa regular do mesmo ano (Coleção Ômega Nº ano);
    # ficam só os materiais próprios da revisão (livros de anos anteriores)
    ps = [p for p in ps if not (p in REVISAO_SO_EXCLUSIVOS and p.replace("Revisão", "Coleção Ômega") in ps)]
    return " | ".join(ps)


def conciliacao_iab_verbum(df, cad):
    """Itens do IAB enviados em consignação para a Verbum: no fim, o saldo deles na Verbum tem que ser ZERO.
    Ajustes de entrada/saída contam como movimento normal (já acertam o estoque).
    Lista os itens do IAB com saldo diferente de zero na Verbum."""
    dono = dict(zip(cad["Codigo"].astype(str), cad["Empresa_item"].fillna("")))
    sinal = df["tipo"].map({"ENTRADA": 1, "SAÍDA": -1}).fillna(0)
    vb = df[df["empresa_nome"].str.startswith("Verbum")].assign(mov=df["qtd"] * sinal)
    out = {}
    for cod, g in vb.groupby("codigo"):
        if dono.get(str(cod)) != "IAB":
            continue
        ent = float(g.loc[g["tipo"].eq("ENTRADA"), "qtd"].sum())
        sai = float(g.loc[g["tipo"].str.startswith("SA"), "qtd"].sum())
        saldo = ent - sai
        acoes = []
        if round(saldo) < 0:
            acoes.append(f"Verbum enviou {abs(saldo):,.0f} ao cliente sem dar entrada no que recebeu do IAB: lançar RECEBIMENTO na Verbum e SAÍDA no IAB de {abs(saldo):,.0f}")
        elif round(saldo) > 0:
            acoes.append(f"Sobra {saldo:,.0f} na Verbum: devolver ao IAB ou conferir (tem que zerar)")
        out[cod] = {"entrada_vb": ent, "saida_vb": sai, "saldo_vb": saldo, "acoes": acoes}
    return out


def aplicar_coala(df, cad, avisar=False):
    if COALA_BRINDES_COMO_EMPRESA:
        coala = set(cad.loc[cad["Grupo"].eq("Brinde Coala"), "Codigo"].astype(str))
        mask = df["empresa_nome"].eq("IAB") & df["codigo"].astype(str).isin(coala)
        df = df.copy(); df.loc[mask, "empresa_nome"] = "Coala"
        if avisar:
            print(f"  Brindes Coala: {df.loc[mask, 'codigo'].nunique()} itens ({int(mask.sum())} movimentos) movidos do IAB para Coala")
    return df


def saidas_por_destino(df, cad, cli):
    """Matriz item x destino: saldo inicial (implantação + ajustes), compras/entradas, e saídas líquidas
    (saída − devolução) por Pessoa Física, Pessoa Jurídica, cada Município e cada Escola."""
    import re
    df = aplicar_coala(df, cad)
    dig = lambda v: re.sub(r"\D", "", str(v or ""))
    info = {}
    if cli is not None and len(cli):
        for _, c in cli.iterrows():
            info[dig(c["CNPJ / CPF"])] = (str(c["tipo"] or ""), str(c["cliente"] or "").strip(),
                                          "" if pd.isna(c["municipio"]) else str(c["municipio"]).strip().title(),
                                          "" if pd.isna(c["uf"]) else str(c["uf"]).strip())
    nat = df["natureza"].str.upper(); parc = df["parceiro"].fillna("").astype(str).str.strip()
    interno = parc.str.upper().str.contains("ALFA E BETO|VERBUM EDUCA", regex=True)
    destinos = {}
    def destino(cnpj, nome):
        d = dig(cnpj); t, n, mun, uf = info.get(d, ("", "", "", ""))
        nome_u = nome.upper()
        if not t:
            if re.search(r"^MUNIC[IÍ]PIO|PREFEITURA|SECRETARIA MUNICIPAL", nome_u): t = "MUNICÍPIO"
            elif re.search(r"COL[EÉ]GIO|ESCOLA|EDUCACIONAL|CENTRO DE EDUCA|INSTITUTO EDUCA|CRECHE", nome_u): t = "ESCOLA"
            elif len(d) == 11: t = "PESSOA FÍSICA"
            elif len(d) == 14: t = "PESSOA JURÍDICA"
            else: t = ""
        if t == "PESSOA FÍSICA": return ("PF", "Pessoa Física", "PF")
        if t == "MUNICÍPIO":
            rot = f"{mun}/{uf}" if mun else (n or nome).title()
            return ("M:" + rot.upper(), rot, "MUN")
        if t == "ESCOLA":
            rot = (n or nome).strip()
            return ("E:" + rot.upper(), rot, "ESC")
        if t: return ("PJ", "Pessoa Jurídica", "PJ")
        return ("NI", "Outras saídas (sem cliente identificado)", "NI")
    cache = {}
    col, tipo_col = [], []
    for cnpj, nome, it, nt in zip(df["cnpj"], parc, interno, nat):
        if "REQUISI" in nt:
            k = ("REQ", "Requisição interna (uso próprio)", "REQ")
        elif it:
            k = ("INT", "Transferências IAB ↔ Verbum", "INT")
        else:
            key = (cnpj, nome)
            if key not in cache: cache[key] = destino(cnpj, nome)
            k = cache[key]
        destinos[k[0]] = (k[1], k[2]); col.append(k[0])
    df = df.assign(dest=col)
    impl = nat.eq(NATUREZA_IMPLANTACAO); aj = nat.str.contains("AJUSTE"); dev = nat.str.contains("DEVOLU")
    loc = nat.str.contains("TRANSFER[EÊ]NCIA ENTRE LOCAIS", regex=True)   # troca de local no mesmo estoque: não entra
    df = df[~loc]; nat = nat[~loc]; impl = impl[~loc]; aj = aj[~loc]; dev = dev[~loc]
    ent = df["tipo"].eq("ENTRADA"); sai = df["tipo"].str.startswith("SA")
    sinal = np.where(ent, 1, np.where(sai, -1, 0))
    q = df["qtd"]
    df = df.assign(
        ini=np.where(impl, q, 0) + np.where(aj, q * sinal, 0),
        compras=np.where(ent & ~impl & ~aj & ~dev, q, 0),
        cli_q=np.where(sai & ~aj, q, 0) - np.where(ent & dev & ~aj, q, 0))
    base = df.groupby(["empresa_nome", "codigo"]).agg(ini=("ini", "sum"), compras=("compras", "sum")).reset_index()
    m = df[df["cli_q"] != 0].groupby(["empresa_nome", "codigo", "dest"])["cli_q"].sum()
    out = {}
    for (e, c, d), v in m.items():
        if round(v) != 0: out.setdefault((e, int(c)), {})[d] = float(v)
    linhas = {(e, int(c)): {"ini": float(i), "comp": float(cp), "dest": out.get((e, int(c)), {})}
              for e, c, i, cp in zip(base["empresa_nome"], base["codigo"], base["ini"], base["compras"])}
    return linhas, {k: {"rot": v[0], "tipo": v[1]} for k, v in destinos.items()}


def margem_vendas(df, cli):
    """Venda x custo. Receita = valor das notas de venda (Relatório de faturamento), já descontadas as devoluções.
    Custo = custo das saídas de venda no estoque (Movimentação), já descontadas as devoluções que voltaram.
    O relatório de faturamento só traz o valor TOTAL da nota (não o preço por item), por isso a margem é
    calculada por nota, por cliente, por tipo de cliente, por empresa e por mês."""
    import re
    if not ARQ_FATURAMENTO.exists():
        print(f"Aviso: {NOME_FATURAMENTO} não encontrado - aba Margem ficará vazia.")
        return None
    f = pd.read_excel(ARQ_FATURAMENTO)
    f.columns = [str(c).strip().upper() for c in f.columns]
    dig = lambda v: re.sub(r"\D", "", str(v or ""))
    f["op"] = f["TIPO_OPERACAO"].fillna("").astype(str).str.strip().str.upper()
    f["emp"] = f["CNPJ_REMETENTE"].map(dig).map(CNPJ_EMPRESAS)
    f["cli"] = f["CPF_CNPJ_DESTINATARIO"].map(dig)
    f["valor"] = pd.to_numeric(f["VALOR_TOTAL_NOTA"], errors="coerce").fillna(0)
    f["data"] = pd.to_datetime(f["DT_EMISSAO_NF"], dayfirst=True, errors="coerce")
    f["nf"] = f["NRO_NFE"].map(lambda x: "" if pd.isna(x) else str(int(x)) if isinstance(x, (int, float)) else str(x).strip())
    interno = lambda c: c[:8] in RAIZES_GRUPO
    f = f[f["emp"].notna() & ~f["cli"].map(interno)]
    rec = f[f["op"].isin(OPS_RECEITA | OPS_DEVOLUCAO_VENDA)].copy()
    rec.loc[rec["op"].isin(OPS_DEVOLUCAO_VENDA), "valor"] = -rec["valor"].abs()   # devolução sempre reduz a receita
    # custo no estoque
    m = df.copy()
    m["nat"] = m["natureza"].str.upper().str.strip()
    m["cli"] = m["cnpj"].map(dig)
    m = m[~m["cli"].map(interno)]
    m["custo"] = m["qtd"] * pd.to_numeric(m["custo_unit"], errors="coerce").fillna(0)
    sai, ent = m["tipo"].str.startswith("SA"), m["tipo"].eq("ENTRADA")
    cmv = m[(sai & m["nat"].isin(NAT_SAIDA_VENDA)) | (ent & m["nat"].isin(NAT_DEVOLUCAO_VENDA))].copy()
    cmv.loc[cmv["tipo"].eq("ENTRADA"), "custo"] *= -1
    bon = m[(sai & m["nat"].str.startswith("REMESSA DE BONIFICA")) | (ent & m["nat"].str.startswith("DEVOLUÇÃO DE REMESSA DE BONIFICA"))].copy()
    bon.loc[bon["tipo"].eq("ENTRADA"), "custo"] *= -1
    for t in (cmv, bon):
        t["emp"] = t["empresa_nome"]
    # cadastro de clientes (tipo, município)
    info = {}
    if cli is not None and len(cli):
        for _, c in cli.iterrows():
            info[dig(c["CNPJ / CPF"])] = (str(c["tipo"] or ""), "" if pd.isna(c["municipio"]) else str(c["municipio"]).strip().title(),
                                          "" if pd.isna(c["uf"]) else str(c["uf"]).strip())
    nomes = dict(zip(rec["cli"], rec["CLIENTE_DESTINATARIO"].fillna("").astype(str).str.strip()))
    for c, n in zip(m["cli"], m["parceiro"].fillna("").astype(str).str.strip()):
        nomes.setdefault(c, n)
    def tipo(c):
        t = info.get(c, ("",))[0]
        if t: return t
        n = nomes.get(c, "").upper()
        if re.search(r"MUNIC[IÍ]P|PREFEITURA|SECRETARIA MUNICIPAL|FUNDO MUNICIPAL", n): return "MUNICÍPIO"
        if re.search(r"COL[EÉ]GIO|ESCOLA|EDUCACIONAL|CRECHE", n): return "ESCOLA"
        return "PESSOA FÍSICA" if len(c) == 11 else "PESSOA JURÍDICA"
    # por cliente x empresa
    k = ["emp", "cli"]
    g = (rec.groupby(k).agg(receita=("valor", "sum"), nnf=("nf", "nunique"), ult=("data", "max")).join(
         cmv.groupby(k)["custo"].sum().rename("custo"), how="outer").join(
         bon.groupby(k)["custo"].sum().rename("bonif"), how="outer")).fillna({"receita": 0, "custo": 0, "bonif": 0, "nnf": 0}).reset_index()
    g = g[(g["receita"].round(2) != 0) | (g["custo"].round(2) != 0) | (g["bonif"].round(2) != 0)]
    clientes = [{"emp": r.emp, "doc": r.cli, "nome": nomes.get(r.cli, ""), "tipo": tipo(r.cli),
                 "mun": info.get(r.cli, ("", "", ""))[1], "uf": info.get(r.cli, ("", "", ""))[2],
                 "nnf": int(r.nnf), "rec": round(float(r.receita), 2), "cus": round(float(r.custo), 2), "bon": round(float(r.bonif), 2),
                 "ult": "" if pd.isna(r.ult) else r.ult.strftime("%d/%m/%Y")} for r in g.itertuples(index=False)]
    # por nota (vendas comuns: custo das saídas com o mesmo número de NF e empresa)
    cnf = cmv[cmv["tipo"].str.startswith("SA")].assign(nf=lambda x: x["nf"].astype(str).str.replace(r"\.0$", "", regex=True)).groupby(["emp", "nf"])["custo"].sum()
    notas = []
    for r in rec.sort_values("data").itertuples(index=False):
        if r.op in OPS_DEVOLUCAO_VENDA: continue
        c = cnf.get((r.emp, r.nf))
        notas.append({"nf": r.nf, "data": "" if pd.isna(r.data) else r.data.strftime("%d/%m/%Y"), "ord": "" if pd.isna(r.data) else r.data.strftime("%Y%m%d"),
                      "emp": r.emp, "doc": r.cli, "nome": nomes.get(r.cli, ""), "tipo": tipo(r.cli), "op": r.op.title(),
                      "rec": round(float(r.valor), 2), "cus": None if c is None else round(float(c), 2)})
    # por mês
    mes = lambda d: d.dt.strftime("%Y-%m")
    mm = pd.concat([rec.assign(mes=mes(rec["data"])).groupby(["emp", "mes"])["valor"].sum().rename("rec"),
                    cmv.assign(mes=mes(cmv["data"])).groupby(["emp", "mes"])["custo"].sum().rename("cus"),
                    bon.assign(mes=mes(bon["data"])).groupby(["emp", "mes"])["custo"].sum().rename("bon")], axis=1).fillna(0).reset_index()
    meses = [{"emp": r.emp, "mes": r.mes, "rec": round(float(r.rec), 2), "cus": round(float(r.cus), 2), "bon": round(float(r.bon), 2)}
             for r in mm.itertuples(index=False) if r.mes and r.mes != "NaT"]
    tr, tc = sum(c["rec"] for c in clientes), sum(c["cus"] for c in clientes)
    print(f"  Margem: receita {tr:,.0f} | custo {tc:,.0f} | {len(clientes)} clientes | {len(notas)} notas de venda".replace(",", "."))
    per = [d for d in (f["data"].min(), f["data"].max())]
    return {"clientes": clientes, "notas": notas, "meses": meses, "fonte": ARQ_FATURAMENTO.name,
            "periodo": ["" if pd.isna(d) else d.strftime("%d/%m/%Y") for d in per]}


def avaliacao_estoque(df, cad):
    df = aplicar_coala(df, cad, avisar=True)
    impl = df["natureza"].str.upper().eq(NATUREZA_IMPLANTACAO)
    ent = df["tipo"].eq("ENTRADA") & ~impl
    sai = df["tipo"].str.startswith("SA")
    df = df.assign(s_ini=df["qtd"].where(impl, 0), s_ent=df["qtd"].where(ent, 0),
                   s_sai=df["qtd"].where(sai, 0), c=df["custo_unit"].where(df["custo_unit"] > 0))
    g = (df.sort_values("data").groupby(["empresa_nome", "codigo"])
           .agg(descricao=("descricao", "last"), saldo_inicial=("s_ini", "sum"),
                entradas=("s_ent", "sum"), saidas=("s_sai", "sum"),
                custo=("c", "last"), ult=("data", "max"))
           .reset_index())
    g["saldo_atual"] = g["saldo_inicial"] + g["entradas"] - g["saidas"]
    cad = cad.copy(); cad["Descricao_cadastro"] = cad.get("Descricao_cadastro", "")
    g = g.merge(cad[["Codigo", "Grupo", "Programas", "Empresa_item", "Descricao_cadastro"]], left_on="codigo",
                right_on="Codigo", how="left").drop(columns="Codigo")
    g["Grupo"] = g["Grupo"].fillna("Não cadastrado")
    desc_up = (g["descricao"].fillna("").astype(str) + " | " + g["Descricao_cadastro"].fillna("").astype(str)).str.upper()
    br = desc_up.str.contains(BRINDES_POR_DESCRICAO, regex=True) & ~g["Grupo"].str.startswith("Brinde")
    g.loc[br, "Grupo"] = np.where(g.loc[br, "empresa_nome"].str.startswith("Verbum"), "Brinde Verbum", "Brinde IAB")
    g.loc[br, "Programas"] = "Brindes"
    for classe, padrao in CLASSES_POR_DESCRICAO.items():
        m = desc_up.str.contains(padrao, regex=True) & ~g["Grupo"].str.startswith("Brinde")
        g.loc[m, "Grupo"] = classe
        g.loc[m, "Programas"] = classe
    # Tudo o que continuar sem programa (e não estiver fora do escopo) vai para "Outros itens"
    sem = g["Programas"].fillna("").str.strip().eq("") & ~g["Grupo"].str.startswith(("Fora do escopo", "Serviço"))
    g.loc[sem, "Programas"] = "Outros itens"
    g["Programas"] = g["Programas"].fillna("")
    conc = conciliacao_iab_verbum(df, cad)
    def alerta(r):
        c = conc.get(r.codigo)
        if c and c["acoes"] and r.empresa_nome.startswith("Verbum"):
            return "Consignação IAB→Verbum: " + " | ".join(c["acoes"])
        dono = str(r.Empresa_item) if pd.notna(r.Empresa_item) else ""
        if dono == "Verbum" and r.empresa_nome == "IAB":
            if round(r.saldo_atual) == 0 and r.entradas > 0:
                return "Item da Verbum lançado no estoque do IAB (entrada e saída se anulam: lançamento corrigido)"
            return "Item da Verbum movimentado no estoque do IAB: conferir empresa ou código do lançamento"
        if r.saldo_atual < 0:
            return f"Saldo negativo: saíram {abs(r.saldo_atual):,.0f} a mais que o saldo inicial + entradas. Lançar ajuste de entrada ou conferir as saídas"
        return ""
    linhas = []
    for r in g.itertuples(index=False):
        progs = [p.strip() for p in r.Programas.split("|") if p.strip()]
        linhas.append({
            "emp": r.empresa_nome, "cod": int(r.codigo), "desc": str(r.descricao).strip(),
            "grupo": r.Grupo, "progs": progs,
            "ini": float(r.saldo_inicial), "ent": float(r.entradas), "sai": float(r.saidas),
            "atual": float(r.saldo_atual), "alerta": alerta(r),
            "custo": None if pd.isna(r.custo) else round(float(r.custo), 2),
            "ult": "" if pd.isna(r.ult) else r.ult.strftime("%d/%m/%Y")})
    return linhas, conc


def carregar_clientes():
    """Lê a aba Lançamentos do Quadro de Envios e monta o cadastro de clientes (1 linha por CNPJ/CPF)."""
    if not ARQ_QUADRO.exists():
        print(f"Aviso: {NOME_QUADRO} não encontrado - aba Clientes ficará vazia.")
        return pd.DataFrame(), []
    q = pd.read_excel(ARQ_QUADRO, sheet_name=ABA_QUADRO)
    q.columns = [str(c).strip() for c in q.columns]
    q = q.dropna(subset=["CNPJ / CPF"]).copy()
    q["CNPJ / CPF"] = q["CNPJ / CPF"].astype(str).str.strip()
    for c in ["Empresa", "Tipo de Cliente", "NATUREZA DA OPERAÇÃO"]:
        q[c] = q[c].fillna("").astype(str).str.strip().str.upper()
    q["Data Coleta"] = pd.to_datetime(q["Data Coleta"], errors="coerce")
    q["NF"] = q["NF"].map(lambda x: "" if pd.isna(x) else (str(int(x)) if isinstance(x, (int, float)) else str(x).strip()))
    q = q.sort_values("Data Coleta")
    moda = lambda x: x[x != ""].mode().iloc[0] if (x != "").any() else ""
    nfs = lambda x: ", ".join(dict.fromkeys(v for v in x if v))
    cli = q.groupby("CNPJ / CPF").agg(
        cliente=("Cliente", "last"), tipo=("Tipo de Cliente", moda), municipio=("Municipio", "last"),
        uf=("UF", "last"), regiao=("Região", "last"), empresas=("Empresa", lambda x: ", ".join(sorted(set(v for v in x if v)))),
        primeiro=("Data Coleta", "min"), ultimo=("Data Coleta", "max"), valor=("Vr NF", "sum"))
    unicas = q[q["NF"] != ""].drop_duplicates(["CNPJ / CPF", "Empresa", "NF"])
    cli["qtd_nfs"] = unicas.groupby("CNPJ / CPF").size()
    for emp in ["IAB", "VERBUM", "COALA"]:
        cli["nfs_" + emp.lower()] = unicas[unicas["Empresa"] == emp].groupby("CNPJ / CPF")["NF"].agg(nfs)
    cli = cli.reset_index().fillna({"qtd_nfs": 0, "nfs_iab": "", "nfs_verbum": "", "nfs_coala": ""})
    # clientes novos = CNPJ/CPF que não estavam no cadastro da atualização anterior
    anteriores = set()
    if ARQ_CLIENTES.exists():
        anteriores = set(pd.read_csv(ARQ_CLIENTES, sep=";", encoding="utf-8-sig", dtype=str)["CNPJ / CPF"].str.strip())
    cli["novo"] = ~cli["CNPJ / CPF"].isin(anteriores) if anteriores else False
    cli.drop(columns="novo").to_csv(ARQ_CLIENTES, sep=";", index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
    d = lambda v: "" if pd.isna(v) else v.strftime("%d/%m/%Y")
    lista = [{"doc": r["CNPJ / CPF"], "nome": str(r.cliente).strip(), "tipo": r.tipo or "NÃO INFORMADO",
              "mun": "" if pd.isna(r.municipio) else str(r.municipio).title(), "uf": "" if pd.isna(r.uf) else str(r.uf),
              "emp": r.empresas, "nnf": int(r.qtd_nfs), "nf_iab": r.nfs_iab, "nf_verbum": r.nfs_verbum, "nf_coala": r.nfs_coala,
              "valor": round(float(r.valor or 0), 2), "prim": d(r.primeiro), "ult": d(r.ultimo),
              "ult_ord": "" if pd.isna(r.ultimo) else r.ultimo.strftime("%Y%m%d"), "novo": bool(r.novo)}
             for _, r in cli.iterrows()]
    return cli, lista


def chave_versao(d):
    """Nome normalizado para achar versões do mesmo item (ignora acento, pontuação, ano e edição)."""
    import unicodedata, re
    d = unicodedata.normalize("NFKD", str(d).lower())
    d = "".join(c for c in d if not unicodedata.combining(c)).replace("°", "º")
    d = re.sub(r"[^a-z0-9º ]", " ", d)
    d = re.sub(r"\b(\d)a ed(icao)?\b|\b20[12]\d\b|\bsite\b", " ", d)
    return re.sub(r"\s+", " ", d).strip()


_GENERICAS = {"manual","do","da","de","dos","das","professor","livro","colecao","iab","ano","e","o","a","aluno",
              "volume","vol","cjto","com","conjunto","programa","alfa","beto","para","no","na","em"}


def palavras_chave(d):
    import unicodedata, re
    d = unicodedata.normalize("NFKD", str(d).lower())
    d = "".join(c for c in d if not unicodedata.combining(c))
    out = set()
    for w in re.sub(r"[^a-z0-9 ]", " ", d).split():
        if w in _GENERICAS or w.isdigit() or len(w) < 2: continue
        out.add(re.sub(r"(es|s)$", "", w) if len(w) > 4 else w)
    return out


def chave_tokens(d):
    """Nome como conjunto de palavras (sem palavras genéricas), para casar 'Coleção IAB de Ciências – Livro do Aluno 1º ano' com 'Ciências 1º ano'."""
    import unicodedata, re
    d = unicodedata.normalize("NFKD", str(d).lower())
    d = "".join(c for c in d if not unicodedata.combining(c)).replace("º", " ").replace("°", " ")
    rom = {"i": "1", "ii": "2", "iii": "3", "iv": "4"}
    gen = {"colecao", "iab", "de", "da", "do", "dos", "das", "e", "o", "a", "livro", "aluno", "ano", "vol", "volume", "2026", "2025", "ed"}
    return frozenset(rom.get(w, w) for w in re.sub(r"[^a-z0-9 ]", " ", d).split() if w not in gen)


def enriquecer_programas(est):
    """Papel (aluno/professor) por programa, versões com o mesmo nome e lista de programas com aba."""
    import re
    pap = {}
    if ARQ_PAPEIS.exists():
        p = pd.read_csv(ARQ_PAPEIS, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
        for c, pr, pa in zip(p["Codigo"].str.strip(), p["Programa"], p["Papel"]):
            pap[(c, alias_programa(pr))] = pa
    kits, kit_nome, kit_tok = {}, {}, {}
    tabela_ok, chaves_ok = {}, {}
    if ARQ_KITS.exists():
        k = pd.read_csv(ARQ_KITS, sep=";", encoding="utf-8-sig", dtype=str).fillna("")
        desc_est = {str(r["cod"]): r["desc"] for r in est}
        conflitos = set()
        tabela_ok, chaves_ok = {}, {}   # programas com divisão por código e nomes cujo código oficial existe no ERP
        for c, pr, kt, it in zip(k["Codigo"].str.strip(), k["Programa"], k["Kit"], k["Item"]):
            pr = alias_programa(pr)
            a, b = palavras_chave(it), palavras_chave(desc_est.get(c, ""))
            if c in desc_est and a and b and not (a & b):
                # o código da tabela aponta para outro item no ERP: vale o nome, não o código
                if (c, pr) not in conflitos:
                    print(f"  Aviso: cód. {c} em '{pr}' é '{it}' na tabela, mas '{desc_est[c]}' no ERP - código ignorado")
                conflitos.add((c, pr))
                kit_nome.setdefault(pr, {}).setdefault(chave_versao(it), set()).add(kt)
                continue
            kits.setdefault((c, pr), []).append(kt)
            if c in desc_est:
                tabela_ok.setdefault(pr, set()).add(c)
                chaves_ok.setdefault(pr, set()).update({chave_versao(it), chave_tokens(it)})
            kit_nome.setdefault(pr, {}).setdefault(chave_versao(it), set()).add(kt)
            kit_tok.setdefault(pr, {}).setdefault(chave_tokens(it), set()).add(kt)
    grupos = {}
    for r in est:
        grupos.setdefault(chave_versao(r["desc"]), set()).add(r["cod"])
    # itens do 1º ano (e suas versões antigas) que a tabela repete na Alfabetização 2º ano
    so_1ano = set()
    if ARQ_KITS.exists():
        cod_de = lambda progs: set(k.loc[k["Programa"].isin(progs), "Codigo"].str.strip())
        nome_de = lambda progs: {f(i) for i in k.loc[k["Programa"].isin(progs), "Item"] for f in (chave_versao, chave_tokens)}
        c1, c2 = cod_de(PROGRAMAS_1ANO), cod_de(["Ensino Estruturado 2º ano"])
        n1, n2 = nome_de(PROGRAMAS_1ANO), nome_de(["Ensino Estruturado 2º ano"])
    for r in est:
        r["abas"] = [] if r["cod"] >= 100000 else sorted({alias_programa(p) for p in r["progs"]} - PROGRAMAS_SEM_ABA)  # códigos 100000+ são kits fechados
        if "Alfabetização 2º ano" in r["abas"] and ARQ_KITS.exists():
            c, nk = str(r["cod"]), {chave_versao(r["desc"]), chave_tokens(r["desc"])}
            do_1 = (c in c1 or bool(nk & n1) or re.search(r"\b1[ºO°] ANO\b", r["desc"].upper())
                    or re.search(TIRAR_DA_ALFA_2ANO, r["desc"].upper()))
            do_2 = c in c2 or bool(nk & n2)
            if do_1 and not do_2:
                r["abas"].remove("Alfabetização 2º ano"); so_1ano.add(r["cod"])
        r["papel"] = {}
        r["kits"] = {}
        for p in r["abas"]:
            pa = pap.get((str(r["cod"]), p))
            if not pa:
                pa = "Professor" if re.search(PALAVRAS_PROFESSOR, r["desc"].upper()) else "Aluno"
            if re.search(MATERIAIS_DA_TURMA, r["desc"].upper()):
                pa = "Professor"
            r["papel"][p] = pa
            if (str(r["cod"]), p) not in kits and p in tabela_ok and not (
                    kit_nome.get(p, {}).get(chave_versao(r["desc"])) or kit_tok.get(p, {}).get(chave_tokens(r["desc"]))):
                # não está na divisão oficial nem é versão (mesmo nome) de um item dela
                r["kits"][p] = ["Outras versões"]
                continue
            kt = (kits.get((str(r["cod"]), p)) or sorted(kit_nome.get(p, {}).get(chave_versao(r["desc"]), []))
                  or sorted(kit_tok.get(p, {}).get(chave_tokens(r["desc"]), [])))
            if not kt:
                kt = ["Aluno"] if pa == "Aluno" else ["Professor"]
            elif re.search(MATERIAIS_DA_TURMA, r["desc"].upper()) and "Aluno" in kt:
                kt = [x for x in kt if x != "Aluno"] or ["Professor"]
            r.setdefault("kits", {})[p] = list(dict.fromkeys(kt))
        r["versoes"] = sorted(grupos[chave_versao(r["desc"])] - {r["cod"]})
    if so_1ano:
        print(f"  Alfabetização 2º ano: {len(so_1ano)} itens do 1º ano retirados da aba")
    return est


def main():
    for arq in (ARQ_MOVIMENTACAO, ARQ_MODELO):
        if not arq.exists():
            sys.exit(f"Arquivo não encontrado: {arq}")
    df = carregar_movimentacao()
    cad = carregar_cadastro()
    cli_df, clientes = carregar_clientes()
    est, conc = avaliacao_estoque(df, cad)
    est = enriquecer_programas(est)
    matriz, destinos = saidas_por_destino(df, cad, cli_df)
    margem = margem_vendas(df, cli_df)
    resumo_dashboard(df, cad, est, margem)
    for r in est:
        mm = matriz.get((r["emp"], r["cod"]), {"ini": 0, "comp": 0, "dest": {}})
        r["m_ini"], r["m_comp"], r["m_dest"] = mm["ini"], mm["comp"], mm["dest"]
    desc = {r["cod"]: r["desc"] for r in est}
    progs = {r["cod"]: r["progs"] for r in est}
    corr = [dict(cod=int(k), desc=desc.get(k, ""), progs=progs.get(k, []), **v) for k, v in conc.items() if v["acoes"]]
    dados = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "periodo": [df["data"].min().strftime("%d/%m/%Y"), df["data"].max().strftime("%d/%m/%Y")],
        "fonte": ARQ_MOVIMENTACAO.name,
        "linhas_mov": int(len(df)),
        "servicos_excluidos": len(df.attrs.get("servicos", [])),
        "linhas_servico": df.attrs.get("linhas_servico", 0),
        "estoque": est,
        "correcoes": corr,
        "destinos": destinos,
        "dono_prog": {p: dono_programa(p) for r in est for p in r.get("abas", [])},
        "clientes": clientes,
        "fonte_clientes": NOME_QUADRO if ARQ_QUADRO.exists() else "",
        "margem": margem,
    }
    html = ARQ_MODELO.read_text(encoding="utf-8")
    import re as _re
    achada = (_re.search(r'name="modelo-versao" content="([^"]+)"', html) or [None, "sem versão"])[1]
    if achada != VERSAO_MODELO:
        sys.exit(f"O arquivo analise_modelo.html está DESATUALIZADO (versão {achada}; o script precisa da {VERSAO_MODELO}).\n"
                 "Substitua o gerar_analise.py E o analise_modelo.html pelos DOIS arquivos mais recentes enviados pelo Claude.")
    html = html.replace("/*__DADOS__*/null", json.dumps(dados, ensure_ascii=False))
    ARQ_SAIDA.write_text(html, encoding="utf-8")
    novos = sum(c["novo"] for c in clientes)
    print(f"Análise gerada: {ARQ_SAIDA}")
    if "--sem-abrir" not in sys.argv:
        import webbrowser
        webbrowser.open(ARQ_SAIDA.as_uri())
    print(f"  Serviços desconsiderados: {dados['servicos_excluidos']} códigos ({dados['linhas_servico']} movimentos)")
    print(f"  {len(dados['estoque'])} itens x empresa | {len(clientes)} clientes ({novos} novos desde a última atualização)")


# ============ RESUMO PARA O DASHBOARD LOGÍSTICO (Visão Geral do Estoque) ============
ARQ_RESUMO_DASH = PASTA / "estoque_resumo.json"
GRUPOS_FORA_VISAO = ("Serviço",)   # serviços nunca entram (não são estoque)
# Itens NÃO comerciais (mídia, brindes, arquivos, equipamentos, uniformes, embalagem, uso/consumo):
# a Visão Geral do Estoque mostra só itens de venda por padrão e tem um botão para incluí-los.
GRUPOS_NAO_COMERCIAIS = ("Brinde", "Fora do escopo")
PROGRAMAS_NAO_COMERCIAIS = {"Brindes", "MIDIA", "ARQUIVOS", "UNIFORMES", "EMBALAGEM", "MAQUINAS E EQUIPAMENTOS"}
_NOME_CAT = {"Brindes": "Brindes", "MIDIA": "Mídia", "ARQUIVOS": "Arquivos", "UNIFORMES": "Uniformes",
             "EMBALAGEM": "Embalagem", "MAQUINAS E EQUIPAMENTOS": "Máquinas e equipamentos"}


def _categoria_nc(grupo, progs):
    """Nome da categoria se o item NÃO é de comercialização; '' se é item de venda."""
    g = str(grupo or "")
    ps = set(progs or [])
    if ps and ps <= PROGRAMAS_NAO_COMERCIAIS:
        return _NOME_CAT.get(sorted(ps)[0], sorted(ps)[0])
    if g.startswith("Brinde"):
        return g
    if g.startswith("Fora do escopo"):
        return "Uso/consumo e insumos"
    return ""


def _ano_ref(dt):
    """Ano Ref = ciclo out→set (out/2025 a set/2026 = 2026)."""
    return dt.dt.year + (dt.dt.month >= 10).astype(int)


def resumo_dashboard(df, cad, est, margem):
    """Indicadores anuais (por Ano Ref) e foto atual do estoque, por empresa (IAB, Verbum, Coala).
    Saldo SEMPRE calculado pela movimentação (implantação + entradas − saídas); a coluna
    SALDO_APOS_MOVIMENTO do ERP não é usada. Itens com saldo negativo contam como zero."""
    d = aplicar_coala(df, cad).copy()
    d = d[d["data"].notna()]
    grupo = dict(zip(cad["Codigo"].astype(int), cad["Grupo"].fillna("").astype(str)))
    d["grupo"] = d["codigo"].map(grupo).fillna("")
    # categoria não comercial por código (pelo programa da Análise e pelo grupo do cadastro)
    cat_nc = {}
    for r in est:
        cat_nc[r["cod"]] = _categoria_nc(r.get("grupo"), r.get("progs"))
    d["cat_nc"] = [cat_nc.get(c, _categoria_nc(g, None)) for c, g in zip(d["codigo"], d["grupo"])]
    d = d[~d["grupo"].str.startswith(GRUPOS_FORA_VISAO)]
    nat = d["natureza"].fillna("").astype(str).str.upper().str.strip()
    d = d[~nat.str.contains("TRANSFER[EÊ]NCIA ENTRE LOCAIS", regex=True)]
    nat = d["natureza"].fillna("").astype(str).str.upper().str.strip()
    d["emp"] = d["empresa_nome"].str.replace(r"\s*\(.*\)", "", regex=True)          # Verbum (Filial/Matriz) -> Verbum
    ent = d["tipo"].eq("ENTRADA"); sai = d["tipo"].str.startswith("SA")
    impl = nat.eq(NATUREZA_IMPLANTACAO)
    cls = np.select(
        [impl,
         ent & nat.str.contains("CONSIGNA"),
         ent & nat.str.startswith("COMPRA"),
         ent & nat.str.contains("DEVOLU"),
         ent & nat.str.contains("AJUSTE"),
         ent,
         sai & nat.isin(NAT_SAIDA_VENDA),
         sai & nat.str.startswith("REMESSA DE BONIFICA"),
         sai & (nat.str.contains("CONSIGNA") | nat.str.contains("TRANSFERENCIA ENTRE FILIAIS")),
         sai & nat.str.contains("AJUSTE"),
         sai & nat.str.contains("REQUISI"),
         sai],
        ["inicial", "consig_ent", "compras", "devolucoes", "ajuste_ent", "outras_ent",
         "venda", "bonif", "transf", "ajuste_sai", "interno", "outras_sai"], default="")
    d["cls"] = cls
    # custo: o da própria linha; se zerado, o último custo positivo do item na empresa
    cu = pd.to_numeric(d["custo_unit"], errors="coerce")
    ult_cu = d.assign(cu=cu.where(cu > 0)).sort_values("data").groupby(["emp", "codigo"])["cu"].last()
    d["cu"] = cu.where(cu > 0)
    d["cu"] = d["cu"].fillna(pd.Series([ult_cu.get((e, c), np.nan) for e, c in zip(d["emp"], d["codigo"])], index=d.index)).fillna(0)
    d["val"] = d["qtd"] * d["cu"]
    d["ano_ref"] = _ano_ref(d["data"]).astype(str)
    d["mov"] = np.where(impl | ent, d["qtd"], np.where(sai, -d["qtd"], 0))
    out = {"gerado": datetime.now().strftime("%d/%m/%Y %H:%M"),
           "periodo": [d["data"].min().strftime("%d/%m/%Y"), d["data"].max().strftime("%d/%m/%Y")],
           "empresas": {}}
    hoje = d["data"].max()
    margem_meses = (margem or {}).get("meses", [])
    out["empresas"] = _resumo_empresas(d, est, hoje, margem_meses, so_comercial=True)
    out["empresas_todos"] = _resumo_empresas(d, est, hoje, margem_meses, so_comercial=False)
    out["nao_comerciais"] = {}
    for emp in ["IAB", "Verbum"]:
        cats = {}
        for r in est:
            if r["emp"].split(" (")[0] != emp or r["atual"] <= 0 or str(r.get("grupo", "")).startswith(GRUPOS_FORA_VISAO):
                continue
            c = _categoria_nc(r.get("grupo"), r.get("progs"))
            if c:
                x = cats.setdefault(c, {"k": c, "itens": 0, "v": 0.0})
                x["itens"] += 1; x["v"] += r["atual"] * (r["custo"] or 0)
        out["nao_comerciais"][emp] = sorted([dict(x, v=round(x["v"], 2)) for x in cats.values()], key=lambda x: -x["v"])
    ARQ_RESUMO_DASH.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    print(f"  Resumo para o Dashboard Logístico: {ARQ_RESUMO_DASH.name}")
    return out


def _resumo_empresas(d, est, hoje, margem_meses, so_comercial):
    """Indicadores por empresa. so_comercial=True deixa fora mídia, brindes, arquivos, equipamentos etc.
    (a Coala só tem brindes, então ela sempre aparece completa)."""
    res = {}
    for emp in ["IAB", "Verbum", "Coala"]:
        e = d[d["emp"] == emp]
        if so_comercial and emp != "Coala":
            e = e[e["cat_nc"] == ""]
        if e.empty:
            continue
        # ---- por Ano Ref ----
        anos = {}
        for a, g in e.groupby("ano_ref"):
            q = g.groupby("cls")["qtd"].sum(); v = g.groupby("cls")["val"].sum()
            anos[a] = {k: {"q": float(q.get(k, 0)), "v": round(float(v.get(k, 0)), 2)}
                       for k in ["inicial", "compras", "devolucoes", "ajuste_ent", "consig_ent", "outras_ent",
                                 "venda", "bonif", "transf", "ajuste_sai", "interno", "outras_sai"]}
        # ---- saldo mês a mês (por item, negativo conta como zero) ----
        e = e.sort_values("data")
        e["mesk"] = e["data"].dt.strftime("%Y-%m")
        meses = sorted(e["mesk"].unique())
        mq = e.pivot_table(index="codigo", columns="mesk", values="mov", aggfunc="sum").reindex(columns=meses).fillna(0).cumsum(axis=1)
        cus = e[e["cu"] > 0].pivot_table(index="codigo", columns="mesk", values="cu", aggfunc="last").reindex(index=mq.index, columns=meses).ffill(axis=1).bfill(axis=1).fillna(0)
        saldo_mes = [{"mes": m, "q": float(mq[m].clip(lower=0).sum()), "v": round(float((mq[m].clip(lower=0) * cus[m]).sum()), 2)} for m in meses]
        for a in anos:
            fim = [s for s in saldo_mes if (int(s["mes"][:4]) + (int(s["mes"][5:]) >= 10)) == int(a)]
            if fim:
                anos[a]["saldo_fim"] = fim[-1]
                anos[a]["saldo_medio_v"] = round(sum(s["v"] for s in fim) / len(fim), 2)
                anos[a]["meses"] = len(fim)
        # ---- margem (Relatório de faturamento x custo) por Ano Ref ----
        for m in margem_meses:
            if not str(m["emp"]).startswith(emp) or not m.get("mes"):
                continue
            a = str(int(m["mes"][:4]) + (int(m["mes"][5:]) >= 10))
            if a in anos:
                mg = anos[a].setdefault("margem", {"rec": 0.0, "cus": 0.0, "bon": 0.0})
                mg["rec"] += m["rec"]; mg["cus"] += m["cus"]; mg["bon"] += m["bon"]
        # ---- foto atual por item ----
        itens = [r for r in est if (r["emp"].split(" (")[0] == emp) and not str(r.get("grupo", "")).startswith(GRUPOS_FORA_VISAO)]
        if so_comercial and emp != "Coala":
            itens = [r for r in itens if not _categoria_nc(r.get("grupo"), r.get("progs"))]
        sai180 = e[(e["data"] >= hoje - pd.Timedelta(days=180)) & e["cls"].isin(["venda", "bonif", "transf", "interno", "outras_sai"])].groupby("codigo")["qtd"].sum()
        sai365 = e[(e["data"] >= hoje - pd.Timedelta(days=365)) & e["cls"].isin(["venda", "bonif", "transf", "interno", "outras_sai"])].groupby("codigo")["qtd"].sum()
        com = [r for r in itens if r["atual"] > 0]
        val = lambda r: r["atual"] * (r["custo"] or 0)
        parados = [r for r in com if sai180.get(r["cod"], 0) <= 0]
        grupos = {}
        for r in com:
            k = _categoria_nc(r.get("grupo"), r.get("progs")) or r["grupo"] or "Não cadastrado"
            if emp == "Coala":
                k = r["grupo"] or "Não cadastrado"
            grupos.setdefault(k, {"k": k, "itens": 0, "q": 0.0, "v": 0.0, "v_parado": 0.0})
            grupos[k]["itens"] += 1; grupos[k]["q"] += r["atual"]; grupos[k]["v"] += val(r)
            if sai180.get(r["cod"], 0) <= 0: grupos[k]["v_parado"] += val(r)
        top = sorted(com, key=lambda r: -val(r))[:15]
        res[emp] = {
            "anos": anos,
            "saldo_mes": saldo_mes,
            "atual": {
                "skus": len(com), "q": float(sum(r["atual"] for r in com)), "v": round(sum(val(r) for r in com), 2),
                "parados_n": len(parados), "parados_v": round(sum(val(r) for r in parados), 2),
                "sai12_q": float(sum(sai365.get(r["cod"], 0) for r in com)),
                "corrigir_n": sum(1 for r in itens if r.get("alerta")),
            },
            "grupos": sorted([dict(g, v=round(g["v"], 2), v_parado=round(g["v_parado"], 2)) for g in grupos.values()], key=lambda g: -g["v"]),
            "top": [{"cod": r["cod"], "desc": r["desc"], "grupo": r["grupo"], "q": r["atual"], "v": round(val(r), 2),
                     "sai12": float(sai365.get(r["cod"], 0)), "ult": r["ult"]} for r in top],
        }
    return res


if __name__ == "__main__":
    # Ao dar dois cliques no arquivo, a janela fica aberta no final para você ler o resultado.
    try:
        main()
        print("\nOK - análise atualizada" + ("." if "--sem-abrir" in sys.argv else " e aberta no navegador."))
    except SystemExit as e:
        print(f"\nERRO: {e}")
    except Exception as e:
        print(f"\nERRO: {type(e).__name__}: {e}")
        print("Dica: feche o Quadro de Envios e a Movimentação de Estoque no Excel e rode de novo.")
    if sys.stdin and sys.stdin.isatty() and "--sem-pausa" not in sys.argv:
        input("\nPressione Enter para fechar...")
