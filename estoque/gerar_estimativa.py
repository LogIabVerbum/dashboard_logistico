# -*- coding: utf-8 -*-
"""
Estimativa de Vendas - Demanda de Vendas x Estoque / Produção Gráfica
----------------------------------------------------------------------
Lê as propostas que o Comercial salva no CRM compartilhado (1 JSON por proposta),
pega a proposta mais nova de cada cliente, soma os itens por código e compara com o
SALDO PRÓPRIO DO IAB calculado pela "Movimentação de Estoque.xlsx".
Gera "Estimativa de Vendas.html" nesta pasta, a partir do modelo "estimativa_modelo.html".

Uso:  python gerar_estimativa.py      (ou dois cliques no ATUALIZAR ESTIMATIVA.bat)
Requisitos: Python 3.9+ com pandas e openpyxl  ->  pip install pandas openpyxl

Regras (combinadas com o Anderson em 14/10/2026):
- Cliente = mesma chave do CRM: categoria + UF + município/instituição + ciclo.
- De cada cliente vale a proposta MAIS NOVA de cada tipo ("Venda" e "Bonificação");
  os "Venda - Aditivo" são complementares e SOMAM. "Versão alternativa" não entra.
- Não entram: propostas inativas, "sem contar" (leads importados), "Lead não convertido",
  "Cliente perdido" e "Material enviado" (o material já saiu e já está fora do saldo).
- "Empenho enviado" conta como FECHADA (100%).
- Estoque = só o saldo próprio do IAB. O que está consignado em parceiro (Verbum) NÃO soma.
- Versões (combinado em 16/10/2026): códigos do IAB com a MESMA descrição são versões do mesmo
  livro; código menor = versão mais antiga. As versões com saldo entram na estimativa junto do item
  da proposta (mesmo grupo), para usar o estoque antigo antes de imprimir a versão nova.
- Descarte previsto (16/10/2026): a lista de "24 - Descartes\\Descarte 2026.xlsx" sai do saldo até o
  descarte ser lançado no ERP (ver ARQ_DESCARTE / DESCARTE_DESDE).
"""
from pathlib import Path
from datetime import datetime, timezone
import json, os, re, sys, unicodedata, warnings
warnings.filterwarnings("ignore")
import pandas as pd

# ============ CONFIGURAÇÃO ============
PASTA = Path(__file__).resolve().parent
ONEDRIVE = Path(r"C:\Users\a2p2t\OneDrive - Instituto Alfa e Beto")


def _achar(*candidatas):
    """Primeira pasta que existir (permite rodar em outros computadores / perfis do OneDrive)."""
    for c in candidatas:
        if c and Path(c).exists():
            return Path(c)
    return Path(candidatas[0])


def _onedrive_base():
    """Procura a pasta do OneDrive do IAB no usuário atual (Anderson, Augusto, Miqueias...)."""
    home = Path.home()
    for p in [ONEDRIVE, home / "OneDrive - Instituto Alfa e Beto"] + sorted(home.glob("OneDrive*Alfa*Beto*")):
        if p.exists():
            return p
    return ONEDRIVE


BASE = _onedrive_base()
PASTA_PROPOSTAS = _achar(BASE / r"Vendas\Controles\2026_2027\CRM_Compartilhado\Dados", PASTA / "Dados")
PASTA_RELATORIOS = _achar(PASTA.parent.parent / "Arquivos de Claude Relatórios - Relatórios Logística",
                          BASE / r"Logística-IAB\10 - DASHIBOARDs\Arquivos de Claude Relatórios - Relatórios Logística", PASTA)
ARQ_MOVIMENTACAO = PASTA_RELATORIOS / "Movimentação de Estoque.xlsx"
ARQ_META_CRM = PASTA_PROPOSTAS / "__crm_meta.json"          # % de estimativa por temperatura (mesmos do CRM)
ARQ_MODELO = PASTA / "estimativa_modelo.html"
ARQ_SAIDA = PASTA / "Estimativa de Vendas.html"
VERSAO_MODELO = "2026-10-16a"   # tem que ser igual à do estimativa_modelo.html

# % padrão por temperatura (se o CRM não tiver configurado). "Fechada" = sempre 100%.
PCT_PADRAO = {"Quente": 90, "Morna": 60, "Fria": 30}

STATUS = {1: "Contato a realizar", 2: "Contato realizado", 3: "Proposta enviada", 4: "Documentação enviada",
          5: "Contrato enviado", 6: "Contrato assinado", 7: "Empenho enviado", 8: "Material enviado",
          9: "Lead não convertido", 10: "Cliente perdido"}
STATUS_FORA = {8: "Material já enviado (já saiu do estoque)", 9: "Lead não convertido", 10: "Cliente perdido"}
STATUS_FECHADA = {7}
# Propostas com "Material enviado" que vão ser ENVIADAS DE NOVO (combinado em 22/10/2026): continuam na demanda.
# Use o código da proposta no CRM. Quando o Comercial abrir a nova proposta, tire daqui.
REENVIAR = {"P5OC43": "Espírito Santo do Pinhal/SP - reenvio (provavelmente bonificado)",
            "PN00HE": "São Roque/SP - reenvio (provavelmente bonificado)"}
# Notas emitidas ANTES do corte que pertencem ao ciclo 2026/2027 (combinado em 04/11/2026): contam como "já saiu"
# da proposta e aparecem no Resumo de vendas "desde o corte". Use o número da NF.
NOTAS_DO_CICLO = {"19117": "Espírito Santo do Pinhal/SP - reenvio bonificado de 26/10/2026, bate com a proposta P5OC43"}

# ---- Estoque ----
# O saldo usa O MESMO código da Análise de Estoque (gerar_analise.py), para os dois dashboards baterem.
# Saldo = implantação + entradas − saídas (soma dos movimentos). A coluna SALDO_APOS_MOVIMENTO do ERP
# está errada e NÃO é usada em nenhum momento.
# Grupos do cadastro que NÃO são estoque de venda/revenda (comprados para uso da empresa)
GRUPOS_FORA_ESTOQUE = ["Fora do escopo (uso/consumo)"]
# As regras de saldo (gerar_analise.py) ficam na mesma pasta (dashboard_logistico\estoque).
PASTA_ANALISE = _achar(PASTA, PASTA.parent.parent / "Analise de Estoque")
# Regras abaixo só valem se o gerar_analise.py não for encontrado (cópia de segurança das mesmas regras)
EMPRESA_IAB = "INSTITUTO ALFA E BETO"
EMPRESAS_IGNORADAS = ["LINCKY CONSULTORIA"]               # erro de lançamento no ERP
NATUREZA_CONSIG_VERBUM = "COMPRA PARA REVENDA POR CONSIGNA"  # entrada na Verbum = saída do IAB
NATUREZA_IMPLANTACAO = "IMPLANTAÇÃO DE SALDOS DE ESTOQUES"
# Naturezas que NÃO são custo de produção/compra (não servem de referência de custo unitário)
NAT_SEM_CUSTO_REF = r"DEVOLU|TRANSFER|AJUSTE|CONSIGNA|RETORNO|IMPLANTA"

# ---- Descarte previsto (combinado em 16/10/2026) ----
# Enquanto o descarte não é lançado no ERP, as quantidades da lista saem do saldo da estimativa.
# Quando o ERP tiver a saída de descarte do item (natureza com "DESCARTE", a partir de DESCARTE_DESDE),
# só desconta o que ainda faltar lançar — não conta em dobro.
ARQ_DESCARTE = BASE / r"Logística-IAB\24 - Descartes\Descarte 2026.xlsx"

# ---- Saldo congelado (combinado em 16/10/2026) ----
# O saldo do IAB fica "congelado" nesta data. Depois dela aparecem, item a item: o que chegou de cada gráfica
# (notas de compra), o que saiu para os clientes das propostas (venda, bonificação, simples remessa) e as
# outras movimentações, fechando no SALDO ATUAL (estoque do dia), que é o usado no cálculo da falta.
CORTE = os.environ.get("ESTIMATIVA_CORTE", "2026-11-01")
NATUREZA_CONSIG_SAIDA = "REMESSA EM CONSIGNAÇÃO PARA A VERBUM (saída criada pelo script)"
ARQ_FATURAMENTO = PASTA_RELATORIOS / "Relatorio de faturamento.xlsx"
ABA_DESCARTE = "Dados para Contabilidade"
DESCARTE_DESDE = "2026-10-16"
# O descarte de 2026 saiu como VENDA DE APARAS para a Eco Primos (NF 19112 de 20/10/2026): saídas para estes
# parceiros também contam como descarte já lançado.
DESCARTE_PARCEIROS = ["ECO PRIMOS"]
RE_DESCARTE_PARCEIRO = "|".join(re.escape(p) for p in DESCARTE_PARCEIROS)
# ======================================


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "").strip().lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


# Versões com NOME DIFERENTE (combinado em 22/10/2026): código antigo -> código da versão nova.
# Quando a descrição é a mesma, a ligação já é automática; aqui só os casos em que o nome mudou.
EQUIVALENCIAS = {
    1709: 1962,   # Agenda do Professor - Alfabetização     -> Caderno de Registro do Professor - Alfabetização
    1706: 1963,   # Agenda do Professor - Ensino Estruturado -> Caderno de Registro do Professor - Ensino Estruturado
    611: 1964,    # Agenda do Professor - Pré-Escola         -> Caderno de Registro do Professor - Pré-Escola
    430: 2092,    # Conte Outra Vez - Livro Gigante          -> Conte Outra Vez - Livro Gigante - 2ª ed. 2025
    # Meu Livro de Atividades Pré I (23/10/2026): "PRÉ-I (VOLUME I)" e "PRÉ I VOL. I" são o mesmo livro
    1799: 2067,   # Meu Livro de Atividades – Pré-I (Volume I)  -> Meu Livro de Atividades Pré I Vol. I
    1798: 2063,   # Meu Livro de Atividades – Pré-I (Volume II) -> Meu Livro de Atividades Pré I Vol. II
    1726: 2061,   # ... (Volume I) - Livro do Professor          -> Livro do Professor do Meu Livro de Atividades Pré I Vol. I
    1727: 2062,   # ... (Volume II) - Livro do Professor         -> Livro do Professor do Meu Livro de Atividades Pré I Vol. II
}
# Saldo ZERADO à força neste relatório (23/10/2026): diferença com o kardex do ERP ainda em análise.
#  611: Agenda do Professor Pré-Escola - simples remessa (6) + ajuste com terceiros (49) de Pilar; ERP dá 0
# 1795: IAB Matemática 1º ano Vol. 1 - 186 no IAB − 177 consignados à Verbum = 9; kardex do ERP dá 0
# Quando a diferença for resolvida, tire o código daqui.
SALDO_ZERADO = {611, 1795}
# Itens que NÃO entram na demanda (nem proposta, nem projeção, nem Verbum)
ITENS_FORA_DEMANDA = {
    10195: "Teste Diagnóstico de Alfabetização - TELCS (não considerar)",
}


def chave_versao(cod, est):
    """Chave do grupo de versões de um código: descrição normalizada, seguindo a tabela de equivalências.
    A equivalência vale para TODOS os códigos com a mesma descrição do código antigo (ex.: 1723 tem o mesmo nome do 1799,
    então também vai para o grupo do 2067)."""
    cache = est.setdefault("_equiv_desc", {})
    if not cache:
        for a, b in EQUIVALENCIAS.items():
            cache[chave_desc(est["desc"].get(a, ""))] = chave_desc(est["desc"].get(b, ""))
        cache[None] = None
    k = chave_desc(est["desc"].get(EQUIVALENCIAS.get(int(cod), int(cod)), ""))
    return cache.get(k, k)


def chave_desc(s):
    """Descrição normalizada para achar versões do mesmo item (ignora acento, caixa, tipo de traço e espaços)."""
    s = norm(s)
    s = re.sub(r"[‐-―\-]+", "-", s)
    s = re.sub(r"\s*([-:.,/()])\s*", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


def chave_cliente(p):
    """Igual ao IABclientKey do CRM de Vendas."""
    cat = norm(p.get("categoria") or "Município")
    uf = str(p.get("uf") or "").strip().upper()
    ciclo = str(p.get("ano") or "").strip()
    if cat == "estado":
        return f"EST|{uf}|{ciclo}"
    if cat == "municipio":
        return f"MUN|{uf}|{norm(p.get('cidade') or p.get('nomeCliente'))}|{ciclo}"
    return f"INST|{cat}|{uf}|{norm(p.get('nomeCliente') or p.get('cidade'))}|{ciclo}"


def nome_cliente(p):
    cat = norm(p.get("categoria") or "Município")
    if cat == "estado":
        return str(p.get("nomeCliente") or f"Estado {p.get('uf', '')}").strip()
    if cat == "municipio":
        return str(p.get("cidade") or p.get("nomeCliente") or "").strip()
    return str(p.get("nomeCliente") or p.get("cidade") or "").strip()


def _data_iso(txt):
    try:
        return datetime.fromisoformat(str(txt).replace("Z", "+00:00"))
    except Exception:
        return None


# ---------------------------------------------------------------- propostas
def carregar_propostas():
    if not PASTA_PROPOSTAS.exists():
        sys.exit(f"ERRO: pasta das propostas não encontrada:\n  {PASTA_PROPOSTAS}")
    regs = []
    for f in sorted(PASTA_PROPOSTAS.glob("*.json")):
        n = f.name
        if n.startswith("__") or n.startswith("conflito__") or n.startswith("_"):
            continue
        try:
            p = json.loads(f.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"  Aviso: não consegui ler {n}: {e}")
            continue
        if not isinstance(p, dict) or not p.get("id"):
            continue
        pg = p.get("propostaGerada") or {}
        mtime = datetime.fromtimestamp(f.stat().st_mtime, tz=timezone.utc)
        criado = _data_iso(pg.get("criadoEm"))
        data = max([d for d in (criado, mtime) if d])   # a mais recente entre "proposta gerada" e "arquivo salvo"
        regs.append({"p": p, "arquivo": n, "data": data, "criado": criado,
                     "itens": [i for i in (pg.get("itens") or []) if i.get("cod") not in (None, "")]})
    return regs


def escolher_propostas(regs):
    """Marca, para cada proposta, se entra na demanda e o porquê."""
    grupos = {}
    for r in regs:
        p = r["p"]
        r["chave"] = chave_cliente(p)
        r["tipo"] = p.get("tipoProposta") or "Venda"
        st = int(p.get("status") or 0)
        motivo = ""
        if p.get("inativo"):
            motivo = "Proposta inativa"
        elif p.get("semContar"):
            motivo = "Marcada para não contar (lead importado)"
        elif r["tipo"] == "Versão alternativa":
            motivo = "Versão alternativa"
        elif st in STATUS_FORA and not (st == 8 and p.get("id") in REENVIAR):
            motivo = STATUS_FORA[st]
        r["motivo"] = motivo
        if not motivo and r["tipo"] != "Venda - Aditivo":
            grupos.setdefault((r["chave"], r["tipo"]), []).append(r)
    for (_, tipo), lst in grupos.items():
        lst.sort(key=lambda r: r["data"], reverse=True)
        for r in lst[1:]:
            r["motivo"] = f"Substituída pela proposta mais nova ({lst[0]['p']['id']} de {lst[0]['data']:%d/%m/%Y})"
    for r in regs:
        r["usada"] = not r["motivo"]
        if r["usada"] and not r["itens"]:
            r["motivo"] = "Sem itens detalhados (proposta não gerada no Gerador)"
    return regs


def temperatura(p):
    st = int(p.get("status") or 0)
    t = str(p.get("temperatura") or "").strip()
    if st in STATUS_FECHADA or t.lower().startswith("venda fechada"):
        return "Fechada"
    return t if t in ("Quente", "Morna", "Fria") else "Fria"


# ---------------------------------------------------------------- estoque
def carregar_estoque():
    """Saldo próprio do IAB calculado pelas regras do gerar_analise.py (Análise de Estoque)."""
    try:
        sys.path.insert(0, str(PASTA_ANALISE))
        import gerar_analise as ga
    except Exception as e:
        print(f"  Aviso: não achei o gerar_analise.py em {PASTA_ANALISE} ({e}). Usando a cópia das regras.")
        return carregar_estoque_copia()
    print(f"  Movimentação: {ga.ARQ_MOVIMENTACAO}")
    df = ga.carregar_movimentacao()                 # Lincky, serviços, e-books, ignorados e consignação Verbum
    cad = ga.carregar_cadastro()
    # material comprado para uso da empresa (micro-ondas, móveis...) não é item de venda/revenda: fica fora do estoque
    uso = set(cad.loc[cad["Grupo"].isin(GRUPOS_FORA_ESTOQUE), "Codigo"].astype(int))
    fora = df["codigo"].isin(uso)
    if fora.any():
        print(f"  Uso/consumo da empresa fora do estoque: {df.loc[fora, 'codigo'].nunique()} códigos ({int(fora.sum())} movimentos)")
        df = df[~fora].copy()
    analise = analise_estoque(ga, df, cad)          # mesma avaliação da aba Estoque Geral (alertas, conciliação)
    df = ga.aplicar_coala(df, cad)                  # brindes Coala saem do IAB
    iab = df[df["empresa_nome"].eq("IAB")].copy()     # já vem sem a baixa dupla de terceiros (ga.anular_duplicidade_terceiros)
    impl = iab["natureza"].str.upper().eq(ga.NATUREZA_IMPLANTACAO)
    ent = iab["tipo"].eq("ENTRADA") & ~impl
    sai = iab["tipo"].str.startswith("SA")
    iab["MOV"] = iab["qtd"].where(impl, 0) + iab["qtd"].where(ent, 0) - iab["qtd"].where(sai, 0)
    saldo = iab.groupby("codigo")["MOV"].sum()
    cons = iab[iab["natureza"].eq(ga.NATUREZA_CONSIG_IAB)]
    consig = cons.groupby("codigo")["qtd"].sum()
    consig_info = {"nfs": int(cons["nf"].nunique()), "ini": cons["data"].min(), "fim": cons["data"].max()}
    desc = df.sort_values("data").groupby("codigo")["descricao"].last().astype(str).str.strip()
    ult = iab.groupby("codigo")["data"].max()
    cu = pd.to_numeric(iab["custo_unit"], errors="coerce")
    ref = iab[ent & (cu > 0) & ~iab["natureza"].str.upper().str.contains(NAT_SEM_CUSTO_REF)].sort_values("data")
    print(f"  Saldo calculado com as regras da Análise de Estoque ({PASTA_ANALISE.name}\\gerar_analise.py)")
    descartes_erp = iab[sai & (iab["natureza"].str.upper().str.contains("DESCARTE")
                               | iab["parceiro"].fillna("").astype(str).str.upper().str.contains(RE_DESCARTE_PARCEIRO))][["codigo", "data", "qtd"]]
    mov = iab[["codigo", "data", "tipo", "natureza", "parceiro", "cidade", "uf", "qtd", "MOV"]].copy()
    mov["nf"] = iab["nf"].map(_nf_txt)
    ult_compra = iab[ent & iab["natureza"].str.upper().str.startswith("COMPRA")].groupby("codigo")["data"].max().to_dict()
    return {"saldo": saldo, "consig": consig, "consig_info": consig_info, "desc": desc, "ult": ult,
            "descartes_erp": descartes_erp, "mov": mov, "ult_compra": ult_compra,
            "grupo": dict(zip(cad["Codigo"].astype(int), cad["Grupo"].fillna("").astype(str))),
            "programas": dict(zip(cad["Codigo"].astype(int), cad["Programas"].fillna("").astype(str))),
            "custo": ref.groupby("codigo")["custo_unit"].last(), "custo_dt": ref.groupby("codigo")["data"].last(),
            "data_base": df["data"].max(), "analise": analise}


def analise_estoque(ga, df, cad):
    """Por código do IAB: saldo inicial, entradas, saídas, saídas dos últimos 90 dias, último movimento,
    alerta da Análise de Estoque e situação da consignação na Verbum (tem que zerar)."""
    try:
        linhas, conc = ga.avaliacao_estoque(df, cad)
    except Exception as e:
        print(f"  Aviso: não consegui montar a análise de estoque por item ({e})")
        return {}
    d = ga.aplicar_coala(df, cad)
    d = d[d["empresa_nome"].eq("IAB") & d["tipo"].str.startswith("SA")]
    corte = d["data"].max() - pd.Timedelta(days=90)
    s90 = d[d["data"] >= corte].groupby("codigo")["qtd"].sum()
    out = {}
    for l in linhas:
        if l["emp"] != "IAB":
            continue
        c = conc.get(l["cod"]) or {}
        out[l["cod"]] = {"i": l["ini"], "e": l["ent"], "sa": l["sai"], "at": l["atual"], "ul": l["ult"],
                         "al": l["alerta"], "s90": float(s90.get(l["cod"], 0)),
                         "vb": c.get("saldo_vb"), "vba": " | ".join(c.get("acoes") or [])}
    return out


def carregar_estoque_copia():
    print(f"  Movimentação: {ARQ_MOVIMENTACAO}")
    if not ARQ_MOVIMENTACAO.exists():
        sys.exit(f"ERRO: arquivo de movimentação não encontrado:\n  {ARQ_MOVIMENTACAO}")
    df = pd.read_excel(ARQ_MOVIMENTACAO)
    df.columns = [str(c).strip().upper() for c in df.columns]
    df["EMP"] = df["EMPRESA_ESTOQUE"].fillna("").astype(str).str.strip().str.upper()
    df = df[~df["EMP"].apply(lambda e: any(x in e for x in EMPRESAS_IGNORADAS))].copy()
    df["COD"] = pd.to_numeric(df["CODIGO_PRODUTO"], errors="coerce")
    df = df.dropna(subset=["COD"]).copy()
    df["COD"] = df["COD"].astype(int)
    df["QTD"] = pd.to_numeric(df["QUANTIDADE"], errors="coerce").fillna(0)
    df["TIPO"] = df["TIPO_MOVIMENTO"].fillna("").astype(str).str.strip().str.upper()
    df["NAT"] = df["NATUREZA_TIPO_USO"].fillna("").astype(str).str.strip().str.upper()
    df["DATA"] = pd.to_datetime(df["DATA_MOVIMENTO"], errors="coerce")
    df["NF"] = df["NUMERO_NF"].astype(str)
    iab = df[df["EMP"].eq(EMPRESA_IAB)].copy()
    # Consignação IAB -> Verbum: o ERP só traz a entrada na Verbum; a saída do IAB é criada aqui
    vb = df[df["EMP"].str.startswith("VERBUM") & df["NAT"].str.contains(NATUREZA_CONSIG_VERBUM)]
    ja = set(zip(iab.loc[iab["TIPO"].str.startswith("SA"), "NF"], iab.loc[iab["TIPO"].str.startswith("SA"), "COD"],
                 iab.loc[iab["TIPO"].str.startswith("SA"), "QTD"]))
    vb = vb[[(n, c, q) not in ja for n, c, q in zip(vb["NF"], vb["COD"], vb["QTD"])]].copy()
    vb["TIPO"] = "SAÍDA"
    vb["NAT"] = "REMESSA EM CONSIGNAÇÃO PARA A VERBUM (saída criada pelo script)"
    iab = pd.concat([iab, vb], ignore_index=True)
    sinal = iab["TIPO"].map(lambda t: 1 if t == "ENTRADA" else (-1 if t.startswith("SA") else 0))
    iab["MOV"] = iab["QTD"] * sinal
    saldo = iab.groupby("COD")["MOV"].sum()
    consig = vb.groupby("COD")["QTD"].sum()
    desc = df.sort_values("DATA").groupby("COD")["DESCRICAO_PRODUTO"].last().astype(str).str.strip()
    ult = iab.groupby("COD")["DATA"].max()
    # custo de referência: última entrada com custo > 0 que seja compra/produção
    ent = iab[iab["TIPO"].eq("ENTRADA") & (pd.to_numeric(iab["CUSTO_UNITARIO"], errors="coerce") > 0)
              & ~iab["NAT"].str.contains(NAT_SEM_CUSTO_REF)].sort_values("DATA")
    custo = ent.groupby("COD")["CUSTO_UNITARIO"].last()
    custo_dt = ent.groupby("COD")["DATA"].last()
    print(f"  Movimentos do IAB: {len(iab):,} | itens com saldo: {(saldo != 0).sum():,} | "
          f"consignado na Verbum (fora do saldo): {int(vb['QTD'].sum()):,} un.".replace(",", "."))
    dsc = iab[iab["TIPO"].str.startswith("SA") & (iab["NAT"].str.contains("DESCARTE")
              | iab["NOME_PARCEIRO"].fillna("").astype(str).str.upper().str.contains(RE_DESCARTE_PARCEIRO))]
    return {"saldo": saldo, "consig": consig, "desc": desc, "ult": ult, "custo": custo, "custo_dt": custo_dt,
            "descartes_erp": dsc.rename(columns={"COD": "codigo", "DATA": "data", "QTD": "qtd"})[["codigo", "data", "qtd"]],
            "consig_info": {"nfs": int(vb["NF"].nunique()), "ini": vb["DATA"].min(), "fim": vb["DATA"].max()},
            "data_base": df["DATA"].max()}


def zerar_negativos(est):
    """Saldo negativo é erro de baixa nosso, corrigido pontualmente no ERP (22/10/2026): não aparece no relatório.
    O item fica com saldo 0 (a demanda dele continua contando como falta)."""
    neg = est["saldo"][est["saldo"] < 0]
    if len(neg):
        print(f"  Saldo negativo (erro de baixa, mostrado como 0): "
              + ", ".join(f"{int(c)} ({int(v)})" for c, v in neg.items()))
    est["saldo"] = est["saldo"].clip(lower=0)
    zerar = [c for c in SALDO_ZERADO if est["saldo"].get(c, 0) != 0]
    if zerar:
        print("  Saldo zerado à força (diferença com o kardex em análise): "
              + ", ".join(f"{c} ({int(est['saldo'][c])} → 0)" for c in zerar))
        for c in zerar:
            est["saldo"].loc[c] = 0
            if c in (est.get("analise") or {}):
                est["analise"][c]["at"] = 0
    for a in (est.get("analise") or {}).values():
        if (a.get("at") or 0) < 0:
            a["at"] = 0


def aplicar_descarte(est):
    """Desconta do saldo a lista de descarte que ainda não foi lançada no ERP. Devolve o resumo."""
    vazio = {"itens": 0, "un": 0, "erp": 0, "arq": ARQ_DESCARTE.name, "por": {}}
    if not ARQ_DESCARTE.exists():
        print(f"  Descarte: lista não encontrada ({ARQ_DESCARTE}) - nada descontado")
        return vazio
    try:
        import shutil, tempfile
        tmp = Path(tempfile.gettempdir()) / "descarte_estimativa.xlsx"
        try:
            shutil.copyfile(ARQ_DESCARTE, tmp)
        except PermissionError:                            # planilha aberta no Excel: o Windows ainda deixa copiar
            import subprocess
            subprocess.run(["powershell", "-NoProfile", "-Command",
                            f"Copy-Item -LiteralPath '{ARQ_DESCARTE}' -Destination '{tmp}' -Force"],
                           check=True, capture_output=True)
        bruto = pd.read_excel(tmp, sheet_name=ABA_DESCARTE, header=None)
        lin = next(i for i, v in bruto[0].items() if str(v).strip().lower().startswith("cód"))
        d = pd.read_excel(tmp, sheet_name=ABA_DESCARTE, header=lin)
    except Exception as e:
        print(f"  Aviso: não consegui ler a lista de descarte ({e}) - nada descontado")
        return vazio
    d["cod"] = pd.to_numeric(d.iloc[:, 0], errors="coerce")
    d["q"] = pd.to_numeric(d["Quantidade"], errors="coerce").fillna(0)
    lista = d.dropna(subset=["cod"]).groupby(d["cod"].dropna().astype(int))["q"].sum()
    erp = est.get("descartes_erp")
    feito = pd.Series(dtype=float)
    if erp is not None and len(erp):
        feito = erp[erp["data"] >= pd.Timestamp(DESCARTE_DESDE)].groupby("codigo")["qtd"].sum()
    por, tot, tot_erp = {}, 0.0, 0.0
    for cod, q in lista.items():
        if q <= 0:
            continue
        ja = min(q, float(feito.get(cod, 0)))
        falta = q - ja
        tot_erp += ja
        saldo_atual = max(0.0, float(est["saldo"].get(cod, 0)))
        if falta > saldo_atual:   # lista maior que o estoque (ex.: 430 com 164 na lista e 163 no saldo): desconta só o saldo
            print(f"  Aviso descarte: cód. {cod} tem {falta:,.0f} na lista e {saldo_atual:,.0f} no saldo - descontado {saldo_atual:,.0f}".replace(",", "."))
            falta = saldo_atual
        if falta > 0:
            est["saldo"].loc[cod] = float(est["saldo"].get(cod, 0)) - falta
            por[int(cod)] = falta
            tot += falta
    print(f"  Descarte previsto: {int(tot):,} un. de {len(por)} itens descontadas do saldo".replace(",", ".")
          + (f" ({int(tot_erp):,} un. já lançadas no ERP)".replace(",", ".") if tot_erp else ""))
    return {"itens": len(por), "un": tot, "erp": tot_erp, "arq": ARQ_DESCARTE.name, "por": por}


def pcts_crm():
    pc = dict(PCT_PADRAO)
    try:
        m = json.loads(ARQ_META_CRM.read_text(encoding="utf-8"))
        for k, v in ((m.get("settings") or {}).get("estimPct") or {}).items():
            if k in pc and v not in (None, ""):
                pc[k] = float(v)
    except Exception:
        pass
    return pc


# ---------------------------------------------------------------- montar dados
def montar(regs, est):
    props, linhas, itens_idx, itens = [], [], {}, []

    def item_id(cod, nome, programa):
        if cod not in itens_idx:
            itens_idx[cod] = len(itens)
            c = est["custo"].get(cod)
            itens.append({"c": cod, "d": est["desc"].get(cod, nome) or nome, "n": nome, "pr": [],
                          "s": float(est["saldo"].get(cod, 0)), "cs": float(est["consig"].get(cod, 0)),
                          "cu": None if c is None or pd.isna(c) else round(float(c), 2),
                          "cud": "" if cod not in est["custo_dt"] else f"{est['custo_dt'][cod]:%d/%m/%Y}",
                          "u": "" if pd.isna(est["ult"].get(cod, pd.NaT)) else f"{est['ult'][cod]:%d/%m/%Y}",
                          "erp": cod in est["saldo"].index,
                          "ds": float(est.get("dsc_por", {}).get(cod, 0)),
                          "uc": "" if cod not in est.get("ult_compra", {}) else f"{est['ult_compra'][cod]:%d/%m/%Y}",
                          "gr": est.get("grupo", {}).get(cod, ""),
                          "an": (est.get("analise") or {}).get(cod)})
        it = itens[itens_idx[cod]]
        if programa and programa not in it["pr"]:
            it["pr"].append(programa)
        return itens_idx[cod]

    regs = sorted(regs, key=lambda r: (nome_cliente(r["p"]).lower(), -r["data"].timestamp()))
    por_desc = {}
    for c in est["saldo"].index:
        por_desc.setdefault(chave_versao(c, est), []).append(int(c))
    trocas = {}
    for r in regs:
        p = r["p"]
        pi = len(props)
        props.append({
            "id": p["id"], "cli": nome_cliente(p), "uf": p.get("uf", ""), "cat": p.get("categoria") or "Município",
            "t": temperatura(p), "tc": p.get("temperatura") or "", "st": int(p.get("status") or 0),
            "stx": STATUS.get(int(p.get("status") or 0), ""), "tipo": r["tipo"],
            "v": float(p.get("proposta2627") or 0), "al": int(p.get("alunos") or 0),
            "vend": p.get("vendedor") or "", "coord": p.get("coordIab") or "",
            "dt": f"{r['data'].astimezone():%d/%m/%Y %H:%M}", "dts": r["data"].timestamp(),
            "ni": len(r["itens"]), "u": r["usada"] and bool(r["itens"]), "m": r["motivo"], "arq": r["arquivo"],
            "progs": p.get("programas") or [],
            "cid": str(p.get("cidade") or "").strip(), "nome": str(p.get("nomeCliente") or "").strip(),
            "cond": " · ".join(str(p.get(k)).strip() for k in ("parcelamento", "fontePagadora", "modalidade")
                               if str(p.get(k) or "").strip()),
        })
        if not props[-1]["u"]:
            continue
        acum = {}
        for i in r["itens"]:
            try:
                cod = int(i["cod"])
            except Exception:
                continue
            kit = str(i.get("kit") or "")
            nova = versao_sem_saldo(cod, est, por_desc)
            if nova != cod:   # versão antiga sem saldo: a quantidade vai para a versão nova (22/10/2026)
                trocas[(cod, nova)] = trocas.get((cod, nova), 0) + float(i.get("qtd") or 0)
                kit, cod = f"{kit} · proposta com cód. {cod} (versão antiga sem saldo)".strip(" ·"), nova
            ii = item_id(cod, str(i.get("nome") or "").strip(), str(i.get("programa") or "").strip())
            k = (ii, kit)
            acum[k] = acum.get(k, 0) + float(i.get("qtd") or 0)
        for (ii, kit), q in acum.items():
            if q:
                linhas.append([pi, ii, q, kit])
    if trocas:
        print(f"  Versões antigas sem saldo nas propostas (quantidade lançada na versão nova): "
              + ", ".join(f"{a}→{b} ({int(q):,} un.)".replace(",", ".") for (a, b), q in sorted(trocas.items())))
    projetar_sem_itens(props, linhas, item_id, est)
    adicionar_consignacao_verbum(props, linhas, item_id, est)
    pos = movimento_pos_corte(props, itens, item_id, est)
    fora = [l for l in linhas if itens[l[1]]["c"] in ITENS_FORA_DEMANDA]
    if fora:
        linhas[:] = [l for l in linhas if itens[l[1]]["c"] not in ITENS_FORA_DEMANDA]
        print(f"  Fora da demanda: {', '.join(str(c) for c in ITENS_FORA_DEMANDA)} ({len(fora)} linhas, "
              f"{int(sum(l[2] for l in fora)):,} un.)".replace(",", "."))
    agrupar_versoes(itens, item_id, est)
    completar_saldos(itens, pos)
    ordenar_por_programa(itens, est)
    itens, linhas = remover_fora_relatorio(props, itens, linhas, pos, est)
    return props, itens, linhas, pos


def remover_fora_relatorio(props, itens, linhas, pos, est):
    """(22/10/2026) Itens SEM PROGRAMA e BRINDES não entram no relatório (nem demanda, nem saldo, nem falta)."""
    def fora(it):
        return it["c"] in ITENS_FORA_DEMANDA or it["pn"] in ("Sem programa", "Brindes") or str(est.get("grupo", {}).get(it["c"], "")).startswith("Brinde")
    novo, novos = {}, []
    sem, bri = [], []
    for i, it in enumerate(itens):
        if fora(it):
            (bri if it["pn"] == "Brindes" or str(est.get("grupo", {}).get(it["c"], "")).startswith("Brinde") else sem).append(it["c"])
            continue
        novo[i] = len(novos)
        novos.append(it)
    linhas = [[pi, novo[ii], q, kit] for pi, ii, q, kit in linhas if ii in novo]
    pos["saidas"] = [[pi, novo[ii], q] for pi, ii, q in pos["saidas"] if ii in novo]
    for pi, p in enumerate(props):
        if p["u"]:
            p["ni"] = sum(1 for l in linhas if l[0] == pi)
    print(f"  Fora do relatório: {len(sem)} itens sem programa ({', '.join(map(str, sem))}) e {len(bri)} brindes")
    return novos, linhas


# ---------------------------------------------------------------- depois do corte (saldo congelado)
NAT_SAIDA_CLIENTE = {"NOTA DE VENDA DE PRODUTO PRODUZIDO", "REMESSA VENDA ENTREGA FUTURA", "REMESSA DE BONIFICAÇÃO - SAIDA",
                     "SIMPLES REMESSA (OUTRAS)", "VENDA CONSIGNADA"}
RE_DEVOLUCAO_CLIENTE = r"DEVOLUÇÃO DE VENDA|DEVOLUÇÃO DE REMESSA|RETORNO SIMPLES REMESSA"
RE_ORGAO_MUNICIPAL = (r"municipio|prefeit|gabinete|\bgab\b|secretaria|\bsec\b|fundo municipal|camara municipal"
                      r"|\bpm\b|\bsme\b|\bfme\b|semed|seduc|semec|smec")
# entidades que têm "municipal" no nome mas NÃO são a prefeitura (caixa escolar, APM, escola...) e órgãos estaduais
RE_NAO_ORGAO = r"^(caixa escolar|circulo|associacao|apm|conselho|escola|colegio|creche|instituto)|estad|governo do"


def chave_cidade(cidade):
    """Nome de cidade comparável: sem acento, sem "de/do/da/dos/das", sem espaços (Santa Cruz do/de Monte Castelo)."""
    n = re.sub(r"\b(de|do|da|dos|das|d)\b", " ", norm(cidade).replace("'", " "))
    return re.sub(r"[^a-z]", "", n)


def mesma_cidade(chave, chaves):
    """Chave igual, ou quase igual (erro de digitação: Santa Carmen/Carmem) entre as chaves da mesma UF."""
    if chave in chaves:
        return chave
    import difflib
    m = difflib.get_close_matches(chave, list(chaves), n=1, cutoff=0.9)
    return m[0] if m else None


def eh_orgao_municipal(nome, cidade):
    """Prefeitura, secretaria, gabinete, fundo... ou o próprio nome da cidade (ex.: "VISCONDE DO RIO BRANCO")."""
    n = norm(nome)
    if re.search(RE_NAO_ORGAO, n):
        return False
    junto = re.sub(r"[^a-z]", "", n)                       # "S E D U C" -> "seduc"
    if re.search(RE_ORGAO_MUNICIPAL, n) or junto.startswith(("seduc", "semed", "semec", "smec")):
        return True
    c = norm(cidade)
    return bool(c) and re.sub(r"[^a-z]", "", n.split(" - ")[0]) == re.sub(r"[^a-z]", "", c)


class Casador:
    """Liga um destinatário (nome, cidade, UF) à proposta do cliente: município pela cidade+UF quando o
    destinatário é órgão municipal; escolas/instituições pelo nome; Verbum pelo nome."""
    def __init__(self, props, so_usadas=True):
        self.mun, self.inst, self.vb = {}, [], None
        prio = (lambda i: (not props[i]["u"], props[i]["tipo"] != "Venda")) if so_usadas else \
               (lambda i: (props[i]["tipo"] != "Venda", not props[i]["u"]))   # no resumo vale a proposta de Venda
        ordem = sorted(range(len(props)), key=lambda i: (*prio(i), -props[i].get("dts", 0)))
        for i in ordem:
            p = props[i]
            if so_usadas and not p["u"]:
                continue
            if p["tipo"] == "Consignação":
                self.vb = self.vb if self.vb is not None else i
            elif norm(p["cat"]) == "municipio":
                self.mun.setdefault(p["uf"], {}).setdefault(chave_cidade(p.get("cid") or p["cli"]), i)
            elif p["tipo"] in ("Venda", "Bonificação", "Venda - Aditivo") or not so_usadas:
                self.inst.append((norm(p.get("nome") or p["cli"]), p["uf"], i))

    def __call__(self, nome, cidade, uf):
        n, uf = norm(nome), str(uf or "").strip().upper()
        if "verbum" in n:
            return self.vb
        for nm, u, i in self.inst:
            if nm and u == uf and (nm in n or n in nm):
                return i
        if eh_orgao_municipal(nome, cidade):
            cidades = self.mun.get(uf, {})
            k = mesma_cidade(chave_cidade(cidade), cidades)
            return cidades[k] if k else None
        return None


def _nf_txt(v):
    """Número da NF como texto ("19117.0" -> "19117")."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    t = str(v).strip()
    return t[:-2] if t.endswith(".0") else t


def _do_ciclo(mov):
    """Linhas de notas anteriores ao corte que contam no ciclo novo (NOTAS_DO_CICLO)."""
    return mov["nf"].isin(NOTAS_DO_CICLO) if "nf" in mov else pd.Series(False, index=mov.index)


def movimento_pos_corte(props, itens, item_id, est):
    """Depois do corte: recebimentos por gráfica (notas de compra) e saídas para os clientes das propostas."""
    corte = pd.Timestamp(CORTE)
    vazio = {"corte": f"{corte:%d/%m/%Y}", "graficas": [], "saidas": [], "rec": {}, "sp": {}, "net": {}, "sem_proposta": 0}
    mov = est.get("mov")
    if mov is None:
        return vazio
    m = mov[(mov["data"] >= corte) | _do_ciclo(mov)]
    if _do_ciclo(mov).any():
        print(f"  Notas antes do corte que contam no ciclo novo: {', '.join(sorted(NOTAS_DO_CICLO))}")
    nat = m["natureza"].astype(str).str.upper()
    # só as compras dos itens da estimativa (e de versões com a mesma descrição): livros, não brindes/embalagens
    chaves = {chave_versao(it["c"], est) for it in itens}
    eh_item = m["codigo"].map(lambda c: chave_versao(c, est) in chaves)
    rec = m[m["tipo"].eq("ENTRADA") & nat.str.startswith("COMPRA") & eh_item]
    graficas = sorted(rec["parceiro"].fillna("SEM FORNECEDOR").astype(str).str.strip().unique())
    gi = {g: i for i, g in enumerate(graficas)}
    recd = {}
    for (cod, g), q in rec.assign(parceiro=rec["parceiro"].fillna("SEM FORNECEDOR").astype(str).str.strip()) \
            .groupby(["codigo", "parceiro"])["qtd"].sum().items():
        recd.setdefault(int(cod), {})[gi[g]] = float(q)
        item_id(int(cod), str(est["desc"].get(cod, "")), "")
    casar = Casador(props)
    sai = m[(m["tipo"].str.startswith("SA") & (nat.isin(NAT_SAIDA_CLIENTE) | nat.eq(NATUREZA_CONSIG_SAIDA.upper())))
            | (m["tipo"].eq("ENTRADA") & nat.str.contains(RE_DEVOLUCAO_CLIENTE))]
    saidas, sp, sem = {}, {}, 0.0
    for r in sai.itertuples():
        pi = casar(r.parceiro, r.cidade, r.uf)
        q = float(r.qtd) * (1 if str(r.tipo).startswith("SA") else -1)
        if pi is None:
            sem += q
            continue
        cod = int(r.codigo)
        ii = item_id(cod, str(est["desc"].get(cod, "")), "")
        saidas[(pi, ii)] = saidas.get((pi, ii), 0) + q
        sp[cod] = sp.get(cod, 0) + q
    net = m.groupby("codigo")["MOV"].sum()
    n = lambda v: f"{int(v):,}".replace(",", ".")
    print(f"  Corte {corte:%d/%m/%Y}: {len(graficas)} gráficas com recebimento, {n(rec['qtd'].sum())} un. recebidas, "
          f"{n(sum(sp.values()))} un. saíram para clientes das propostas"
          + (f", {n(sem)} un. para clientes sem proposta" if sem else ""))
    return {"corte": f"{corte:%d/%m/%Y}", "graficas": graficas, "saidas": [[pi, ii, q] for (pi, ii), q in saidas.items() if q],
            "rec": recd, "sp": sp, "net": {int(k): float(v) for k, v in net.items()}, "sem_proposta": sem}


def completar_saldos(itens, pos):
    """Saldo no corte + recebido das gráficas − saiu para as propostas ± outras movimentações = saldo atual."""
    for it in itens:
        c = it["c"]
        net = pos["net"].get(c, 0.0)
        rg = pos["rec"].get(c, {})
        sp = pos["sp"].get(c, 0.0)
        it["sc"] = it["s"] - net
        it["rg"] = [[g, q] for g, q in rg.items()]
        it["sp"] = sp
        it["om"] = net - sum(rg.values()) + sp


# ---------------------------------------------------------------- sequência dos programas (combinado em 22/10/2026)
def programa_canonico(nome):
    """Nome do programa (proposta ou cadastro) -> (posição na sequência, rótulo único)."""
    n = norm(nome)
    d = re.search(r"(\d)\s*[ºo°]?\s*ano", n)
    ano = int(d.group(1)) if d else None
    if re.search(r"enraizar|plenitude|movimente|germinar", n):
        return (14, "Programas Verbum")
    if "zero a tres" in n:
        return (1, "Zero a Três")
    if re.search(r"pre-?\s?escola ii|pre ii\b", n):
        return (3, "Pré-Escola II")
    if re.search(r"pre-?\s?escola i|pre i\b", n):
        return (2, "Pré-Escola I")
    if "intensivo" in n:
        return (5, "Intensivo de Alfabetização")
    if "alfabetizacao" in n and ("3" in n and "5" in n) and "revis" not in n:
        return (7, "Alfabetização 3º ao 5º ano")
    if "alfabetizacao" in n and ano == 2:
        return (6, "Alfabetização 2º ano")
    if "alfabetizacao" in n and "revis" not in n and "avalia" not in n and "telcs" not in n:
        return (4, "Alfabetização 1º ano")
    if "ensino estruturado" in n and ano:
        return (8 + ano / 10, f"Ensino Estruturado {ano}º ano")
    if "anos iniciais ii" in n:
        return (9.2, "Revisão dos Anos Iniciais II")
    if "anos iniciais" in n:
        return (9.1, "Revisão dos Anos Iniciais I")
    if "revis" in n and ano and ano <= 5:
        return (9.3 + ano / 100, f"Revisão {ano}º ano")
    if ("revis" in n or "omega" in n) and ano and ano >= 6:
        return (10 + ano / 10, f"Revisão / Coleção Ômega {ano}º ano")
    if "leituras" in n:
        return (11, "Coleção Leituras")
    if "pequenos leitores" in n:
        return (11.1, "Coleção Pequenos Leitores")
    if "bj_pnld" in n or "bom jesus" in n:
        return (12.5, "Bom Jesus (BJ_PNLD)")
    if "brinde" in n:
        return (13, "Brindes")
    return (12, "Outros programas")


def ordenar_por_programa(itens, est):
    """Cada item fica no PRIMEIRO programa da sequência em que aparece (proposta; se não tiver, cadastro).
    Versões do mesmo item ficam no programa da versão que vier primeiro."""
    cadp = est.get("programas", {})
    for it in itens:
        cad = [programa_canonico(p) for p in cadp.get(it["c"], "").split("|") if p.strip()]
        prop = [programa_canonico(p) for p in it["pr"]]
        # (22/10/2026) Se o item vem em vários programas nas propostas (ex.: Grafismo Letras de Forma no kit da Pré I e da
        # Pré II), vale o programa que também está no cadastro da Análise de Estoque. Sem coincidência: o das propostas.
        cad_n = {c[1] for c in cad}
        prog = [p for p in prop if p[1] in cad_n] or prop or cad
        if "BJ_PNLD" in str(it["d"]).upper():   # Cadernos de Testes BJ_PNLD: são do cliente Bom Jesus (22/10/2026)
            prog = [programa_canonico("BJ_PNLD")]
        if not prog and str(est.get("grupo", {}).get(it["c"], "")).startswith("Brinde"):
            prog = [programa_canonico("Brindes")]
        it["po"], it["pn"] = min(prog, default=(15, "Sem programa"))
    grupos = {}
    for it in itens:
        if "g" in it:
            grupos.setdefault(it["g"], []).append(it)
    for ms in grupos.values():
        po, pn = min((m["po"], m["pn"]) for m in ms)
        for m in ms:
            m["po"], m["pn"] = po, pn


def projetar_sem_itens(props, linhas, item_id, est):
    """Propostas SEM itens no Gerador (combinado em 22/10/2026): a demanda é projetada pelo que o IAB enviou a esse
    cliente no ciclo passado (12 meses antes do corte: venda, entrega futura, bonificação e simples remessa, menos
    devoluções), sempre na VERSÃO MAIS NOVA de cada item. Ficam marcadas como projeção ("pj") para destacar."""
    mov = est.get("mov")
    # cliente que já tem outra proposta com itens na demanda não é projetado (não conta em dobro)
    com_itens = {(chave_cidade(p.get("cid") or p["cli"]), p["uf"]) for p in props if p["u"]}
    cands = [i for i, p in enumerate(props) if not p["u"] and p["m"].startswith("Sem itens")
             and (chave_cidade(p.get("cid") or p["cli"]), p["uf"]) not in com_itens]
    if mov is None or not cands:
        return
    corte = pd.Timestamp(CORTE)
    ini = corte - pd.DateOffset(years=1)
    m = mov[(mov["data"] >= ini) & (mov["data"] < corte) & ~_do_ciclo(mov)]
    nat = m["natureza"].astype(str).str.upper()
    m = m[(m["tipo"].str.startswith("SA") & nat.isin(NAT_SAIDA_CLIENTE))
          | (m["tipo"].eq("ENTRADA") & nat.str.contains(RE_DEVOLUCAO_CLIENTE))]
    sub = [dict(props[i], u=True) for i in cands]          # o casador só olha propostas "usadas"
    casar = Casador(sub)
    por_desc = {}
    for c in est["saldo"].index:
        por_desc.setdefault(chave_versao(c, est), []).append(int(c))
    acc = {}
    for r in m.itertuples():
        k = casar(r.parceiro, r.cidade, r.uf)
        if k is None:
            continue
        cod = versao_nova(int(r.codigo), est, por_desc)
        q = float(r.qtd) * (1 if str(r.tipo).startswith("SA") else -1)
        acc[(k, cod)] = acc.get((k, cod), 0) + q
    dmin, dmax = (m["data"].min(), m["data"].max()) if len(m) else (ini, corte)
    kit = f"Projeção: enviado de {max(dmin, ini):%d/%m/%Y} a {min(dmax, corte):%d/%m/%Y}"
    n_prop, n_un = set(), 0.0
    for (k, cod), q in acc.items():
        if q <= 0:
            continue
        pi = cands[k]
        linhas.append([pi, item_id(cod, str(est["desc"].get(cod, "")), ""), q, kit])
        n_prop.add(pi); n_un += q
    for pi in n_prop:
        p = props[pi]
        p.update(u=True, pj=True, m="", ni=sum(1 for l in linhas if l[0] == pi))
    print(f"  Projeção (propostas sem itens): {len(n_prop)} de {len(cands)} clientes com envio no ciclo passado, "
          f"{f'{int(n_un):,}'.replace(',', '.')} un. (versão mais nova de cada item)")


def versao_sem_saldo(cod, est, por_desc=None):
    """Regra (22/10/2026): se `cod` é uma versão ANTIGA e está sem saldo no IAB (zero ou negativo), a demanda vai para
    a versão nova e o código antigo não aparece no relatório. Se a versão antiga tem saldo, fica nela (usa antes)."""
    nova = versao_nova(cod, est, por_desc)
    return nova if nova != cod and float(est["saldo"].get(cod, 0)) <= 0 else cod


def versao_nova(cod, est, por_desc=None):
    """Código mais novo (maior) do IAB com a mesma descrição de `cod`."""
    if por_desc is None:
        por_desc = {}
        for c in est["saldo"].index:
            por_desc.setdefault(chave_versao(c, est), []).append(int(c))
    return max(por_desc.get(chave_versao(cod, est), [cod]) + [cod])


def adicionar_consignacao_verbum(props, linhas, item_id, est):
    """A Verbum Educação é cliente do IAB sem proposta no CRM: tudo o que o IAB enviou a ela em consignação
    entra como uma "proposta" (coluna própria no Item x Clientes), contada como fechada.
    Regra (16/10/2026): se o item enviado é de uma versão antiga que está ZERADA no IAB, a quantidade
    vai para a versão nova (é ela que vai ser enviada da próxima vez)."""
    cons = est.get("consig")
    mov = est.get("mov")
    if mov is not None:   # só o enviado ANTES do corte; o que for enviado depois aparece como "saiu"
        mc = mov[mov["natureza"].eq(NATUREZA_CONSIG_SAIDA) & (mov["data"] < pd.Timestamp(CORTE))]
        cons = mc.groupby("codigo")["qtd"].sum()
    if cons is None or not len(cons):
        return
    por_desc = {}
    for c in est["saldo"].index:
        por_desc.setdefault(chave_versao(c, est), []).append(int(c))
    acum, trocas = {}, []
    for cod, q in cons.items():
        cod, q = int(cod), float(q)
        if q <= 0:
            continue
        destino, kit = cod, "Consignação Verbum"
        nova = versao_sem_saldo(cod, est, por_desc)
        if nova != cod:
            destino, kit = nova, f"Consignação Verbum (enviado como cód. {cod}, versão antiga zerada)"
            trocas.append((cod, nova, q))
        acum[(destino, kit)] = acum.get((destino, kit), 0) + q
    info = est.get("consig_info") or {}
    fim = info.get("fim")
    pi = len(props)
    props.append({
        "id": "CONSIG-VERBUM", "cli": "Verbum Educação (consignação)", "uf": "", "cat": "Parceiro",
        "t": "Fechada", "tc": "", "st": 0, "stx": "Consignação enviada", "tipo": "Consignação",
        "v": 0.0, "al": 0, "vend": "", "coord": "", "cond": "Consignação",
        "dt": "" if fim is None or pd.isna(fim) else f"{fim:%d/%m/%Y}",
        "dts": 0 if fim is None or pd.isna(fim) else fim.timestamp(),
        "ni": len(acum), "u": True, "m": "", "arq": "", "progs": [],
        "cons": {"nfs": info.get("nfs", 0), "un": float(sum(acum.values())),
                 "ini": "" if pd.isna(info.get("ini", pd.NaT)) else f"{info['ini']:%d/%m/%Y}",
                 "fim": "" if fim is None or pd.isna(fim) else f"{fim:%d/%m/%Y}", "trocas": len(trocas)},
    })
    for (cod, kit), q in acum.items():
        linhas.append([pi, item_id(cod, str(est["desc"].get(cod, "")), ""), q, kit])
    print(f"  Verbum (consignação): {int(sum(acum.values())):,} un. em {len(acum)} linhas | ".replace(",", ".")
          + f"{len(trocas)} itens de versão antiga zerada somados na versão nova"
          + (f": {', '.join(f'{a}→{b}' for a, b, _ in trocas[:12])}" if trocas else ""))


def agrupar_versoes(itens, item_id, est):
    """Junta no mesmo grupo ("g") os códigos do IAB com a mesma descrição de um item das propostas.
    Entram as outras versões com saldo positivo (as zeradas não ajudam a desovar)."""
    por_desc = {}
    for cod in est["saldo"].index:
        por_desc.setdefault(chave_versao(cod, est), []).append(int(cod))
    grupos = {}
    for it in list(itens):
        k = chave_versao(it["c"], est)
        outros = [c for c in por_desc.get(k, []) if c != it["c"] and float(est["saldo"].get(c, 0)) > 0]
        if outros:
            grupos.setdefault(k, set()).update(outros + [it["c"]])
    for g, (k, cods) in enumerate(sorted(grupos.items(), key=lambda kv: max(kv[1]))):
        for c in sorted(cods):
            itens[item_id(c, str(est["desc"].get(c, "")), "")]["g"] = g
    n_ext = sum(1 for it in itens if "g" in it and not it["pr"])
    print(f"  Versões: {len(grupos)} itens das propostas têm outra versão em estoque "
          f"({n_ext} códigos de outras versões incluídos)")


# ---------------------------------------------------------------- resumo de vendas (Relatório de faturamento)
# Tipo de operação da nota -> coluna do resumo (sinal: devolução/retorno subtrai)
OPS_RESUMO = {
    "NOTA DE VENDA DE PRODUTO PRODUZIDO": ("v", 1), "VENDA ENTREGA FUTURA - SIMPLES FAT": ("v", 1),
    "NF DEFINIÇÃO DE PREÇO": ("v", 1), "NOTA DE VENDA DE EBOOK": ("v", 1), "VENDA CONSIGNADA": ("v", 1),
    "DEVOLUÇÃO DE VENDA - NF PRÓPRIA": ("v", -1), "DEVOLUÇÃO DE VENDA - NF TERCEIROS": ("v", -1),
    "ESTORNO VDA - ENTREGA FUTURA SIMPLES FAT": ("v", -1), "RETORNO MERC CONSIG COMERCIALIZADA": ("v", -1),
    "REMESSA VENDA ENTREGA FUTURA": ("e", 1), "DEVOLUÇÃO DE REMESSA - ENTREGA FUTURA": ("e", -1),
    "REMESSA DE BONIFICAÇÃO - SAIDA": ("b", 1), "DEVOLUÇÃO DE REMESSA DE BONIFICAÇÃO": ("b", -1),
    "SIMPLES REMESSA (OUTRAS)": ("o", 1), "RETORNO SIMPLES REMESSA (OUTRAS) PROPRIO": ("o", -1),
    "REMESSA EM CONSIGNAÇÃO - P/ VENDA": ("o", 1),
    # "AJUSTE DE ESTOQUE - SAÍDA DE ESTOQUE COM TERCEIROS" NÃO entra (22/10/2026): é baixa do saldo que estava em poder
    # de terceiros (ajuste de estoque), não é envio nem venda ao cliente. No saldo do IAB ela é só uma saída normal.
}
CNPJ_IAB = "08.458.084/0001-13"
# Clientes que recebem como "Simples remessa (outras)" mas é VENDA (combinado em 22/10/2026):
# Pilar/AL - o material é doado e outra empresa paga por recibo (outro documento fiscal).
SIMPLES_REMESSA_E_VENDA = {("pilar", "AL")}


def resumo_vendas(props):
    """Notas do Relatório de faturamento por cliente (destinatário + cidade + UF): venda faturada, entrega futura
    enviada, bonificação e simples remessa/outras saídas, com a condição de pagamento da proposta do cliente."""
    if not ARQ_FATURAMENTO.exists():
        print(f"  Aviso: relatório de faturamento não encontrado ({ARQ_FATURAMENTO}) - aba Resumo de vendas vazia")
        return {"clientes": [], "notas": [], "periodo": ["", ""]}
    try:
        import shutil, tempfile
        tmp = Path(tempfile.gettempdir()) / "faturamento_estimativa.xlsx"
        try:
            shutil.copyfile(ARQ_FATURAMENTO, tmp)
        except PermissionError:
            import subprocess
            subprocess.run(["powershell", "-NoProfile", "-Command",
                            f"Copy-Item -LiteralPath '{ARQ_FATURAMENTO}' -Destination '{tmp}' -Force"], check=True, capture_output=True)
        f = pd.read_excel(tmp)
    except Exception as e:
        print(f"  Aviso: não consegui ler o relatório de faturamento ({e})")
        return {"clientes": [], "notas": [], "periodo": ["", ""]}
    f["op"] = f["TIPO_OPERACAO"].astype(str).str.strip().str.upper()
    f = f[f["op"].isin(OPS_RESUMO) & ~f["CPF_CNPJ_DESTINATARIO"].astype(str).str.strip().eq(CNPJ_IAB)].copy()
    f["dt"] = pd.to_datetime(f["DT_EMISSAO_NF"], dayfirst=True, errors="coerce")
    f["valor"] = pd.to_numeric(f["VALOR_TOTAL_NOTA"], errors="coerce").fillna(0)
    casar = Casador(props, so_usadas=False)
    clientes, idx, notas = [], {}, []
    for r in f.itertuples():
        nome = str(r.CLIENTE_DESTINATARIO or "").strip()
        cid, uf = str(r.CIDADE_DESTINO or "").strip(), str(r.UF_DESTINO or "").strip().upper()
        orgao = eh_orgao_municipal(nome, cid)
        # órgãos municipais (prefeitura, secretaria, fundo...) do mesmo município = um cliente só
        k = ("MUN", chave_cidade(cid), uf) if orgao else (norm(nome), uf)
        if k not in idx:
            pi = casar(nome, cid, uf)
            idx[k] = len(clientes)
            clientes.append({"n": f"Município de {cid.title()}" if orgao else nome, "cid": cid.title(), "uf": uf,
                             "doc": str(r.CPF_CNPJ_DESTINATARIO or ""), "pi": -1 if pi is None else pi, "dest": []})
        c = clientes[idx[k]]
        if c["pi"] < 0:
            pi = casar(nome, cid, uf)
            c["pi"] = -1 if pi is None else pi
        if nome not in c["dest"]:
            c["dest"].append(nome)
        col, sinal = OPS_RESUMO[r.op]
        if col == "o" and "SIMPLES REMESSA" in r.op and (chave_cidade(cid), uf) in {(chave_cidade(c), u) for c, u in SIMPLES_REMESSA_E_VENDA}:
            col = "v"
        nfx = _nf_txt(r.NRO_NFE)
        notas.append([idx[k], 0 if pd.isna(r.dt) else int(r.dt.timestamp()) + 43200, col, round(sinal * float(r.valor), 2),
                      nfx, 1 if nfx in NOTAS_DO_CICLO else 0])
    per = [f"{f['dt'].min():%d/%m/%Y}", f"{f['dt'].max():%d/%m/%Y}"] if len(f) else ["", ""]
    print(f"  Resumo de vendas: {len(notas)} notas de {len(clientes)} clientes ({per[0]} a {per[1]}), "
          f"{sum(1 for c in clientes if c['pi'] >= 0)} ligados a uma proposta")
    return {"clientes": clientes, "notas": notas, "periodo": per}


def main():
    print("Estimativa de Vendas - gerando...")
    print(f"  Propostas: {PASTA_PROPOSTAS}")
    regs = escolher_propostas(carregar_propostas())
    print(f"  {len(regs)} propostas lidas | {sum(r['usada'] and bool(r['itens']) for r in regs)} entram na demanda (com itens)")
    est = carregar_estoque()
    zerar_negativos(est)
    descarte = aplicar_descarte(est)
    est["dsc_por"] = descarte["por"]
    props, itens, linhas, pos = montar(regs, est)
    vendas = resumo_vendas(props)
    sem_erp = [i["c"] for i in itens if not i["erp"]]
    if sem_erp:
        print(f"  Aviso: {len(sem_erp)} códigos das propostas não existem na movimentação: {sem_erp[:15]}")
    dados = {
        "versao": VERSAO_MODELO,
        "gerado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "base_estoque": "" if pd.isna(est["data_base"]) else f"{est['data_base']:%d/%m/%Y}",
        "pcts": pcts_crm(),
        "descarte": {k: v for k, v in descarte.items() if k != "por"},
        "corte": pos["corte"], "graficas": pos["graficas"], "saidas": pos["saidas"], "sem_proposta": pos["sem_proposta"],
        "vendas": vendas,
        "props": props, "itens": itens, "linhas": linhas,
    }
    if not ARQ_MODELO.exists():
        sys.exit(f"ERRO: modelo não encontrado: {ARQ_MODELO}")
    html = ARQ_MODELO.read_text(encoding="utf-8")
    m = re.search(r'data-versao="([^"]+)"', html)
    if m and m.group(1) != VERSAO_MODELO:
        print(f"  Aviso: versão do modelo ({m.group(1)}) diferente da do script ({VERSAO_MODELO})")
    payload = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    if "/*__DADOS__*/null" not in html:
        sys.exit("ERRO: marcador /*__DADOS__*/null não encontrado no modelo")
    html = html.replace("/*__DADOS__*/null", payload, 1)
    ARQ_SAIDA.write_text(html, encoding="utf-8")
    print(f"OK -> {ARQ_SAIDA}")


if __name__ == "__main__":
    main()
