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
# 06/11/2026: os dois reenvios já saíram e estão fora do saldo, então saíram da lista:
#   P5OC43 Espírito Santo do Pinhal/SP - NF 19117 de 26/10/2026 (R$ 435.306)
#   PN00HE São Roque/SP - NF 19079 de 06/10/2026 (R$ 19.586, mesmo valor da proposta)
REENVIAR = {}
# Notas emitidas ANTES do corte que pertencem ao ciclo 2026/2027: aparecem no Resumo de vendas "desde o corte".
# (06/11/2026) Não entram mais no "já saiu" das propostas: a nota é anterior ao corte, então já está fora do saldo.
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
CORTE = os.environ.get("ESTIMATIVA_CORTE", "2026-10-01")   # 06/11/2026: corte passou de 01/11 para 01/10/2026
NATUREZA_CONSIG_SAIDA = "REMESSA EM CONSIGNAÇÃO PARA A VERBUM (saída criada pelo script)"
ARQ_FATURAMENTO = PASTA_RELATORIOS / "Relatorio de faturamento.xlsx"
ABA_DESCARTE = "Dados para Contabilidade"
# Itens com SALDO ZERO na estimativa (não podem ser enviados; a baixa no ERP vem depois). Código: motivo.
ITENS_SALDO_ZERO = {1803: "item com erro - não pode ser enviado; baixa no estoque futuramente (10/11/2026)"}
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
def saldo_embalagens(ga, df, cad):
    """(11/11/2026) Saldo atual dos itens do programa EMBALAGEM (caixas, fitas, filmes, saquinhos) por empresa
    (IAB; Verbum = Filial + Matriz), pela mesma regra do saldo: implantação + entradas − saídas."""
    try:
        rx = getattr(ga, "CLASSES_POR_DESCRICAO", {}).get("EMBALAGEM") or r"CAIXA|FITAS? ADESIVAS?|FILME|STRE[TS]C?H|EMBALAGEM"
        cods = set(cad.loc[cad["Programas"].fillna("").astype(str).str.upper().str.contains("EMBALAGEM"), "Codigo"].astype(int))
        cods |= set(df.loc[df["descricao"].fillna("").astype(str).str.upper().str.contains(rx, regex=True), "codigo"].astype(int))
        e = df[df["codigo"].isin(cods)].copy()
        if e.empty: return [], [], []
        impl = e["natureza"].str.upper().eq(ga.NATUREZA_IMPLANTACAO)
        ent = e["tipo"].eq("ENTRADA") & ~impl
        sai = e["tipo"].astype(str).str.startswith("SA")
        e["MOV"] = e["qtd"].where(impl, 0) + e["qtd"].where(ent, 0) - e["qtd"].where(sai, 0)
        e["EMPX"] = e["empresa_nome"].astype(str).apply(lambda x: "Verbum" if x.startswith("Verbum") else x)
        e = e[e["EMPX"].isin(["IAB", "Verbum"])]
        g = e.groupby(["EMPX", "codigo"]).agg(s=("MOV", "sum"), ult=("data", "max"), d=("descricao", "last")).reset_index()
        out = [{"emp": r.EMPX, "cod": int(r.codigo), "d": str(r.d).split("|")[0].strip(), "s": round(float(r.s), 2),
                "ult": "" if pd.isna(r.ult) else f"{r.ult:%d/%m/%Y}"} for r in g.itertuples()]
        # (11/11/2026) notas de COMPRA das embalagens (quantidade e valor unitário) + último custo conhecido de cada item
        cu = pd.to_numeric(e["custo_unit"], errors="coerce")
        comp = e[ent & e["natureza"].str.upper().str.startswith("COMPRA")].copy()
        compras = [{"dt": f"{r.data:%d/%m/%Y}" if not pd.isna(r.data) else "", "emp": r.EMPX, "nf": _nf_txt(r.nf),
                    "forn": r.parceiro.strip() if isinstance(r.parceiro, str) else "", "cod": int(r.codigo), "d": str(r.descricao).split("|")[0].strip(),
                    "q": float(r.qtd), "cu": None if pd.isna(c) or c <= 0 else round(float(c), 4)}
                   for r, c in zip(comp.sort_values("data").itertuples(), cu[comp.sort_values("data").index])]
        ref = e[(ent | impl) & (cu > 0)].assign(cu=cu).sort_values("data").groupby("codigo").agg(cu=("cu", "last"), dt=("data", "last"), nat=("natureza", "last"))
        for o in out:
            if o["cod"] in ref.index:
                r = ref.loc[o["cod"]]
                o["cu"] = round(float(r["cu"]), 4); o["cud"] = f"{r['dt']:%d/%m/%Y}"; o["cun"] = str(r["nat"]).strip().title()
        # (11/11/2026) consumo = saídas no período da movimentação (baixa por ajuste/descarte, requisição interna e outras)
        sx = e[sai].copy()
        nat = sx["natureza"].str.upper()
        sx["grp"] = "ou"
        sx.loc[nat.str.contains("REQUISI", na=False), "grp"] = "rq"
        sx.loc[nat.str.contains("AJUSTE|DESCARTE", na=False), "grp"] = "aj"
        cs = sx.pivot_table(index=["EMPX", "codigo"], columns="grp", values="qtd", aggfunc="sum", fill_value=0)
        for o in out:
            k = (o["emp"], o["cod"])
            if k in cs.index:
                r = cs.loc[k]
                o["c_aj"], o["c_rq"], o["c_ou"] = (float(r.get(x, 0)) for x in ("aj", "rq", "ou"))
        per = [f"{df['data'].min():%d/%m/%Y}", f"{df['data'].max():%d/%m/%Y}"] if len(df) else ["", ""]
        print(f"  Embalagens: saldo de {len(out)} itens (IAB e Verbum) | {len(compras)} entradas de compra | consumo {per[0]} a {per[1]}")
        return out, compras, per
    except Exception as ex:
        print(f"  Aviso: saldo das embalagens não calculado ({type(ex).__name__}: {ex})")
        return [], [], []


# (12/11/2026) Etiquetas (papel couchê por cor de tarja + etiqueta branca 10x8). O gerar_analise ignora "PAPEL COUCH",
# então o saldo é lido direto da planilha de Movimentação. Compras vieram ora em etiqueta (R$ 0,10), ora em rolo de
# 400 (R$ 39,48): movimento com custo unitário >= R$ 10 é tratado como ROLO e convertido em etiquetas.
ETIQUETAS_COD = {2263: "ETIQUETA BRANCA 10X8 CM"}
RE_ETIQUETA = r"PAPEL COUCH|ETIQUETA BRANCA"
ETQ_POR_ROLO = 400


def saldo_etiquetas(ga):
    try:
        arq = ga.ARQ_MOVIMENTACAO
        m = pd.read_excel(arq, usecols=["CODIGO_PRODUTO", "DESCRICAO_PRODUTO", "EMPRESA_ESTOQUE", "TIPO_MOVIMENTO",
                                        "NATUREZA_TIPO_USO", "QUANTIDADE", "CUSTO_UNITARIO", "DATA_MOVIMENTO",
                                        "NUMERO_NF", "NOME_PARCEIRO"])
        m = m[~m["EMPRESA_ESTOQUE"].astype(str).str.upper().str.contains("LINCKY")]
        m["cod"] = pd.to_numeric(m["CODIGO_PRODUTO"], errors="coerce")
        m = m[m["cod"].isin(list(ETIQUETAS_COD)) | m["DESCRICAO_PRODUTO"].astype(str).str.upper().str.contains(RE_ETIQUETA, regex=True)
              & ~m["DESCRICAO_PRODUTO"].astype(str).str.upper().str.contains("IMPRESSORA|TONER|QUE ETIQUETA", regex=True)].copy()
        per = []
        if len(m):
            dt = pd.to_datetime(m["DATA_MOVIMENTO"], errors="coerce")
            per = [f"{dt.min():%d/%m/%Y}", f"{dt.max():%d/%m/%Y}"]
        cu = pd.to_numeric(m["CUSTO_UNITARIO"], errors="coerce")
        q = pd.to_numeric(m["QUANTIDADE"], errors="coerce").fillna(0)
        fat = cu.ge(10).map({True: ETQ_POR_ROLO, False: 1})          # rolo -> etiquetas
        m["et"] = q * fat
        m["cu_et"] = (cu / fat).where(cu > 0)
        nat = m["NATUREZA_TIPO_USO"].astype(str).str.upper()
        tp = m["TIPO_MOVIMENTO"].astype(str).str.upper()
        impl = nat.str.contains("IMPLANTA")
        ent = tp.eq("ENTRADA") & ~impl
        sai = tp.str.startswith("SA")
        m["mov"] = m["et"].where(impl | ent, 0) - m["et"].where(sai, 0)
        m["sai"] = m["et"].where(sai, 0)
        m["dt"] = pd.to_datetime(m["DATA_MOVIMENTO"], errors="coerce")
        out = []
        cods = sorted(set(m["cod"].dropna().astype(int)) | set(ETIQUETAS_COD))
        for c in cods:
            g = m[m["cod"] == c]
            d = str(g["DESCRICAO_PRODUTO"].iloc[-1]).split("|")[0].strip() if len(g) else ETIQUETAS_COD.get(c, "")
            comp = g[ent.reindex(g.index, fill_value=False) & nat.reindex(g.index).str.startswith("COMPRA") & g["cu_et"].notna()].sort_values("dt")
            o = {"cod": int(c), "d": d, "mov": int(len(g)), "s": round(float(g["mov"].sum()), 0), "c": round(float(g["sai"].sum()), 0),
                 "ult": "" if g.empty or pd.isna(g["dt"].max()) else f"{g['dt'].max():%d/%m/%Y}"}
            if len(comp):
                r = comp.iloc[-1]
                o.update(cu=round(float(r["cu_et"]), 4), cud=f"{r['dt']:%d/%m/%Y}", nf=_nf_txt(r["NUMERO_NF"]),
                         forn=str(r["NOME_PARCEIRO"]).strip(), ce=round(float(comp["et"].sum()), 0))
            out.append(o)
        print(f"  Etiquetas: {len(out)} itens (saldo e consumo em etiquetas; rolo = {ETQ_POR_ROLO})")
        return out, per
    except Exception as ex:
        print(f"  Aviso: etiquetas não calculadas ({type(ex).__name__}: {ex})")
        return [{"cod": c, "d": d, "mov": 0} for c, d in ETIQUETAS_COD.items()], []


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
    emb, emb_compras, emb_per = saldo_embalagens(ga, df, cad)
    etq, etq_per = saldo_etiquetas(ga)
    # (19/11/2026) saldo da Verbum (Filial + Matriz) por código, para a página Produção Verbum
    try:
        vbx = df[df["empresa_nome"].astype(str).str.startswith("Verbum")]
        impv = vbx["natureza"].str.upper().eq(ga.NATUREZA_IMPLANTACAO)
        entv = vbx["tipo"].eq("ENTRADA") & ~impv
        saiv = vbx["tipo"].astype(str).str.startswith("SA")
        saldo_vb = (vbx["qtd"].where(impv, 0) + vbx["qtd"].where(entv, 0) - vbx["qtd"].where(saiv, 0)).groupby(vbx["codigo"]).sum()
        saldo_vb = {int(k): float(v) for k, v in saldo_vb.items()}
    except Exception as ex:
        print(f"  Aviso: saldo da Verbum por código não calculado ({ex})"); saldo_vb = {}              # (11/11/2026) caixas, fitas, saquinhos - para o quadro de insumos
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
    # saídas da Verbum para os clientes dela (escolas): usadas para casar com a consignação IAB → Verbum (06/11/2026)
    vbm = df[df["empresa_nome"].astype(str).str.startswith("Verbum") & df["tipo"].astype(str).str.upper().str.startswith("SA")
             & ~df["parceiro"].fillna("").astype(str).str.upper().str.contains("ALFA E BETO|VERBUM EDUCA", regex=True)]
    mov_vb = vbm[["codigo", "data", "natureza", "parceiro", "cidade", "uf", "qtd"]].copy()
    ult_compra = iab[ent & iab["natureza"].str.upper().str.startswith("COMPRA")].groupby("codigo")["data"].max().to_dict()
    return {"saldo": saldo, "consig": consig, "consig_info": consig_info, "desc": desc, "ult": ult, "emb": emb, "emb_compras": emb_compras, "emb_per": emb_per, "etq": etq, "etq_per": etq_per, "saldo_vb": saldo_vb,
            "descartes_erp": descartes_erp, "mov": mov, "mov_vb": mov_vb, "ult_compra": ult_compra,
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
    pos = movimento_pos_corte(props, itens, item_id, est, linhas)
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


RE_ESCOLA = r"col[eé]gio|escola|educacional|centro de educa|instituto educa|creche|ensino|pedag"


def eh_escola(nome, cidade=""):
    """Escola / colégio particular (cliente da Verbum). Escolas municipais/estaduais e órgãos públicos não entram."""
    n = norm(nome)
    return (bool(re.search(norm(RE_ESCOLA), n)) and not re.search(r"municipa|estadua|federa|prefeit|secretaria", n)
            and not eh_orgao_municipal(nome, cidade))


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


def movimento_pos_corte(props, itens, item_id, est, linhas=None):
    """Depois do corte: recebimentos por gráfica (notas de compra) e saídas para os clientes das propostas."""
    corte = pd.Timestamp(CORTE)
    vazio = {"corte": f"{corte:%d/%m/%Y}", "graficas": [], "saidas": [], "rec": {}, "sp": {}, "net": {}, "sem_proposta": 0}
    mov = est.get("mov")
    if mov is None:
        return vazio
    m = mov[mov["data"] >= corte]
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
    # consignação IAB → Verbum NÃO entra direto: só abate quando a Verbum fatura o mesmo item para a escola (abaixo)
    sai = m[(m["tipo"].str.startswith("SA") & nat.isin(NAT_SAIDA_CLIENTE))
            | (m["tipo"].eq("ENTRADA") & nat.str.contains(RE_DEVOLUCAO_CLIENTE))]
    saidas, sp, sem = {}, {}, 0.0
    # (06/11/2026) Enviou outra versão do mesmo livro (ex.: pediu 1973, saiu 1392): abate a linha da proposta mesmo assim,
    # porque a proposta pode ser ajustada no envio. Vale quando a proposta não tem o código exato que saiu.
    na_prop, por_versao = set(), {}
    for l in (linhas or []):
        na_prop.add((l[0], l[1]))
        por_versao.setdefault((l[0], chave_versao(itens[l[1]]["c"], est)), l[1])
    trocadas = 0.0
    vb = casar.vb   # "proposta" da consignação Verbum
    # colunas nomeadas antes do saldo: escolas via Verbum e clientes com proposta "Material enviado" (fora da demanda)
    esc_nomes, esc = [], {}
    def _col(nome, cod, q):
        if nome not in esc_nomes: esc_nomes.append(nome)
        j = esc_nomes.index(nome)
        esc.setdefault(cod, {})[j] = esc.get(cod, {}).get(j, 0) + q
    # (10/11/2026) proposta "Material enviado" fora da demanda (ex.: Pinhal, São Roque): o envio depois do corte
    # aparece numa coluna com o nome do cliente, em vez de ficar escondido em "± Outras mov."
    env_idx = [i for i, p in enumerate(props) if not p.get("u") and p.get("st") == 8]
    casar_env = Casador([dict(props[i], u=True) for i in env_idx]) if env_idx else None
    enviados = 0.0
    for r in sai.itertuples():
        pi = casar(r.parceiro, r.cidade, r.uf)
        q = float(r.qtd) * (1 if str(r.tipo).startswith("SA") else -1)
        if pi is None:
            k = casar_env(r.parceiro, r.cidade, r.uf) if casar_env else None
            if k is not None:
                p_ = props[env_idx[k]]
                _col(f"{p_['cli']} / {p_['uf']} (material enviado)", int(r.codigo), q)
                sp[int(r.codigo)] = sp.get(int(r.codigo), 0) + q
                enviados += q
                continue
            sem += q
            continue
        cod = int(r.codigo)
        ii = item_id(cod, str(est["desc"].get(cod, "")), "")
        if (pi, ii) not in na_prop:
            alt = por_versao.get((pi, chave_versao(cod, est)))
            if alt is not None:
                ii = alt; trocadas += q
        saidas[(pi, ii)] = saidas.get((pi, ii), 0) + q
        sp[cod] = sp.get(cod, 0) + q
    # (06/11/2026) Consignação IAB → Verbum depois do corte casada com a NF da Verbum para a escola (mesmo item):
    # a quantidade casada abate a demanda da Verbum e aparece numa coluna com o nome da escola, antes do saldo.
    escolas, cons_sem = 0.0, 0.0
    cons = m[m["tipo"].str.startswith("SA") & nat.eq(NATUREZA_CONSIG_SAIDA.upper())]
    mvb = est.get("mov_vb")
    if len(cons):
        fila = {}
        for r in cons.sort_values("data").itertuples():
            fila.setdefault(int(r.codigo), []).append([r.data, float(r.qtd)])
        if mvb is not None and len(mvb) and vb is not None:
            nvb = mvb["natureza"].astype(str).str.upper()
            vs = mvb[(mvb["data"] >= corte) & nvb.isin(NAT_SAIDA_CLIENTE)].sort_values("data")
            for r in vs.itertuples():
                cod = int(r.codigo); falta = float(r.qtd)
                for e in fila.get(cod, []):
                    if falta <= 0: break
                    if e[0] > r.data or e[1] <= 0: continue
                    q = min(falta, e[1]); e[1] -= q; falta -= q
                    _col(str(r.parceiro or "").strip().title() + " (via Verbum)", cod, q)
                    ii = item_id(cod, str(est["desc"].get(cod, "")), "")
                    if (vb, ii) not in na_prop:
                        alt = por_versao.get((vb, chave_versao(cod, est)))
                        if alt is not None: ii = alt
                    saidas[(vb, ii)] = saidas.get((vb, ii), 0) + q
                    sp[cod] = sp.get(cod, 0) + q
                    escolas += q
        cons_sem = sum(e[1] for l in fila.values() for e in l)
    net = m.groupby("codigo")["MOV"].sum()
    # (06/11/2026) Ajustes de estoque depois do corte (entrada, saída/descarte, com terceiros) entram no SALDO DO CORTE,
    # como se fossem implantação: não aparecem em "± Outras mov." e não se misturam com clientes e fornecedores.
    aj = m[nat.str.contains("AJUSTE")].groupby("codigo")["MOV"].sum()
    n = lambda v: f"{int(v):,}".replace(",", ".")
    print(f"  Corte {corte:%d/%m/%Y}: {len(graficas)} gráficas com recebimento, {n(rec['qtd'].sum())} un. recebidas, "
          f"{n(sum(sp.values()))} un. saíram para clientes das propostas"
          + (f", {n(sem)} un. para clientes sem proposta" if sem else "")
          + (f", {n(trocadas)} un. enviadas em outra versão abatidas da proposta" if trocadas else "")
          + (f", {n(enviados)} un. para clientes com proposta Material enviado (coluna própria)" if enviados else "")
          + (f", {n(escolas)} un. consignadas à Verbum e já faturadas para escolas" if escolas else "")
          + (f", {n(cons_sem)} un. consignadas à Verbum ainda sem NF da Verbum para escola" if cons_sem else ""))
    return {"corte": f"{corte:%d/%m/%Y}", "graficas": graficas, "saidas": [[pi, ii, q] for (pi, ii), q in saidas.items() if q],
            "rec": recd, "sp": sp, "net": {int(k): float(v) for k, v in net.items()}, "sem_proposta": sem,
            "esc_nomes": esc_nomes, "esc": esc, "aj": {int(k): float(v) for k, v in aj.items()}}


def completar_saldos(itens, pos):
    """Saldo no corte (+ ajustes de estoque feitos depois dele) + recebido das gráficas − saiu para as propostas
    ± outras movimentações = saldo atual."""
    for it in itens:
        c = it["c"]
        net = pos["net"].get(c, 0.0) - pos.get("aj", {}).get(c, 0.0)   # sem os ajustes: eles ficam no saldo do corte
        rg = pos["rec"].get(c, {})
        sp = pos["sp"].get(c, 0.0)
        it["sc"] = it["s"] - net
        it["rg"] = [[g, q] for g, q in rg.items()]
        it["sp"] = sp
        it["ve"] = [[j, q] for j, q in pos.get("esc", {}).get(c, {}).items()]
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
# Notas que NÃO entram no resumo de vendas (exceção combinada com Anderson): (nº NF, trecho do destinatário) -> motivo
NOTAS_IGNORAR_VENDAS = {
    ("19098", "AMARANTE"): "devolução de Amarante (15/10/2026) - a troca já foi enviada na NF 18711 de 23/06/2026",
    ("19099", "AMARANTE"): "devolução de Amarante (15/10/2026) - a troca já foi enviada na NF 18711 de 23/06/2026",
    ("19101", "AMARANTE"): "devolução de Amarante (15/10/2026) - a troca já foi enviada na NF 18711 de 23/06/2026",
}
# Clientes que recebem como "Simples remessa (outras)" mas é VENDA (combinado em 22/10/2026):
# Pilar/AL - o material é doado e outra empresa paga por recibo (outro documento fiscal).
SIMPLES_REMESSA_E_VENDA = {("pilar", "AL")}
# Colégios da Verbum que recebem por SIMPLES REMESSA / OUTRAS SAÍDAS mas é VENDA (05/11/2026): (trecho do nome, UF)
CLIENTES_SR_VENDA = [("CENTRO EDUCACIONAL CAMINHO", "RJ"), ("ELENA GUERRA", "RJ"), ("MAR ABERTO", "RJ"),
                     ("MONTE CARMELO", "MG"), ("SAFARI KIDS", "SP")]


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
    clientes, idx, notas, excl = [], {}, [], []
    for r in f.itertuples():
        nome = str(r.CLIENTE_DESTINATARIO or "").strip()
        ign = next((m for (nf, t), m in NOTAS_IGNORAR_VENDAS.items() if _nf_txt(r.NRO_NFE) == nf and t in norm(nome).upper()), None)
        if ign:
            excl.append({"nf": _nf_txt(r.NRO_NFE), "cli": nome, "dt": "" if pd.isna(r.dt) else f"{r.dt:%d/%m/%Y}",
                         "v": round(float(r.valor), 2), "motivo": ign})
            continue
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
        if col == "o" and "SIMPLES REMESSA" in r.op and ((chave_cidade(cid), uf) in {(chave_cidade(c), u) for c, u in SIMPLES_REMESSA_E_VENDA}
                                                       or any(norm(t) in norm(nome) and uf == u for t, u in CLIENTES_SR_VENDA)):
            col = "v"
        nfx = _nf_txt(r.NRO_NFE)
        # (23/11/2026) o relatório já traz devolução/retorno com valor NEGATIVO: o sinal vale sobre o valor absoluto
        # (antes, sinal × valor negativo virava positivo e a devolução SOMAVA na venda/bonificação)
        notas.append([idx[k], 0 if pd.isna(r.dt) else int(r.dt.timestamp()) + 43200, col, round(sinal * abs(float(r.valor)), 2),
                      nfx, 1 if nfx in NOTAS_DO_CICLO else 0, 1 if ("DEVOLU" in r.op or "RETORNO" in r.op) else 0])
    per = [f"{f['dt'].min():%d/%m/%Y}", f"{f['dt'].max():%d/%m/%Y}"] if len(f) else ["", ""]
    print(f"  Resumo de vendas: {len(notas)} notas de {len(clientes)} clientes ({per[0]} a {per[1]}), "
          f"{sum(1 for c in clientes if c['pi'] >= 0)} ligados a uma proposta")
    if excl:
        print(f"  Resumo de vendas: {len(excl)} notas fora por exceção ({', '.join(e['nf'] for e in excl)})")
    return {"clientes": clientes, "notas": notas, "periodo": per, "excl": excl}


ARQ_INSUMOS = PASTA / "Base Insumos - Itens por Caixa.xlsx"


def carregar_insumos():
    """(11/11/2026) Parâmetros de embalagem (itens por caixa, modelo, fita, etiquetas) levantados dos romaneios
    de 2025/2026 - arquivo 'Base Insumos - Itens por Caixa.xlsx' (as células amarelas podem ser ajustadas)."""
    if not ARQ_INSUMOS.exists():
        print(f"  Aviso: base de insumos não encontrada ({ARQ_INSUMOS.name}) - quadro de insumos fica vazio")
        return None
    cache = PASTA / "insumos_base.json"
    try:
        import openpyxl, shutil, tempfile
        # copia antes de ler: funciona mesmo com a planilha aberta no Excel / sincronizando no OneDrive
        tmp = Path(tempfile.gettempdir()) / "base_insumos_tmp.xlsx"
        shutil.copyfile(ARQ_INSUMOS, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True)
        num = lambda v: float(v) if isinstance(v, (int, float)) else None
        mods = {}
        for r in wb["Modelos de Caixa"].iter_rows(min_row=2, values_only=True):
            if r[0] is None: continue
            mods[str(r[0]).strip()] = {"emp": r[1], "cod": r[2], "m": r[3], "f": num(r[7]) or 0, "e": num(r[8]) or 1,
                                       "a": num(r[4]), "l": num(r[5]), "c": num(r[6])}   # medidas em cm (A x L x C)
        itens = {}
        for r in wb["Itens por Caixa"].iter_rows(min_row=2, values_only=True):
            if r[0] is None or not num(r[5]): continue
            itens[str(int(r[0]))] = {"k": str(r[4] or "").strip(), "cap": num(r[5]), "pk": num(r[9]), "kgu": num(r[10]), "cx": num(r[11]) or 0}
        kits = {}
        for r in wb["Kits por Caixa"].iter_rows(min_row=2, values_only=True):
            if r[0] is None: continue
            kits[str(r[0]).strip()] = {"k": str(r[2] or "").strip(), "n": num(r[4]), "pk": num(r[6]), "e": num(r[10]) if len(r) > 10 else None}
        print(f"  Base de insumos: {len(itens)} itens, {len(kits)} kits, {len(mods)} modelos de caixa")
        res = {"mods": mods, "itens": itens, "kits": kits}
        try: cache.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
        except Exception: pass
        return res
    except Exception as e:
        print(f"  Aviso: não foi possível ler a base de insumos ({type(e).__name__}: {e})")
        if cache.exists():
            print("  Usando a última base de insumos lida com sucesso (insumos_base.json)")
            return json.loads(cache.read_text(encoding="utf-8"))
        return None


def carregar_ecommerce_pf():
    """(11/11/2026) Envios para pessoa física com até 2 volumes (saquinho de e-commerce), gravado pelo gerar_dashboard.py
    em ecommerce_pf.json. Projeção do ciclo = volumes do Ano Ref anterior completo (out-set)."""
    arq = PASTA / "ecommerce_pf.json"
    if not arq.exists():
        return None
    try:
        j = json.loads(arq.read_text(encoding="utf-8"))
        hoje = datetime.now()
        atual = hoje.year + 1 if hoje.month >= 10 else hoje.year
        out = {}
        for emp, anos in (j.get("por_emp") or {}).items():
            ref = anos.get(str(atual - 1), {"envios": 0, "volumes": 0})
            cur = anos.get(str(atual), {"envios": 0, "volumes": 0})
            out["Verbum" if emp.upper().startswith("VERBUM") else "IAB"] = {
                "ref": str(atual - 1), "envios": ref["envios"], "volumes": ref["volumes"],
                "atual": str(atual), "envios_at": cur["envios"], "volumes_at": cur["volumes"]}
        print("  E-commerce PF (até 2 volumes): " + " | ".join(f"{k} {v['ref']}: {v['volumes']} saquinhos" for k, v in out.items()))
        return {"gerado": j.get("gerado", ""), "por_emp": out}
    except Exception as e:
        print(f"  Aviso: não foi possível ler ecommerce_pf.json ({e})")
        return None


def carregar_demanda_verbum(est):
    """(19/11/2026) Previsão de demanda da Verbum (planilha 'Demanda Verbum *.xlsx' na pasta estoque). A Verbum não usa as
    propostas do CRM: a produção é a coluna 'Sugestão de Produção' de cada aba (Resumo fica de fora)."""
    arqs = sorted(PASTA.glob("Demanda*Verbum*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not arqs:
        return None
    arq = arqs[-1]
    try:
        import openpyxl, shutil, tempfile
        tmp = Path(tempfile.gettempdir()) / "demanda_verbum_tmp.xlsx"
        shutil.copyfile(arq, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True)
        num = lambda v: float(v) if isinstance(v, (int, float)) else None
        sv = est.get("saldo_vb", {}) or {}
        itens, notas = [], []
        for ws in wb:
            linhas = list(ws.iter_rows(values_only=True))
            if not linhas:
                continue
            if ws.title.strip().lower() == "resumo":
                notas = [str(r[0]).strip() for r in linhas if r and isinstance(r[0], str) and (r[0].startswith("•") or r[0].startswith("Arquivo"))]
                continue
            cab = [str(c or "").strip().lower() for c in linhas[0]]
            def col(*nomes):
                for n in nomes:
                    for i, c in enumerate(cab):
                        if c.startswith(n):
                            return i
                return None
            ic, inn, icat, ifm, inec, iest, isug, ireg = (col("código", "codigo"), col("nome"), col("categoria"), col("formato"),
                                                        col("necessidade"), col("estoque"), col("sugestão", "sugestao"), col("regra"))
            if ic is None or isug is None:
                continue
            for r in linhas[1:]:
                g = lambda i: r[i] if i is not None and i < len(r) else None
                try:
                    cod = int(str(g(ic)).strip())
                except Exception:
                    continue
                itens.append({"aba": ws.title, "c": cod, "n": str(g(inn) or "").strip(), "cat": str(g(icat) or "").strip(),
                              "fmt": str(g(ifm) or "").strip(), "nec": num(g(inec)) or 0, "sug": num(g(isug)) or 0,
                              "ep": num(g(iest)), "reg": str(g(ireg) or "").strip(), "s": sv.get(cod)})
        print(f"  Demanda Verbum: {len(itens)} itens ({arq.name}) | produção sugerida {int(sum(i['sug'] for i in itens)):,} un.".replace(",", "."))
        return {"arq": arq.name, "dt": datetime.fromtimestamp(arq.stat().st_mtime).strftime("%d/%m/%Y %H:%M"), "itens": itens, "notas": notas}
    except Exception as e:
        print(f"  Aviso: não foi possível ler a demanda da Verbum ({type(e).__name__}: {e})")
        return None


# (19/11/2026) Caixas da produção Verbum: os códigos da Verbum não aparecem nos romaneios, então cada item usa a
# quantidade por caixa dos MESMOS títulos já enviados pelo IAB (códigos de referência abaixo, lidos da base de insumos).
# Livro do professor (LP) usa o parâmetro do livro do aluno. Caixa IAB 220 → Verbum 220 (2025); Germinar → Verbum 515 (2024).
VB_CX_REGRAS = [
    # (aba regex, nome regex, códigos de referência na base, descrição da origem)
    (r"Enraizar - LPM", r"", [1863, 1865, 1884, 1886, 1905, 1907, 1926], "Enraizar LPM já enviado (vol. III/IV)"),
    (r"Enraizar - HGC", r"", [1864, 1866, 1885, 1887, 1906], "Enraizar HGC já enviado (vol. III/IV)"),
    (r"Germinar", r"", [2031, 2032, 2035, 2036], "Germinar já enviado (Pré I/II vol. III/IV)"),
    (r"Plenitude", r"", [2006], "Plenitude 2º ano vol. II"),
    (r"Movimente", r"", [2011], "Movimente 3º ano vol. II"),
    (r"Consci", r"", [1864, 1866, 1885, 1887, 1906], "estimado: sem envio próprio, usa o Enraizar HGC (mesmo formato)"),
    (r"Grafismo", r"Introdu", [1380], "Grafismo e Caligrafia Introdução (IAB)"),
    (r"Grafismo", r"Cursiva", [1381], "Grafismo e Caligrafia Letras Cursivas (IAB)"),
    (r"Grafismo", r"Forma", [1379], "Grafismo e Caligrafia Letras de Forma (IAB)"),
    (r"Grafismo", r"Mini ?livro", [2059], "Minilivros (IAB)"),
    (r"Kits", r"EVA \(54", [25], "Letras do Alfabeto em EVA (IAB)"),
    (r"Kits", r"Cartaz", [23], "Cartaz do Alfabeto (IAB)"),
    (r"Kits", r"Cartela", [24], "Cartelas de Letras do Alfabeto (IAB)"),
]
VB_CX_MODELO = {"236": "2025", "547": "2026", "562": "2027", "2023": "2024", "240": "2024", "2022": "2024"}


def caixas_demanda_verbum(dv, ins):
    """Anexa a cada item da demanda Verbum o parâmetro de caixa {k, cap, kgu, ref}. Sem regra → sem parâmetro (fica de fora)."""
    if not dv or not ins:
        return dv
    from statistics import median
    base = ins.get("itens", {})
    for x in dv["itens"]:
        x["cx"] = None
        for ra, rn, refs, orig in VB_CX_REGRAS:
            if re.search(ra, x["aba"], re.I) and (not rn or re.search(rn, x["n"], re.I)):
                ps = [base[str(c)] for c in refs if str(c) in base and base[str(c)].get("cap")]
                if not ps:
                    break
                ks = [p["k"] for p in ps]
                k = max(set(ks), key=ks.count)
                kgu = [p["kgu"] for p in ps if p.get("kgu")]
                caps = sorted(p["cap"] for p in ps)
                x["cx"] = {"k": VB_CX_MODELO.get(k, k if k.startswith("20") else "VB:" + k), "cap": caps[len(caps) // 2],
                           "kgu": median(kgu) if kgu else 0, "ref": orig + " · cód. " + ", ".join(str(c) for c in refs)}
                break
    sem = [x for x in dv["itens"] if not x["cx"] and x["nec"] > 0]
    if sem:
        print(f"  Demanda Verbum: {len(sem)} itens sem parâmetro de caixa: " + ", ".join(f"{x['c']} {x['n']}" for x in sem))
    return dv


def main():
    print("Estimativa de Vendas - gerando...")
    print(f"  Propostas: {PASTA_PROPOSTAS}")
    regs = escolher_propostas(carregar_propostas())
    print(f"  {len(regs)} propostas lidas | {sum(r['usada'] and bool(r['itens']) for r in regs)} entram na demanda (com itens)")
    est = carregar_estoque()
    zerar_negativos(est)
    descarte = aplicar_descarte(est)
    for cod, motivo in ITENS_SALDO_ZERO.items():   # tratado como descarte previsto (sai do saldo até a baixa no ERP)
        q = max(0.0, float(est["saldo"].get(cod, 0)))
        est["saldo"].loc[cod] = float(est["saldo"].get(cod, 0)) - q
        if q > 0:
            descarte["por"][int(cod)] = descarte["por"].get(int(cod), 0) + q
            descarte["un"] += q; descarte["itens"] = len(descarte["por"])
        print(f"  Saldo zerado na estimativa: cód. {cod} ({str(est['desc'].get(cod, '')).strip()}) - "
              f"{int(q):,} un. fora do saldo - {motivo}".replace(",", "."))
    est["dsc_por"] = descarte["por"]
    props, itens, linhas, pos = montar(regs, est)
    vendas = resumo_vendas(props)
    sem_erp = [i["c"] for i in itens if not i["erp"]]
    if sem_erp:
        print(f"  Aviso: {len(sem_erp)} códigos das propostas não existem na movimentação: {sem_erp[:15]}")
    ins = carregar_insumos()
    dados = {
        "versao": VERSAO_MODELO,
        "gerado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "base_estoque": "" if pd.isna(est["data_base"]) else f"{est['data_base']:%d/%m/%Y}",
        "pcts": pcts_crm(),
        "descarte": {k: v for k, v in descarte.items() if k != "por"},
        "corte": pos["corte"], "graficas": pos["graficas"], "saidas": pos["saidas"], "sem_proposta": pos["sem_proposta"],
        "escolas_vb": pos.get("esc_nomes", []),
        "vendas": vendas,
        "props": props, "itens": itens, "linhas": linhas,
        "insumos": ins,
        "ecommerce_pf": carregar_ecommerce_pf(),
        "embalagens": est.get("emb", []),
        "embalagens_compras": est.get("emb_compras", []),
        "embalagens_periodo": est.get("emb_per", []),
        "etiquetas": est.get("etq", [{"cod": c, "d": d, "mov": 0} for c, d in ETIQUETAS_COD.items()]),
        "etiquetas_periodo": est.get("etq_per", []),
        "demanda_verbum": caixas_demanda_verbum(carregar_demanda_verbum(est), ins),
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
