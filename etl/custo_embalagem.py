# custo_embalagem.py
# (12/11/2026) Custo de embalagem por NF para o Quadro de Envios (coluna AA "Custo Embalagem").
#
# Como calcula (só nas linhas com a coluna vazia — nunca mexe no que foi lançado à mão):
#   1. Volumes da NF = coluna Z do Quadro (vem do QTD_VOLUME do Relatório de Faturamento).
#   2. Envio pequeno (até 2 volumes), como já é lançado hoje: até 2 livros -> saco de e-commerce (1 por volume);
#      até 5 itens por volume -> caixa padrão (1 por volume; definido por Anderson em 12/11/2026). Sem itens na Movimentação: pessoa física -> saco.
#   3. Demais NFs: itens da nota na Movimentação de Estoque (saídas da mesma empresa e NF) -> caixas por modelo
#      pela base de insumos (estoque\insumos_base.json, montada dos romaneios 2025/2026):
#        - item com parâmetro próprio: unidades ÷ qtd padrão por caixa, no modelo do item;
#        - item que só sai em kit: unidades ÷ itens por caixa do kit do professor do segmento, no modelo do kit;
#        - troca de modelos (12/11/2026): 340 e 520 -> Modelo 250 nova; 530 -> 515.
#      O total de caixas é ajustado aos Volumes da NF, mantendo a proporção entre os modelos.
#      NF sem itens na Movimentação: todos os volumes na caixa 220.
#      Notas sem envio físico (ajustes, definição de preço, venda de aparas, e-book, devolução/retorno) ficam em branco.
#      Se os Volumes parecem ser unidades (>= 80% das unidades da nota ou mais de 3x o que o peso comporta a 14,5 kg/caixa),
#      usa as caixas estimadas / pelo peso.
#   4. Custo = caixas x preço da caixa + fita (metros por caixa do modelo x preço do metro) + etiquetas x preço.
#      Preços da aba "Custo Embalagem" do próprio Quadro (D1 = rolo de fita 100 m; D2 = rolo de 400 etiquetas;
#      coluna C = valor unitário de cada caixa). Kits do 1º ano: 2 etiquetas por caixa; demais 1.

import os
import re
import json
import math
import logging
import unicodedata

import pandas as pd

log = logging.getLogger(__name__)

COL_DATA, COL_DOC, COL_TIPO, COL_EMP, COL_NF, COL_NAT, COL_PESO, COL_VOL, COL_CUSTO = 1, 4, 11, 14, 15, 16, 20, 26, 27   # A D K N O P T Z AA
# notas sem envio físico (ajustes, definição de preço, venda de aparas) ficam sem custo de embalagem
RE_SEM_ENVIO = re.compile(r"AJUSTE|DEFINIC|DESCARTE|TRANSFERENCIA|EBOOK|E-BOOK|DEVOLUCAO|RETORNO|REQUISIC|ENTRADA|ESTORNO|COMPRA")
RE_SEM_ENVIO_QE = re.compile(RE_SEM_ENVIO.pattern + r"|PRODUTO PRODUZIDO")   # no Quadro, "produto produzido" = venda de aparas
KG_CAIXA = 14.5      # peso médio de uma caixa cheia (romaneios)
INICIO_PADRAO = "2025-10-27"     # início da Movimentação de Estoque (antes disso não há itens por NF)

# linha da aba "Custo Embalagem" (coluna A) -> chave do modelo na base de insumos
ABA_PARA_CHAVE = {"220": "236", "340": "237", "450": "239", "1": "240", "530": "240", "170": "547",
                  "AVULSO": "562", "PADRAO": "562", "520": "1391", "530 N": "2023", "515": "2023", "250": "250N",
                  "SACO E-COMMERCE": "SACO"}
RE_EMB = re.compile(r"CAIXA|FITA|FILME|STRETCH|EMBALAGEM|ETIQUETA|SACOLA")
LIM_PADRAO = 5       # caixa padrão leva no máximo 5 itens (Anderson, 12/11/2026)
REDIR = {"237": "250N", "1391": "250N", "240": "2023", "2022": "2024"}
VB_EQUIV = {"2025": "236", "2026": "547", "2027": "562", "2024": "2023", "2022": "2023"}   # preço da caixa Verbum = equivalente IAB


def _sa(t):
    return unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().upper().strip()


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _precos(wb):
    """Preços da aba 'Custo Embalagem' do Quadro (valores calculados; fórmulas simples são avaliadas)."""
    ws = wb["Custo Embalagem"]
    def val(c):
        v = ws[c].value
        if isinstance(v, str) and v.startswith("="):
            try:
                return float(eval(v[1:].replace(",", ".")))   # ex.: =0.0987*400
            except Exception:
                return None
        return _num(v)
    fita_rolo = val("D1") or 15.0
    etq_rolo = val("D2") or 39.48
    caixas = {}
    for r in range(5, 40):
        a, c = ws.cell(r, 1).value, _num(ws.cell(r, 3).value)
        if a is None or c is None:
            continue
        k = ABA_PARA_CHAVE.get(_sa(a).replace("CAIXA ", "").replace("MODELO ", "").strip())
        if k:
            caixas[k] = c
    return {"fita_m": fita_rolo / 100.0, "etq": etq_rolo / 400.0, "cx": caixas}


def _seg(desc):
    d = _sa(desc)
    if re.search(r"PRE[- ]?(ESCOLA )?II\b|PRE II", d): return "Pré II"
    if re.search(r"PRE[- ]?(ESCOLA )?I\b|PRE I\b", d): return "Pré I"
    m = re.search(r"(\d)\s*[º°O]?\s*ANO", d)
    return f"{m.group(1)}º ano" if m else None


def _itens_por_nf(arq_mov, inicio):
    df = pd.read_excel(arq_mov, usecols=["CODIGO_PRODUTO", "DATA_MOVIMENTO", "DESCRICAO_PRODUTO", "EMPRESA_ESTOQUE",
                                         "NATUREZA_TIPO_USO", "NUMERO_NF", "QUANTIDADE", "TIPO_MOVIMENTO"])
    df = df[df["TIPO_MOVIMENTO"].astype(str).str.upper().str.startswith("SA")
            & ~df["NATUREZA_TIPO_USO"].astype(str).map(_sa).str.contains(RE_SEM_ENVIO)].copy()
    df["emp"] = df["EMPRESA_ESTOQUE"].astype(str).str.upper().str.contains("VERBUM").map({True: "VERBUM", False: "IAB"})
    df["nf"] = pd.to_numeric(df["NUMERO_NF"], errors="coerce")
    df["cod"] = pd.to_numeric(df["CODIGO_PRODUTO"], errors="coerce")
    df["q"] = pd.to_numeric(df["QUANTIDADE"], errors="coerce").fillna(0)
    df = df[df["nf"].notna() & df["cod"].notna() & (df["q"] > 0)]
    g = df.groupby(["emp", "nf", "cod"]).agg(q=("q", "sum"), d=("DESCRICAO_PRODUTO", "first")).reset_index()
    out = {}
    for r in g.itertuples():
        out.setdefault((r.emp, int(r.nf)), []).append((int(r.cod), float(r.q), str(r.d)))
    ini = pd.to_datetime(df["DATA_MOVIMENTO"], errors="coerce").min()
    return out, (ini if pd.notna(ini) else pd.Timestamp(inicio))


def _caixas_estimadas(itens, base, emp):
    """{chave_modelo: caixas (fracionário), ...} e nº de caixas de kit do 1º ano (2 etiquetas)."""
    it, kits = base["itens"], base["kits"]
    est, kit_units = {}, {}
    for cod, q, d in itens:
        if RE_EMB.search(_sa(d)):
            continue
        p = it.get(str(cod))
        if p and p.get("cap"):
            k = REDIR.get(p["k"], p["k"])
            est[k] = est.get(k, 0) + q / p["cap"]
        else:
            s = _seg(d) or "?"
            kit_units[s] = kit_units.get(s, 0) + q
    kit1 = 0.0
    for s, u in kit_units.items():
        kp = kits.get(f"{s} – Kit Professor") or {"k": "562", "n": 14}
        n = kp.get("n") or 14
        k = REDIR.get(kp["k"], kp["k"])
        cx = u / n
        est[k] = est.get(k, 0) + cx
        if s == "1º ano":
            kit1 += cx
    if emp == "VERBUM":
        vb = {"236": "2025", "547": "2026", "562": "2027", "2023": "2024"}
        est = {vb.get(k, k): v for k, v in est.items()}
    return est, kit1


def _distribuir(est, vol):
    """Ajusta as caixas estimadas ao total de volumes (maiores restos)."""
    tot = sum(est.values())
    if not vol or vol <= 0:
        return {k: math.ceil(v - 1e-9) for k, v in est.items() if v > 0}
    if tot <= 0:
        return {"236": int(vol)}
    alvo = {k: v / tot * vol for k, v in est.items()}
    base = {k: int(math.floor(v)) for k, v in alvo.items()}
    resto = int(vol) - sum(base.values())
    for k, _ in sorted(alvo.items(), key=lambda kv: -(kv[1] - math.floor(kv[1])))[:max(0, resto)]:
        base[k] += 1
    return {k: v for k, v in base.items() if v > 0}


def custo_nf(itens, vol, pf, emp, base, precos, peso=None):
    """Retorna (custo, detalhe) de uma NF."""
    mods, pc = base["mods"], precos["cx"]
    vol = int(vol) if vol and vol > 0 else 0
    # volumes que parecem unidades (lançamento errado): mais de 3x as caixas que o peso da nota comporta
    if peso and peso > 0 and vol > 20 and vol > 3 * math.ceil(peso / KG_CAIXA):
        vol = math.ceil(peso / KG_CAIXA)
    un = sum(q for _, q, d in (itens or []) if not RE_EMB.search(_sa(d))) if itens else None
    # envio pequeno (até 2 volumes), como é feito hoje: até 2 livros -> saco de e-commerce; até 5 por volume -> caixa padrão
    if 1 <= vol <= 2 and (un is None or un <= LIM_PADRAO * vol):
        if (un is not None and un <= 2) or (un is None and pf):
            return round(vol * (pc.get("SACO", 0.66) + precos["etq"]), 2), f"{vol} saco(s) e-commerce"
        k = "2027" if emp == "VERBUM" else "562"
        m = mods.get(k, {})
        c = vol * (pc.get("562", 2.10) + (m.get("f") or 0.861) * precos["fita_m"] + precos["etq"])
        return round(c, 2), f"{vol} caixa(s) padrão"
    if itens:
        est, kit1 = _caixas_estimadas(itens, base, emp)
        if un and vol > 20 and vol >= 0.8 * un:    # volumes = unidades (lançamento errado): usa as caixas estimadas
            vol = 0
        cx = _distribuir(est, vol)
    else:
        cx, kit1 = ({"236": vol} if vol else {}), 0
    if not cx:
        return None, "sem volumes e sem itens"
    total, partes = 0.0, []
    totcx = sum(cx.values())
    for k, n in cx.items():
        m = mods.get(k, {})
        preco = pc.get(k, pc.get(VB_EQUIV.get(k, ""), pc.get("236", 4.35)))
        if k == "250N" and "250N" not in pc:
            preco = pc.get("1391", preco)          # 250 nova ainda sem preço na aba: usa o da 520
        fita = (m.get("f") or 1.386) * precos["fita_m"]
        total += n * (preco + fita + precos["etq"])
        partes.append(f"{n} cx {m.get('m', k)}")
    if kit1 and totcx:
        total += min(kit1, totcx) * precos["etq"]   # 2ª etiqueta dos kits do 1º ano
    return round(total, 2), " + ".join(partes)


def preencher_custo_embalagem(quadro, arq_mov, arq_base_insumos, salvar=True, inicio=INICIO_PADRAO):
    """Preenche a coluna 'Custo Embalagem' das NFs vazias do Quadro. Retorna nº de linhas preenchidas."""
    from openpyxl import load_workbook
    if not os.path.exists(arq_base_insumos):
        log.warning(f"Custo de embalagem: base de insumos não encontrada ({arq_base_insumos}) - etapa ignorada")
        return 0
    base = json.load(open(arq_base_insumos, encoding="utf-8"))
    itens_nf, ini = ({}, pd.Timestamp(inicio))
    if os.path.exists(arq_mov):
        itens_nf, ini = _itens_por_nf(arq_mov, inicio)
    else:
        log.warning(f"Custo de embalagem: Movimentação não encontrada ({arq_mov}) - uso só os volumes")
    wb = load_workbook(quadro)
    precos = _precos(wb)
    ws = wb["Lançamentos"]
    if str(ws.cell(1, COL_CUSTO).value or "").strip().upper() != "CUSTO EMBALAGEM":
        log.warning("Custo de embalagem: coluna AA não é 'Custo Embalagem' - etapa ignorada")
        return 0
    n, soma = 0, 0.0
    for r in range(2, ws.max_row + 1):
        nf = _num(ws.cell(r, COL_NF).value)
        if not nf or ws.cell(r, COL_CUSTO).value not in (None, ""):
            continue
        dt = ws.cell(r, COL_DATA).value
        try:
            if dt is None or pd.Timestamp(dt) < ini:
                continue
        except Exception:
            continue
        if RE_SEM_ENVIO_QE.search(_sa(ws.cell(r, COL_NAT).value)):   # _sa tira acentos
            continue
        emp = "VERBUM" if "VERBUM" in _sa(ws.cell(r, COL_EMP).value) else "IAB"
        doc = re.sub(r"\D", "", str(ws.cell(r, COL_DOC).value or ""))
        vol = _num(ws.cell(r, COL_VOL).value) or 0
        peso = _num(ws.cell(r, COL_PESO).value)
        c, det = custo_nf(itens_nf.get((emp, int(nf))), vol, len(doc) == 11, emp, base, precos, peso)
        if c is None:
            continue
        ws.cell(r, COL_CUSTO, c)
        n += 1; soma += c
    if n and salvar:
        wb.save(quadro)
    wb.close()
    log.info(f"Custo de embalagem calculado em {n} NF(s) sem valor (total R$ {soma:,.2f})".replace(",", "X").replace(".", ",").replace("X", "."))
    return n
