# VERSAO: 2.0-VE — Com Visão Executiva e pós-processamento
import sys, os
# Limpar cache Python para garantir versão mais recente
_this_dir = os.path.dirname(os.path.abspath(__file__))
_cache_dir = os.path.join(_this_dir, '__pycache__')
if os.path.exists(_cache_dir):
    import shutil as _shutil
    _shutil.rmtree(_cache_dir, ignore_errors=True)

# =============================================================================
# gerar_dashboard.py — Atualiza o dashboard.html com os dados mais recentes
# Executado automaticamente pelo ATUALIZAR.bat após o ETL
# =============================================================================

import sys, os, json, math, base64, io, shutil
import pandas as pd
import numpy as np
from datetime import datetime
from PIL import Image

# ── CAMINHOS ──
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DASH_DIR    = os.path.dirname(BASE_DIR)
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

PASTA_DASHBOARDS = _encontrar_onedrive_dashboards()          # pasta pai do etl/
PARQUET     = os.path.join(BASE_DIR, "data", "processed", "lancamentos.parquet")
DASH_OUT    = os.path.join(PASTA_DASHBOARDS, "dashboard_logistico", "dashboard.html")

# Imagens (na mesma pasta que o dashboard.html)
# Pasta onde estão as imagens (mesma pasta do dashboard.html)
DASH_DIR    = os.path.dirname(os.path.abspath(DASH_OUT)) if DASH_OUT else BASE_DIR
IMG_DIR     = DASH_DIR
_img_candidatos_extra = [
    DASH_DIR,                               # dashboard_logistico/  ← imagens aqui
    os.path.dirname(DASH_DIR),             # pai de dashboard_logistico/ = 10-DASHIBOARDs/
    PASTA_DASHBOARDS,                       # 10 - DASHIBOARDs/
    os.path.join(BASE_DIR, ".."),           # etl/../ = dashboard_logistico/
    os.path.join(BASE_DIR, "..", ".."),     # etl/../../ = 10-DASHIBOARDs/
    os.path.join(BASE_DIR, "..", "..", "..", "dashboard_logistico"),  # extra
]
IMGS = {
    "LI": (os.path.join(IMG_DIR, "logo_iab.png"),       260, 85),
    "LV": (os.path.join(IMG_DIR, "logo_verbum.png"),    220, 85),
    "CI": (os.path.join(IMG_DIR, "Carreta_iab.png"),    380, 78),
    "CV": (os.path.join(IMG_DIR, "Carreta_verbum.png"), 380, 78),
}

# ── HELPERS ──
def webp(path, max_w, q=78):
    img = Image.open(path).convert("RGBA")
    w, h = img.size
    if w > max_w:
        img = img.resize((max_w, int(h * max_w / w)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="WEBP", quality=q, method=6)
    return "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()

def sf(v, default="—"):
    try:
        if v is None or (isinstance(v, float) and math.isnan(v)): return default
        s = str(v).strip()
        return s if s else default
    except: return default

def fmt_d(v):
    try: return v.strftime("%d/%m/%Y") if not pd.isna(v) else "—"
    except: return "—"


def gerar():
    print("=" * 55)
    print("  GERANDO DASHBOARD ATUALIZADO")
    print("=" * 55)

    # ── 1. CARREGAR DADOS ──
    print("[ 1 ] Carregando dados processados...")
    if not os.path.exists(PARQUET):
        print("ERRO: lancamentos.parquet não encontrado. Rode o ETL primeiro.")
        sys.exit(1)

    df = pd.read_parquet(PARQUET)
    hoje = pd.Timestamp.now().normalize()
    print(f"      {len(df)} registros carregados.")

    # ── 2. IMAGENS ──
    print("[ 2 ] Processando imagens...")
    imgs = {}
    for key, (path, max_w, q) in IMGS.items():
        # Tentar múltiplos caminhos
        _nome = os.path.basename(path)
        candidatos = [path] + [
            os.path.join(d, _nome)
            for d in _img_candidatos_extra if d
        ]
        # Candidatos adicionais baseados no DASH_OUT
        if DASH_OUT:
            _dash_dir = os.path.dirname(os.path.abspath(DASH_OUT))
            candidatos += [
                os.path.join(_dash_dir, _nome),
                os.path.join(os.path.dirname(_dash_dir), _nome),
            ]
        encontrado = None
        for c in candidatos:
            if os.path.exists(c):
                encontrado = c
                break
        if not encontrado:
            print(f"      AVISO: {os.path.basename(path)} não encontrado.")
            print(f"      Caminhos tentados:")
            for c in candidatos:
                print(f"        - {c}")
            # Usar caminho relativo como fallback para quando aberto sem BAT
            _nomes_rel = {
                "LI": "logo_iab.png",
                "LV": "logo_verbum.png",
                "CI": "Carreta_iab.png",
                "CV": "Carreta_verbum.png",
            }
            imgs[key] = _nomes_rel.get(key, "")
            print(f"      Usando caminho relativo: {imgs[key]}")
            continue
        imgs[key] = webp(encontrado, max_w, q)
    LI = imgs.get("LI",""); LV = imgs.get("LV","")
    CI = imgs.get("CI",""); CV = imgs.get("CV","")

    # ── 3. PAYLOAD COMPRIMIDO (Visão Geral) ──
    print("[ 3 ] Gerando dados da Visão Geral...")
    nfs_map = df.groupby("id_entrega")["nf"].apply(
        lambda x: " / ".join(str(int(v)) for v in sorted(x.dropna().unique()))
    ).to_dict()

    df_pri = df[df["is_primeira_linha_entrega"] == 1].copy()
    df_pri["obs"] = df_pri["obs"].fillna("—")

    # Lookup comprimido
    emps  = sorted(df_pri["empresa"].dropna().unique().tolist())
    anos  = sorted(df_pri["ano_ref"].dropna().unique().tolist())
    tcs   = sorted(df_pri["tipo_cliente"].dropna().unique().tolist())
    nats  = sorted(df_pri["natureza_operacao"].dropna().unique().tolist())
    trs   = sorted(df_pri["transportadora_norm"].dropna().unique().tolist())
    emp_i = {v:i for i,v in enumerate(emps)}
    ano_i = {v:i for i,v in enumerate(anos)}
    tc_i  = {v:i for i,v in enumerate(tcs)}
    nat_i = {v:i for i,v in enumerate(nats)}
    tr_i  = {v:i for i,v in enumerate(trs)}

    rows_comp = []
    for _, r in df.iterrows():
        ei = emp_i.get(r["empresa"], 0)
        ai = ano_i.get(str(r["ano_ref"]), 0)
        ti = tc_i.get(r["tipo_cliente"], 0)
        ni = nat_i.get(r["natureza_operacao"], 0)
        ri = tr_i.get(r["transportadora_norm"], 0)
        rows_comp.append([
            ei, ai, ti, ni, ri,
            round(float(r["vr_nf"]), 2),
            round(float(r["frete"]), 2),
            round(float(r["custo_embalagem"]), 2),
            round(float(r["total_impostos_calc"]), 2),
            int(r["is_primeira_linha_entrega"]),
            int(r["entrega_no_prazo"]),
            int(r["entrega_em_atraso"]),
        ])

    PAYLOAD_GER = json.dumps(
        {"E":emps,"A":anos,"T":tcs,"N":nats,"R":trs,"D":rows_comp},
        ensure_ascii=False, separators=(',',':')
    )

    # ── 4. PAYLOAD ENTREGAS ──
    print("[ 4 ] Gerando dados da aba Entregas...")

    da_list, dv_list = [], []
    for _, row in df_pri.iterrows():
        prev = row["prev_entrega_cliente"]
        efet = row["data_efetiva_entrega"]
        try:
            if pd.isna(prev): da_list.append(0); dv_list.append(None); continue
            ref = efet if not pd.isna(efet) else hoje
            da_list.append(max(0, int((ref - prev).days)))
            dv_list.append(int((prev - hoje).days) if pd.isna(efet) else None)
        except: da_list.append(0); dv_list.append(None)
    df_pri["_da"] = da_list
    df_pri["_dv"] = dv_list
    df_pri["nfs_ag"] = df_pri["id_entrega"].map(nfs_map)

    def rdict(row, c2=False):
        nf_str = sf(row.get("nfs_ag", ""))
        if nf_str == "—":
            nf_val = row["nf"]
            nf_str = str(int(nf_val)) if not pd.isna(nf_val) else "—"
        d = {
            "emp": sf(row["empresa"]), "nf": nf_str,
            "cli": sf(row["cliente"]), "tr": sf(row["transportadora_norm"]),
            "mun": sf(row["municipio"]), "uf": sf(row["uf"]),
            "reg": sf(row["regiao"]), "col": fmt_d(row["data_coleta"]),
            "prev": fmt_d(row["prev_entrega_cliente"]),
            "da": int(row["_da"]),
            "dv": int(row["_dv"]) if (row["_dv"] is not None and not (isinstance(row["_dv"], float) and math.isnan(row["_dv"]))) else None,
            "obs": sf(row["obs"]), "st": sf(row["status_prazo"]),
        }
        if c2: d["ef"] = fmt_d(row["data_efetiva_entrega"])
        return d

    df26 = df_pri[df_pri["ano_ref"] == "2026"].copy()
    q1 = df26[df26["data_efetiva_entrega"].isna()].sort_values("prev_entrega_cliente")
    q2 = df26[df26["entrega_em_atraso"] == 1].sort_values("data_efetiva_entrega", ascending=False)

    def agg(d):
        def g(col):
            return [{"k": str(r[col]), "total": int(r["total"]),
                     "prazo": int(r["prazo"]), "atraso": int(r["atraso"])}
                    for _, r in d.groupby(col).agg(
                        total=("entrega_no_prazo","count"),
                        prazo=("entrega_no_prazo","sum"),
                        atraso=("entrega_em_atraso","sum")
                    ).reset_index().iterrows()]
        mes_rows = []
        for _, r in d.groupby(["ano","mes"]).agg(
            total=("entrega_no_prazo","count"),
            prazo=("entrega_no_prazo","sum"),
            atraso=("entrega_em_atraso","sum")
        ).reset_index().sort_values(["ano","mes"]).iterrows():
            mes_rows.append({"ano":int(r["ano"]),"mes":int(r["mes"]),
                             "total":int(r["total"]),"prazo":int(r["prazo"]),"atraso":int(r["atraso"])})
        return {"total":int(d.shape[0]),"prazo":int(d["entrega_no_prazo"].sum()),
                "atraso":int(d["entrega_em_atraso"].sum()),
                "reg":g("regiao"),"tc":g("tipo_cliente"),
                "tr":g("transportadora_norm"),"mes":mes_rows}

    G = {
        "TODAS": {"TODOS": agg(df_pri)},
        "IAB": {
            "TODOS": agg(df_pri[df_pri["empresa"]=="IAB"]),
            "2025":  agg(df_pri[(df_pri["empresa"]=="IAB")&(df_pri["ano_ref"]=="2025")]),
            "2026":  agg(df_pri[(df_pri["empresa"]=="IAB")&(df_pri["ano_ref"]=="2026")]),
        },
        "VERBUM": {
            "TODOS": agg(df_pri[df_pri["empresa"]=="VERBUM"]),
            "2025":  agg(df_pri[(df_pri["empresa"]=="VERBUM")&(df_pri["ano_ref"]=="2025")]),
            "2026":  agg(df_pri[(df_pri["empresa"]=="VERBUM")&(df_pri["ano_ref"]=="2026")]),
        },
    }

    PAYLOAD_ENT = json.dumps(
        {"hoje": hoje.strftime("%Y-%m-%d"),
         "q1": [rdict(r) for _, r in q1.iterrows()],
         "q2": [rdict(r, c2=True) for _, r in q2.iterrows()],
         "G": G},
        ensure_ascii=False, separators=(',',':')
    )

    anos_ref = sorted(df_pri["ano_ref"].dropna().unique().tolist())
    print(f"      Q1: {len(q1)} pendentes | Q2: {len(q2)} atrasos")

    # ── 5. LER TEMPLATE DO DASHBOARD ──
    print("[ 5 ] Lendo template do dashboard...")
    template_path = os.path.join(BASE_DIR, "dashboard_template.html")
    if not os.path.exists(template_path):
        print("ERRO: dashboard_template.html não encontrado na pasta etl/")
        sys.exit(1)

    with open(template_path, encoding="utf-8") as f:
        html = f.read()

    # ── 6. INJETAR DADOS ──
    print("[ 6 ] Injetando dados no template...")

    # Botões de ano dinâmicos
    ano_btns = "".join([
        f'<button class="ano-btn" id="ab{a}" onclick="fAno(\'{a}\')">{a}</button>'
        for a in anos_ref
    ])
    eano_btns = "".join([
        f'<button class="ano-btn" id="eab{a}" onclick="eAno(\'{a}\')">{a}</button>'
        for a in anos_ref
    ])


    # ── Payload Visão Executiva ──
    import numpy as _np

    def _kex(dp_x, d_x):
        if len(dp_x)==0: return None
        total=len(dp_x); prazo=int(dp_x["entrega_no_prazo"].sum()); atraso=int(dp_x["entrega_em_atraso"].sum())
        ns=round(prazo/total*100,1)
        lt_v=dp_x["lead_time"].dropna() if "lead_time" in dp_x.columns else pd.Series(dtype=float)
        lt=round(float(lt_v.mean()),1) if len(lt_v) else 0
        frete=float(d_x["frete"].sum())
        fat_b=float(d_x[d_x["natureza_operacao"]=="VENDA"]["vr_nf"].sum())
        dev=float(d_x[d_x["natureza_operacao"]=="DEVOLUÇÃO"]["vr_nf"].sum())
        bon=float(d_x[d_x["natureza_operacao"]=="BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l=fat_b-dev-bon
        peso=float(d_x["peso_kg"].sum()) if "peso_kg" in d_x.columns else 0
        c_kg=round(frete/peso,4) if peso else 0
        return {"total":total,"prazo":prazo,"atraso":atraso,"ns":ns,
                "taxa_at":round(atraso/total*100,1),"lt":lt,"lt_med":lt,
                "fat_b":round(fat_b),"fat_l":round(fat_l),"dev":round(dev),"bon":round(bon),
                "frete":round(frete),"fsf":round(frete/fat_l*100,2) if fat_l else 0,
                "pct_dev":round(dev/fat_b*100,1) if fat_b else 0,
                "c_ent":round(frete/total,2) if total else 0,
                "peso":round(peso,1),"c_kg":c_kg}

    df_pri_ve = df[df["is_primeira_linha_entrega"]==1].copy()
    if "lead_time" not in df_pri_ve.columns:
        df_pri_ve["lead_time"] = (df_pri_ve["data_efetiva_entrega"] - df_pri_ve["data_coleta"]).dt.days.clip(lower=0)
    df["regiao"] = df["regiao"].fillna("").str.strip().str.upper()
    df_pri_ve["regiao"] = df_pri_ve["regiao"].fillna("").str.strip().str.upper()

    _VE = {}
    for _emp in ["TODAS","IAB","VERBUM"]:
        _VE[_emp] = {}
        _de  = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        _dpe = df_pri_ve if _emp=="TODAS" else df_pri_ve[df_pri_ve["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da  = _de  if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _dpa = _dpe if _ano=="TODOS" else _dpe[_dpe["ano_ref"].astype(str)==_ano]
            _regs=[]; _transp=[]; _mes=[]
            for _r,_g in _dpa.groupby("regiao"):
                if not _r or _r=="NAN": continue
                _k=_kex(_g,_da[_da["regiao"]==_r])
                if _k: _k["k"]=str(_r); _regs.append(_k)
            for _t,_g in _dpa.groupby("transportadora_norm"):
                _k=_kex(_g,_da[_da["transportadora_norm"]==_t])
                if _k: _k["k"]=str(_t); _transp.append(_k)
            for (_a,_m),_g in _dpa.groupby(["ano","mes"]):
                _k=_kex(_g,_da[(_da["ano"]==_a)&(_da["mes"]==_m)])
                if _k: _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            _VE[_emp][_ano]={
                "total":_kex(_dpa,_da),
                "regiao":sorted(_regs,key=lambda x:-x["total"]),
                "transp":sorted(_transp,key=lambda x:-x["total"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"]))}

    class _NpEnc(json.JSONEncoder):
        def default(self,o):
            if isinstance(o,_np.integer): return int(o)
            if isinstance(o,_np.floating): return float(o)
            return super().default(o)

    PAYLOAD_VE  = json.dumps(_VE, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))
    MESES_VE_JS = json.dumps(["","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"], ensure_ascii=False)

    # ── Payload Financeiro ──
    def _kfin(d_x):
        fat_b=float(d_x[d_x["natureza_operacao"]=="VENDA"]["vr_nf"].sum())
        dev=float(d_x[d_x["natureza_operacao"]=="DEVOLUÇÃO"]["vr_nf"].sum())
        bon=float(d_x[d_x["natureza_operacao"]=="BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l=fat_b-dev-bon
        frete=float(d_x["frete"].sum())
        emb=float(d_x["custo_embalagem"].sum())
        imp=float(d_x["total_impostos_calc"].sum())
        custo=frete+emb
        margem_op=fat_l-custo-imp
        ents=len(d_x[d_x["is_primeira_linha_entrega"]==1])
        ticket_med=round(fat_l/ents) if ents else 0
        custo_med=round(custo/ents) if ents else 0
        return {"fat_b":round(fat_b),"dev":round(dev),"bon":round(bon),
                "fat_l":round(fat_l),"frete":round(frete),"emb":round(emb),"imp":round(imp),
                "custo":round(custo),"margem_op":round(margem_op),"ents":ents,
                "ticket_med":ticket_med,"custo_med":custo_med,
                "pct_dev":round(dev/fat_b*100,1) if fat_b else 0,
                "pct_bon":round(bon/fat_b*100,1) if fat_b else 0,
                "pct_frete":round(frete/fat_l*100,2) if fat_l else 0,
                "pct_custo":round(custo/fat_l*100,2) if fat_l else 0,
                "pct_imp":round(imp/fat_l*100,2) if fat_l else 0,
                "pct_margem":round(margem_op/fat_l*100,1) if fat_l else 0}

    _FIN = {}
    for _emp in ["TODAS","IAB","VERBUM"]:
        _FIN[_emp] = {}
        _de = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da = _de if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _regs=[]; _trs=[]; _mes=[]; _clis=[]
            for _r,_g in _da.groupby("regiao"):
                if not _r: continue
                _k=_kfin(_g); _k["k"]=str(_r); _regs.append(_k)
            for _t,_g in _da.groupby("transportadora_norm"):
                _k=_kfin(_g); _k["k"]=str(_t); _trs.append(_k)
            for (_a,_m),_g in _da.groupby(["ano","mes"]):
                _k=_kfin(_g); _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            for _c,_g in _da.groupby("cliente"):
                _k=_kfin(_g); _k["k"]=str(_c)
                # Adicionar cidade e UF do cliente
                _row1 = _g.iloc[0] if len(_g) > 0 else None
                if _row1 is not None:
                    _k["mun"] = str(_row1.get("municipio","")) if "municipio" in _g.columns else ""
                    _k["uf"]  = str(_row1.get("uf","")) if "uf" in _g.columns else ""
                _clis.append(_k)
            _FIN[_emp][_ano]={
                "total":_kfin(_da),
                "regiao":sorted(_regs,key=lambda x:-x["fat_l"]),
                "transp":sorted(_trs,key=lambda x:-x["frete"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"])),
                "clientes":sorted(_clis,key=lambda x:-x["fat_l"])[:20]}

    PAYLOAD_FIN = json.dumps(_FIN, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))

    # ── Payload Transportadoras ──
    _TR = {}
    df_pri_tr = df[df["is_primeira_linha_entrega"]==1].copy()
    if "lead_time" not in df_pri_tr.columns:
        df_pri_tr["lead_time"] = (df_pri_tr["data_efetiva_entrega"] - df_pri_tr["data_coleta"]).dt.days.clip(lower=0)
    for _emp in ["TODAS","IAB","VERBUM"]:
        _TR[_emp] = {}
        _de  = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        _dpe = df_pri_tr if _emp=="TODAS" else df_pri_tr[df_pri_tr["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da  = _de  if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _dpa = _dpe if _ano=="TODOS" else _dpe[_dpe["ano_ref"].astype(str)==_ano]
            _trs=[]; _ufs=[]; _mes=[]
            for _t,_g in _dpa.groupby("transportadora_norm"):
                _k=_kex(_g,_da[_da["transportadora_norm"]==_t])
                if _k: _k["k"]=str(_t); _trs.append(_k)
            for _u,_g in _dpa.groupby("uf"):
                _k=_kex(_g,_da[_da["uf"]==_u])
                if _k: _k["k"]=str(_u); _ufs.append(_k)
            for (_a,_m),_g in _dpa.groupby(["ano","mes"]):
                _k=_kex(_g,_da[(_da["ano"]==_a)&(_da["mes"]==_m)])
                if _k: _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            # Agregar por região para TR
            _regs_tr=[]
            _dpa_r = _dpa.copy()
            _dpa_r["regiao"] = _dpa_r["regiao"].fillna("").str.strip().str.upper()
            for _r,_g in _dpa_r.groupby("regiao"):
                if not _r or _r=="NAN": continue
                _k=_kex(_g,_da[_da["regiao"]==_r])
                if _k: _k["k"]=str(_r); _regs_tr.append(_k)
            _TR[_emp][_ano]={
                "total":_kex(_dpa,_da),
                "transp":sorted(_trs,key=lambda x:-x["total"]),
                "uf":sorted(_ufs,key=lambda x:-x["total"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"])),
                "regiao":sorted(_regs_tr,key=lambda x:-x["total"])}

    PAYLOAD_TR = json.dumps(_TR, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))

    # ── Substituições (com fallback para placeholders não definidos) ──
    def _safe_replace(h, placeholder, value):
        try:
            return h.replace(placeholder, value)
        except Exception as e:
            print(f"AVISO: Erro ao substituir {placeholder}: {e}")
            return h

    html = _safe_replace(html, "___LI___", LI)
    html = _safe_replace(html, "___LV___", LV)
    html = _safe_replace(html, "___CI___", CI)
    html = _safe_replace(html, "___CV___", CV)
    html = _safe_replace(html, "___PAYLOAD_GER___", PAYLOAD_GER)
    html = _safe_replace(html, "___PAYLOAD_VE___", PAYLOAD_VE)
    html = _safe_replace(html, "___MESES_VE___", MESES_VE_JS)
    html = _safe_replace(html, "___PAYLOAD_ENT___", PAYLOAD_ENT)
    html = _safe_replace(html, "___ANO_BTNS___", ano_btns)
    html = _safe_replace(html, "___EANO_BTNS___", eano_btns)
    html = _safe_replace(html, "___GERADO_EM___", datetime.now().strftime("%d/%m/%Y %H:%M"))
    # Remover quaisquer placeholders restantes para evitar erros JS
    import re as _re3; html = _re3.sub(r'___[A-Z_]+___', 'null', html)


    # ── 7. VALIDAR E SALVAR ──
    print("[ 7 ] Validando e salvando dashboard.html...")
    import re
    placeholders_restantes = re.findall(r'___[A-Z_]+___', html)
    if placeholders_restantes:
        print(f"ERRO CRÍTICO: Placeholders não substituídos: {set(placeholders_restantes)}")
        print("Verifique se o dashboard_template.html esta na pasta etl/")
        sys.exit(1)

    # ══════════════════════════════════════════════
    # PÓS-PROCESSAMENTO COMPLETO
    # ══════════════════════════════════════════════
    import re as _re

    # 1. CSS da Visão Executiva
    _css_ve = '<style>\n.ve-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}\n.ve-card{background:var(--white);border-radius:var(--r);padding:16px 18px;box-shadow:var(--shadow);border-left:5px solid transparent;cursor:pointer;transition:box-shadow .15s;display:flex;flex-direction:column}\n.ve-card:hover{box-shadow:0 4px 20px rgba(0,0,0,.12)}\n.ve-card.ok{border-color:var(--ok)}.ve-card.warn{border-color:var(--alert)}.ve-card.bad{border-color:var(--err)}.ve-card.neu{border-color:var(--ac)}\n.ve-card .vc-ico{font-size:16px;margin-bottom:4px}\n.ve-card .vc-lbl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px}\n.ve-card .vc-val{font-size:26px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:8px}\n.ve-card .vc-badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:700;margin-bottom:6px}\n.ve-card .vc-yoy{font-size:11px;font-weight:700;margin-bottom:6px;min-height:18px}\n.ve-card .vc-drill{font-size:9px;color:var(--gl);cursor:pointer;margin-top:auto}\n.ve-alerta{border-radius:var(--r);padding:14px 16px}\n.ve-alerta.crit{background:#FFF5F5;border:1.5px solid #FFCDD2}\n.ve-alerta.aviso{background:#FFFDE7;border:1.5px solid #FFF176}\n.ve-alerta-tit{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}\n.ve-alerta-item{font-size:12px;padding:5px 0;border-bottom:1px solid rgba(0,0,0,.06);display:flex;justify-content:space-between;align-items:center}\n.ve-alerta-item:last-child{border:none}\n.ve-alerta-vazio{font-size:12px;color:var(--gl);font-style:italic;text-align:center;padding:10px 0}\n.tc{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow);margin-bottom:14px}\n</style>'
    html = html.replace('</head>', _css_ve + '\n</head>')
    # Cache-busting: adicionar timestamp no titulo
    _ts = datetime.now().strftime('%d/%m %H:%M')
    html = html.replace('<title>Dashboard Logístico — IAB & Verbum</title>',
                       f'<title>Dashboard Logístico — IAB & Verbum ({_ts})</title>')

    # 2. Substituir menu
    html = html.replace(
        "onclick=\"navPage('overview');fMenu()\">📊 Visão Geral",
        "onclick=\"navPage('visao-executiva');fMenu()\">📊 Visão Executiva"
    )
    html = html.replace('id="ni-overview"', 'id="ni-visao-executiva"')
    html = _re.sub(r'\s*<a class="ni"[^>]+ni-ent[^>]*>📦 Entregas</a>', '', html)
    html = html.replace("let PAGINA = 'overview'", "let PAGINA = 'visao-executiva'")
    # Corrigir onclick dos itens do menu que só têm fMenu()
    html = html.replace(
        'id="ni-financeiro" onclick="fMenu()">💰 Financeiro',
        'id="ni-financeiro" onclick="navPage(\'financeiro\');fMenu()">💰 Financeiro'
    )
    html = html.replace(
        'id="ni-transportadoras" onclick="fMenu()">🚛 Transportadoras',
        'id="ni-transportadoras" onclick="navPage(\'transportadoras\');fMenu()">🚛 Transportadoras'
    )
    html = html.replace(
        'id="ni-acomp"',
        'id="ni-acomp" onclick="navPage(\'acompanhamento\');fMenu()"'
    ).replace(
        'onclick="navPage(\'acompanhamento\');fMenu()" onclick="navPage(\'acompanhamento\');fMenu()"',
        'onclick="navPage(\'acompanhamento\');fMenu()"'
    )
    # Reordenar menu: VE → Financeiro → Transportadoras → ↳Acompanhamento → Estoque
    html = _re.sub(
        r'(<a class="ni[^"]*" id="ni-acomp"[^>]*>↳ Acompanhamento</a>)'
        r'(\s*<a class="ni[^"]*" id="ni-financeiro"[^>]*>💰 Financeiro</a>)'
        r'(\s*<a class="ni[^"]*" id="ni-transportadoras"[^>]*>🚛 Transportadoras</a>)',
        r'\2\3\n    \1\n    <a class="ni" id="ni-estoque" onclick="abrirEstoque();fMenu()">📦 Estoque</a>',
        html
    )
    # Remover seção "Filtrar por Empresa"
    html = _re.sub(
        r'\s*<div class="ns">Filtrar por Empresa</div>.*?<a[^>]*>🏢 Todas as Empresas</a>',
        '', html, flags=_re.DOTALL
    )

    # 2a. Injetar CSS completo (financeiro, transportadoras, etc.)
    _css_bka = "\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\n:root{\n  --iab:#0056A8;--iab-dark:#003d7a;\n  --verbum:#6E1184;--verbum-dark:#4a0a5a;\n  --alert:#F39200;\n  --bg:#F5F5F5;--white:#fff;\n  --gd:#333333;--gm:#666666;--gl:#999999;--gl2:#E0E0E0;\n  --ok:#107C41;--err:#D13438;\n  --shadow:0 2px 12px rgba(0,0,0,.08);\n  --r:12px;--r2:8px;\n  --font:'Montserrat',sans-serif;\n  --ac:#0056A8;--ac-rgb:0,86,168;\n}\nbody{font-family:var(--font);background:var(--bg);color:var(--gd);min-height:100vh;overflow-x:hidden}\n#ov{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:200;opacity:0;pointer-events:none;transition:opacity .3s}\n#ov.open{opacity:1;pointer-events:all}\n#sb{position:fixed;left:-280px;top:0;bottom:0;width:280px;background:#1E2A3A;z-index:201;transition:left .3s cubic-bezier(.4,0,.2,1);box-shadow:4px 0 24px rgba(0,0,0,.15);display:flex;flex-direction:column}\n#sb.open{left:0}\n.sbh{padding:20px;border-bottom:1px solid rgba(255,255,255,.1);display:flex;align-items:center;justify-content:space-between}\n.sbh img{height:30px;object-fit:contain}\n.sbx{background:none;border:none;cursor:pointer;font-size:20px;color:rgba(255,255,255,.6);padding:4px;line-height:1}\n.sbn{flex:1;padding:8px 0;overflow-y:auto}\n.ni{display:flex;align-items:center;gap:11px;padding:12px 22px;cursor:pointer;color:rgba(255,255,255,.75);font-size:13px;font-weight:500;transition:all .18s;border-left:3px solid transparent;text-decoration:none;user-select:none}\n.ni:hover,.ni.on{background:rgba(255,255,255,.08);color:#fff;border-left-color:var(--ac);font-weight:600}\n.ns{padding:14px 22px 4px;font-size:10px;font-weight:700;color:rgba(255,255,255,.35);letter-spacing:1.5px;text-transform:uppercase}\n#hd{background:var(--white);padding:0 20px;height:60px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.07);position:sticky;top:0;z-index:100}\n.hl{display:flex;align-items:center;gap:14px}\n.mb{background:none;border:none;cursor:pointer;padding:7px;border-radius:7px;display:flex;flex-direction:column;gap:5px;transition:background .2s}\n.mb:hover{background:var(--bg)}\n.mb span{display:block;width:20px;height:2px;background:var(--gd);border-radius:2px}\n.ht{font-size:14px;font-weight:700;line-height:1.2}\n.hs{font-size:10px;color:var(--gl);font-weight:600;margin-top:1px}\n.hr{display:flex;align-items:center;gap:12px}\n.ef{display:flex;align-items:center;gap:7px}\n.eb{background:var(--bg);border:2px solid var(--gl2);border-radius:9px;padding:5px 11px;cursor:pointer;transition:all .22s;line-height:0}\n.eb img{height:23px;object-fit:contain;display:block}\n.eb.ai{border-color:var(--iab);background:#EFF6FF;box-shadow:0 0 0 3px rgba(0,86,168,.12)}\n.eb.av{border-color:var(--verbum);background:#F5EEF8;box-shadow:0 0 0 3px rgba(110,17,132,.12)}\n.cw img{height:42px;object-fit:contain;transition:opacity .35s}\n.chip{background:rgba(var(--ac-rgb),.12);color:var(--ac);border-radius:20px;padding:3px 10px;font-size:10px;font-weight:700;display:none}\n.chip.show{display:inline-block}\n#mn{padding:16px 20px 48px;max-width:1400px;margin:0 auto}\n.ano-bar{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap}\n.ano-label{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:1px;white-space:nowrap}\n.ano-btn{padding:6px 18px;border-radius:20px;border:2px solid var(--gl2);background:var(--white);font-family:var(--font);font-size:12px;font-weight:700;cursor:pointer;color:var(--gm);transition:all .2s}\n.ano-btn:hover{border-color:var(--ac);color:var(--ac)}\n.ano-btn.on{background:var(--ac);border-color:var(--ac);color:var(--white)}\n.fb{background:var(--white);border-radius:var(--r);padding:12px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;box-shadow:var(--shadow)}\n.fl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:1px;white-space:nowrap}\nselect{border:1.5px solid var(--gl2);border-radius:var(--r2);padding:7px 10px;font-family:var(--font);font-size:12px;color:var(--gd);background:var(--bg);cursor:pointer;outline:none;min-width:155px;transition:border-color .2s}\nselect:focus{border-color:var(--ac)}\nselect.ativo{border-color:var(--ac);background:rgba(var(--ac-rgb),.04);font-weight:600}\n.fd{width:1px;height:24px;background:var(--gl2)}\n.bc{background:none;border:1.5px solid var(--gl2);border-radius:var(--r2);padding:7px 13px;font-family:var(--font);font-size:11px;color:var(--gm);cursor:pointer;font-weight:600;transition:all .2s}\n.bc:hover{border-color:var(--alert);color:var(--alert)}\n.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:13px;margin-bottom:14px}\n.kc{background:var(--white);border-radius:var(--r);padding:17px;box-shadow:var(--shadow);border-top:3px solid var(--ac);transition:transform .2s}\n.kc:hover{transform:translateY(-2px)}\n.ki{font-size:19px;margin-bottom:6px}\n.kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px}\n.kv{font-size:26px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:4px}\n.kv.sm{font-size:19px}\n.kd{font-size:10px;font-weight:600}\n.up{color:var(--ok)}.dn{color:var(--err)}.nt{color:var(--gl)}\n.cg{display:grid;grid-template-columns:2fr 1fr;gap:13px;margin-bottom:14px}\n.cc{background:var(--white);border-radius:var(--r);padding:18px;box-shadow:var(--shadow)}\n.ct{font-size:13px;font-weight:700;margin-bottom:2px}\n.cs{font-size:10px;color:var(--gl);margin-bottom:13px}\n.bl{display:flex;flex-direction:column;gap:10px}\n.bi{display:flex;flex-direction:column;gap:3px}\n.bh{display:flex;justify-content:space-between;align-items:center}\n.bn{font-size:11px;font-weight:600}\n.bv{font-size:11px;font-weight:700;color:var(--gm)}\n.bt{height:7px;background:var(--bg);border-radius:4px;overflow:hidden}\n.bf{height:100%;border-radius:4px;background:var(--ac);transition:width .6s cubic-bezier(.4,0,.2,1)}\n.gw{display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;padding:6px 0}\n.gv{font-size:32px;font-weight:800;text-align:center;margin-top:-8px}\n.glt{font-size:10px;color:var(--gl);text-align:center;margin-top:1px}\n.gd2{display:flex;gap:22px;margin-top:12px}\n.gdi{text-align:center}\n.gdv{font-size:16px;font-weight:800}\n.gdl{font-size:9px;color:var(--gl);font-weight:700;text-transform:uppercase;letter-spacing:.5px}\n.tc{background:var(--white);border-radius:var(--r);padding:18px;box-shadow:var(--shadow);margin-bottom:14px}\n.tw{overflow-x:auto}\ntable{width:100%;border-collapse:collapse;font-size:12px}\nthead th{padding:9px 11px;text-align:left;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;border-bottom:2px solid var(--gl2)}\ntbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\ntbody tr:last-child{border-bottom:none}\ntbody tr:hover{background:rgba(var(--ac-rgb),.03)}\ntbody td{padding:9px 11px;font-weight:500}\n.bg{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700}\n.bp{background:#E3EEF9;color:#003d7a}\n.ba{background:#FFEBEE;color:#7B0000}\n.bn2{background:#FFF3E0;color:#E65100}\n#wl{position:fixed;inset:0;background:var(--white);z-index:500;display:flex;flex-direction:column;align-items:center;justify-content:center;transition:opacity .5s,transform .5s}\n#wl.hide{opacity:0;transform:scale(1.03);pointer-events:none}\n.wlogos{display:flex;align-items:center;gap:32px;margin-bottom:22px}\n.wlogos img{height:42px;object-fit:contain}\n.wsep{width:1px;height:34px;background:var(--gl2)}\n.wtruck{width:100%;max-width:480px;margin-bottom:20px;text-align:center}\n.wtruck img{width:100%;object-fit:contain}\n.wtitle{font-size:20px;font-weight:800;text-align:center;margin-bottom:4px}\n.wsub{font-size:13px;color:var(--gl);text-align:center;margin-bottom:24px}\n.wbtns{display:flex;gap:12px;flex-wrap:wrap;justify-content:center}\n.wb{padding:12px 26px;border-radius:var(--r);font-family:var(--font);font-size:13px;font-weight:700;cursor:pointer;border:none;transition:all .2s;display:flex;align-items:center;gap:8px}\n.wb img{height:20px;object-fit:contain}\n.wi{background:var(--iab);color:#fff}.wi:hover{background:var(--iab-dark);transform:translateY(-2px)}\n.wv{background:var(--verbum);color:#fff}.wv:hover{background:var(--verbum-dark);transform:translateY(-2px)}\n.wa{background:var(--bg);color:var(--gd);border:2px solid var(--gl2)}.wa:hover{border-color:var(--alert);color:var(--alert);transform:translateY(-2px)}\n\n\n/* GRÁFICO DE LINHA */\n.linha-wrap{position:relative;width:100%;height:180px;margin-top:8px}\n.linha-svg{width:100%;height:100%;overflow:visible}\n.linha-grid{stroke:var(--gl2);stroke-width:0.5;stroke-dasharray:3 3}\n.linha-path{fill:none;stroke-width:2.5;stroke-linecap:round;stroke-linejoin:round}\n.linha-dot{r:4;cursor:pointer;transition:r .15s}\n.linha-dot:hover{r:6}\n.linha-label{font-size:9px;fill:var(--gl);font-family:'Montserrat',sans-serif}\n.linha-val{font-size:9px;font-weight:700;font-family:'Montserrat',sans-serif}\n.linha-tooltip{position:absolute;background:var(--gd);color:#fff;padding:5px 9px;border-radius:6px;font-size:11px;font-weight:600;pointer-events:none;opacity:0;transition:opacity .2s;white-space:nowrap;z-index:10}\n/* BARRAS HORIZONTAIS */\n.hbar-wrap{display:flex;flex-direction:column;gap:8px;margin-top:4px}\n.hbar-row{display:grid;grid-template-columns:110px 1fr 80px;align-items:center;gap:8px}\n.hbar-label{font-size:11px;font-weight:600;color:var(--gd);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.hbar-track{height:10px;background:var(--bg);border-radius:5px;overflow:hidden;position:relative}\n.hbar-fill{height:100%;border-radius:5px;transition:width .7s cubic-bezier(.4,0,.2,1)}\n.hbar-pct{font-size:11px;font-weight:700;white-space:nowrap}\n.hbar-badge{font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;margin-left:4px}\n\n/* ABA FINANCEIRO */\n.fin-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:14px}\n.fin-kpi{background:var(--white);border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow)}\n.fin-kpi.principal{border-top:3px solid var(--ac)}\n.fin-kpi.deducao{border-top:3px solid var(--err)}\n.fin-kpi.resultado{border-top:3px solid var(--ok)}\n.fin-kpi.custo{border-top:3px solid var(--alert)}\n.fin-kpi .ki{font-size:16px;margin-bottom:4px}\n.fin-kpi .kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px}\n.fin-kpi .kv{font-size:18px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:3px}\n.kpct{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700}\n.kpct.neg{background:#FFEBEE;color:#7B0000}\n.kpct.warn{background:#FFF3E0;color:#E65100}\n.fin-table{width:100%;border-collapse:collapse;font-size:12px}\n.fin-table thead th{padding:9px 11px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;border-bottom:2px solid var(--gl2)}\n.fin-table thead th:first-child,.fin-table thead th:nth-child(2),.fin-table thead th:nth-child(3){text-align:left}\n.fin-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.fin-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.fin-table tbody td{padding:9px 11px;text-align:right}\n.fin-table tbody td:first-child{text-align:left;font-weight:600;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.fin-table tbody td:nth-child(2){text-align:left;color:var(--gm)}\n.fin-table tbody td:nth-child(3){text-align:center;font-weight:700;color:var(--ac)}\n.fin-table tfoot td{padding:9px 11px;text-align:right;font-weight:800;border-top:2px solid var(--gl2)}\n.fin-table tfoot td:first-child{text-align:left}\n.pct-pill{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700}\n.pct-ok{background:#E3EEF9;color:#003d7a}\n.pct-warn{background:#FFF3E0;color:#E65100}\n.pct-bad{background:#FFEBEE;color:#7B0000}\n\n/* ═══ ABA FINANCEIRO ═══ */\n.fin-grid-kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px;align-items:start}.fin-grid-kpi .fk{display:flex!important;flex-direction:row!important;align-items:flex-start;box-sizing:border-box;height:auto}\n.fin-grid-kpi2{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px}\n.fk{background:var(--white);border-radius:var(--r);padding:12px 14px;box-shadow:var(--shadow);border-left:4px solid transparent;display:flex;flex-direction:row;align-items:flex-start;gap:10px}\n.fk.fb{border-color:var(--ac)}\n.fk.fd{border-color:#e57373;background:#FFF5F5}\n.fk.fl{border-color:var(--ok)}\n.fk.fc{border-color:var(--alert)}\n.fk.fm{border-color:#6A1B9A}\n.fk.fi{border-color:#00838F}\n.fk.ft{border-color:#37474F}\n.fk .fki{font-size:20px;flex-shrink:0;margin-top:2px}\n.fk .fkl{font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;margin-bottom:2px}\n.fk .fkv{font-size:14px;font-weight:800;color:var(--gd);line-height:1.2}\n.fk .fkd{font-size:10px;font-weight:600;margin-top:3px}\n.fk .fkpct{display:inline-block;padding:1px 7px;border-radius:10px;font-size:9px;font-weight:700}\n.pct-ok{background:#E3EEF9;color:#003d7a}\n.pct-warn{background:#FFF3E0;color:#E65100}\n.pct-bad{background:#FFEBEE;color:#7B0000}\n.pct-pur{background:#F3E5F5;color:#6A1B9A}\n.pct-tl{background:#E0F7FA;color:#00838F}\n.fin-charts2{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.fin-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n.fin-table{width:100%;border-collapse:collapse;font-size:12px}\n.fin-table thead th{padding:8px 10px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;border-bottom:2px solid var(--gl2)}\n.fin-table thead th:first-child,.fin-table thead th:nth-child(2),.fin-table thead th:nth-child(3){text-align:left}\n.fin-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.fin-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.fin-table tbody td{padding:8px 10px;text-align:right;font-size:12px}\n.fin-table tbody td:first-child{text-align:left;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.fin-table tbody td:nth-child(2){text-align:left;color:var(--gm);font-size:11px}\n.fin-table tbody td:nth-child(3){text-align:center;font-weight:700;color:var(--ac)}\n.fin-table tfoot td{padding:8px 10px;text-align:right;font-weight:800;border-top:2px solid var(--gl2);color:var(--gd)}\n.fin-table tfoot td:first-child{text-align:left}\n\n/* ═══ ABA TRANSPORTADORAS ═══ */\n.tr-grid-kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}\n.tr-grid-kpi2{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px}\n.tr-charts2{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.tr-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n.tr-table{width:100%;border-collapse:collapse;font-size:12px}\n.tr-table thead th{padding:8px 10px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;border-bottom:2px solid var(--gl2);white-space:nowrap}\n.tr-table thead th:first-child{text-align:left}\n.tr-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.tr-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.tr-table tbody td{padding:8px 10px;text-align:right;font-size:12px;white-space:nowrap}\n.tr-table tbody td:first-child{text-align:left;font-weight:600}\n.tr-table tfoot td{padding:8px 10px;text-align:right;font-weight:800;border-top:2px solid var(--gl2)}\n.tr-table tfoot td:first-child{text-align:left}\n.rank-badge{display:inline-block;width:20px;height:20px;border-radius:50%;font-size:10px;font-weight:800;text-align:center;line-height:20px;margin-right:6px}\n.rank-1{background:#FFD700;color:#333}\n.rank-2{background:#C0C0C0;color:#333}\n.rank-3{background:#CD7F32;color:#fff}\n/* ABA ENTREGAS */\n.page{display:none}.page.on{display:block}\n/* TABELA ACOMPANHAMENTO */\n.qtitle{font-size:14px;font-weight:700;margin-bottom:3px}\n.qsub{font-size:11px;color:var(--gl);margin-bottom:14px}\n.qtable-wrap{overflow-x:auto;border-radius:var(--r2);border:1px solid var(--gl2)}\n.qtable{width:100%;border-collapse:collapse;font-size:12px;min-width:900px}\n.qtable thead th{padding:10px 12px;text-align:left;font-size:9px;font-weight:700;color:#fff;text-transform:uppercase;letter-spacing:.8px;background:#2d3a4a;border-bottom:2px solid #1a2330;white-space:nowrap}\n.qtable tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.qtable tbody tr:last-child{border-bottom:none}\n.qtable tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.qtable tbody td{padding:10px 12px;vertical-align:middle}.qtable tbody td.obs-cell{vertical-align:top}\n.qtable tbody td.obs-cell{max-width:260px;font-size:11px;color:var(--gm);line-height:1.4}\n/* STATUS TAGS */\n.tag{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap}\n.tag-atrasado{background:#FFEBEE;color:#7B0000}\n.tag-alerta{background:#FFF3E0;color:#E65100}\n.tag-ok{background:#E3EEF9;color:#003d7a}\n.tag-prazo{background:#E8F5E9;color:#2E7D32}\n/* EMPRESA BADGE */\n.emp-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}\n.emp-iab{background:#EFF6FF;color:var(--iab)}\n.emp-verbum{background:#F5EEF8;color:var(--verbum)}\n/* KPIs ENTREGAS */\n.ent-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:14px}\n.ent-kpi{background:var(--white);border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow);border-left:3px solid var(--ac)}\n.ent-kpi .ki{font-size:16px;margin-bottom:4px}\n.ent-kpi .kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px}\n.ent-kpi .kv{font-size:22px;font-weight:800;color:var(--gd);line-height:1}\n.ent-kpi .kd{font-size:10px;font-weight:600;color:var(--gl);margin-top:2px}\n/* GRÁFICOS ENTREGAS */\n.ent-charts{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.ent-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n/* TABS internas */\n.itab-bar{display:flex;gap:4px;margin-bottom:16px;background:var(--bg);border-radius:var(--r2);padding:4px}\n.itab{padding:7px 16px;border-radius:6px;font-family:var(--font);font-size:12px;font-weight:600;cursor:pointer;border:none;background:none;color:var(--gm);transition:all .2s}\n.itab.on{background:var(--white);color:var(--ac);box-shadow:0 1px 4px rgba(0,0,0,.1)}\n/* Dias atraso coloridos */\n.da-ok{color:var(--gl);font-weight:500}\n.da-warn{color:#E65100;font-weight:700}\n.da-err{color:#7B0000;font-weight:700}\n@media(max-width:768px){\n  .ent-charts{grid-template-columns:1fr}\n  .ent-kpi-row{grid-template-columns:1fr 1fr}\n}\n@media(max-width:768px){\n  #mn{padding:10px 10px 32px}\n  .kg{grid-template-columns:1fr 1fr}\n  .kv{font-size:20px}.kv.sm{font-size:15px}\n  .cg{grid-template-columns:1fr}\n  select{min-width:128px;font-size:11px}\n  .eb img{height:19px}\n  .cw img{height:32px}\n}\n@media(max-width:440px){\n  .kg{grid-template-columns:1fr}\n  .wbtns{flex-direction:column;width:90%;align-items:stretch}\n  .wb{justify-content:center}\n}\n"
    html = html.replace('</style>', _css_bka + '\n</style>', 1)

    # 2b. Adicionar indicador de aba ativa no header
    _aba_nomes = {
        "visao-executiva": "📊 Visão Executiva",
        "financeiro":      "💰 Financeiro",
        "transportadoras": "🚛 Transportadoras",
        "acompanhamento":  "↳ Acompanhamento",
        "estoque":         "📦 Estoque"
    }
    _aba_js = "const _ABA_NOMES = " + repr(_aba_nomes).replace("'", '"') + ";"

    # Inserir elemento no header (após .hs)
    html = html.replace(
        '</div>\n    <span class="chip" id="chip-ano">',
        '</div>\n    <span id="aba-ativa" style="font-size:11px;font-weight:600;color:var(--ac);background:#E3EEF9;padding:3px 10px;border-radius:12px;margin-left:8px"></span>\n    <span class="chip" id="chip-ano">'
    )

    # Adicionar JS para atualizar o nome da aba no navPage
    _js_aba = """
""" + _aba_js + """
// Atualizar nome da aba no header
const _updateAbaNome = (p) => {
  const el = document.getElementById('aba-ativa');
  if(el) el.textContent = _ABA_NOMES[p] || '';
};
"""

    # 3. Substituir page-overview → page-visao-executiva
    _page_ve = '<div id="page-visao-executiva" class="page on">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="veabTODOS" onclick="veAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"   id="veab2025"  onclick="veAno(\'2025\')">2025</button>\n    <button class="ano-btn"   id="veab2026"  onclick="veAno(\'2026\')">2026</button>\n  </div>\n  <div style="font-size:10px;color:var(--gl);margin-bottom:12px;padding:8px 12px;background:var(--white);border-radius:8px;box-shadow:var(--shadow);display:flex;gap:16px;flex-wrap:wrap">\n    <span><b>NS%</b> = Nível de Serviço</span><span><b>FSF%</b> = Frete ÷ Fat. Líquido</span>\n    <span><b>LT</b> = Lead Time médio (dias)</span><span><b>▲▼</b> = Variação vs Ano Ref anterior</span>\n    <span>🔍 = Clique para ver detalhes</span>\n  </div>\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">🎯 Confiabilidade &amp; Eficiência</div>\n  <div id="ve-kpi1" class="ve-grid3"></div>\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">💰 Resultado Financeiro</div>\n  <div id="ve-kpi2" class="ve-grid3"></div>\n  <div style="font-size:10px;font-weight:700;color:var(--err);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">🚨 Alertas de Exceção</div>\n  <div id="ve-alertas" class="ve-grid3" style="margin-bottom:14px"></div>\n  <div style="display:grid;grid-template-columns:2fr 1fr;gap:13px;margin-bottom:14px">\n    <div class="tc"><div class="ct">Evolução — NS% e Faturamento Líquido</div><div class="cs" id="ve-ctx-evol">—</div><div id="ve-chart-evol"></div></div>\n    <div class="tc"><div class="ct">Resultado por Região</div><div class="cs">NS% · FSF% · Entregas</div><div id="ve-chart-reg"></div></div>\n  </div>\n  <div class="tc">\n    <div class="ct">Top Transportadoras — Visão Executiva</div>\n    <div class="cs" id="ve-ctx-tr">—</div>\n    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">\n      <thead><tr style="border-bottom:2px solid var(--gl2)">\n        <th style="text-align:left;padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase">Transportadora</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">Entregas</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">NS %</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">LT (dias)</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">FSF %</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">Custo/Entrega</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:center">Status</th>\n      </tr></thead>\n      <tbody id="ve-tbody"></tbody>\n    </table></div>\n  </div>\n</div>'
    _idx_ov = html.find('<div id="page-overview"')
    _idx_after = html.find('\n<div id="page-', _idx_ov + 10)
    if _idx_ov >= 0:
        html = html[:_idx_ov] + _page_ve + '\n' + html[_idx_after:]

    # 4. Remover page-entregas
    _idx_ent = html.find('<div id="page-entregas"')
    if _idx_ent >= 0:
        _idx_after_ent = html.find('\n<div id="page-', _idx_ent + 10)
        html = html[:_idx_ent] + html[_idx_after_ent:]

    # 4b. Adicionar páginas Financeiro e Transportadoras (antes de page-acompanhamento)
    _page_fin = '<div id="page-financeiro" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="fabTODOS" onclick="ffinAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="fab2025"  onclick="ffinAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="fab2026"  onclick="ffinAno(\'2026\')">2026</button>\n  </div>\n\n  <!-- Bloco 1: Faturamento -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">💰 Faturamento</div>\n  <div class="fin-grid-kpi" id="fin-kpi-fat"></div>\n\n  <!-- Bloco 2: Custos -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">🚚 Custos Operacionais</div>\n  <div class="fin-grid-kpi" id="fin-kpi-custo"></div>\n\n  <!-- Bloco 3: Margem + Fiscal + Ticket -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">📊 Eficiência</div>\n  <div class="fin-grid-kpi2" id="fin-kpi-efic"></div>\n\n  <!-- Gráficos linha 1 -->\n  <div class="fin-charts2">\n    <div class="fin-chart">\n      <div class="ct">Evolução Acumulada</div>\n      <div class="cs" id="fin-ctx-mes" style="margin-bottom:10px">Fat. Bruto vs Custo Total acumulado</div>\n      <div id="fin-chart-mes"></div>\n    </div>\n    <div class="fin-chart">\n      <div class="ct">Top 10 — Maiores Faturamentos</div>\n      <div class="cs" id="fin-ctx-liq" style="margin-bottom:10px">Fat. Bruto - Devoluções por cliente</div>\n      <div id="fin-chart-liq"></div>\n    </div>\n  </div>\n\n  <!-- Gráficos linha 2 -->\n  <div class="fin-charts2">\n    <div class="fin-chart">\n      <div class="ct">Resultado por Região</div>\n      <div class="cs" style="margin-bottom:10px">Fat. Bruto vs Líquido e % Custo Op.</div>\n      <div id="fin-chart-reg"></div>\n    </div>\n    <div class="fin-chart">\n      <div class="ct">Frete por Transportadora</div>\n      <div class="cs" style="margin-bottom:10px">Custo de frete · % sobre Fat. Líquido</div>\n      <div id="fin-chart-tr" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Tabela clientes -->\n  <div class="tc">\n    <div class="ct">Top Clientes — Resultado Financeiro</div>\n    <div class="cs" id="fin-ctx-tab" style="margin-bottom:12px">Ordenado por maior faturamento</div>\n    <div class="tw">\n      <table class="fin-table">\n        <thead><tr>\n          <th style="text-align:left">Cliente</th>\n          <th style="text-align:left">Cidade</th>\n          <th style="text-align:center">UF</th>\n          <th>Fat. Bruto</th><th>Devoluções</th><th>% Dev</th>\n          <th>Bonificações</th><th>% Bon</th>\n          <th>Fat. Líquido</th><th>Custo Op.</th><th>% Custo/Líq</th>\n          <th>Margem Op.</th>\n        </tr></thead>\n        <tbody id="fin-tbody"></tbody>\n        <tfoot id="fin-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n</div>\n\n\n<!-- ═══════════════ PÁGINA ACOMPANHAMENTO ═══════════════ -->\n<div id="page-acompanhamento" class="page">\n  <div style="background:var(--white);border-radius:var(--r);padding:16px 20px;margin-bottom:16px;box-shadow:var(--shadow);display:flex;align-items:center;gap:12px;flex-wrap:wrap">\n    <span style="font-size:12px;color:var(--gl);font-weight:600">⚠️ Este quadro exibe dados do <strong style="color:var(--gd)">Ano Ref 2026</strong> e não é afetado pelos filtros globais.</span>\n    <span id="acomp-hoje" style="margin-left:auto;font-size:11px;color:var(--gl)"></span>\n  </div>\n  <div style="margin-bottom:20px">\n    <div class="qtitle">📋 Quadro 1 — Entregas Pendentes e em Alerta</div>\n    <div class="qsub">Atrasadas ou com previsão de entrega nos próximos 2 dias</div>\n    <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">\n      <span class="tag tag-atrasado">🔴 Atrasado</span>\n      <span class="tag tag-alerta">⚠️ Vence em até 2 dias</span>\n      <span class="tag tag-ok">✅ No prazo</span>\n    </div>\n    <div class="qtable-wrap">\n      <table class="qtable">\n        <thead><tr>\n          <th>Empresa</th><th>NF</th><th>Cliente</th><th>Cidade</th><th>UF</th><th>Transportadora</th>\n          <th>Data Coleta</th><th>Prev. Entrega</th><th>Dias p/ Vencer</th><th>Observação</th>\n        </tr></thead>\n        <tbody id="q1-tbody"></tbody>\n      </table>\n    </div>\n  </div>\n  <div>\n    <div class="qtitle">📁 Quadro 2 — Histórico de Atrasos (Ano Ref 2026)</div>\n    <div class="qsub">Entregas concluídas com atraso · Mais recentes primeiro</div>\n    <div class="qtable-wrap">\n      <table class="qtable">\n        <thead><tr>\n          <th>Empresa</th><th>NF</th><th>Cliente</th><th>Transportadora</th>\n          <th>Data Coleta</th><th>Prev. Entrega</th><th>Data Efetiva</th><th>Dias Atraso</th><th>Observação</th>\n        </tr></thead>\n        <tbody id="q2-tbody"></tbody>\n      </table>\n    </div>\n  </div>\n</div><!-- fim page-acompanhamento -->\n\n<div id="page-financeiro" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="fabTODOS" onclick="ffinAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="fab2025"  onclick="ffinAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="fab2026"  onclick="ffinAno(\'2026\')">2026</button>\n  </div>\n  <div class="fin-kpi-row" id="fin-kpis"></div>\n  <div class="ent-charts" style="grid-template-columns:1fr 1fr;margin-bottom:14px">\n    <div class="ent-chart">\n      <div class="ct">Evolução Acumulada</div>\n      <div class="cs" id="fin-ctx-mes" style="margin-bottom:12px">Fat. Bruto acumulado vs Custo Total acumulado</div>\n      <div id="fin-chart-mes"></div>\n    </div>\n    <div class="ent-chart">\n      <div class="ct">Top 10 — Maiores Faturamentos</div>\n      <div class="cs" id="fin-ctx-liq" style="margin-bottom:12px">Fat. Bruto - Devoluções por cliente</div>\n      <div id="fin-chart-liq"></div>\n    </div>\n  </div>\n  <div class="ent-charts" style="grid-template-columns:1fr 1fr;margin-bottom:14px">\n    <div class="ent-chart">\n      <div class="ct">Resultado por Região</div>\n      <div class="cs" style="margin-bottom:12px">Fat. Bruto vs Líquido e % Custo Operacional</div>\n      <div id="fin-chart-reg"></div>\n    </div>\n    <div class="ent-chart">\n      <div class="ct">Custo Operacional por Transportadora</div>\n      <div class="cs" style="margin-bottom:12px">Apenas Frete · % sobre Fat. Bruto</div>\n      <div id="fin-chart-tr" class="hbar-wrap"></div>\n    </div>\n  </div>\n  <div class="tc">\n    <div class="ct">Top Clientes — Resultado Financeiro</div>\n    <div class="cs" style="margin-bottom:13px" id="fin-ctx-tab">Faturamento, deduções e custos por cliente</div>\n    <div class="tw">\n      <table class="fin-table">\n        <thead><tr>\n          <th style="text-align:left">Cliente</th>\n          <th style="text-align:left">Cidade</th>\n          <th style="text-align:center">UF</th>\n          <th>Fat. Bruto</th><th>Devoluções</th><th>% Dev</th>\n          <th>Bonificações</th><th>% Bon</th>\n          <th>Fat. Líquido</th><th>Custo Op.</th><th>% Custo</th>\n        </tr></thead>\n        <tbody id="fin-tbody"></tbody>\n        <tfoot id="fin-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n</div>\n\n</main>'
    _page_tr  = '<div id="page-transportadoras" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="trabTODOS" onclick="trAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="trab2025"  onclick="trAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="trab2026"  onclick="trAno(\'2026\')">2026</button>\n  </div>\n\n  <!-- KPIs Linha 1: Operacionais -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">🚛 Performance Operacional</div>\n  <div class="tr-grid-kpi" id="tr-kpi-op"></div>\n\n  <!-- KPIs Linha 2: Custo -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">💰 Performance de Custo</div>\n  <div class="tr-grid-kpi2" id="tr-kpi-custo"></div>\n\n  <!-- Gráficos linha 1 -->\n  <div class="tr-charts2">\n    <div class="tr-chart">\n      <div class="ct">Nível de Serviço por Transportadora</div>\n      <div class="cs" id="tr-ctx-ns" style="margin-bottom:10px">% entregas no prazo · ordenado do melhor ao pior</div>\n      <div id="tr-chart-ns" class="hbar-wrap"></div>\n    </div>\n    <div class="tr-chart">\n      <div class="ct">Lead Time Médio por Transportadora</div>\n      <div class="cs" style="margin-bottom:10px">Dias médios entre coleta e entrega efetiva</div>\n      <div id="tr-chart-lt" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Gráficos linha 2 -->\n  <div class="tr-charts2">\n    <div class="tr-chart">\n      <div class="ct">Evolução Mensal — Nível de Serviço</div>\n      <div class="cs" id="tr-ctx-mes" style="margin-bottom:10px">NS% e volume de entregas por mês</div>\n      <div id="tr-chart-mes"></div>\n    </div>\n    <div class="tr-chart">\n      <div class="ct">Prazo Médio por Região</div>\n      <div class="cs" style="margin-bottom:10px">Lead Time médio (dias) · % NS por região</div>\n      <div id="tr-chart-reg" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Tabela detalhada -->\n  <div class="tc">\n    <div class="ct">Ranking Detalhado — Transportadoras</div>\n    <div class="cs" id="tr-ctx-tab" style="margin-bottom:12px">Todos os indicadores por transportadora</div>\n    <div class="tw">\n      <table class="tr-table">\n        <thead><tr>\n          <th style="text-align:left">Transportadora</th>\n          <th>Entregas</th><th>No Prazo</th><th>Atrasos</th>\n          <th>NS %</th><th>Lead Time</th>\n          <th>Frete Total</th><th>Custo/Entrega</th><th>Custo/KG</th><th>FSF %</th>\n        </tr></thead>\n        <tbody id="tr-tbody"></tbody>\n        <tfoot id="tr-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n\n  <!-- Tabela por UF -->\n  <div class="tc" style="margin-top:14px">\n    <div class="ct">Prazo Médio e Custo por UF</div>\n    <div class="cs" id="tr-ctx-uf" style="margin-bottom:12px">Análise por estado de destino</div>\n    <div class="tw">\n      <table class="tr-table">\n        <thead><tr>\n          <th style="text-align:left">UF</th>\n          <th>Entregas</th><th>NS %</th><th>Lead Time</th>\n          <th>Frete Total</th><th>Custo/Entrega</th><th>Custo/KG</th><th>FSF %</th>\n        </tr></thead>\n        <tbody id="tr-tbody-uf"></tbody>\n      </table>\n    </div>\n  </div>\n</div>\n\n\n<!-- ═══════════════ PÁGINA ACOMPANHAMENTO ═══════════════ -->'
    # Remover TODAS as páginas antigas e reconstruir na ordem correta
    import re as _re_pg
    # Extrair page-acompanhamento
    _idx_ac_s = html.find('<div id="page-acompanhamento"')
    _idx_ac_e = html.find('\n<script>', _idx_ac_s)
    _page_ac_content = html[_idx_ac_s:_idx_ac_e].rstrip() if _idx_ac_s >= 0 else ''
    # Remover tudo entre page-visao-executiva e <script>
    _idx_ve_e = html.find('\n<div id="page-', html.find('<div id="page-visao-executiva"'))
    if _idx_ve_e < 0:
        _idx_ve_e = html.find('\n<script>', html.find('<div id="page-visao-executiva"'))
    # Reconstruir na ordem correta: VE + Fin + Tr + Ac + script
    html = (html[:_idx_ve_e] + '\n\n' +
            _page_fin + '\n\n' +
            _page_tr + '\n\n' +
            _page_ac_content + '\n\n' +
            html[_idx_ac_e:])

# VERSAO: 2.0-VE — Com Visão Executiva e pós-processamento
import sys, os
# Limpar cache Python para garantir versão mais recente
_this_dir = os.path.dirname(os.path.abspath(__file__))
_cache_dir = os.path.join(_this_dir, '__pycache__')
if os.path.exists(_cache_dir):
    import shutil as _shutil
    _shutil.rmtree(_cache_dir, ignore_errors=True)

# =============================================================================
# gerar_dashboard.py — Atualiza o dashboard.html com os dados mais recentes
# Executado automaticamente pelo ATUALIZAR.bat após o ETL
# =============================================================================

import sys, os, json, math, base64, io, shutil
import pandas as pd
import numpy as np
from datetime import datetime
from PIL import Image

# ── CAMINHOS ──
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DASH_DIR    = os.path.dirname(BASE_DIR)
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

PASTA_DASHBOARDS = _encontrar_onedrive_dashboards()          # pasta pai do etl/
PARQUET     = os.path.join(BASE_DIR, "data", "processed", "lancamentos.parquet")
DASH_OUT    = os.path.join(PASTA_DASHBOARDS, "dashboard_logistico", "dashboard.html")

# Imagens (na mesma pasta que o dashboard.html)
# Pasta onde estão as imagens (mesma pasta do dashboard.html)
DASH_DIR    = os.path.dirname(os.path.abspath(DASH_OUT)) if DASH_OUT else BASE_DIR
IMG_DIR     = DASH_DIR
_img_candidatos_extra = [
    DASH_DIR,                               # dashboard_logistico/  ← imagens aqui
    os.path.dirname(DASH_DIR),             # pai de dashboard_logistico/ = 10-DASHIBOARDs/
    PASTA_DASHBOARDS,                       # 10 - DASHIBOARDs/
    os.path.join(BASE_DIR, ".."),           # etl/../ = dashboard_logistico/
    os.path.join(BASE_DIR, "..", ".."),     # etl/../../ = 10-DASHIBOARDs/
    os.path.join(BASE_DIR, "..", "..", "..", "dashboard_logistico"),  # extra
]
IMGS = {
    "LI": (os.path.join(IMG_DIR, "logo_iab.png"),       260, 85),
    "LV": (os.path.join(IMG_DIR, "logo_verbum.png"),    220, 85),
    "CI": (os.path.join(IMG_DIR, "Carreta_iab.png"),    380, 78),
    "CV": (os.path.join(IMG_DIR, "Carreta_verbum.png"), 380, 78),
}

# ── HELPERS ──
def webp(path, max_w, q=78):
    img = Image.open(path).convert("RGBA")
    w, h = img.size
    if w > max_w:
        img = img.resize((max_w, int(h * max_w / w)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="WEBP", quality=q, method=6)
    return "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()

def sf(v, default="—"):
    try:
        if v is None or (isinstance(v, float) and math.isnan(v)): return default
        s = str(v).strip()
        return s if s else default
    except: return default

def fmt_d(v):
    try: return v.strftime("%d/%m/%Y") if not pd.isna(v) else "—"
    except: return "—"


def gerar():
    print("=" * 55)
    print("  GERANDO DASHBOARD ATUALIZADO")
    print("=" * 55)

    # ── 1. CARREGAR DADOS ──
    print("[ 1 ] Carregando dados processados...")
    if not os.path.exists(PARQUET):
        print("ERRO: lancamentos.parquet não encontrado. Rode o ETL primeiro.")
        sys.exit(1)

    df = pd.read_parquet(PARQUET)
    hoje = pd.Timestamp.now().normalize()
    print(f"      {len(df)} registros carregados.")

    # ── 2. IMAGENS ──
    print("[ 2 ] Processando imagens...")
    imgs = {}
    for key, (path, max_w, q) in IMGS.items():
        # Tentar múltiplos caminhos
        _nome = os.path.basename(path)
        candidatos = [path] + [
            os.path.join(d, _nome)
            for d in _img_candidatos_extra if d
        ]
        # Candidatos adicionais baseados no DASH_OUT
        if DASH_OUT:
            _dash_dir = os.path.dirname(os.path.abspath(DASH_OUT))
            candidatos += [
                os.path.join(_dash_dir, _nome),
                os.path.join(os.path.dirname(_dash_dir), _nome),
            ]
        encontrado = None
        for c in candidatos:
            if os.path.exists(c):
                encontrado = c
                break
        if not encontrado:
            print(f"      AVISO: {os.path.basename(path)} não encontrado.")
            print(f"      Caminhos tentados:")
            for c in candidatos:
                print(f"        - {c}")
            imgs[key] = ""
            continue
        imgs[key] = webp(encontrado, max_w, q)
    LI = imgs.get("LI",""); LV = imgs.get("LV","")
    CI = imgs.get("CI",""); CV = imgs.get("CV","")

    # ── 3. PAYLOAD COMPRIMIDO (Visão Geral) ──
    print("[ 3 ] Gerando dados da Visão Geral...")
    nfs_map = df.groupby("id_entrega")["nf"].apply(
        lambda x: " / ".join(str(int(v)) for v in sorted(x.dropna().unique()))
    ).to_dict()

    df_pri = df[df["is_primeira_linha_entrega"] == 1].copy()
    df_pri["obs"] = df_pri["obs"].fillna("—")

    # Lookup comprimido
    emps  = sorted(df_pri["empresa"].dropna().unique().tolist())
    anos  = sorted(df_pri["ano_ref"].dropna().unique().tolist())
    tcs   = sorted(df_pri["tipo_cliente"].dropna().unique().tolist())
    nats  = sorted(df_pri["natureza_operacao"].dropna().unique().tolist())
    trs   = sorted(df_pri["transportadora_norm"].dropna().unique().tolist())
    emp_i = {v:i for i,v in enumerate(emps)}
    ano_i = {v:i for i,v in enumerate(anos)}
    tc_i  = {v:i for i,v in enumerate(tcs)}
    nat_i = {v:i for i,v in enumerate(nats)}
    tr_i  = {v:i for i,v in enumerate(trs)}

    rows_comp = []
    for _, r in df.iterrows():
        ei = emp_i.get(r["empresa"], 0)
        ai = ano_i.get(str(r["ano_ref"]), 0)
        ti = tc_i.get(r["tipo_cliente"], 0)
        ni = nat_i.get(r["natureza_operacao"], 0)
        ri = tr_i.get(r["transportadora_norm"], 0)
        rows_comp.append([
            ei, ai, ti, ni, ri,
            round(float(r["vr_nf"]), 2),
            round(float(r["frete"]), 2),
            round(float(r["custo_embalagem"]), 2),
            round(float(r["total_impostos_calc"]), 2),
            int(r["is_primeira_linha_entrega"]),
            int(r["entrega_no_prazo"]),
            int(r["entrega_em_atraso"]),
        ])

    PAYLOAD_GER = json.dumps(
        {"E":emps,"A":anos,"T":tcs,"N":nats,"R":trs,"D":rows_comp},
        ensure_ascii=False, separators=(',',':')
    )

    # ── 4. PAYLOAD ENTREGAS ──
    print("[ 4 ] Gerando dados da aba Entregas...")

    da_list, dv_list = [], []
    for _, row in df_pri.iterrows():
        prev = row["prev_entrega_cliente"]
        efet = row["data_efetiva_entrega"]
        try:
            if pd.isna(prev): da_list.append(0); dv_list.append(None); continue
            ref = efet if not pd.isna(efet) else hoje
            da_list.append(max(0, int((ref - prev).days)))
            dv_list.append(int((prev - hoje).days) if pd.isna(efet) else None)
        except: da_list.append(0); dv_list.append(None)
    df_pri["_da"] = da_list
    df_pri["_dv"] = dv_list
    df_pri["nfs_ag"] = df_pri["id_entrega"].map(nfs_map)

    def rdict(row, c2=False):
        nf_str = sf(row.get("nfs_ag", ""))
        if nf_str == "—":
            nf_val = row["nf"]
            nf_str = str(int(nf_val)) if not pd.isna(nf_val) else "—"
        d = {
            "emp": sf(row["empresa"]), "nf": nf_str,
            "cli": sf(row["cliente"]), "tr": sf(row["transportadora_norm"]),
            "mun": sf(row["municipio"]), "uf": sf(row["uf"]),
            "reg": sf(row["regiao"]), "col": fmt_d(row["data_coleta"]),
            "prev": fmt_d(row["prev_entrega_cliente"]),
            "da": int(row["_da"]),
            "dv": int(row["_dv"]) if (row["_dv"] is not None and not (isinstance(row["_dv"], float) and math.isnan(row["_dv"]))) else None,
            "obs": sf(row["obs"]), "st": sf(row["status_prazo"]),
        }
        if c2: d["ef"] = fmt_d(row["data_efetiva_entrega"])
        return d

    df26 = df_pri[df_pri["ano_ref"] == "2026"].copy()
    q1 = df26[df26["data_efetiva_entrega"].isna()].sort_values("prev_entrega_cliente")
    q2 = df26[df26["entrega_em_atraso"] == 1].sort_values("data_efetiva_entrega", ascending=False)

    def agg(d):
        def g(col):
            return [{"k": str(r[col]), "total": int(r["total"]),
                     "prazo": int(r["prazo"]), "atraso": int(r["atraso"])}
                    for _, r in d.groupby(col).agg(
                        total=("entrega_no_prazo","count"),
                        prazo=("entrega_no_prazo","sum"),
                        atraso=("entrega_em_atraso","sum")
                    ).reset_index().iterrows()]
        mes_rows = []
        for _, r in d.groupby(["ano","mes"]).agg(
            total=("entrega_no_prazo","count"),
            prazo=("entrega_no_prazo","sum"),
            atraso=("entrega_em_atraso","sum")
        ).reset_index().sort_values(["ano","mes"]).iterrows():
            mes_rows.append({"ano":int(r["ano"]),"mes":int(r["mes"]),
                             "total":int(r["total"]),"prazo":int(r["prazo"]),"atraso":int(r["atraso"])})
        return {"total":int(d.shape[0]),"prazo":int(d["entrega_no_prazo"].sum()),
                "atraso":int(d["entrega_em_atraso"].sum()),
                "reg":g("regiao"),"tc":g("tipo_cliente"),
                "tr":g("transportadora_norm"),"mes":mes_rows}

    G = {
        "TODAS": {"TODOS": agg(df_pri)},
        "IAB": {
            "TODOS": agg(df_pri[df_pri["empresa"]=="IAB"]),
            "2025":  agg(df_pri[(df_pri["empresa"]=="IAB")&(df_pri["ano_ref"]=="2025")]),
            "2026":  agg(df_pri[(df_pri["empresa"]=="IAB")&(df_pri["ano_ref"]=="2026")]),
        },
        "VERBUM": {
            "TODOS": agg(df_pri[df_pri["empresa"]=="VERBUM"]),
            "2025":  agg(df_pri[(df_pri["empresa"]=="VERBUM")&(df_pri["ano_ref"]=="2025")]),
            "2026":  agg(df_pri[(df_pri["empresa"]=="VERBUM")&(df_pri["ano_ref"]=="2026")]),
        },
    }

    PAYLOAD_ENT = json.dumps(
        {"hoje": hoje.strftime("%Y-%m-%d"),
         "q1": [rdict(r) for _, r in q1.iterrows()],
         "q2": [rdict(r, c2=True) for _, r in q2.iterrows()],
         "G": G},
        ensure_ascii=False, separators=(',',':')
    )

    anos_ref = sorted(df_pri["ano_ref"].dropna().unique().tolist())
    print(f"      Q1: {len(q1)} pendentes | Q2: {len(q2)} atrasos")

    # ── 5. LER TEMPLATE DO DASHBOARD ──
    print("[ 5 ] Lendo template do dashboard...")
    template_path = os.path.join(BASE_DIR, "dashboard_template.html")
    if not os.path.exists(template_path):
        print("ERRO: dashboard_template.html não encontrado na pasta etl/")
        sys.exit(1)

    with open(template_path, encoding="utf-8") as f:
        html = f.read()

    # ── 6. INJETAR DADOS ──
    print("[ 6 ] Injetando dados no template...")

    # Botões de ano dinâmicos
    ano_btns = "".join([
        f'<button class="ano-btn" id="ab{a}" onclick="fAno(\'{a}\')">{a}</button>'
        for a in anos_ref
    ])
    eano_btns = "".join([
        f'<button class="ano-btn" id="eab{a}" onclick="eAno(\'{a}\')">{a}</button>'
        for a in anos_ref
    ])


    # ── Payload Visão Executiva ──
    import numpy as _np

    def _kex(dp_x, d_x):
        if len(dp_x)==0: return None
        total=len(dp_x); prazo=int(dp_x["entrega_no_prazo"].sum()); atraso=int(dp_x["entrega_em_atraso"].sum())
        ns=round(prazo/total*100,1)
        lt_v=dp_x["lead_time"].dropna() if "lead_time" in dp_x.columns else pd.Series(dtype=float)
        lt=round(float(lt_v.mean()),1) if len(lt_v) else 0
        frete=float(d_x["frete"].sum())
        fat_b=float(d_x[d_x["natureza_operacao"]=="VENDA"]["vr_nf"].sum())
        dev=float(d_x[d_x["natureza_operacao"]=="DEVOLUÇÃO"]["vr_nf"].sum())
        bon=float(d_x[d_x["natureza_operacao"]=="BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l=fat_b-dev-bon
        peso=float(d_x["peso_kg"].sum()) if "peso_kg" in d_x.columns else 0
        c_kg=round(frete/peso,4) if peso else 0
        return {"total":total,"prazo":prazo,"atraso":atraso,"ns":ns,
                "taxa_at":round(atraso/total*100,1),"lt":lt,"lt_med":lt,
                "fat_b":round(fat_b),"fat_l":round(fat_l),"dev":round(dev),"bon":round(bon),
                "frete":round(frete),"fsf":round(frete/fat_l*100,2) if fat_l else 0,
                "pct_dev":round(dev/fat_b*100,1) if fat_b else 0,
                "c_ent":round(frete/total,2) if total else 0,
                "peso":round(peso,1),"c_kg":c_kg}

    df_pri_ve = df[df["is_primeira_linha_entrega"]==1].copy()
    if "lead_time" not in df_pri_ve.columns:
        df_pri_ve["lead_time"] = (df_pri_ve["data_efetiva_entrega"] - df_pri_ve["data_coleta"]).dt.days.clip(lower=0)
    df["regiao"] = df["regiao"].fillna("").str.strip().str.upper()
    df_pri_ve["regiao"] = df_pri_ve["regiao"].fillna("").str.strip().str.upper()

    _VE = {}
    for _emp in ["TODAS","IAB","VERBUM"]:
        _VE[_emp] = {}
        _de  = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        _dpe = df_pri_ve if _emp=="TODAS" else df_pri_ve[df_pri_ve["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da  = _de  if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _dpa = _dpe if _ano=="TODOS" else _dpe[_dpe["ano_ref"].astype(str)==_ano]
            _regs=[]; _transp=[]; _mes=[]
            for _r,_g in _dpa.groupby("regiao"):
                if not _r or _r=="NAN": continue
                _k=_kex(_g,_da[_da["regiao"]==_r])
                if _k: _k["k"]=str(_r); _regs.append(_k)
            for _t,_g in _dpa.groupby("transportadora_norm"):
                _k=_kex(_g,_da[_da["transportadora_norm"]==_t])
                if _k: _k["k"]=str(_t); _transp.append(_k)
            for (_a,_m),_g in _dpa.groupby(["ano","mes"]):
                _k=_kex(_g,_da[(_da["ano"]==_a)&(_da["mes"]==_m)])
                if _k: _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            _VE[_emp][_ano]={
                "total":_kex(_dpa,_da),
                "regiao":sorted(_regs,key=lambda x:-x["total"]),
                "transp":sorted(_transp,key=lambda x:-x["total"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"]))}

    class _NpEnc(json.JSONEncoder):
        def default(self,o):
            if isinstance(o,_np.integer): return int(o)
            if isinstance(o,_np.floating): return float(o)
            return super().default(o)

    PAYLOAD_VE  = json.dumps(_VE, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))
    MESES_VE_JS = json.dumps(["","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"], ensure_ascii=False)

    # ── Payload Financeiro ──
    def _kfin(d_x):
        fat_b=float(d_x[d_x["natureza_operacao"]=="VENDA"]["vr_nf"].sum())
        dev=float(d_x[d_x["natureza_operacao"]=="DEVOLUÇÃO"]["vr_nf"].sum())
        bon=float(d_x[d_x["natureza_operacao"]=="BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l=fat_b-dev-bon
        frete=float(d_x["frete"].sum())
        emb=float(d_x["custo_embalagem"].sum())
        imp=float(d_x["total_impostos_calc"].sum())
        custo=frete+emb
        margem_op=fat_l-custo-imp
        ents=len(d_x[d_x["is_primeira_linha_entrega"]==1])
        ticket_med=round(fat_l/ents) if ents else 0
        custo_med=round(custo/ents) if ents else 0
        return {"fat_b":round(fat_b),"dev":round(dev),"bon":round(bon),
                "fat_l":round(fat_l),"frete":round(frete),"emb":round(emb),"imp":round(imp),
                "custo":round(custo),"margem_op":round(margem_op),"ents":ents,
                "ticket_med":ticket_med,"custo_med":custo_med,
                "pct_dev":round(dev/fat_b*100,1) if fat_b else 0,
                "pct_bon":round(bon/fat_b*100,1) if fat_b else 0,
                "pct_frete":round(frete/fat_l*100,2) if fat_l else 0,
                "pct_custo":round(custo/fat_l*100,2) if fat_l else 0,
                "pct_imp":round(imp/fat_l*100,2) if fat_l else 0,
                "pct_margem":round(margem_op/fat_l*100,1) if fat_l else 0}

    _FIN = {}
    for _emp in ["TODAS","IAB","VERBUM"]:
        _FIN[_emp] = {}
        _de = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da = _de if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _regs=[]; _trs=[]; _mes=[]; _clis=[]
            for _r,_g in _da.groupby("regiao"):
                if not _r: continue
                _k=_kfin(_g); _k["k"]=str(_r); _regs.append(_k)
            for _t,_g in _da.groupby("transportadora_norm"):
                _k=_kfin(_g); _k["k"]=str(_t); _trs.append(_k)
            for (_a,_m),_g in _da.groupby(["ano","mes"]):
                _k=_kfin(_g); _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            for _c,_g in _da.groupby("cliente"):
                _k=_kfin(_g); _k["k"]=str(_c)
                # Adicionar cidade e UF do cliente
                _row1 = _g.iloc[0] if len(_g) > 0 else None
                if _row1 is not None:
                    _k["mun"] = str(_row1.get("municipio","")) if "municipio" in _g.columns else ""
                    _k["uf"]  = str(_row1.get("uf","")) if "uf" in _g.columns else ""
                _clis.append(_k)
            _FIN[_emp][_ano]={
                "total":_kfin(_da),
                "regiao":sorted(_regs,key=lambda x:-x["fat_l"]),
                "transp":sorted(_trs,key=lambda x:-x["frete"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"])),
                "clientes":sorted(_clis,key=lambda x:-x["fat_l"])[:20]}

    PAYLOAD_FIN = json.dumps(_FIN, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))

    # ── Payload Transportadoras ──
    _TR = {}
    df_pri_tr = df[df["is_primeira_linha_entrega"]==1].copy()
    if "lead_time" not in df_pri_tr.columns:
        df_pri_tr["lead_time"] = (df_pri_tr["data_efetiva_entrega"] - df_pri_tr["data_coleta"]).dt.days.clip(lower=0)
    for _emp in ["TODAS","IAB","VERBUM"]:
        _TR[_emp] = {}
        _de  = df if _emp=="TODAS" else df[df["empresa"]==_emp]
        _dpe = df_pri_tr if _emp=="TODAS" else df_pri_tr[df_pri_tr["empresa"]==_emp]
        for _ano in ["TODOS","2025","2026"]:
            _da  = _de  if _ano=="TODOS" else _de[_de["ano_ref"].astype(str)==_ano]
            _dpa = _dpe if _ano=="TODOS" else _dpe[_dpe["ano_ref"].astype(str)==_ano]
            _trs=[]; _ufs=[]; _mes=[]
            for _t,_g in _dpa.groupby("transportadora_norm"):
                _k=_kex(_g,_da[_da["transportadora_norm"]==_t])
                if _k: _k["k"]=str(_t); _trs.append(_k)
            for _u,_g in _dpa.groupby("uf"):
                _k=_kex(_g,_da[_da["uf"]==_u])
                if _k: _k["k"]=str(_u); _ufs.append(_k)
            for (_a,_m),_g in _dpa.groupby(["ano","mes"]):
                _k=_kex(_g,_da[(_da["ano"]==_a)&(_da["mes"]==_m)])
                if _k: _k["ano"]=int(_a); _k["mes"]=int(_m); _mes.append(_k)
            # Agregar por região para TR
            _regs_tr=[]
            _dpa_r = _dpa.copy()
            _dpa_r["regiao"] = _dpa_r["regiao"].fillna("").str.strip().str.upper()
            for _r,_g in _dpa_r.groupby("regiao"):
                if not _r or _r=="NAN": continue
                _k=_kex(_g,_da[_da["regiao"]==_r])
                if _k: _k["k"]=str(_r); _regs_tr.append(_k)
            _TR[_emp][_ano]={
                "total":_kex(_dpa,_da),
                "transp":sorted(_trs,key=lambda x:-x["total"]),
                "uf":sorted(_ufs,key=lambda x:-x["total"]),
                "mes":sorted(_mes,key=lambda x:(x["ano"],x["mes"])),
                "regiao":sorted(_regs_tr,key=lambda x:-x["total"])}

    PAYLOAD_TR = json.dumps(_TR, cls=_NpEnc, ensure_ascii=False, separators=(",",":"))

    # ── Substituições (com fallback para placeholders não definidos) ──
    def _safe_replace(h, placeholder, value):
        try:
            return h.replace(placeholder, value)
        except Exception as e:
            print(f"AVISO: Erro ao substituir {placeholder}: {e}")
            return h

    html = _safe_replace(html, "___LI___", LI)
    html = _safe_replace(html, "___LV___", LV)
    html = _safe_replace(html, "___CI___", CI)
    html = _safe_replace(html, "___CV___", CV)
    html = _safe_replace(html, "___PAYLOAD_GER___", PAYLOAD_GER)
    html = _safe_replace(html, "___PAYLOAD_VE___", PAYLOAD_VE)
    html = _safe_replace(html, "___MESES_VE___", MESES_VE_JS)
    html = _safe_replace(html, "___PAYLOAD_ENT___", PAYLOAD_ENT)
    html = _safe_replace(html, "___ANO_BTNS___", ano_btns)
    html = _safe_replace(html, "___EANO_BTNS___", eano_btns)
    html = _safe_replace(html, "___GERADO_EM___", datetime.now().strftime("%d/%m/%Y %H:%M"))
    # Remover quaisquer placeholders restantes para evitar erros JS
    import re as _re3; html = _re3.sub(r'___[A-Z_]+___', 'null', html)


    # ── 7. VALIDAR E SALVAR ──
    print("[ 7 ] Validando e salvando dashboard.html...")
    import re
    placeholders_restantes = re.findall(r'___[A-Z_]+___', html)
    if placeholders_restantes:
        print(f"ERRO CRÍTICO: Placeholders não substituídos: {set(placeholders_restantes)}")
        print("Verifique se o dashboard_template.html esta na pasta etl/")
        sys.exit(1)

    # ══════════════════════════════════════════════
    # PÓS-PROCESSAMENTO COMPLETO
    # ══════════════════════════════════════════════
    import re as _re

    # 1. CSS da Visão Executiva
    _css_ve = '<style>\n.ve-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}\n.ve-card{background:var(--white);border-radius:var(--r);padding:16px 18px;box-shadow:var(--shadow);border-left:5px solid transparent;cursor:pointer;transition:box-shadow .15s;display:flex;flex-direction:column}\n.ve-card:hover{box-shadow:0 4px 20px rgba(0,0,0,.12)}\n.ve-card.ok{border-color:var(--ok)}.ve-card.warn{border-color:var(--alert)}.ve-card.bad{border-color:var(--err)}.ve-card.neu{border-color:var(--ac)}\n.ve-card .vc-ico{font-size:16px;margin-bottom:4px}\n.ve-card .vc-lbl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px}\n.ve-card .vc-val{font-size:26px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:8px}\n.ve-card .vc-badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:10px;font-weight:700;margin-bottom:6px}\n.ve-card .vc-yoy{font-size:11px;font-weight:700;margin-bottom:6px;min-height:18px}\n.ve-card .vc-drill{font-size:9px;color:var(--gl);cursor:pointer;margin-top:auto}\n.ve-alerta{border-radius:var(--r);padding:14px 16px}\n.ve-alerta.crit{background:#FFF5F5;border:1.5px solid #FFCDD2}\n.ve-alerta.aviso{background:#FFFDE7;border:1.5px solid #FFF176}\n.ve-alerta-tit{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}\n.ve-alerta-item{font-size:12px;padding:5px 0;border-bottom:1px solid rgba(0,0,0,.06);display:flex;justify-content:space-between;align-items:center}\n.ve-alerta-item:last-child{border:none}\n.ve-alerta-vazio{font-size:12px;color:var(--gl);font-style:italic;text-align:center;padding:10px 0}\n.tc{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow);margin-bottom:14px}\n</style>'
    html = html.replace('</head>', _css_ve + '\n</head>')
    # Cache-busting: adicionar timestamp no titulo
    _ts = datetime.now().strftime('%d/%m %H:%M')
    html = html.replace('<title>Dashboard Logístico — IAB & Verbum</title>',
                       f'<title>Dashboard Logístico — IAB & Verbum ({_ts})</title>')

    # 2. Substituir menu
    html = html.replace(
        "onclick=\"navPage('overview');fMenu()\">📊 Visão Geral",
        "onclick=\"navPage('visao-executiva');fMenu()\">📊 Visão Executiva"
    )
    html = html.replace('id="ni-overview"', 'id="ni-visao-executiva"')
    html = _re.sub(r'\s*<a class="ni"[^>]+ni-ent[^>]*>📦 Entregas</a>', '', html)
    html = html.replace("let PAGINA = 'overview'", "let PAGINA = 'visao-executiva'")
    # Corrigir onclick dos itens do menu que só têm fMenu()
    html = html.replace(
        'id="ni-financeiro" onclick="fMenu()">💰 Financeiro',
        'id="ni-financeiro" onclick="navPage(\'financeiro\');fMenu()">💰 Financeiro'
    )
    html = html.replace(
        'id="ni-transportadoras" onclick="fMenu()">🚛 Transportadoras',
        'id="ni-transportadoras" onclick="navPage(\'transportadoras\');fMenu()">🚛 Transportadoras'
    )
    html = html.replace(
        'id="ni-acomp"',
        'id="ni-acomp" onclick="navPage(\'acompanhamento\');fMenu()"'
    ).replace(
        'onclick="navPage(\'acompanhamento\');fMenu()" onclick="navPage(\'acompanhamento\');fMenu()"',
        'onclick="navPage(\'acompanhamento\');fMenu()"'
    )
    # Reordenar menu: VE → Financeiro → Transportadoras → ↳Acompanhamento → Estoque
    html = _re.sub(
        r'(<a class="ni[^"]*" id="ni-acomp"[^>]*>↳ Acompanhamento</a>)'
        r'(\s*<a class="ni[^"]*" id="ni-financeiro"[^>]*>💰 Financeiro</a>)'
        r'(\s*<a class="ni[^"]*" id="ni-transportadoras"[^>]*>🚛 Transportadoras</a>)',
        r'\2\3\n    \1\n    <a class="ni" id="ni-estoque" onclick="abrirEstoque();fMenu()">📦 Estoque</a>',
        html
    )
    # Remover seção "Filtrar por Empresa"
    html = _re.sub(
        r'\s*<div class="ns">Filtrar por Empresa</div>.*?<a[^>]*>🏢 Todas as Empresas</a>',
        '', html, flags=_re.DOTALL
    )

    # 2a. Injetar CSS completo (financeiro, transportadoras, etc.)
    _css_bka = "\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\n:root{\n  --iab:#0056A8;--iab-dark:#003d7a;\n  --verbum:#6E1184;--verbum-dark:#4a0a5a;\n  --alert:#F39200;\n  --bg:#F5F5F5;--white:#fff;\n  --gd:#333333;--gm:#666666;--gl:#999999;--gl2:#E0E0E0;\n  --ok:#107C41;--err:#D13438;\n  --shadow:0 2px 12px rgba(0,0,0,.08);\n  --r:12px;--r2:8px;\n  --font:'Montserrat',sans-serif;\n  --ac:#0056A8;--ac-rgb:0,86,168;\n}\nbody{font-family:var(--font);background:var(--bg);color:var(--gd);min-height:100vh;overflow-x:hidden}\n#ov{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:200;opacity:0;pointer-events:none;transition:opacity .3s}\n#ov.open{opacity:1;pointer-events:all}\n#sb{position:fixed;left:-280px;top:0;bottom:0;width:280px;background:#1E2A3A;z-index:201;transition:left .3s cubic-bezier(.4,0,.2,1);box-shadow:4px 0 24px rgba(0,0,0,.15);display:flex;flex-direction:column}\n#sb.open{left:0}\n.sbh{padding:20px;border-bottom:1px solid rgba(255,255,255,.1);display:flex;align-items:center;justify-content:space-between}\n.sbh img{height:30px;object-fit:contain}\n.sbx{background:none;border:none;cursor:pointer;font-size:20px;color:rgba(255,255,255,.6);padding:4px;line-height:1}\n.sbn{flex:1;padding:8px 0;overflow-y:auto}\n.ni{display:flex;align-items:center;gap:11px;padding:12px 22px;cursor:pointer;color:rgba(255,255,255,.75);font-size:13px;font-weight:500;transition:all .18s;border-left:3px solid transparent;text-decoration:none;user-select:none}\n.ni:hover,.ni.on{background:rgba(255,255,255,.08);color:#fff;border-left-color:var(--ac);font-weight:600}\n.ns{padding:14px 22px 4px;font-size:10px;font-weight:700;color:rgba(255,255,255,.35);letter-spacing:1.5px;text-transform:uppercase}\n#hd{background:var(--white);padding:0 20px;height:60px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.07);position:sticky;top:0;z-index:100}\n.hl{display:flex;align-items:center;gap:14px}\n.mb{background:none;border:none;cursor:pointer;padding:7px;border-radius:7px;display:flex;flex-direction:column;gap:5px;transition:background .2s}\n.mb:hover{background:var(--bg)}\n.mb span{display:block;width:20px;height:2px;background:var(--gd);border-radius:2px}\n.ht{font-size:14px;font-weight:700;line-height:1.2}\n.hs{font-size:10px;color:var(--gl);font-weight:600;margin-top:1px}\n.hr{display:flex;align-items:center;gap:12px}\n.ef{display:flex;align-items:center;gap:7px}\n.eb{background:var(--bg);border:2px solid var(--gl2);border-radius:9px;padding:5px 11px;cursor:pointer;transition:all .22s;line-height:0}\n.eb img{height:23px;object-fit:contain;display:block}\n.eb.ai{border-color:var(--iab);background:#EFF6FF;box-shadow:0 0 0 3px rgba(0,86,168,.12)}\n.eb.av{border-color:var(--verbum);background:#F5EEF8;box-shadow:0 0 0 3px rgba(110,17,132,.12)}\n.cw img{height:42px;object-fit:contain;transition:opacity .35s}\n.chip{background:rgba(var(--ac-rgb),.12);color:var(--ac);border-radius:20px;padding:3px 10px;font-size:10px;font-weight:700;display:none}\n.chip.show{display:inline-block}\n#mn{padding:16px 20px 48px;max-width:1400px;margin:0 auto}\n.ano-bar{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap}\n.ano-label{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:1px;white-space:nowrap}\n.ano-btn{padding:6px 18px;border-radius:20px;border:2px solid var(--gl2);background:var(--white);font-family:var(--font);font-size:12px;font-weight:700;cursor:pointer;color:var(--gm);transition:all .2s}\n.ano-btn:hover{border-color:var(--ac);color:var(--ac)}\n.ano-btn.on{background:var(--ac);border-color:var(--ac);color:var(--white)}\n.fb{background:var(--white);border-radius:var(--r);padding:12px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;box-shadow:var(--shadow)}\n.fl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:1px;white-space:nowrap}\nselect{border:1.5px solid var(--gl2);border-radius:var(--r2);padding:7px 10px;font-family:var(--font);font-size:12px;color:var(--gd);background:var(--bg);cursor:pointer;outline:none;min-width:155px;transition:border-color .2s}\nselect:focus{border-color:var(--ac)}\nselect.ativo{border-color:var(--ac);background:rgba(var(--ac-rgb),.04);font-weight:600}\n.fd{width:1px;height:24px;background:var(--gl2)}\n.bc{background:none;border:1.5px solid var(--gl2);border-radius:var(--r2);padding:7px 13px;font-family:var(--font);font-size:11px;color:var(--gm);cursor:pointer;font-weight:600;transition:all .2s}\n.bc:hover{border-color:var(--alert);color:var(--alert)}\n.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:13px;margin-bottom:14px}\n.kc{background:var(--white);border-radius:var(--r);padding:17px;box-shadow:var(--shadow);border-top:3px solid var(--ac);transition:transform .2s}\n.kc:hover{transform:translateY(-2px)}\n.ki{font-size:19px;margin-bottom:6px}\n.kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px}\n.kv{font-size:26px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:4px}\n.kv.sm{font-size:19px}\n.kd{font-size:10px;font-weight:600}\n.up{color:var(--ok)}.dn{color:var(--err)}.nt{color:var(--gl)}\n.cg{display:grid;grid-template-columns:2fr 1fr;gap:13px;margin-bottom:14px}\n.cc{background:var(--white);border-radius:var(--r);padding:18px;box-shadow:var(--shadow)}\n.ct{font-size:13px;font-weight:700;margin-bottom:2px}\n.cs{font-size:10px;color:var(--gl);margin-bottom:13px}\n.bl{display:flex;flex-direction:column;gap:10px}\n.bi{display:flex;flex-direction:column;gap:3px}\n.bh{display:flex;justify-content:space-between;align-items:center}\n.bn{font-size:11px;font-weight:600}\n.bv{font-size:11px;font-weight:700;color:var(--gm)}\n.bt{height:7px;background:var(--bg);border-radius:4px;overflow:hidden}\n.bf{height:100%;border-radius:4px;background:var(--ac);transition:width .6s cubic-bezier(.4,0,.2,1)}\n.gw{display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;padding:6px 0}\n.gv{font-size:32px;font-weight:800;text-align:center;margin-top:-8px}\n.glt{font-size:10px;color:var(--gl);text-align:center;margin-top:1px}\n.gd2{display:flex;gap:22px;margin-top:12px}\n.gdi{text-align:center}\n.gdv{font-size:16px;font-weight:800}\n.gdl{font-size:9px;color:var(--gl);font-weight:700;text-transform:uppercase;letter-spacing:.5px}\n.tc{background:var(--white);border-radius:var(--r);padding:18px;box-shadow:var(--shadow);margin-bottom:14px}\n.tw{overflow-x:auto}\ntable{width:100%;border-collapse:collapse;font-size:12px}\nthead th{padding:9px 11px;text-align:left;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;border-bottom:2px solid var(--gl2)}\ntbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\ntbody tr:last-child{border-bottom:none}\ntbody tr:hover{background:rgba(var(--ac-rgb),.03)}\ntbody td{padding:9px 11px;font-weight:500}\n.bg{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700}\n.bp{background:#E3EEF9;color:#003d7a}\n.ba{background:#FFEBEE;color:#7B0000}\n.bn2{background:#FFF3E0;color:#E65100}\n#wl{position:fixed;inset:0;background:var(--white);z-index:500;display:flex;flex-direction:column;align-items:center;justify-content:center;transition:opacity .5s,transform .5s}\n#wl.hide{opacity:0;transform:scale(1.03);pointer-events:none}\n.wlogos{display:flex;align-items:center;gap:32px;margin-bottom:22px}\n.wlogos img{height:42px;object-fit:contain}\n.wsep{width:1px;height:34px;background:var(--gl2)}\n.wtruck{width:100%;max-width:480px;margin-bottom:20px;text-align:center}\n.wtruck img{width:100%;object-fit:contain}\n.wtitle{font-size:20px;font-weight:800;text-align:center;margin-bottom:4px}\n.wsub{font-size:13px;color:var(--gl);text-align:center;margin-bottom:24px}\n.wbtns{display:flex;gap:12px;flex-wrap:wrap;justify-content:center}\n.wb{padding:12px 26px;border-radius:var(--r);font-family:var(--font);font-size:13px;font-weight:700;cursor:pointer;border:none;transition:all .2s;display:flex;align-items:center;gap:8px}\n.wb img{height:20px;object-fit:contain}\n.wi{background:var(--iab);color:#fff}.wi:hover{background:var(--iab-dark);transform:translateY(-2px)}\n.wv{background:var(--verbum);color:#fff}.wv:hover{background:var(--verbum-dark);transform:translateY(-2px)}\n.wa{background:var(--bg);color:var(--gd);border:2px solid var(--gl2)}.wa:hover{border-color:var(--alert);color:var(--alert);transform:translateY(-2px)}\n\n\n/* GRÁFICO DE LINHA */\n.linha-wrap{position:relative;width:100%;height:180px;margin-top:8px}\n.linha-svg{width:100%;height:100%;overflow:visible}\n.linha-grid{stroke:var(--gl2);stroke-width:0.5;stroke-dasharray:3 3}\n.linha-path{fill:none;stroke-width:2.5;stroke-linecap:round;stroke-linejoin:round}\n.linha-dot{r:4;cursor:pointer;transition:r .15s}\n.linha-dot:hover{r:6}\n.linha-label{font-size:9px;fill:var(--gl);font-family:'Montserrat',sans-serif}\n.linha-val{font-size:9px;font-weight:700;font-family:'Montserrat',sans-serif}\n.linha-tooltip{position:absolute;background:var(--gd);color:#fff;padding:5px 9px;border-radius:6px;font-size:11px;font-weight:600;pointer-events:none;opacity:0;transition:opacity .2s;white-space:nowrap;z-index:10}\n/* BARRAS HORIZONTAIS */\n.hbar-wrap{display:flex;flex-direction:column;gap:8px;margin-top:4px}\n.hbar-row{display:grid;grid-template-columns:110px 1fr 80px;align-items:center;gap:8px}\n.hbar-label{font-size:11px;font-weight:600;color:var(--gd);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.hbar-track{height:10px;background:var(--bg);border-radius:5px;overflow:hidden;position:relative}\n.hbar-fill{height:100%;border-radius:5px;transition:width .7s cubic-bezier(.4,0,.2,1)}\n.hbar-pct{font-size:11px;font-weight:700;white-space:nowrap}\n.hbar-badge{font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;margin-left:4px}\n\n/* ABA FINANCEIRO */\n.fin-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:14px}\n.fin-kpi{background:var(--white);border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow)}\n.fin-kpi.principal{border-top:3px solid var(--ac)}\n.fin-kpi.deducao{border-top:3px solid var(--err)}\n.fin-kpi.resultado{border-top:3px solid var(--ok)}\n.fin-kpi.custo{border-top:3px solid var(--alert)}\n.fin-kpi .ki{font-size:16px;margin-bottom:4px}\n.fin-kpi .kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px}\n.fin-kpi .kv{font-size:18px;font-weight:800;color:var(--gd);line-height:1;margin-bottom:3px}\n.kpct{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700}\n.kpct.neg{background:#FFEBEE;color:#7B0000}\n.kpct.warn{background:#FFF3E0;color:#E65100}\n.fin-table{width:100%;border-collapse:collapse;font-size:12px}\n.fin-table thead th{padding:9px 11px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;border-bottom:2px solid var(--gl2)}\n.fin-table thead th:first-child,.fin-table thead th:nth-child(2),.fin-table thead th:nth-child(3){text-align:left}\n.fin-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.fin-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.fin-table tbody td{padding:9px 11px;text-align:right}\n.fin-table tbody td:first-child{text-align:left;font-weight:600;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.fin-table tbody td:nth-child(2){text-align:left;color:var(--gm)}\n.fin-table tbody td:nth-child(3){text-align:center;font-weight:700;color:var(--ac)}\n.fin-table tfoot td{padding:9px 11px;text-align:right;font-weight:800;border-top:2px solid var(--gl2)}\n.fin-table tfoot td:first-child{text-align:left}\n.pct-pill{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700}\n.pct-ok{background:#E3EEF9;color:#003d7a}\n.pct-warn{background:#FFF3E0;color:#E65100}\n.pct-bad{background:#FFEBEE;color:#7B0000}\n\n/* ═══ ABA FINANCEIRO ═══ */\n.fin-grid-kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px;align-items:start}.fin-grid-kpi .fk{display:flex!important;flex-direction:row!important;align-items:flex-start;box-sizing:border-box;height:auto}\n.fin-grid-kpi2{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px}\n.fk{background:var(--white);border-radius:var(--r);padding:12px 14px;box-shadow:var(--shadow);border-left:4px solid transparent;display:flex;flex-direction:row;align-items:flex-start;gap:10px}\n.fk.fb{border-color:var(--ac)}\n.fk.fd{border-color:#e57373;background:#FFF5F5}\n.fk.fl{border-color:var(--ok)}\n.fk.fc{border-color:var(--alert)}\n.fk.fm{border-color:#6A1B9A}\n.fk.fi{border-color:#00838F}\n.fk.ft{border-color:#37474F}\n.fk .fki{font-size:20px;flex-shrink:0;margin-top:2px}\n.fk .fkl{font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.7px;margin-bottom:2px}\n.fk .fkv{font-size:14px;font-weight:800;color:var(--gd);line-height:1.2}\n.fk .fkd{font-size:10px;font-weight:600;margin-top:3px}\n.fk .fkpct{display:inline-block;padding:1px 7px;border-radius:10px;font-size:9px;font-weight:700}\n.pct-ok{background:#E3EEF9;color:#003d7a}\n.pct-warn{background:#FFF3E0;color:#E65100}\n.pct-bad{background:#FFEBEE;color:#7B0000}\n.pct-pur{background:#F3E5F5;color:#6A1B9A}\n.pct-tl{background:#E0F7FA;color:#00838F}\n.fin-charts2{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.fin-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n.fin-table{width:100%;border-collapse:collapse;font-size:12px}\n.fin-table thead th{padding:8px 10px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;border-bottom:2px solid var(--gl2)}\n.fin-table thead th:first-child,.fin-table thead th:nth-child(2),.fin-table thead th:nth-child(3){text-align:left}\n.fin-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.fin-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.fin-table tbody td{padding:8px 10px;text-align:right;font-size:12px}\n.fin-table tbody td:first-child{text-align:left;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.fin-table tbody td:nth-child(2){text-align:left;color:var(--gm);font-size:11px}\n.fin-table tbody td:nth-child(3){text-align:center;font-weight:700;color:var(--ac)}\n.fin-table tfoot td{padding:8px 10px;text-align:right;font-weight:800;border-top:2px solid var(--gl2);color:var(--gd)}\n.fin-table tfoot td:first-child{text-align:left}\n\n/* ═══ ABA TRANSPORTADORAS ═══ */\n.tr-grid-kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}\n.tr-grid-kpi2{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px}\n.tr-charts2{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.tr-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n.tr-table{width:100%;border-collapse:collapse;font-size:12px}\n.tr-table thead th{padding:8px 10px;text-align:right;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;border-bottom:2px solid var(--gl2);white-space:nowrap}\n.tr-table thead th:first-child{text-align:left}\n.tr-table tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.tr-table tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.tr-table tbody td{padding:8px 10px;text-align:right;font-size:12px;white-space:nowrap}\n.tr-table tbody td:first-child{text-align:left;font-weight:600}\n.tr-table tfoot td{padding:8px 10px;text-align:right;font-weight:800;border-top:2px solid var(--gl2)}\n.tr-table tfoot td:first-child{text-align:left}\n.rank-badge{display:inline-block;width:20px;height:20px;border-radius:50%;font-size:10px;font-weight:800;text-align:center;line-height:20px;margin-right:6px}\n.rank-1{background:#FFD700;color:#333}\n.rank-2{background:#C0C0C0;color:#333}\n.rank-3{background:#CD7F32;color:#fff}\n/* ABA ENTREGAS */\n.page{display:none}.page.on{display:block}\n/* TABELA ACOMPANHAMENTO */\n.qtitle{font-size:14px;font-weight:700;margin-bottom:3px}\n.qsub{font-size:11px;color:var(--gl);margin-bottom:14px}\n.qtable-wrap{overflow-x:auto;border-radius:var(--r2);border:1px solid var(--gl2)}\n.qtable{width:100%;border-collapse:collapse;font-size:12px;min-width:900px}\n.qtable thead th{padding:10px 12px;text-align:left;font-size:9px;font-weight:700;color:#fff;text-transform:uppercase;letter-spacing:.8px;background:#2d3a4a;border-bottom:2px solid #1a2330;white-space:nowrap}\n.qtable tbody tr{border-bottom:1px solid var(--gl2);transition:background .15s}\n.qtable tbody tr:last-child{border-bottom:none}\n.qtable tbody tr:hover{background:rgba(var(--ac-rgb),.03)}\n.qtable tbody td{padding:10px 12px;vertical-align:middle}.qtable tbody td.obs-cell{vertical-align:top}\n.qtable tbody td.obs-cell{max-width:260px;font-size:11px;color:var(--gm);line-height:1.4}\n/* STATUS TAGS */\n.tag{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;white-space:nowrap}\n.tag-atrasado{background:#FFEBEE;color:#7B0000}\n.tag-alerta{background:#FFF3E0;color:#E65100}\n.tag-ok{background:#E3EEF9;color:#003d7a}\n.tag-prazo{background:#E8F5E9;color:#2E7D32}\n/* EMPRESA BADGE */\n.emp-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}\n.emp-iab{background:#EFF6FF;color:var(--iab)}\n.emp-verbum{background:#F5EEF8;color:var(--verbum)}\n/* KPIs ENTREGAS */\n.ent-kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:14px}\n.ent-kpi{background:var(--white);border-radius:var(--r);padding:14px 16px;box-shadow:var(--shadow);border-left:3px solid var(--ac)}\n.ent-kpi .ki{font-size:16px;margin-bottom:4px}\n.ent-kpi .kl{font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px}\n.ent-kpi .kv{font-size:22px;font-weight:800;color:var(--gd);line-height:1}\n.ent-kpi .kd{font-size:10px;font-weight:600;color:var(--gl);margin-top:2px}\n/* GRÁFICOS ENTREGAS */\n.ent-charts{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:14px}\n.ent-chart{background:var(--white);border-radius:var(--r);padding:16px;box-shadow:var(--shadow)}\n/* TABS internas */\n.itab-bar{display:flex;gap:4px;margin-bottom:16px;background:var(--bg);border-radius:var(--r2);padding:4px}\n.itab{padding:7px 16px;border-radius:6px;font-family:var(--font);font-size:12px;font-weight:600;cursor:pointer;border:none;background:none;color:var(--gm);transition:all .2s}\n.itab.on{background:var(--white);color:var(--ac);box-shadow:0 1px 4px rgba(0,0,0,.1)}\n/* Dias atraso coloridos */\n.da-ok{color:var(--gl);font-weight:500}\n.da-warn{color:#E65100;font-weight:700}\n.da-err{color:#7B0000;font-weight:700}\n@media(max-width:768px){\n  .ent-charts{grid-template-columns:1fr}\n  .ent-kpi-row{grid-template-columns:1fr 1fr}\n}\n@media(max-width:768px){\n  #mn{padding:10px 10px 32px}\n  .kg{grid-template-columns:1fr 1fr}\n  .kv{font-size:20px}.kv.sm{font-size:15px}\n  .cg{grid-template-columns:1fr}\n  select{min-width:128px;font-size:11px}\n  .eb img{height:19px}\n  .cw img{height:32px}\n}\n@media(max-width:440px){\n  .kg{grid-template-columns:1fr}\n  .wbtns{flex-direction:column;width:90%;align-items:stretch}\n  .wb{justify-content:center}\n}\n"
    html = html.replace('</style>', _css_bka + '\n</style>', 1)

    # 2b. Adicionar indicador de aba ativa no header
    _aba_nomes = {
        "visao-executiva": "📊 Visão Executiva",
        "financeiro":      "💰 Financeiro",
        "transportadoras": "🚛 Transportadoras",
        "acompanhamento":  "↳ Acompanhamento",
        "estoque":         "📦 Estoque"
    }
    _aba_js = "const _ABA_NOMES = " + repr(_aba_nomes).replace("'", '"') + ";"

    # Inserir elemento no header (após .hs)
    html = html.replace(
        '</div>\n    <span class="chip" id="chip-ano">',
        '</div>\n    <span id="aba-ativa" style="font-size:11px;font-weight:600;color:var(--ac);background:#E3EEF9;padding:3px 10px;border-radius:12px;margin-left:8px"></span>\n    <span class="chip" id="chip-ano">'
    )

    # Adicionar JS para atualizar o nome da aba no navPage
    _js_aba = """
""" + _aba_js + """
// Atualizar nome da aba no header
const _updateAbaNome = (p) => {
  const el = document.getElementById('aba-ativa');
  if(el) el.textContent = _ABA_NOMES[p] || '';
};
"""

    # 3. Substituir page-overview → page-visao-executiva
    _page_ve = '<div id="page-visao-executiva" class="page on">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="veabTODOS" onclick="veAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"   id="veab2025"  onclick="veAno(\'2025\')">2025</button>\n    <button class="ano-btn"   id="veab2026"  onclick="veAno(\'2026\')">2026</button>\n  </div>\n  <div style="font-size:10px;color:var(--gl);margin-bottom:12px;padding:8px 12px;background:var(--white);border-radius:8px;box-shadow:var(--shadow);display:flex;gap:16px;flex-wrap:wrap">\n    <span><b>NS%</b> = Nível de Serviço</span><span><b>FSF%</b> = Frete ÷ Fat. Líquido</span>\n    <span><b>LT</b> = Lead Time médio (dias)</span><span><b>▲▼</b> = Variação vs Ano Ref anterior</span>\n    <span>🔍 = Clique para ver detalhes</span>\n  </div>\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">🎯 Confiabilidade &amp; Eficiência</div>\n  <div id="ve-kpi1" class="ve-grid3"></div>\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">💰 Resultado Financeiro</div>\n  <div id="ve-kpi2" class="ve-grid3"></div>\n  <div style="font-size:10px;font-weight:700;color:var(--err);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">🚨 Alertas de Exceção</div>\n  <div id="ve-alertas" class="ve-grid3" style="margin-bottom:14px"></div>\n  <div style="display:grid;grid-template-columns:2fr 1fr;gap:13px;margin-bottom:14px">\n    <div class="tc"><div class="ct">Evolução — NS% e Faturamento Líquido</div><div class="cs" id="ve-ctx-evol">—</div><div id="ve-chart-evol"></div></div>\n    <div class="tc"><div class="ct">Resultado por Região</div><div class="cs">NS% · FSF% · Entregas</div><div id="ve-chart-reg"></div></div>\n  </div>\n  <div class="tc">\n    <div class="ct">Top Transportadoras — Visão Executiva</div>\n    <div class="cs" id="ve-ctx-tr">—</div>\n    <div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">\n      <thead><tr style="border-bottom:2px solid var(--gl2)">\n        <th style="text-align:left;padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase">Transportadora</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">Entregas</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">NS %</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">LT (dias)</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">FSF %</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:right">Custo/Entrega</th>\n        <th style="padding:7px 10px;font-size:9px;font-weight:700;color:var(--gl);text-transform:uppercase;text-align:center">Status</th>\n      </tr></thead>\n      <tbody id="ve-tbody"></tbody>\n    </table></div>\n  </div>\n</div>'
    _idx_ov = html.find('<div id="page-overview"')
    _idx_after = html.find('\n<div id="page-', _idx_ov + 10)
    if _idx_ov >= 0:
        html = html[:_idx_ov] + _page_ve + '\n' + html[_idx_after:]

    # 4. Remover page-entregas
    _idx_ent = html.find('<div id="page-entregas"')
    if _idx_ent >= 0:
        _idx_after_ent = html.find('\n<div id="page-', _idx_ent + 10)
        html = html[:_idx_ent] + html[_idx_after_ent:]

    # 4b. Adicionar páginas Financeiro e Transportadoras (antes de page-acompanhamento)
    _page_fin = '<div id="page-financeiro" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="fabTODOS" onclick="ffinAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="fab2025"  onclick="ffinAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="fab2026"  onclick="ffinAno(\'2026\')">2026</button>\n  </div>\n\n  <!-- Bloco 1: Faturamento -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">💰 Faturamento</div>\n  <div class="fin-grid-kpi" id="fin-kpi-fat"></div>\n\n  <!-- Bloco 2: Custos -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">🚚 Custos Operacionais</div>\n  <div class="fin-grid-kpi" id="fin-kpi-custo"></div>\n\n  <!-- Bloco 3: Margem + Fiscal + Ticket -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">📊 Eficiência</div>\n  <div class="fin-grid-kpi2" id="fin-kpi-efic"></div>\n\n  <!-- Gráficos linha 1 -->\n  <div class="fin-charts2">\n    <div class="fin-chart">\n      <div class="ct">Evolução Acumulada</div>\n      <div class="cs" id="fin-ctx-mes" style="margin-bottom:10px">Fat. Bruto vs Custo Total acumulado</div>\n      <div id="fin-chart-mes"></div>\n    </div>\n    <div class="fin-chart">\n      <div class="ct">Top 10 — Maiores Faturamentos</div>\n      <div class="cs" id="fin-ctx-liq" style="margin-bottom:10px">Fat. Bruto - Devoluções por cliente</div>\n      <div id="fin-chart-liq"></div>\n    </div>\n  </div>\n\n  <!-- Gráficos linha 2 -->\n  <div class="fin-charts2">\n    <div class="fin-chart">\n      <div class="ct">Resultado por Região</div>\n      <div class="cs" style="margin-bottom:10px">Fat. Bruto vs Líquido e % Custo Op.</div>\n      <div id="fin-chart-reg"></div>\n    </div>\n    <div class="fin-chart">\n      <div class="ct">Frete por Transportadora</div>\n      <div class="cs" style="margin-bottom:10px">Custo de frete · % sobre Fat. Líquido</div>\n      <div id="fin-chart-tr" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Tabela clientes -->\n  <div class="tc">\n    <div class="ct">Top Clientes — Resultado Financeiro</div>\n    <div class="cs" id="fin-ctx-tab" style="margin-bottom:12px">Ordenado por maior faturamento</div>\n    <div class="tw">\n      <table class="fin-table">\n        <thead><tr>\n          <th style="text-align:left">Cliente</th>\n          <th style="text-align:left">Cidade</th>\n          <th style="text-align:center">UF</th>\n          <th>Fat. Bruto</th><th>Devoluções</th><th>% Dev</th>\n          <th>Bonificações</th><th>% Bon</th>\n          <th>Fat. Líquido</th><th>Custo Op.</th><th>% Custo/Líq</th>\n          <th>Margem Op.</th>\n        </tr></thead>\n        <tbody id="fin-tbody"></tbody>\n        <tfoot id="fin-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n</div>\n\n\n<!-- ═══════════════ PÁGINA ACOMPANHAMENTO ═══════════════ -->\n<div id="page-acompanhamento" class="page">\n  <div style="background:var(--white);border-radius:var(--r);padding:16px 20px;margin-bottom:16px;box-shadow:var(--shadow);display:flex;align-items:center;gap:12px;flex-wrap:wrap">\n    <span style="font-size:12px;color:var(--gl);font-weight:600">⚠️ Este quadro exibe dados do <strong style="color:var(--gd)">Ano Ref 2026</strong> e não é afetado pelos filtros globais.</span>\n    <span id="acomp-hoje" style="margin-left:auto;font-size:11px;color:var(--gl)"></span>\n  </div>\n  <div style="margin-bottom:20px">\n    <div class="qtitle">📋 Quadro 1 — Entregas Pendentes e em Alerta</div>\n    <div class="qsub">Atrasadas ou com previsão de entrega nos próximos 2 dias</div>\n    <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">\n      <span class="tag tag-atrasado">🔴 Atrasado</span>\n      <span class="tag tag-alerta">⚠️ Vence em até 2 dias</span>\n      <span class="tag tag-ok">✅ No prazo</span>\n    </div>\n    <div class="qtable-wrap">\n      <table class="qtable">\n        <thead><tr>\n          <th>Empresa</th><th>NF</th><th>Cliente</th><th>Cidade</th><th>UF</th><th>Transportadora</th>\n          <th>Data Coleta</th><th>Prev. Entrega</th><th>Dias p/ Vencer</th><th>Observação</th>\n        </tr></thead>\n        <tbody id="q1-tbody"></tbody>\n      </table>\n    </div>\n  </div>\n  <div>\n    <div class="qtitle">📁 Quadro 2 — Histórico de Atrasos (Ano Ref 2026)</div>\n    <div class="qsub">Entregas concluídas com atraso · Mais recentes primeiro</div>\n    <div class="qtable-wrap">\n      <table class="qtable">\n        <thead><tr>\n          <th>Empresa</th><th>NF</th><th>Cliente</th><th>Transportadora</th>\n          <th>Data Coleta</th><th>Prev. Entrega</th><th>Data Efetiva</th><th>Dias Atraso</th><th>Observação</th>\n        </tr></thead>\n        <tbody id="q2-tbody"></tbody>\n      </table>\n    </div>\n  </div>\n</div><!-- fim page-acompanhamento -->\n\n<div id="page-financeiro" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="fabTODOS" onclick="ffinAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="fab2025"  onclick="ffinAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="fab2026"  onclick="ffinAno(\'2026\')">2026</button>\n  </div>\n  <div class="fin-kpi-row" id="fin-kpis"></div>\n  <div class="ent-charts" style="grid-template-columns:1fr 1fr;margin-bottom:14px">\n    <div class="ent-chart">\n      <div class="ct">Evolução Acumulada</div>\n      <div class="cs" id="fin-ctx-mes" style="margin-bottom:12px">Fat. Bruto acumulado vs Custo Total acumulado</div>\n      <div id="fin-chart-mes"></div>\n    </div>\n    <div class="ent-chart">\n      <div class="ct">Top 10 — Maiores Faturamentos</div>\n      <div class="cs" id="fin-ctx-liq" style="margin-bottom:12px">Fat. Bruto - Devoluções por cliente</div>\n      <div id="fin-chart-liq"></div>\n    </div>\n  </div>\n  <div class="ent-charts" style="grid-template-columns:1fr 1fr;margin-bottom:14px">\n    <div class="ent-chart">\n      <div class="ct">Resultado por Região</div>\n      <div class="cs" style="margin-bottom:12px">Fat. Bruto vs Líquido e % Custo Operacional</div>\n      <div id="fin-chart-reg"></div>\n    </div>\n    <div class="ent-chart">\n      <div class="ct">Custo Operacional por Transportadora</div>\n      <div class="cs" style="margin-bottom:12px">Apenas Frete · % sobre Fat. Bruto</div>\n      <div id="fin-chart-tr" class="hbar-wrap"></div>\n    </div>\n  </div>\n  <div class="tc">\n    <div class="ct">Top Clientes — Resultado Financeiro</div>\n    <div class="cs" style="margin-bottom:13px" id="fin-ctx-tab">Faturamento, deduções e custos por cliente</div>\n    <div class="tw">\n      <table class="fin-table">\n        <thead><tr>\n          <th style="text-align:left">Cliente</th>\n          <th style="text-align:left">Cidade</th>\n          <th style="text-align:center">UF</th>\n          <th>Fat. Bruto</th><th>Devoluções</th><th>% Dev</th>\n          <th>Bonificações</th><th>% Bon</th>\n          <th>Fat. Líquido</th><th>Custo Op.</th><th>% Custo</th>\n        </tr></thead>\n        <tbody id="fin-tbody"></tbody>\n        <tfoot id="fin-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n</div>\n\n</main>'
    _page_tr  = '<div id="page-transportadoras" class="page">\n  <div class="ano-bar">\n    <span class="ano-label">Ano Ref</span>\n    <button class="ano-btn on" id="trabTODOS" onclick="trAno(\'TODOS\')">Todos</button>\n    <button class="ano-btn"    id="trab2025"  onclick="trAno(\'2025\')">2025</button>\n    <button class="ano-btn"    id="trab2026"  onclick="trAno(\'2026\')">2026</button>\n  </div>\n\n  <!-- KPIs Linha 1: Operacionais -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px">🚛 Performance Operacional</div>\n  <div class="tr-grid-kpi" id="tr-kpi-op"></div>\n\n  <!-- KPIs Linha 2: Custo -->\n  <div style="font-size:10px;font-weight:700;color:var(--gl);text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;margin-top:4px">💰 Performance de Custo</div>\n  <div class="tr-grid-kpi2" id="tr-kpi-custo"></div>\n\n  <!-- Gráficos linha 1 -->\n  <div class="tr-charts2">\n    <div class="tr-chart">\n      <div class="ct">Nível de Serviço por Transportadora</div>\n      <div class="cs" id="tr-ctx-ns" style="margin-bottom:10px">% entregas no prazo · ordenado do melhor ao pior</div>\n      <div id="tr-chart-ns" class="hbar-wrap"></div>\n    </div>\n    <div class="tr-chart">\n      <div class="ct">Lead Time Médio por Transportadora</div>\n      <div class="cs" style="margin-bottom:10px">Dias médios entre coleta e entrega efetiva</div>\n      <div id="tr-chart-lt" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Gráficos linha 2 -->\n  <div class="tr-charts2">\n    <div class="tr-chart">\n      <div class="ct">Evolução Mensal — Nível de Serviço</div>\n      <div class="cs" id="tr-ctx-mes" style="margin-bottom:10px">NS% e volume de entregas por mês</div>\n      <div id="tr-chart-mes"></div>\n    </div>\n    <div class="tr-chart">\n      <div class="ct">Prazo Médio por Região</div>\n      <div class="cs" style="margin-bottom:10px">Lead Time médio (dias) · % NS por região</div>\n      <div id="tr-chart-reg" class="hbar-wrap"></div>\n    </div>\n  </div>\n\n  <!-- Tabela detalhada -->\n  <div class="tc">\n    <div class="ct">Ranking Detalhado — Transportadoras</div>\n    <div class="cs" id="tr-ctx-tab" style="margin-bottom:12px">Todos os indicadores por transportadora</div>\n    <div class="tw">\n      <table class="tr-table">\n        <thead><tr>\n          <th style="text-align:left">Transportadora</th>\n          <th>Entregas</th><th>No Prazo</th><th>Atrasos</th>\n          <th>NS %</th><th>Lead Time</th>\n          <th>Frete Total</th><th>Custo/Entrega</th><th>Custo/KG</th><th>FSF %</th>\n        </tr></thead>\n        <tbody id="tr-tbody"></tbody>\n        <tfoot id="tr-tfoot"></tfoot>\n      </table>\n    </div>\n  </div>\n\n  <!-- Tabela por UF -->\n  <div class="tc" style="margin-top:14px">\n    <div class="ct">Prazo Médio e Custo por UF</div>\n    <div class="cs" id="tr-ctx-uf" style="margin-bottom:12px">Análise por estado de destino</div>\n    <div class="tw">\n      <table class="tr-table">\n        <thead><tr>\n          <th style="text-align:left">UF</th>\n          <th>Entregas</th><th>NS %</th><th>Lead Time</th>\n          <th>Frete Total</th><th>Custo/Entrega</th><th>Custo/KG</th><th>FSF %</th>\n        </tr></thead>\n        <tbody id="tr-tbody-uf"></tbody>\n      </table>\n    </div>\n  </div>\n</div>\n\n\n<!-- ═══════════════ PÁGINA ACOMPANHAMENTO ═══════════════ -->'
    # Remover TODAS as páginas antigas e reconstruir na ordem correta
    import re as _re_pg
    # Extrair page-acompanhamento
    _idx_ac_s = html.find('<div id="page-acompanhamento"')
    _idx_ac_e = html.find('\n<script>', _idx_ac_s)
    _page_ac_content = html[_idx_ac_s:_idx_ac_e].rstrip() if _idx_ac_s >= 0 else ''
    # Remover tudo entre page-visao-executiva e <script>
    _idx_ve_e = html.find('\n<div id="page-', html.find('<div id="page-visao-executiva"'))
    if _idx_ve_e < 0:
        _idx_ve_e = html.find('\n<script>', html.find('<div id="page-visao-executiva"'))
    # Reconstruir na ordem correta: VE + Fin + Tr + Ac + script
    html = (html[:_idx_ve_e] + '\n\n' +
            _page_fin + '\n\n' +
            _page_tr + '\n\n' +
            _page_ac_content + '\n\n' +
            html[_idx_ac_e:])


    # 4c. Adicionar função abrirEstoque se não existir
    if 'function abrirEstoque' not in html:
        html = html.replace(
            "// ── INIT ──\nfAno('TODOS');",
            "function abrirEstoque() {\n  window.open('dashboard_estoque.html', '_blank');\n}\n\n// ── INIT ──\nfAno('TODOS');"
        )

    # 5. Inserir JS da VE + navPage + helpers antes do INIT
    _js_ve = '// ═══════════════ VISÃO EXECUTIVA ═══════════════\n\nfunction veAno(ano) {\n  ANO_VE = ano;\n  [\'TODOS\',\'2025\',\'2026\'].forEach(a=>{\n    const b=document.getElementById(\'veab\'+a);\n    if(b) b.classList.toggle(\'on\', a===ano);\n  });\n  renderVE();\n}\n\nfunction getVED() {\n  const g = VE[EMP] || VE[\'TODAS\'];\n  const d = g[ANO_VE];\n  return (d && d.total !== undefined ? d : (g[\'TODOS\'] || {}));\n}\nfunction getVET(ano) {\n  const g = VE[EMP] || VE[\'TODAS\'];\n  const dv = g[ano]; return dv !== undefined ? (dv.total || null) : null;\n}\n\nfunction veFmt(n) {\n  if(n>=1e6) return \'R$\\u00a0\'+(n/1e6).toFixed(1).replace(\'.\',\',\')+\'M\';\n  if(n>=1e3) return \'R$\\u00a0\'+(n/1e3).toFixed(0)+\'K\';\n  return \'R$\\u00a0\'+Math.round(n).toLocaleString(\'pt-BR\');\n}\nfunction veD(n) { return (n||0).toFixed(1).replace(\'.\',\',\'); }\nfunction veYoy(cur, prev, inv) {\n  if(prev==null||prev===0) return \'\';\n  const p = ((cur-prev)/Math.abs(prev)*100);\n  if(Math.abs(p)<0.1) return \'<span style="color:var(--gl)">→ sem variação</span>\';\n  const bom = inv ? p<0 : p>0;\n  const cor = bom ? \'var(--ok)\' : \'var(--err)\';\n  return \'<span style="color:\'+cor+\';font-weight:700">\'+(p>0?\'▲\':\'▼\')+\' \'+(p>0?\'+\':\'\')+p.toFixed(1).replace(\'.\',\',\')+\'% vs Ano Ref ant.</span>\';\n}\nfunction veCard(cor, ico, lbl, val, badge, yoy, page) {\n  const badgeBg  = cor===\'ok\'?\'#E8F5E9\':cor===\'warn\'?\'#FFF3E0\':cor===\'bad\'?\'#FFEBEE\':\'#E3EEF9\';\n  const badgeCor = cor===\'ok\'?\'#2E7D32\':cor===\'warn\'?\'#E65100\':cor===\'bad\'?\'#7B0000\':\'#003d7a\';\n  const drill = page ? \'<div class="vc-drill" onclick="navPage(\\\'\'+page+\'\\\');event.stopPropagation()">🔍 Ver detalhes em \'+page+\' →</div>\' : \'\';\n  const click = page ? \'onclick="navPage(\\\'\'+page+\'\\\')"\' : \'\';\n  return \'<div class="ve-card \'+cor+\'" \'+click+\'>\'+\n    \'<div class="vc-ico">\'+ico+\'</div>\'+\n    \'<div class="vc-lbl">\'+lbl+\'</div>\'+\n    \'<div class="vc-val">\'+val+\'</div>\'+\n    \'<div><span class="vc-badge" style="background:\'+badgeBg+\';color:\'+badgeCor+\'">\'+badge+\'</span></div>\'+\n    \'<div class="vc-yoy">\'+yoy+\'</div>\'+\n    drill+\'</div>\';\n}\n\nfunction renderVE() {\n  const d = getVED();\n  const t = d.total;\n  if(!t) return;\n\n  const tPrev = ANO_VE===\'2026\' ? getVET(\'2025\') : ANO_VE===\'2025\' ? null : getVET(\'2025\');\n  const empNm = {IAB:\'IAB\',VERBUM:\'Verbum\',TODAS:\'Todas\'}[EMP]||EMP;\n  const anoNm = ANO_VE===\'TODOS\' ? \'Todos os anos\' : \'Ano Ref \'+ANO_VE;\n  const ctx = empNm+\' · \'+anoNm;\n  [\'ve-ctx-evol\',\'ve-ctx-tr\'].forEach(id=>{ const el=document.getElementById(id); if(el) el.textContent=ctx; });\n\n  // ── KPI Confiabilidade ──\n  const nsCor = t.ns>=98.8?\'ok\':t.ns>=97?\'warn\':\'bad\';\n  const ltCor = t.lt<=7?\'ok\':t.lt<=12?\'warn\':\'bad\';\n  const ceCor = t.c_ent<=300?\'ok\':t.c_ent<=450?\'warn\':\'bad\';\n  const veK1 = document.getElementById(\'ve-kpi1\');\n  if(veK1) veK1.innerHTML =\n    veCard(nsCor,\'🎯\',\'Nível de Serviço\', veD(t.ns)+\'%\',\n      (nsCor===\'ok\'?\'Ótimo\':nsCor===\'warn\'?\'Atenção\':\'Crítico\')+\' · Meta ≥ 98,8%\',\n      veYoy(t.ns, tPrev?.ns), \'transportadoras\') +\n    veCard(ltCor,\'⏱️\',\'Lead Time Médio\', veD(t.lt)+\' dias\',\n      (ltCor===\'ok\'?\'Ótimo\':ltCor===\'warn\'?\'Atenção\':\'Crítico\')+\' · Meta ≤ 7 dias\',\n      veYoy(t.lt, tPrev?.lt, true), \'transportadoras\') +\n    veCard(ceCor,\'📊\',\'Custo Médio/Entrega\', \'R$\\u00a0\'+t.c_ent.toFixed(2).replace(\'.\',\',\'),\n      (ceCor===\'ok\'?\'Ótimo\':ceCor===\'warn\'?\'Atenção\':\'Crítico\')+\' · Meta ≤ R$ 300\',\n      veYoy(t.c_ent, tPrev?.c_ent, true), \'transportadoras\');\n\n  // ── KPI Financeiro ──\n  const fsfCor = t.fsf<=2?\'ok\':t.fsf<=3.5?\'warn\':\'bad\';\n  const devCor = t.pct_dev<=2?\'ok\':t.pct_dev<=5?\'warn\':\'bad\';\n  const veK2 = document.getElementById(\'ve-kpi2\');\n  if(veK2) veK2.innerHTML =\n    veCard(\'neu\',\'💰\',\'Faturamento Líquido\', veFmt(t.fat_l), anoNm,\n      veYoy(t.fat_l, tPrev?.fat_l), \'financeiro\') +\n    veCard(fsfCor,\'📉\',\'FSF — Frete / Fat. Líq.\', veD(t.fsf)+\'%\',\n      (fsfCor===\'ok\'?\'Ótimo\':fsfCor===\'warn\'?\'Atenção\':\'Crítico\')+\' · Meta ≤ 2,0%\',\n      veYoy(t.fsf, tPrev?.fsf, true), \'financeiro\') +\n    veCard(devCor,\'↩️\',\'Taxa de Devolução\', veD(t.pct_dev)+\'%\',\n      (devCor===\'ok\'?\'Normal\':devCor===\'warn\'?\'Atenção\':\'Crítico\')+\' · Meta ≤ 2,0%\',\n      veYoy(t.pct_dev, tPrev?.pct_dev, true), \'financeiro\');\n\n  // ── Alertas ──\n  const tr = d.transp || [];\n  const reg = d.regiao || [];\n  const trCrit  = tr.filter(r=>r.ns<97&&r.total>=30).sort((a,b)=>a.ns-b.ns);\n  const regCrit = reg.filter(r=>r.ns<95).sort((a,b)=>a.ns-b.ns);\n  const fsfCrit = tr.filter(r=>r.fsf>3.5&&r.total>=20).sort((a,b)=>b.fsf-a.fsf);\n\n  function alerta(tipo, tit, itens, campo) {\n    const cls = tipo===\'crit\' ? \'ve-alerta crit\' : \'ve-alerta aviso\';\n    const cor = tipo===\'crit\' ? \'#D13438\' : \'#F57F17\';\n    const corpo = itens.length===0\n      ? \'<div class="ve-alerta-vazio">✅ Nenhuma ocorrência</div>\'\n      : itens.slice(0,4).map(r=>{\n          const val = campo===\'fsf\' ? veD(r.fsf)+\'% FSF\' : veD(r.ns)+\'% NS\';\n          const vc  = campo===\'fsf\' ? (r.fsf>5?\'#D13438\':\'#E65100\') : (r.ns<95?\'#D13438\':\'#E65100\');\n          return \'<div class="ve-alerta-item"><span>\'+r.k+\'</span><span style="font-weight:700;color:\'+vc+\'">\'+val+\'</span></div>\';\n        }).join(\'\');\n    return \'<div class="\'+cls+\'"><div class="ve-alerta-tit" style="color:\'+cor+\'">\'+tit+\'</div>\'+corpo+\'</div>\';\n  }\n\n  const veAl = document.getElementById(\'ve-alertas\');\n  if(veAl) veAl.innerHTML =\n    alerta(\'crit\',\'🚨 Transportadoras NS &lt; 97%\', trCrit, \'ns\') +\n    alerta(\'crit\',\'🚨 Regiões NS &lt; 95%\', regCrit, \'ns\') +\n    alerta(\'aviso\',\'⚠️ FSF Acima do Limite (> 3,5%)\', fsfCrit, \'fsf\');\n\n  // ── Gráficos ──\n  veRenderEvol(d.mes || []);\n  veRenderReg(reg);\n  veRenderTransp(tr);\n}\n\nfunction veRenderEvol(rows) {\n  const el=document.getElementById(\'ve-chart-evol\');\n  if(!el||!rows.length) return;\n  const W=el.offsetWidth||520, H=200;\n  const PAD={top:26,right:54,bottom:34,left:50};\n  const cW=W-PAD.left-PAD.right, cH=H-PAD.top-PAD.bottom, n=rows.length;\n  const maxFat=Math.max(...rows.map(r=>r.fat_l),1);\n  const nsMin=85, nsMax=101;\n  const xB=i=>PAD.left+i*(cW/(n-1||1));\n  const yF=v=>PAD.top+cH-(v/maxFat)*cH;\n  const yN=v=>PAD.top+cH-((v-nsMin)/(nsMax-nsMin))*cH;\n  const step=Math.ceil(n/8);\n  const yMeta=yN(98).toFixed(1);\n  const grid=[0,.25,.5,.75,1].map(p=>{\n    const y=(PAD.top+cH-p*cH).toFixed(1);\n    return \'<line stroke="var(--gl2)" stroke-width=".5" stroke-dasharray="3 3" x1="\'+PAD.left+\'" y1="\'+y+\'" x2="\'+(PAD.left+cW)+\'" y2="\'+y+\'"/>\'+\n           \'<text font-size="9" fill="var(--gl)" font-family="Montserrat" x="\'+(PAD.left-4)+\'" y="\'+y+\'" text-anchor="end" dominant-baseline="central">\'+\n           (p*maxFat>=1e6?(p*maxFat/1e6).toFixed(1)+\'M\':p*maxFat>=1e3?(p*maxFat/1e3).toFixed(0)+\'K\':Math.round(p*maxFat))+\'</text>\';\n  }).join(\'\');\n  const nsGrid=[90,95,98,100].map(v=>\n    \'<text font-size="9" fill="var(--ac)" font-family="Montserrat" x="\'+(PAD.left+cW+4)+\'" y="\'+yN(v).toFixed(1)+\'" dominant-baseline="central">\'+v+\'%</text>\'\n  ).join(\'\');\n  const areaF=\'M\'+PAD.left+\',\'+(PAD.top+cH)+\' \'+rows.map((r,i)=>xB(i).toFixed(1)+\',\'+yF(r.fat_l).toFixed(1)).join(\' L\')+\' L\'+(PAD.left+cW)+\',\'+(PAD.top+cH)+\' Z\';\n  const pathF=rows.map((r,i)=>(i===0?\'M\':\'L\')+xB(i).toFixed(1)+\',\'+yF(r.fat_l).toFixed(1)).join(\' \');\n  const pathN=rows.map((r,i)=>(i===0?\'M\':\'L\')+xB(i).toFixed(1)+\',\'+yN(r.ns).toFixed(1)).join(\' \');\n  const dotsN=rows.map((r,i)=>{\n    const sl=n<=10||i%step===0||i===n-1;\n    const cor=r.ns>=98.8?\'var(--ok)\':r.ns>=97?\'var(--alert)\':\'var(--err)\';\n    return \'<circle cx="\'+xB(i).toFixed(1)+\'" cy="\'+yN(r.ns).toFixed(1)+\'" r="3.5" fill="\'+cor+\'" stroke="#fff" stroke-width="1.5"/>\'+\n           (sl?\'<text font-size="8" font-weight="700" fill="\'+cor+\'" font-family="Montserrat" text-anchor="middle" x="\'+xB(i).toFixed(1)+\'" y="\'+(yN(r.ns)-7).toFixed(1)+\'">\'+veD(r.ns)+\'%</text>\':\'\');\n  }).join(\'\');\n  const xlbl=rows.map((r,i)=>{\n    if(n>12&&i%step!==0&&i!==n-1) return \'\';\n    return \'<text font-size="9" fill="var(--gl)" font-family="Montserrat" text-anchor="middle" x="\'+xB(i).toFixed(1)+\'" y="\'+(H-4)+\'">\'+MESES_VE[r.mes]+\'/\'+String(r.ano).slice(2)+\'</text>\';\n  }).join(\'\');\n  el.innerHTML=\'<svg viewBox="0 0 \'+W+\' \'+H+\'" style="width:100%;height:\'+H+\'px;overflow:visible">\'+\n    \'<defs><linearGradient id="gVE" x1="0" y1="0" x2="0" y2="1">\'+\n    \'<stop offset="0%" stop-color="var(--ac)" stop-opacity=".12"/><stop offset="100%" stop-color="var(--ac)" stop-opacity=".01"/></linearGradient></defs>\'+\n    grid+nsGrid+\n    \'<path d="\'+areaF+\'" fill="url(#gVE)"/>\'+\n    \'<path d="\'+pathF+\'" fill="none" stroke="var(--ac)" stroke-width="1.5" opacity=".4" stroke-linecap="round" stroke-linejoin="round"/>\'+\n    \'<line x1="\'+PAD.left+\'" y1="\'+yMeta+\'" x2="\'+(PAD.left+cW)+\'" y2="\'+yMeta+\'" stroke="var(--alert)" stroke-width="1" stroke-dasharray="5 3"/>\'+\n    \'<path d="\'+pathN+\'" fill="none" stroke="var(--ac)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>\'+\n    dotsN+xlbl+\'</svg>\'+\n    \'<div style="display:flex;gap:14px;justify-content:center;margin-top:6px">\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--ac);display:flex;align-items:center;gap:4px"><span style="width:18px;height:2.5px;background:var(--ac);display:inline-block"></span>NS %</span>\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--ac);opacity:.4;display:flex;align-items:center;gap:4px"><span style="width:18px;height:7px;background:rgba(var(--ac-rgb),.15);display:inline-block;border-radius:2px"></span>Fat. Líquido</span>\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--alert);display:flex;align-items:center;gap:4px"><span style="width:16px;border-top:2px dashed var(--alert);display:inline-block"></span>Meta 98%</span></div>\';\n}\n\nfunction veRenderReg(rows) {\n  const el=document.getElementById(\'ve-chart-reg\');\n  if(!el) return;\n  el.innerHTML=rows.filter(r=>r.total>0).map(r=>{\n    const corNS=r.ns>=98.8?\'var(--ok)\':r.ns>=97?\'var(--alert)\':\'var(--err)\';\n    const corFSF=r.fsf<=2?\'var(--ok)\':r.fsf<=3.5?\'var(--alert)\':\'var(--err)\';\n    return \'<div style="margin-bottom:11px">\'+\n      \'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px">\'+\n      \'<span style="font-size:12px;font-weight:600">\'+r.k+\'</span>\'+\n      \'<div style="display:flex;gap:8px">\'+\n      \'<span style="font-size:10px;color:\'+corFSF+\';font-weight:700">FSF \'+veD(r.fsf)+\'%</span>\'+\n      \'<span style="font-size:11px;font-weight:800;color:\'+corNS+\'">\'+veD(r.ns)+\'%</span></div></div>\'+\n      \'<div style="height:8px;background:var(--bg);border-radius:4px;overflow:hidden">\'+\n      \'<div style="height:100%;width:\'+Math.round(r.ns)+\'%;background:\'+corNS+\';border-radius:4px"></div></div>\'+\n      \'<div style="font-size:9px;color:var(--gl);margin-top:2px">\'+r.total+\' entregas</div></div>\';\n  }).join(\'\');\n}\n\nfunction veRenderTransp(rows) {\n  const tbody=document.getElementById(\'ve-tbody\');\n  if(!tbody) return;\n  tbody.innerHTML=[...rows].filter(r=>r.total>=10).sort((a,b)=>b.total-a.total).map(r=>{\n    const nsBg=r.ns>=98.8?\'#E8F5E9\':r.ns>=97?\'#FFF3E0\':\'#FFEBEE\';\n    const nsCor=r.ns>=98.8?\'var(--ok)\':r.ns>=97?\'var(--alert)\':\'var(--err)\';\n    const fsfCor=r.fsf<=2?\'var(--ok)\':r.fsf<=3.5?\'var(--alert)\':\'var(--err)\';\n    const ltCor=r.lt<=7?\'var(--ok)\':r.lt<=12?\'var(--alert)\':\'var(--err)\';\n    const st=r.ns>=98.8\n      ?\'<span style="background:#E8F5E9;color:var(--ok);padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700">✅ Ótimo</span>\'\n      :r.ns>=97\n      ?\'<span style="background:#FFF3E0;color:var(--alert);padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700">⚠️ Atenção</span>\'\n      :\'<span style="background:#FFEBEE;color:var(--err);padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700">🚨 Crítico</span>\';\n    return \'<tr style="border-bottom:1px solid var(--gl2)">\'+\n      \'<td style="padding:8px 10px;font-weight:600;text-align:left">\'+r.k+\'</td>\'+\n      \'<td style="padding:8px 10px;text-align:right">\'+r.total+\'</td>\'+\n      \'<td style="padding:8px 10px;text-align:right"><span style="background:\'+nsBg+\';color:\'+nsCor+\';padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700">\'+veD(r.ns)+\'%</span></td>\'+\n      \'<td style="padding:8px 10px;text-align:right;color:\'+ltCor+\';font-weight:600">\'+veD(r.lt||0)+\'d</td>\'+\n      \'<td style="padding:8px 10px;text-align:right;color:\'+fsfCor+\';font-weight:600">\'+veD(r.fsf)+\'%</td>\'+\n      \'<td style="padding:8px 10px;text-align:right">R$\\u00a0\'+r.c_ent.toFixed(2).replace(\'.\',\',\')+\'</td>\'+\n      \'<td style="padding:8px 10px;text-align:center">\'+st+\'</td></tr>\';\n  }).join(\'\');\n}\n\n\nfunction _st(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}\nfunction _sh(id,v){const e=document.getElementById(id);if(e)e.innerHTML=v;}\nfunction _sc(id,p,v){const e=document.getElementById(id);if(e)e.style[p]=v;}\n\n// ── FILTROS TRANSPORTADORAS ──\n\n\nfunction trPopularFiltros() {\n  const selNat = document.getElementById(\'tr-fnat\');\n  const selTc  = document.getElementById(\'tr-ftc\');\n  if(!selNat || !selTc) return;\n  // Pegar dados do payload P (globais)\n  const nats = N || [];\n  const tcs  = T || [];\n  selNat.innerHTML = \'<option value="">Natureza da Operação</option>\' +\n    nats.map(n => \'<option value="\'+n+\'" \'+(TR_FNAT===n?\'selected\':\'\')+\'>\'+n+\'</option>\').join(\'\');\n  selTc.innerHTML = \'<option value="">Tipo de Cliente</option>\' +\n    tcs.map(t => \'<option value="\'+t+\'" \'+(TR_FTC===t?\'selected\':\'\')+\'>\'+t+\'</option>\').join(\'\');\n}\n\nfunction trFiltrar() {\n  const selNat = document.getElementById(\'tr-fnat\');\n  const selTc  = document.getElementById(\'tr-ftc\');\n  TR_FNAT = selNat ? selNat.value : \'\';\n  TR_FTC  = selTc  ? selTc.value  : \'\';\n  renderTransportadoras();\n}\n\nfunction trLimparFiltros() {\n  TR_FNAT = \'\'; TR_FTC = \'\';\n  const selNat = document.getElementById(\'tr-fnat\');\n  const selTc  = document.getElementById(\'tr-ftc\');\n  if(selNat) selNat.value = \'\';\n  if(selTc)  selTc.value  = \'\';\n  renderTransportadoras();\n}\n\n\nfunction _st(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}\nfunction _sh(id,v){const e=document.getElementById(id);if(e)e.innerHTML=v;}\nfunction _sc(id,p,v){const e=document.getElementById(id);if(e)e.style[p]=v;}\n\nfunction _st(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}\nfunction _sh(id,v){const e=document.getElementById(id);if(e)e.innerHTML=v;}\nfunction _sc(id,p,v){const e=document.getElementById(id);if(e)e.style[p]=v;}'
    _helpers = "\nfunction _st(id,v){const e=document.getElementById(id);if(e)e.textContent=v;}\nfunction _sh(id,v){const e=document.getElementById(id);if(e)e.innerHTML=v;}\nfunction _sc(id,p,v){const e=document.getElementById(id);if(e)e.style[p]=v;}\n"
    _meses_js = json.dumps(["","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"], ensure_ascii=False)
    _js_fin = 'let FANO_FIN = \'TODOS\';\n\nfunction ffinAno(ano) {\n  FANO_FIN = ano;\n  [\'TODOS\',\'2025\',\'2026\'].forEach(a => {\n    const b = document.getElementById(\'fab\'+a);\n    if(b) b.classList.toggle(\'on\', a===ano);\n  });\n  renderFinanceiro();\n}\n\nfunction getFinDados() {\n  const g = FIN[EMP] || FIN[\'TODAS\'];\n  const d = g[FANO_FIN];\n  return (d !== undefined ? d : (g[\'TODOS\'] || {}));\n}\n\nfunction fmtM(n) {\n  if(n>=1000000) return \'R$ \'+(n/1000000).toFixed(1).replace(\'.\',\',\')+\'M\';\n  if(n>=1000)    return \'R$ \'+(n/1000).toFixed(0)+\'K\';\n  return \'R$ \'+Math.round(n).toLocaleString(\'pt-BR\');\n}\nfunction fmtFin(n) { return \'R$ \'+Math.round(n).toLocaleString(\'pt-BR\'); }\n\nfunction renderFinanceiro() {\n  const d = getFinDados();\n  if(!d || !d.total) return;\n  const t = d.total;\n  const emp = {IAB:\'IAB\',VERBUM:\'Verbum\',TODAS:\'Todas\'}[EMP]||EMP;\n  const ano = FANO_FIN===\'TODOS\'?\'Todos os anos\':\'Ano Ref \'+FANO_FIN;\n  const ctx = emp+\' · \'+ano;\n  [\'fin-ctx-mes\',\'fin-ctx-liq\',\'fin-ctx-tab\'].forEach(id=>{\n    const el=document.getElementById(id); if(el) el.textContent=ctx;\n  });\n\n  // ── BLOCO 1: FATURAMENTO ──\n  document.getElementById(\'fin-kpi-fat\').innerHTML =\n    fk(\'fb\',\'💰\',\'Faturamento Bruto\',   fmtM(t.fat_b),  \'Σ Valor NFs de Venda\',\'\')  +\n    fk(\'fd\',\'🔄\',\'Devoluções\', fmtM(t.dev), pctPill(t.pct_dev,\'neg\')+\' · Fat.Bruto\',\'\') +\n    fk(\'fd\',\'🎁\',\'Bonificações\', fmtM(t.bon), pctPill(t.pct_bon,\'warn\')+\' · Fat.Bruto\',\'\') +\n    fk(\'fl\',\'✅\',\'Faturamento Líquido\', fmtM(t.fat_l),  \'Bruto - Dev. - Bon.\',\'var(--ok)\');\n\n  // ── BLOCO 2: CUSTOS ──\n  document.getElementById(\'fin-kpi-custo\').innerHTML =\n    fk(\'fc\',\'🚚\',\'Frete\',              fmtFin(t.frete), pctPill(t.pct_frete,\'warn\')+\' · FSF\', \'\') +\n    fk(\'fc\',\'📦\',\'Embalagem\',          fmtFin(t.emb),   \'Σ Custo de embalagens\',\'\') +\n    fk(\'fc\',\'⚙️\',\'Custo Operacional\',  fmtFin(t.custo), pctPill(t.pct_custo,\'warn\')+\' · Frete+Emb.\', \'\') +\n    fk(\'fi\',\'🧾\',\'Impostos\',           fmtFin(t.imp),   pctPill(t.pct_imp,\'tl\')+\' · ICMS+Difal\', \'\');\n\n  // ── BLOCO 3: EFICIÊNCIA ──\n  const corMargem = t.pct_margem >= 95 ? \'var(--ok)\' : t.pct_margem >= 85 ? \'var(--alert)\' : \'var(--err)\';\n  document.getElementById(\'fin-kpi-efic\').innerHTML =\n    fk(\'fm\',\'📈\',\'Margem Operacional\',  fmtFin(t.margem_op), \'<span style="color:\'+corMargem+\';font-weight:700">\'+t.pct_margem+\'%</span> · Fat. Líq - Custo Op.\', \'\') +\n    fk(\'ft\',\'🎯\',\'Ticket Médio/Entrega\',fmtFin(t.ticket_med),\'Fat. Bruto ÷ Nº Entregas\',\'\') +\n    fk(\'ft\',\'📊\',\'Custo Médio/Entrega\', fmtFin(t.custo_med), \'Custo Op. ÷ Nº Entregas\',\'\');\n\n  renderFinMes(d.mes || []);\n  renderFinLiq([]);\n  renderFinReg(d.regiao || []);\n  renderFinTr(d.transp || []);\n  renderFinTabela(d.clientes || [], t);\n}\n\nfunction fk(cls, icon, label, valor, sub, subCor) {\n  return \'<div class="fk \'+cls+\'">\'+\n    \'<div class="fki">\'+icon+\'</div>\'+\n    \'<div style="flex:1;min-width:0">\'+ \'<div class="fkl">\'+label+\'</div>\'+\n    \'<div class="fkv" style="\'+(subCor?\'color:\'+subCor:\'\')+\'">\'+valor+\'</div>\'+\n    \'<div class="fkd">\'+sub+\'</div>\'+\'</div></div>\';\n}\n\nfunction pctPill(v, tipo) {\n  const cls = tipo===\'neg\'?(v<=3?\'pct-ok\':v<=8?\'pct-warn\':\'pct-bad\')\n            : tipo===\'warn\'?(v<=2?\'pct-ok\':v<=4?\'pct-warn\':\'pct-bad\')\n            : tipo===\'tl\'?\'pct-tl\':\'pct-ok\';\n  return \'<span class="fkpct \'+cls+\'">\'+String(v).replace(\'.\',\',\')+\'%</span>\';\n}\n\nfunction renderFinMes(rows) {\n  const el=document.getElementById(\'fin-chart-mes\');\n  if(!el||!rows.length) return;\n  const W=el.offsetWidth||520, H=200;\n  const PAD={top:28,right:16,bottom:36,left:56};\n  const cW=W-PAD.left-PAD.right, cH=H-PAD.top-PAD.bottom;\n  const n=rows.length;\n  let acFat=0,acCusto=0;\n  const acRows=rows.map(r=>{acFat+=r.fat_b;acCusto+=r.custo;return{mes:r.mes,ano:r.ano,acFat,acCusto};});\n  const maxV=Math.max(...acRows.map(r=>r.acFat));\n  const xB=i=>PAD.left+i*(cW/(n-1||1));\n  const yB=v=>PAD.top+cH-(v/maxV)*cH;\n  const grid=[0,0.25,0.5,0.75,1].map(p=>{\n    const y=(PAD.top+cH-p*cH).toFixed(1);\n    return \'<line stroke="var(--gl2)" stroke-width="0.5" stroke-dasharray="3 3" x1="\'+PAD.left+\'" y1="\'+y+\'" x2="\'+(PAD.left+cW)+\'" y2="\'+y+\'"/>\'+\n           \'<text font-size="9" fill="var(--gl)" font-family="Montserrat,sans-serif" x="\'+(PAD.left-5)+\'" y="\'+y+\'" text-anchor="end" dominant-baseline="central">\'+fmtM(p*maxV)+\'</text>\';\n  }).join(\'\');\n  const areaFat=\'M\'+PAD.left+\',\'+(PAD.top+cH)+\' \'+acRows.map((r,i)=>xB(i).toFixed(1)+\',\'+yB(r.acFat).toFixed(1)).join(\' L\')+\' L\'+(PAD.left+cW)+\',\'+(PAD.top+cH)+\' Z\';\n  const pathFat=acRows.map((r,i)=>(i===0?\'M\':\'L\')+xB(i).toFixed(1)+\',\'+yB(r.acFat).toFixed(1)).join(\' \');\n  const pathCusto=acRows.map((r,i)=>(i===0?\'M\':\'L\')+xB(i).toFixed(1)+\',\'+yB(r.acCusto).toFixed(1)).join(\' \');\n  const step=Math.ceil(n/8);\n  const dotsF=acRows.map((r,i)=>{\n    const showL=n<=8||i%step===0||i===n-1;\n    const lbl=MESES_FIN[r.mes]+\'/\'+String(r.ano).slice(2);\n    return \'<circle cx="\'+xB(i).toFixed(1)+\'" cy="\'+yB(r.acFat).toFixed(1)+\'" r="3.5" fill="var(--ac)" stroke="#fff" stroke-width="1.5" data-tip="\'+lbl+\': \'+fmtFin(r.acFat)+\'"/>\'+\n           (showL?\'<text font-size="8" font-weight="700" fill="var(--ac)" font-family="Montserrat,sans-serif" text-anchor="middle" x="\'+xB(i).toFixed(1)+\'" y="\'+(yB(r.acFat)-8).toFixed(1)+\'">\'+fmtM(r.acFat)+\'</text>\':\'\');\n  }).join(\'\');\n  const dotsC=acRows.map((r,i)=>{\n    const showL=n<=8||i%step===0||i===n-1;\n    const lbl=MESES_FIN[r.mes]+\'/\'+String(r.ano).slice(2);\n    return \'<circle cx="\'+xB(i).toFixed(1)+\'" cy="\'+yB(r.acCusto).toFixed(1)+\'" r="3" fill="var(--alert)" stroke="#fff" stroke-width="1.5" data-tip="\'+lbl+\': \'+fmtFin(r.acCusto)+\'"/>\'+\n           (showL?\'<text font-size="8" font-weight="700" fill="var(--alert)" font-family="Montserrat,sans-serif" text-anchor="middle" x="\'+xB(i).toFixed(1)+\'" y="\'+(yB(r.acCusto)-8).toFixed(1)+\'">\'+fmtM(r.acCusto)+\'</text>\':\'\');\n  }).join(\'\');\n  const xlbls=acRows.map((r,i)=>{\n    if(n>12&&i%step!==0&&i!==n-1)return\'\';\n    return \'<text font-size="9" fill="var(--gl)" font-family="Montserrat,sans-serif" text-anchor="middle" x="\'+xB(i).toFixed(1)+\'" y="\'+(H-6)+\'">\'+MESES_FIN[r.mes]+\'/\'+String(r.ano).slice(2)+\'</text>\';\n  }).join(\'\');\n  el.innerHTML=\'<svg viewBox="0 0 \'+W+\' \'+H+\'" style="width:100%;height:\'+H+\'px;overflow:visible">\'+\n    \'<defs><linearGradient id="gradFat" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="var(--ac)" stop-opacity="0.15"/><stop offset="100%" stop-color="var(--ac)" stop-opacity="0.01"/></linearGradient></defs>\'+\n    grid+\'<path d="\'+areaFat+\'" fill="url(#gradFat)"/>\'+\n    \'<path d="\'+pathFat+\'" fill="none" stroke="var(--ac)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>\'+\n    \'<path d="\'+pathCusto+\'" fill="none" stroke="var(--alert)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" stroke-dasharray="6 3"/>\'+\n    dotsF+dotsC+xlbls+\'</svg>\'+\n    \'<div style="display:flex;gap:20px;justify-content:center;margin-top:6px">\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--ac);display:flex;align-items:center;gap:5px"><span style="width:20px;height:2.5px;background:var(--ac);display:inline-block;border-radius:2px"></span>Fat. Bruto Acum.</span>\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--alert);display:flex;align-items:center;gap:5px"><span style="width:20px;height:2px;background:var(--alert);display:inline-block;border-radius:2px"></span>Custo Op. Acum.</span></div>\';\n}\n\nfunction renderFinLiq(rows_unused) {\n  const el=document.getElementById(\'fin-chart-liq\');\n  if(!el) return;\n  const d=getFinDados();\n  if(!d||!d.clientes){el.innerHTML=\'<p style="color:var(--gl);font-size:12px">Sem dados.</p>\';return;}\n  const top=d.clientes.map(r=>({...r,fat_sem_dev:r.fat_b-r.dev})).filter(r=>r.fat_sem_dev>0).sort((a,b)=>b.fat_sem_dev-a.fat_sem_dev).slice(0,10);\n  if(!top.length){el.innerHTML=\'<p style="color:var(--gl);font-size:12px">Sem dados.</p>\';return;}\n  const maxVal=top[0].fat_sem_dev;\n  el.innerHTML=\'<div style="display:flex;flex-direction:column;gap:8px">\'+top.map((r,i)=>{\n    const pct=Math.round(r.fat_sem_dev/maxVal*100);\n    const pctLiq=r.fat_sem_dev>0?Math.round(r.fat_l/r.fat_sem_dev*100):0;\n    const nome=r.k.length>30?r.k.slice(0,30)+\'…\':r.k;\n    const cor=i===0?\'var(--ac)\':i<=2?\'#4A90E2\':\'rgba(var(--ac-rgb),0.45)\';\n    return \'<div style="display:flex;flex-direction:column;gap:3px">\'+\n      \'<div style="display:flex;justify-content:space-between;align-items:center">\'+\n      \'<div style="display:flex;align-items:center;gap:6px">\'+\n      \'<span style="font-size:10px;font-weight:800;color:var(--gl);min-width:16px;text-align:right">\'+(i+1)+\'</span>\'+\n      \'<span style="font-size:11px;font-weight:600;color:var(--gd)" title="\'+r.k+\'">\'+nome+\'</span>\'+\n      (r.dev>0?\'<span style="font-size:9px;color:var(--err);font-weight:700">-\'+fmtM(r.dev).replace(\'R$ \',\'\')+\'</span>\':\'\')+\'</div>\'+\n      \'<div style="text-align:right;white-space:nowrap">\'+\n      \'<span style="font-size:12px;font-weight:800;color:var(--ac)">\'+fmtM(r.fat_sem_dev)+\'</span>\'+\n      \'<span style="font-size:9px;color:var(--gl);margin-left:5px">\'+pctLiq+\'% líq</span></div></div>\'+\n      \'<div style="height:7px;background:var(--bg);border-radius:4px;overflow:hidden">\'+\n      \'<div style="height:100%;width:\'+pct+\'%;background:\'+cor+\';border-radius:4px"></div></div></div>\';\n  }).join(\'\')+\'</div>\';\n}\n\nfunction renderFinReg(rows) {\n  const el=document.getElementById(\'fin-chart-reg\');\n  if(!el) return;\n  const maxB=Math.max(...rows.map(r=>r.fat_b),1);\n  el.innerHTML=rows.filter(r=>r.fat_b>0).map(r=>{\n    const pctB=Math.round(r.fat_b/maxB*100),pctL=Math.round(r.fat_l/maxB*100);\n    const cor=r.pct_custo<=2?\'var(--ok)\':r.pct_custo<=4?\'var(--alert)\':\'var(--err)\';\n    return \'<div style="margin-bottom:10px">\'+\n      \'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px">\'+\n      \'<span style="font-size:11px;font-weight:600">\'+r.k+\'</span>\'+\n      \'<span style="font-size:10px;color:var(--gl)">\'+fmtM(r.fat_b)+\' · <span style="color:\'+cor+\';font-weight:700">\'+r.pct_margem+\'% margem</span></span></div>\'+\n      \'<div style="height:8px;background:var(--bg);border-radius:4px;overflow:hidden;position:relative">\'+\n      \'<div style="position:absolute;height:100%;width:\'+pctB+\'%;background:rgba(var(--ac-rgb),0.2);border-radius:4px"></div>\'+\n      \'<div style="position:absolute;height:100%;width:\'+pctL+\'%;background:var(--ac);border-radius:4px;opacity:0.85"></div></div>\'+\n      \'<div style="display:flex;justify-content:space-between;margin-top:2px">\'+\n      \'<span style="font-size:9px;color:var(--gl)">Líq: \'+fmtM(r.fat_l)+\'</span>\'+\n      \'<span style="font-size:9px;color:var(--alert)">Custo: \'+r.pct_custo+\'%</span></div></div>\';\n  }).join(\'\');\n}\n\nfunction renderFinTr(rows) {\n  const el=document.getElementById(\'fin-chart-tr\');\n  if(!el) return;\n  const d=getFinDados();\n  const fatLiq=d&&d.total?d.total.fat_l:1;\n  const top=rows.filter(r=>r.frete>0).sort((a,b)=>b.frete-a.frete).slice(0,10);\n  const maxF=Math.max(...top.map(r=>r.frete),1);\n  el.innerHTML=\'<div class="hbar-wrap">\'+top.map(r=>{\n    const pct=Math.round(r.frete/maxF*100);\n    const pctLiq=fatLiq>0?(r.frete/fatLiq*100):0;\n    const cor=pctLiq<=1?\'var(--ok)\':pctLiq<=3?\'var(--alert)\':\'var(--err)\';\n    return \'<div class="hbar-row">\'+\n      \'<span class="hbar-label" title="\'+r.k+\'">\'+r.k+\'</span>\'+\n      \'<div class="hbar-track"><div class="hbar-fill" style="width:\'+pct+\'%;background:var(--alert)"></div></div>\'+\n      \'<span class="hbar-pct" style="font-size:10px">\'+fmtM(r.frete)+\' <span style="color:\'+cor+\';font-weight:700">\'+pctLiq.toFixed(1).replace(\'.\',\',\')+\'%</span></span></div>\';\n  }).join(\'\')+\'</div>\';\n}\n\nfunction renderFinTabela(rows, tot) {\n  const tbody=document.getElementById(\'fin-tbody\');\n  const tfoot=document.getElementById(\'fin-tfoot\');\n  if(!tbody) return;\n  function pill(v,tipo){\n    const c=tipo===\'dev\'?(v<=3?\'pct-ok\':v<=8?\'pct-warn\':\'pct-bad\'):(v<=2?\'pct-ok\':v<=5?\'pct-warn\':\'pct-bad\');\n    return \'<span class="pct-pill \'+c+\'">\'+v.toFixed(1).replace(\'.\',\',\')+\'%</span>\';\n  }\n  const sorted=[...rows].filter(r=>r.fat_b>0).sort((a,b)=>b.fat_b-a.fat_b);\n  tbody.innerHTML=sorted.map(r=>{\n    const nome=r.k.length>30?r.k.slice(0,30)+\'…\':r.k;\n    const mun=(r.mun||\'—\').length>16?(r.mun||\'—\').slice(0,16)+\'…\':(r.mun||\'—\');\n    const pctCustoLiq=r.fat_l>0?r.custo/r.fat_l*100:0;\n    return \'<tr>\'+\n      \'<td title="\'+r.k+\'">\'+nome+\'</td>\'+\n      \'<td title="\'+(r.mun||\'\')+\'">\'+mun+\'</td>\'+\n      \'<td style="text-align:center;font-weight:700;color:var(--ac)">\'+(r.uf||\'—\')+\'</td>\'+\n      \'<td>\'+fmtFin(r.fat_b)+\'</td>\'+\n      \'<td style="color:var(--err)">\'+(r.dev>0?\'- \'+fmtFin(r.dev):\'—\')+\'</td>\'+\n      \'<td>\'+(r.dev>0?pill(r.pct_dev,\'dev\'):\'—\')+\'</td>\'+\n      \'<td style="color:var(--alert)">\'+(r.bon>0?\'- \'+fmtFin(r.bon):\'—\')+\'</td>\'+\n      \'<td>\'+(r.bon>0?pill(r.pct_bon,\'bon\'):\'—\')+\'</td>\'+\n      \'<td style="color:var(--ok);font-weight:700">\'+fmtFin(r.fat_l)+\'</td>\'+\n      \'<td>\'+fmtFin(r.custo)+\'</td>\'+\n      \'<td>\'+pill(pctCustoLiq,\'bon\')+\'</td>\'+\n      \'<td style="color:#6A1B9A;font-weight:700">\'+fmtFin(r.margem_op)+\'</td></tr>\';\n  }).join(\'\');\n  if(tfoot&&tot){\n    const pctCL=tot.fat_l>0?(tot.custo/tot.fat_l*100):0;\n    tfoot.innerHTML=\'<tr><td colspan="3"><strong>TOTAL GERAL</strong></td>\'+\n      \'<td><strong>\'+fmtFin(tot.fat_b)+\'</strong></td>\'+\n      \'<td style="color:var(--err)"><strong>- \'+fmtFin(tot.dev)+\'</strong></td>\'+\n      \'<td><strong>\'+tot.pct_dev+\'%</strong></td>\'+\n      \'<td style="color:var(--alert)"><strong>- \'+fmtFin(tot.bon)+\'</strong></td>\'+\n      \'<td><strong>\'+tot.pct_bon+\'%</strong></td>\'+\n      \'<td style="color:var(--ok)"><strong>\'+fmtFin(tot.fat_l)+\'</strong></td>\'+\n      \'<td><strong>\'+fmtFin(tot.custo)+\'</strong></td>\'+\n      \'<td><strong>\'+pctCL.toFixed(1).replace(\'.\',\',\')+\'%</strong></td>\'+\n      \'<td style="color:#6A1B9A"><strong>\'+fmtFin(tot.margem_op)+\'</strong></td></tr>\';\n  }\n}\n\n'
    _js_tr  = 'let ANO_TR = \'TODOS\';\n\nfunction trAno(ano) {\n  ANO_TR = ano;\n  [\'TODOS\',\'2025\',\'2026\'].forEach(a => {\n    const b = document.getElementById(\'trab\'+a);\n    if(b) b.classList.toggle(\'on\', a===ano);\n  });\n  renderTransportadoras();\n}\n\nfunction getTrDados() {\n  const g = TR[EMP] || TR[\'TODAS\'];\n  const d = g[ANO_TR];\n  return (d !== undefined ? d : (g[\'TODOS\'] || {}));\n}\n\nfunction nsColor(v) { return v>=98.8?\'#003d7a\':v>=97?\'#E65100\':\'#7B0000\'; }\nfunction nsLabel(v) { return v>=98.8?\'Ótimo\':v>=97?\'Bom\':\'Ruim\'; }\nfunction trFmtD(v)  { return v.toFixed(1).replace(\'.\',\',\'); }\nfunction trFmtR(v)  { return \'R$ \'+Math.round(v).toLocaleString(\'pt-BR\'); }\nfunction trFmtR2(v) { return \'R$ \'+v.toFixed(2).replace(\'.\',\',\'); }\n\nfunction renderTransportadoras() {\n  const d = getTrDados();\n  if(!d || !d.total) return;\n  const t = d.total;\n  const emp = {IAB:\'IAB\',VERBUM:\'Verbum\',TODAS:\'Todas\'}[EMP]||EMP;\n  const ano = ANO_TR===\'TODOS\'?\'Todos os anos\':\'Ano Ref \'+ANO_TR;\n  const ctx = emp+\' · \'+ano;\n  [\'tr-ctx-ns\',\'tr-ctx-mes\',\'tr-ctx-tab\',\'tr-ctx-uf\'].forEach(id=>{\n    const el=document.getElementById(id); if(el) el.textContent=ctx;\n  });\n\n  const corNS   = nsColor(t.ns);\n  const corTaxa = t.taxa_at<=1?\'var(--ok)\':t.taxa_at<=3?\'var(--alert)\':\'var(--err)\';\n  const corLT   = t.lt_med<=7?\'var(--ok)\':t.lt_med<=12?\'var(--alert)\':\'var(--err)\';\n  const corFSF  = t.fsf<=2?\'var(--ok)\':t.fsf<=4?\'var(--alert)\':\'var(--err)\';\n\n  // KPIs Operacionais\n  document.getElementById(\'tr-kpi-op\').innerHTML =\n    trKpi(\'fb\',\'📦\',\'Total Entregas\',      fmt(t.total),          \'Entregas únicas\',\'\',\'\') +\n    trKpi(\'fl\',\'✅\',\'Nível de Serviço\',    trFmtD(t.ns)+\'%\',        \'Entregas no prazo ÷ Total\',\'\',corNS) +\n    trKpi(\'fc\',\'⏱️\',\'Lead Time Médio\',     trFmtD(t.lt_med)+\' dias\',\'Data Efetiva - Data Coleta\',\'\',corLT) +\n    trKpi(\'fm\',\'🎯\',\'Entregas no Prazo\',   fmt(t.prazo),          fmt(t.prazo)+\' de \'+fmt(t.total),\'\',\'var(--ok)\');\n\n  // KPIs Custo\n  document.getElementById(\'tr-kpi-custo\').innerHTML =\n    trKpi(\'fc\',\'💸\',\'Frete Total\',         trFmtR(t.frete),         \'Σ Custo de Frete\',\'\',\'\') +\n    trKpi(\'fc\',\'📊\',\'Custo/Entrega\',       trFmtR2(t.c_ent),        \'Frete ÷ Nº Entregas\',\'\',\'\') +\n    trKpi(\'fc\',\'⚖️\',\'Custo/KG\',           trFmtR2(t.c_kg),         \'Frete ÷ Peso Total (kg)\',\'\',\'\') +\n    trKpi(\'fi\',\'📉\',\'FSF\',                 trFmtD(t.fsf)+\'%\',       \'Frete ÷ Fat. Líquido × 100\',\'\',corFSF) +\n    trKpi(\'fd\',\'⚠\',\'Tx. Atraso\',          trFmtD(t.taxa_at)+\'%\',   \'Atrasos ÷ Total\',\'\',corTaxa);\n\n  // Gráficos\n  trRenderNS(d.transp || []);\n  trRenderLT(d.transp || []);\n  trRenderMes(d.mes || []);\n  trRenderReg(d.regiao || []);\n  trRenderTabela(d.transp || [], t);\n  trRenderUF(d.uf || []);\n}\n\nfunction trKpi(cls, icon, label, valor, sub, pct, cor) {\n  return \'<div class="fk \'+cls+\'">\'+\n    \'<div class="fki">\'+icon+\'</div>\'+\n    \'<div style=\"flex:1;min-width:0\">\'+ \'<div class="fkl">\'+label+\'</div>\'+\n    \'<div class="fkv" style="\'+(cor?\'color:\'+cor:\'\')+\'">\'+valor+\'</div>\'+\n    \'<div class="fkd">\'+sub+\'</div>\'+ \'</div></div>\';\n}\n\nfunction trRenderNS(rows) {\n  const el=document.getElementById(\'tr-chart-ns\');\n  if(!el) return;\n  const sorted=[...rows].filter(r=>r.total>=3).sort((a,b)=>b.ns-a.ns);\n  el.innerHTML=\'<div class="hbar-wrap">\'+sorted.map(r=>{\n    const cor=nsColor(r.ns);\n    const bg=r.ns>=98.8?\'#E3EEF9\':r.ns>=97?\'#FFF3E0\':\'#FFEBEE\';\n    return \'<div class="hbar-row">\'+\n      \'<span class="hbar-label" title="\'+r.k+\'">\'+r.k+\'</span>\'+\n      \'<div class="hbar-track"><div class="hbar-fill" style="width:\'+r.ns+\'%;background:\'+cor+\'"></div></div>\'+\n      \'<span class="hbar-pct" style="font-size:10px;color:\'+cor+\';font-weight:700">\'+trFmtD(r.ns)+\'%\'+\n      \'<span style="background:\'+bg+\';color:\'+cor+\';padding:1px 5px;border-radius:8px;font-size:9px;margin-left:4px">\'+nsLabel(r.ns)+\'</span></span></div>\';\n  }).join(\'\')+\'</div>\';\n}\n\nfunction trRenderLT(rows) {\n  const el=document.getElementById(\'tr-chart-lt\');\n  if(!el) return;\n  const sorted=[...rows].filter(r=>r.total>=3&&r.lt_med>0).sort((a,b)=>a.lt_med-b.lt_med);\n  const maxLT=Math.max(...sorted.map(r=>r.lt_med),1);\n  el.innerHTML=\'<div class="hbar-wrap">\'+sorted.map(r=>{\n    const pct=Math.round(r.lt_med/maxLT*100);\n    const cor=r.lt_med<=7?\'var(--ok)\':r.lt_med<=12?\'var(--alert)\':\'var(--err)\';\n    return \'<div class="hbar-row">\'+\n      \'<span class="hbar-label" title="\'+r.k+\'">\'+r.k+\'</span>\'+\n      \'<div class="hbar-track"><div class="hbar-fill" style="width:\'+pct+\'%;background:\'+cor+\'"></div></div>\'+\n      \'<span class="hbar-pct" style="font-size:10px;color:\'+cor+\';font-weight:700">\'+trFmtD(r.lt_med)+\'d</span></div>\';\n  }).join(\'\')+\'</div>\';\n}\n\nfunction trRenderMes(rows) {\n  const el=document.getElementById(\'tr-chart-mes\');\n  if(!el||!rows.length) return;\n  const W=el.offsetWidth||520, H=200;\n  const PAD={top:28,right:52,bottom:36,left:42};\n  const cW=W-PAD.left-PAD.right, cH=H-PAD.top-PAD.bottom;\n  const n=rows.length;\n  const maxT=Math.max(...rows.map(r=>r.total));\n  const nsMin=85, nsMax=101;\n  const xB=i=>PAD.left+i*(cW/(n-1||1));\n  const yBar=v=>PAD.top+cH-(v/maxT)*cH;\n  const yNS=v=>PAD.top+cH-((v-nsMin)/(nsMax-nsMin))*cH;\n  const barW=Math.max(4,Math.floor(cW/n)-2);\n  const step=Math.ceil(n/8);\n\n  const grid=[0,0.25,0.5,0.75,1].map(p=>{\n    const y=(PAD.top+cH-p*cH).toFixed(1);\n    return \'<line stroke="var(--gl2)" stroke-width="0.5" stroke-dasharray="3 3" x1="\'+PAD.left+\'" y1="\'+y+\'" x2="\'+(PAD.left+cW)+\'" y2="\'+y+\'"/>\'+\n           \'<text font-size="9" fill="var(--gl)" font-family="Montserrat,sans-serif" x="\'+(PAD.left-4)+\'" y="\'+y+\'" text-anchor="end" dominant-baseline="central">\'+Math.round(p*maxT)+\'</text>\';\n  }).join(\'\');\n\n  const nsGrid=[90,95,98,100].map(v=>{\n    const y=yNS(v).toFixed(1);\n    return \'<text font-size="9" fill="var(--ac)" font-family="Montserrat,sans-serif" x="\'+(PAD.left+cW+5)+\'" y="\'+y+\'" dominant-baseline="central">\'+v+\'%</text>\';\n  }).join(\'\');\n\n  const yMeta=yNS(98).toFixed(1);\n  const metaLine=\'<line x1="\'+PAD.left+\'" y1="\'+yMeta+\'" x2="\'+(PAD.left+cW)+\'" y2="\'+yMeta+\'" stroke="#F39200" stroke-width="1" stroke-dasharray="5 3" opacity="0.8"/>\'+\n    \'<text font-size="8" fill="#F39200" font-family="Montserrat,sans-serif" x="\'+(PAD.left+cW+5)+\'" y="\'+yMeta+\'" dominant-baseline="central" font-weight="700">98%</text>\';\n\n  const bars=rows.map((r,i)=>{\n    const x=(PAD.left+i*(cW/n)+(cW/n-barW)/2).toFixed(1);\n    const y=yBar(r.total).toFixed(1);\n    const bH=(cH-(yBar(r.total)-PAD.top)).toFixed(1);\n    const cor=nsColor(r.ns);\n    const lbl=MESES_TR[r.mes]+\'/\'+String(r.ano).slice(2);\n    const showL=n<=12||i%step===0||i===n-1;\n    return \'<rect x="\'+x+\'" y="\'+y+\'" width="\'+barW+\'" height="\'+bH+\'" rx="2" fill="\'+cor+\'" opacity="0.2"/>\'+\n           (showL?\'<text font-size="8" fill="var(--gl)" font-family="Montserrat,sans-serif" text-anchor="middle" x="\'+(parseFloat(x)+barW/2).toFixed(1)+\'" y="\'+(H-6)+\'">\'+lbl+\'</text>\':\'\');\n  }).join(\'\');\n\n  const pathNS=rows.map((r,i)=>(i===0?\'M\':\'L\')+(PAD.left+i*(cW/n)+barW/2).toFixed(1)+\',\'+yNS(r.ns).toFixed(1)).join(\' \');\n  const dotsNS=rows.map((r,i)=>{\n    const cx=(PAD.left+i*(cW/n)+barW/2).toFixed(1);\n    const cy=yNS(r.ns).toFixed(1);\n    const cor=nsColor(r.ns);\n    const showL=n<=12||i%step===0||i===n-1;\n    return \'<circle cx="\'+cx+\'" cy="\'+cy+\'" r="3.5" fill="\'+cor+\'" stroke="#fff" stroke-width="1.5"/>\'+\n           (showL?\'<text font-size="8" font-weight="700" fill="\'+cor+\'" font-family="Montserrat,sans-serif" text-anchor="middle" x="\'+cx+\'" y="\'+(parseFloat(cy)-7).toFixed(1)+\'">\'+trFmtD(r.ns)+\'%</text>\':\'\');\n  }).join(\'\');\n\n  el.innerHTML=\'<svg viewBox="0 0 \'+W+\' \'+H+\'" style="width:100%;height:\'+H+\'px;overflow:visible">\'+\n    grid+nsGrid+bars+metaLine+\n    \'<path d="\'+pathNS+\'" fill="none" stroke="var(--ac)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>\'+\n    dotsNS+\'</svg>\'+\n    \'<div style="display:flex;gap:16px;justify-content:center;margin-top:6px">\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--gl);display:flex;align-items:center;gap:4px"><span style="width:12px;height:10px;background:rgba(0,61,122,0.2);display:inline-block;border-radius:2px"></span>Volume</span>\'+\n    \'<span style="font-size:10px;font-weight:600;color:var(--ac);display:flex;align-items:center;gap:4px"><span style="width:20px;height:2.5px;background:var(--ac);display:inline-block;border-radius:2px"></span>NS %</span>\'+\n    \'<span style="font-size:10px;font-weight:600;color:#F39200;display:flex;align-items:center;gap:4px"><span style="width:20px;height:1px;border-top:2px dashed #F39200;display:inline-block"></span>Meta 98%</span></div>\';\n}\n\nfunction trRenderReg(rows) {\n  const el=document.getElementById(\'tr-chart-reg\');\n  if(!el) return;\n  const maxLT=Math.max(...rows.map(r=>r.lt_med||0),1);\n  el.innerHTML=\'<div class="hbar-wrap">\'+rows.filter(r=>r.total>=3).map(r=>{\n    const pct=Math.round((r.lt_med||0)/maxLT*100);\n    const corLT=r.lt_med<=7?\'var(--ok)\':r.lt_med<=12?\'var(--alert)\':\'var(--err)\';\n    const corNS=nsColor(r.ns);\n    return \'<div class="hbar-row" style="grid-template-columns:100px 1fr 120px">\'+\n      \'<span class="hbar-label">\'+r.k+\'</span>\'+\n      \'<div class="hbar-track"><div class="hbar-fill" style="width:\'+pct+\'%;background:\'+corLT+\'"></div></div>\'+\n      \'<span style="font-size:10px;white-space:nowrap">\'+\n      \'<span style="color:\'+corLT+\';font-weight:700">\'+trFmtD(r.lt_med||0)+\'d</span>\'+\n      \' <span style="color:\'+corNS+\';font-size:9px">NS \'+trFmtD(r.ns)+\'%</span></span></div>\';\n  }).join(\'\')+\'</div>\';\n}\n\nfunction trRenderTabela(rows, tot) {\n  const tbody=document.getElementById(\'tr-tbody\');\n  const tfoot=document.getElementById(\'tr-tfoot\');\n  if(!tbody) return;\n  const sorted=[...rows].sort((a,b)=>b.total-a.total);\n  tbody.innerHTML=sorted.map((r,i)=>{\n    const cor=nsColor(r.ns);\n    const bg=r.ns>=98.8?\'#E3EEF9\':r.ns>=97?\'#FFF3E0\':\'#FFEBEE\';\n    const rankCls=i===0?\'rank-1\':i===1?\'rank-2\':i===2?\'rank-3\':\'\';\n    const badge=rankCls?\'<span class="rank-badge \'+rankCls+\'">\'+(i+1)+\'</span>\':\'<span style="display:inline-block;width:20px;text-align:center;margin-right:6px;font-size:11px;color:var(--gl)">\'+(i+1)+\'</span>\';\n    return \'<tr>\'+\n      \'<td>\'+badge+r.k+\'</td>\'+\n      \'<td>\'+fmt(r.total)+\'</td>\'+\n      \'<td style="color:var(--ok)">\'+fmt(r.prazo)+\'</td>\'+\n      \'<td style="color:\'+( r.atraso>0?\'var(--err)\':\'var(--gl)\')+\'">\'+r.atraso+\'</td>\'+\n      \'<td><span style="background:\'+bg+\';color:\'+cor+\';padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700">\'+trFmtD(r.ns)+\'%</span></td>\'+\n      \'<td style="color:\'+(r.lt_med<=7?\'var(--ok)\':r.lt_med<=12?\'var(--alert)\':\'var(--err)\')+\'">\'+trFmtD(r.lt_med||0)+\'d</td>\'+\n      \'<td>\'+trFmtR(r.frete)+\'</td>\'+\n      \'<td>\'+trFmtR2(r.c_ent)+\'</td>\'+\n      \'<td>\'+trFmtR2(r.c_kg)+\'</td>\'+\n      \'<td style="color:\'+(r.fsf<=2?\'var(--ok)\':r.fsf<=4?\'var(--alert)\':\'var(--err)\')+\'">\'+trFmtD(r.fsf)+\'%</td></tr>\';\n  }).join(\'\');\n  if(tfoot&&tot){\n    tfoot.innerHTML=\'<tr><td><strong>TOTAL GERAL</strong></td>\'+\n      \'<td><strong>\'+fmt(tot.total)+\'</strong></td>\'+\n      \'<td style="color:var(--ok)"><strong>\'+fmt(tot.prazo)+\'</strong></td>\'+\n      \'<td style="color:var(--err)"><strong>\'+tot.atraso+\'</strong></td>\'+\n      \'<td><strong>\'+trFmtD(tot.ns)+\'%</strong></td>\'+\n      \'<td><strong>\'+trFmtD(tot.lt_med)+\'d</strong></td>\'+\n      \'<td><strong>\'+trFmtR(tot.frete)+\'</strong></td>\'+\n      \'<td><strong>\'+trFmtR2(tot.c_ent)+\'</strong></td>\'+\n      \'<td><strong>\'+trFmtR2(tot.c_kg)+\'</strong></td>\'+\n      \'<td><strong>\'+trFmtD(tot.fsf)+\'%</strong></td></tr>\';\n  }\n}\n\nfunction trRenderUF(rows) {\n  const tbody=document.getElementById(\'tr-tbody-uf\');\n  if(!tbody) return;\n  const sorted=[...rows].sort((a,b)=>b.total-a.total);\n  tbody.innerHTML=sorted.map(r=>{\n    const cor=nsColor(r.ns);\n    const bg=r.ns>=98.8?\'#E3EEF9\':r.ns>=97?\'#FFF3E0\':\'#FFEBEE\';\n    return \'<tr>\'+\n      \'<td style="font-weight:700;color:var(--ac)">\'+r.k+\'</td>\'+\n      \'<td>\'+fmt(r.total)+\'</td>\'+\n      \'<td><span style="background:\'+bg+\';color:\'+cor+\';padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700">\'+trFmtD(r.ns)+\'%</span></td>\'+\n      \'<td style="color:\'+(r.lt_med<=7?\'var(--ok)\':r.lt_med<=12?\'var(--alert)\':\'var(--err)\')+\'">\'+trFmtD(r.lt_med||0)+\'d</td>\'+\n      \'<td>\'+trFmtR(r.frete)+\'</td>\'+\n      \'<td>\'+trFmtR2(r.c_ent)+\'</td>\'+\n      \'<td>\'+trFmtR2(r.c_kg)+\'</td>\'+\n      \'<td style="color:\'+(r.fsf<=2?\'var(--ok)\':r.fsf<=4?\'var(--alert)\':\'var(--err)\')+\'">\'+trFmtD(r.fsf)+\'%</td></tr>\';\n  }).join(\'\');\n}\nfunction scrollAcomp() {\n  setTimeout(() => {\n    const el = document.querySelector(\'#page-transportadoras [style*="margin-top:20px"]\');\n    if(el) el.scrollIntoView({behavior:\'smooth\', block:\'start\'});\n    // Renderizar quadros se ainda não foram\n    renderAcompanhamento2();\n  }, 100);\n}\n\nfunction renderAcompanhamento2() {\n  const hojeEl = document.getElementById(\'acomp-hoje2\');\n  if(hojeEl) {\n    const d = new Date(EP.hoje);\n    hojeEl.textContent = \'Atualizado em: \'+d.toLocaleDateString(\'pt-BR\');\n  }\n  const q1 = EP.q1.filter(r => r.da > 0 || (r.dv !== null && r.dv <= 2));\n  const q2 = EP.q2;\n\n  const tbody1 = document.getElementById(\'q1-tbody2\');\n  const tbody2 = document.getElementById(\'q2-tbody2\');\n  if(!tbody1 || !tbody2) return;\n\n  tbody1.innerHTML = q1.length ? q1.map(r => {\n    const dv=r.dv, da=r.da;\n    let tag, dvTxt;\n    if(da>0){tag=\'<span class="tag tag-atrasado">🔴 Atrasado</span>\';dvTxt=\'<span class="da-err">+\'+da+\'d</span>\';}\n    else if(dv!==null&&dv<=2){tag=\'<span class="tag tag-alerta">⚠️ Alerta</span>\';dvTxt=dv===0?\'<span class="da-warn">Hoje!</span>\':\'<span class="da-warn">\'+dv+\'d</span>\';}\n    else{tag=\'<span class="tag tag-ok">✅ No prazo</span>\';dvTxt=\'<span class="da-ok">\'+(dv!==null?dv+\'d\':\'—\')+\'</span>\';}\n    const empBadge=r.emp===\'IAB\'?\'<span class="emp-badge emp-iab">IAB</span>\':\'<span class="emp-badge emp-verbum">Verbum</span>\';\n    return \'<tr><td>\'+empBadge+\'</td><td style="font-weight:700">\'+r.nf+\'</td><td style="font-weight:600">\'+r.cli+\'</td><td>\'+r.tr+\'</td><td>\'+r.col+\'</td><td>\'+r.prev+\' \'+tag+\'</td><td style="text-align:center">\'+dvTxt+\'</td><td class="obs-cell">\'+r.obs+\'</td></tr>\';\n  }).join(\'\') : \'<tr><td colspan="8" style="text-align:center;color:var(--gl);padding:24px">Nenhuma entrega pendente. ✅</td></tr>\';\n\n  tbody2.innerHTML = q2.length ? q2.map(r => {\n    const empBadge=r.emp===\'IAB\'?\'<span class="emp-badge emp-iab">IAB</span>\':\'<span class="emp-badge emp-verbum">Verbum</span>\';\n    const daClass=r.da<=2?\'da-warn\':\'da-err\';\n    return \'<tr><td>\'+empBadge+\'</td><td style="font-weight:700">\'+r.nf.split(\' / \').reduce((a,n,i)=>i>0&&i%3===0?a+\'<br>\'+n:i>0?a+\' / \'+n:n,\'\')+\'</td><td style="font-weight:600">\'+r.cli+\'</td><td>\'+r.tr+\'</td><td>\'+r.col+\'</td><td>\'+r.prev+\'</td><td style="font-weight:700">\'+r.ef+\'</td><td style="text-align:center"><span class="\'+daClass+\'">+\'+r.da+\'d</span></td><td class="obs-cell">\'+r.obs+\'</td></tr>\';\n  }).join(\'\') : \'<tr><td colspan="9" style="text-align:center;color:var(--gl);padding:24px">Sem atrasos registrados.</td></tr>\';\n}\n\nfunction abrirEstoque() {\n  window.open(\'dashboard_estoque.html\', \'_blank\');\n}\n\n'
    _ve_init = (
        _helpers +
        "const VE = " + PAYLOAD_VE + ";\n" +
        "const MESES_VE = " + _meses_js + ";\n" +
        "let ANO_VE = 'TODOS';\n\n" +
        _js_ve + "\n\n" +
        "// ═══════════════ ABA FINANCEIRO ═══════════════\n" +
        "const FIN = " + PAYLOAD_FIN + ";\n" +
        "const MESES_FIN = " + _meses_js + ";\n" +
        _js_fin + "\n\n" +
        "// ═══════════════ ABA TRANSPORTADORAS ═══════════════\n" +
        "const TR = " + PAYLOAD_TR + ";\n" +
        "const MESES_TR = " + _meses_js + ";\n" +
        _js_tr + "\n\n"
    )
    _old_init = "// ── INIT ──\nfAno('TODOS');"
    html = html.replace(_old_init, _ve_init + _old_init + "\nsetTimeout(renderVE, 150);", 1)

    # 6. navPage
    html = html.replace(
        "  if (p === 'financeiro') { renderFinanceiro(); }",
        "  if (p === 'visao-executiva') { renderVE(); }\n  if (p === 'financeiro') { renderFinanceiro(); }"
    )
    html = html.replace(
        "  if (PAGINA === 'financeiro') renderFinanceiro();",
        "  if (PAGINA === 'visao-executiva') renderVE();\n  if (PAGINA === 'financeiro') renderFinanceiro();"
    )
    html = html.replace(
        "  if (PAGINA === 'entregas') eAtualizar();",
        "  if (PAGINA === 'entregas') eAtualizar();\n  if (PAGINA === 'visao-executiva') renderVE();"
    )
    # Adicionar _updateAbaNome e renderers no navPage
    html = html.replace(
        "function navPage(p) {\n  PAGINA = p;",
        "function navPage(p) {\n  PAGINA = p;\n  _updateAbaNome(p);"
    )
    # Adicionar renderers para cada aba no navPage
    html = html.replace(
        "  if (p === 'entregas') { eAtualizar(); }\n  if (p === 'acompanhamento') { renderAcompanhamento(); }",
        "  if (p === 'visao-executiva') { renderVE(); }\n  if (p === 'financeiro') { renderFinanceiro(); }\n  if (p === 'transportadoras') { renderTransportadoras(); }\n  if (p === 'entregas') { eAtualizar(); }\n  if (p === 'acompanhamento') { renderAcompanhamento(); }"
    )
    # Inserir JS da aba antes do INIT
    html = html.replace(
        "// ── INIT ──\nfAno('TODOS');",
        _js_aba + "// ── INIT ──\nfAno('TODOS');\n_updateAbaNome(PAGINA);"
    )
    # 6b. Corrigir fEmp para renderizar aba ativa
    _femp_old = "  if (PAGINA === 'entregas') eAtualizar();\n}"
    _femp_new = "  if (PAGINA === 'entregas') eAtualizar();\n  if (PAGINA === 'visao-executiva') renderVE();\n}"
    _femp_old2 = "  if (PAGINA === 'entregas') eAtualizar();\n  if (PAGINA === 'visao-executiva') renderVE();\n}"
    _femp_new2 = "  if (PAGINA === 'entregas') eAtualizar();\n  if (PAGINA === 'visao-executiva') renderVE();\n  if (PAGINA === 'financeiro') renderFinanceiro();\n  if (PAGINA === 'transportadoras') renderTransportadoras();\n}"
    html = html.replace(_femp_old, _femp_new2)
    html = html.replace(_femp_old2, _femp_new2)


    # 7. Guards
    _ids = set(_re.findall(r'id="([^"]+)"', html))
    _lns = html.split("\n"); _out = []
    for _l in _lns:
        _m = _re.match(r'^(\s*)document\.getElementById\(\'([^\']+)\'\)\.textContent\s*=\s*(.+)', _l)
        if _m and _m.group(2) not in _ids:
            _l = _m.group(1) + "_st('" + _m.group(2) + "', " + _m.group(3).rstrip(";") + ");"
        _out.append(_l)
    html = "\n".join(_out)
    for _o, _n in [
        ("function renderKPI(d) {\n  document.getElementById('k1')",
         "function renderKPI(d) {\n  if(!document.getElementById('k1')) return;\n  document.getElementById('k1')"),
        ("function renderKPI(d) {\n  _st('k1'",
         "function renderKPI(d) {\n  if(!document.getElementById('k1')) return;\n  _st('k1'"),
        ("function renderTransp(tr) {\n  const top=tr.slice(0,8)",
         "function renderTransp(tr) {\n  if(!document.getElementById('cht')) return;\n  const top=tr.slice(0,8)"),
        ("function renderTabela(tr) {\n  document.getElementById('ttransp').innerHTML",
         "function renderTabela(tr) {\n  if(!document.getElementById('ttransp')) return;\n  document.getElementById('ttransp').innerHTML"),
        ("function eAtualizar() {\n  const d = getDadosEnt();",
         "function eAtualizar() {\n  if(!document.getElementById('ek1')) return;\n  const d = getDadosEnt();"),
        ("    const el = document.getElementById(selId);\n    const label = el.options[0].text;",
         "    const el = document.getElementById(selId);\n    if(!el) return;\n    const label = el.options[0].text;"),
        ("    document.getElementById('ab'+a).classList.toggle('on',a===ano);",
         "    const _ab=document.getElementById('ab'+a); if(_ab) _ab.classList.toggle('on',a===ano);"),
        ("  const arc = document.getElementById('garc');\n  arc.style.strokeDashoffset",
         "  const arc = document.getElementById('garc');\n  if(arc) arc.style.strokeDashoffset"),
        ("  const arc = document.getElementById('garc');\n  if(arc) arc.style.strokeDashoffset = 220-(220*pct/100);\n  // Mesma escala das transportadoras\n  const pctNum",
         "  const arc = document.getElementById('garc');\n  if(arc) arc.style.strokeDashoffset = 220-(220*pct/100);\n  // Mesma escala das transportadoras\n  const pctNum"),
        ("  arc.style.stroke = gaugeColor;",
         "  if(arc) arc.style.stroke = gaugeColor;"),
        ("  const ek4 = document.getElementById('ek4');\n  ek4.style.color",
         "  const ek4 = document.getElementById('ek4');\n  if(ek4) ek4.style.color"),
    ]:
        if _o in html: html = html.replace(_o, _n)


    # 7b. Guard VE sem vendas
    html = html.replace(
        '  const t = d.total;\n  if(!t) return;',
        '  const t = d.total;\n  if(!t) return;\n  if(!t.fat_b || t.fat_b === 0) {\n    const veK2 = document.getElementById(\'ve-kpi2\');\n    if(veK2) veK2.innerHTML = \'<div style="grid-column:1/-1;background:#FFF9E6;border:1.5px solid #FFE082;border-radius:12px;padding:20px;text-align:center;color:#E65100;font-weight:600">⚠️ Sem registros de venda para este período.</div>\';\n  }'
    )

    # 8. Deduplicar páginas e garantir ordem correta
    import re as _re_dedup
    _pages_m = list(_re_dedup.finditer(r'<div id="page-([^"]+)"', html))
    _all_s = [m.start() for m in _pages_m]
    _idx_scr = html.find("\n<script>")
    _vistas = {}; _pages_u = {}
    for i, m in enumerate(_pages_m):
        _pid = m.group(1)
        if _pid not in _vistas:
            _vistas[_pid] = m.start()
            _nxt = [s for s in _all_s if s > m.start()]
            _end = min(_nxt) if _nxt else _idx_scr
            _pages_u[_pid] = (m.start(), _end)
    _idx_fp = min(s for s,e in _pages_u.values())
    _html_antes = html[:_idx_fp]
    _ordem = ["visao-executiva","financeiro","transportadoras","acompanhamento"]
    _html_pgs = ""
    for _pid in _ordem:
        if _pid in _pages_u:
            _s, _e = _pages_u[_pid]
            _html_pgs += html[_s:_e].rstrip() + "\n\n"
    html = _html_antes + _html_pgs + html[_idx_scr:]







    # PATCH: Q1 — adicionar colunas Cidade e UF na linha do tbody
    html = html.replace(
        '<td style="font-weight:600;min-width:160px">${r.cli}</td>\\n          <td style="white-space:nowrap">${r.tr}</td>',
        '<td style="font-weight:600;min-width:160px">${r.cli}</td>\\n          <td style="white-space:nowrap;max-width:140px;overflow:hidden;text-overflow:ellipsis">${r.mun||\'—\'}</td>\\n          <td style="font-weight:700;color:var(--ac);text-align:center">${r.uf||\'—\'}</td>\\n          <td style="white-space:nowrap">${r.tr}</td>',
        1
    )





    html = html.replace(
        'fk(\'fd\',\'🎁\',\'Bonificações\',        fmtM(t.bon), \'<span style="white-space:nowrap">\'+pctPill(t.pct_bon,\'warn\')+\' · Fat.Bruto</span>\', \'\') +\n    fk(\'fd\',\'🎁\',\'Bonificações\', fmtM(t.bon), \'<span style="white-space:nowrap">\'+ pctPill(t.pct_bon,\'warn\')+\' · Fat.Bruto</span>\',\'\')',
        'fk(\'fd\',\'🎁\',\'Bonificações\', fmtM(t.bon), \'<span style="white-space:nowrap">\'+pctPill(t.pct_bon,\'warn\')+\' · Fat.Bruto</span>\',\'\')',
        1
    )

    html = html.replace(
        'function renderVE() {\n  const d = getVED();\n  const t = d.total;\n  if(!t) return;',
        'function renderVE() {\n  const d = getVED();\n  const t = d.total;\n  if(!t) {\n    [\'ve-kpi1\',\'ve-kpi2\',\'ve-alertas\',\'ve-chart-evol\',\'ve-chart-reg\',\'ve-tbody\'].forEach(id=>{\n      const el=document.getElementById(id); if(el) el.innerHTML=\'\';\n    });\n    const _vk=document.getElementById(\'ve-kpi1\'); if(_vk) _vk.innerHTML=\'<div style="grid-column:1/-1;text-align:center;padding:40px;color:#999;font-size:14px;font-weight:600;background:#fafafa;border-radius:12px;border:1.5px dashed #ddd">Sem dados para o período selecionado.</div>\';\n    return;\n  }',
        1
    )
    html = html.replace(
        'function renderFinanceiro() {\n  const d = getFinDados();\n  if(!d || !d.total) return;',
        'function renderFinanceiro() {\n  const d = getFinDados();\n  if(!d || !d.total) {\n    [\'fin-kpi-fat\',\'fin-kpi-custo\',\'fin-kpi-efic\',\'fin-chart-mes\',\'fin-chart-reg\',\'fin-chart-tr\',\'fin-chart-liq\',\'fin-tbody\',\'fin-tfoot\'].forEach(id=>{\n      const el=document.getElementById(id); if(el) el.innerHTML=\'\';\n    });\n    const _fk=document.getElementById(\'fin-kpi-fat\'); if(_fk) _fk.innerHTML=\'<div style="grid-column:1/-1;text-align:center;padding:40px;color:#999;font-size:14px;font-weight:600;background:#fafafa;border-radius:12px;border:1.5px dashed #ddd">Sem dados para o período selecionado.</div>\';\n    return;\n  }',
        1
    )

    html = html.replace(
        'function renderTransportadoras() {\n  const d = getTrDados();\n  if(!d || !d.total) return;',
        'function renderTransportadoras() {\n  const d = getTrDados();\n  if(!d || !d.total) {\n    [\'tr-kpi-op\',\'tr-kpi-custo\',\'tr-chart-ns\',\'tr-chart-lt\',\'tr-chart-mes\',\'tr-chart-reg\',\'tr-tbody-uf\',\'tr-tbody\',\'tr-tfoot\'].forEach(id=>{\n      const el=document.getElementById(id); if(el) el.innerHTML=\'\';\n    });\n    const _tk=document.getElementById(\'tr-kpi-op\'); if(_tk) _tk.innerHTML=\'<div style="grid-column:1/-1;text-align:center;padding:40px;color:#999;font-size:14px;font-weight:600;background:#fafafa;border-radius:12px;border:1.5px dashed #ddd">Sem dados para o período selecionado.</div>\';\n    return;\n  }',
        1
    )

    html = html.replace(
        "function renderFinMes(rows) {\n  const el=document.getElementById('fin-chart-mes');\n  if(!el||!rows.length) return;",
        "function renderFinMes(rows) {\n  const el=document.getElementById('fin-chart-mes');\n  if(!el) return;\n  if(!rows.length) { el.innerHTML=''; return; }",
        1
    )

    # PATCH: corrigir % duplicado nos cards Devoluções e Bonificações
    html = html.replace("pctPill(t.pct_dev,'neg')+'% · Fat.Bruto'", "pctPill(t.pct_dev,'neg')+' · Fat.Bruto'")
    html = html.replace("pctPill(t.pct_bon,'warn')+'% · Fat.Bruto'", "pctPill(t.pct_bon,'warn')+' · Fat.Bruto'")
    # 9. Remover placeholders restantes
    import re as _re3
    html = _re3.sub(r'___[A-Z_]+___', 'null', html)


    with open(DASH_OUT, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(DASH_OUT) / 1024
    print(f"\n{'='*55}")
    print(f"  ✅ DASHBOARD ATUALIZADO COM SUCESSO")
    print(f"  📁 {DASH_OUT}")
    print(f"  📦 {size_kb:.1f} KB")
    print(f"  🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"{'='*55}")


if __name__ == "__main__":
    gerar()
