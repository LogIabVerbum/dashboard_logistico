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
IMG_DIR     = PASTA_DASHBOARDS
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
        candidatos = [
            path,
            os.path.join(BASE_DIR, "..", os.path.basename(path)),
            os.path.join(BASE_DIR, os.path.basename(path)),
            os.path.join(os.path.dirname(DASH_DIR), os.path.basename(path)),
        ]
        encontrado = None
        for c in candidatos:
            if os.path.exists(c):
                encontrado = c
                break
        if not encontrado:
            print(f"      AVISO: {os.path.basename(path)} não encontrado em nenhum caminho.")
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

    html = (html
        .replace("___LI___", LI)
        .replace("___LV___", LV)
        .replace("___CI___", CI)
        .replace("___CV___", CV)
        .replace("___PAYLOAD_GER___", PAYLOAD_GER)
        .replace("___PAYLOAD_ENT___", PAYLOAD_ENT)
        .replace("___ANO_BTNS___", ano_btns)
        .replace("___EANO_BTNS___", eano_btns)
        .replace("___GERADO_EM___", datetime.now().strftime("%d/%m/%Y %H:%M"))
    )

    # ── 7. VALIDAR E SALVAR ──
    print("[ 7 ] Validando e salvando dashboard.html...")
    import re
    placeholders_restantes = re.findall(r'___[A-Z_]+___', html)
    if placeholders_restantes:
        print(f"ERRO CRÍTICO: Placeholders não substituídos: {set(placeholders_restantes)}")
        print("Verifique se o dashboard_template.html esta na pasta etl/")
        sys.exit(1)
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
