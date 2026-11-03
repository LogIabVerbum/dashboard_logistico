# =============================================================================
# gerar_dashboard.py  — Gerador do Dashboard Logístico IAB & Verbum
# Versão limpa: uma única função gerar(), zero pós-processamento,
# todos os dados injetados via placeholders no template.
# =============================================================================
import sys, os, json, math, base64, io
import pandas as pd
import numpy as np
from datetime import datetime
from PIL import Image

# ── CAMINHOS ──────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _achar_onedrive():
    candidatos = [
        os.environ.get("OneDriveCommercial", ""),
        os.environ.get("OneDrive", ""),
        os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive - Instituto Alfa e Beto"),
        os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive"),
    ]
    subfixos = [
        os.path.join("Logística-IAB",  "10 - DASHIBOARDs"),
        os.path.join("Logistica-IAB",  "10 - DASHIBOARDs"),
        "10 - DASHIBOARDs",
    ]
    for base in candidatos:
        if not base:
            continue
        for sub in subfixos:
            pasta = os.path.join(base, sub)
            if os.path.isdir(pasta):
                return pasta
    # fallback: etl/../../ = 10-DASHIBOARDs/
    return os.path.normpath(os.path.join(BASE_DIR, "..", ".."))

PASTA_DASHBOARDS = _achar_onedrive()
PARQUET  = os.path.join(BASE_DIR, "data", "processed", "lancamentos.parquet")
DASH_OUT = os.path.join(PASTA_DASHBOARDS, "dashboard_logistico", "dashboard.html")
DASH_DIR = os.path.dirname(os.path.abspath(DASH_OUT))

def _achar_img(nomes, qualidade=85):
    """Retorna (caminho, qualidade) para o primeiro arquivo encontrado na lista de nomes."""
    for nome in nomes:
        caminho = os.path.join(DASH_DIR, nome)
        if os.path.exists(caminho):
            return (caminho, qualidade)
    return (os.path.join(DASH_DIR, nomes[0]), qualidade)  # fallback: primeiro nome

IMGS = {
    # (caminho, max_largura, qualidade)  — max_largura é o 2º campo usado em webp()
    "LI": (*_achar_img(["logo_iab.png", "logo_iab.PNG"]),          ),
    "LV": (*_achar_img(["logo_verbum.png", "logo_verbum.PNG"]),     ),
    "LS": (*_achar_img(["logo_serena.png", "logo_serena.PNG"]),     ),
    "CI": (*_achar_img(["Carreta_iab.png", "Carreta_bem_vindo.png", "Carreta_IAB.png"]),  ),
    "CV": (*_achar_img(["Carreta_verbum.png", "Carreta_Verbum.png"]),                     ),
    "BG": (*_achar_img(["imagem_navegador_paginainicial.png", "tela_boas_vindas.jpg",
                         "imagem_navegador_paginainicial.PNG"]),     ),
}
# Ajustar max_largura e qualidade conforme o tipo de imagem
IMGS["LI"] = (IMGS["LI"][0], 260, 85)
IMGS["LV"] = (IMGS["LV"][0], 220, 85)
IMGS["LS"] = (IMGS["LS"][0], 220, 85)
IMGS["CI"] = (IMGS["CI"][0], 1920, 100)
IMGS["CV"] = (IMGS["CV"][0], 380, 78)
IMGS["BG"] = (IMGS["BG"][0], 1920, 85)

# ── ARQUIVO DE ESTOQUE (TI — atualizado diariamente) ─────────────────────────
def _achar_estoque():
    """Localiza Movimentação de Estoque.xlsx na pasta do TI (OneDrive)."""
    candidatos_base = [
        os.environ.get("OneDriveCommercial", ""),
        os.environ.get("OneDrive", ""),
        os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive - Instituto Alfa e Beto"),
        os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive"),
    ]
    subfixos = [
        os.path.join("Logística-IAB", "10 - DASHIBOARDs",
                     "Arquivos de Claude Relatórios - Relatórios Logística",
                     "Movimentação de Estoque.xlsx"),
        os.path.join("Logistica-IAB", "10 - DASHIBOARDs",
                     "Arquivos de Claude Relatórios - Relatórios Logística",
                     "Movimentação de Estoque.xlsx"),
        os.path.join("Logística-IAB", "10 - DASHIBOARDs",
                     "Movimentação de Estoque.xlsx"),
    ]
    for base in candidatos_base:
        if not base:
            continue
        for sub in subfixos:
            caminho = os.path.join(base, sub)
            if os.path.isfile(caminho):
                return caminho
    # Fallback: pasta do script
    return os.path.join(BASE_DIR, "Movimentação de Estoque.xlsx")

ESTOQUE_XLSX = _achar_estoque()

# ── QUADRO DE ENVIOS (aba "Base Distâncias": km de Uberlândia até cada município) ──
def _achar_quadro():
    for c in [os.path.join(PASTA_DASHBOARDS, "Quadro de Envios.xlsx"),
              os.path.join(BASE_DIR, "Quadro de Envios.xlsx"),
              os.path.normpath(os.path.join(BASE_DIR, "..", "Quadro de Envios.xlsx")),
              os.path.normpath(os.path.join(BASE_DIR, "..", "..", "Quadro de Envios.xlsx"))]:
        if os.path.isfile(c):
            return c
    return None

def _chave_mun(mun, uf):
    import unicodedata as _ud
    t = lambda v: _ud.normalize("NFD", str(v or "")).encode("ascii", "ignore").decode().upper().strip()
    return t(mun) + "|" + t(uf)

def _completar_distancias(df):
    """Preenche distancia_km (zerada/ausente) pelo município+UF, usando a aba 'Base Distâncias'
    do Quadro de Envios e, se faltar, a coluna 'Distância (km)' da aba Lançamentos."""
    import warnings
    q = _achar_quadro()
    if not q:
        print("      AVISO: Quadro de Envios.xlsx não encontrado — km não pôde ser completado.")
        return df
    mapa = {}
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            lanc = pd.read_excel(q, sheet_name="Lançamentos", usecols=lambda c: str(c).strip() in ("Municipio", "UF", "Distância (km)"))
            lanc = lanc.dropna(subset=["Distância (km)"])
            for m, u, d in zip(lanc["Municipio"], lanc["UF"], lanc["Distância (km)"]):
                v = pd.to_numeric(d, errors="coerce")
                if pd.notna(v) and v > 0: mapa[_chave_mun(m, u)] = float(v)
        except Exception as e:
            print(f"      AVISO: não li 'Distância (km)' da aba Lançamentos ({e})")
        try:
            base = pd.read_excel(q, sheet_name="Base Distâncias")
            cm = next(c for c in base.columns if str(c).upper().startswith("MUNIC"))
            cd = next(c for c in base.columns if "DIST" in str(c).upper())
            for m, u, d in zip(base[cm], base["UF"], base[cd]):
                v = pd.to_numeric(d, errors="coerce")
                if pd.notna(v) and v > 0: mapa[_chave_mun(m, u)] = float(v)
        except Exception as e:
            print(f"      AVISO: não li a aba 'Base Distâncias' ({e})")
    if "distancia_km" not in df.columns:
        df["distancia_km"] = 0.0
    df["distancia_km"] = pd.to_numeric(df["distancia_km"], errors="coerce").fillna(0.0)
    falta = df["distancia_km"] <= 0
    novos = [mapa.get(_chave_mun(m, u), 0.0) for m, u in zip(df.loc[falta, "municipio"], df.loc[falta, "uf"])]
    df.loc[falta, "distancia_km"] = novos
    sem = df.loc[df["distancia_km"] <= 0]
    print(f"      distancia_km: {int(falta.sum())} linhas completadas pelo município | "
          f"{len(sem)} ainda sem km (município fora da Base Distâncias)")
    return df

# ── HELPERS ───────────────────────────────────────────────────────────────────
def webp(path, max_w, q=78):
    img = Image.open(path).convert("RGBA")
    w, h = img.size
    if w > max_w:
        img = img.resize((max_w, int(h * max_w / w)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="WEBP", quality=q)
    return "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()

def sf(v, default="—"):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return default
    s = str(v).strip()
    return s if s else default

def fmt_d(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    try:
        return pd.Timestamp(v).strftime("%d/%m/%Y")
    except Exception:
        return "—"

class _NpEnc(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, np.integer):  return int(o)
        if isinstance(o, np.floating): return None if (np.isnan(o) or np.isinf(o)) else float(o)
        return super().default(o)
    def iterencode(self, obj, _one_shot=False):
        # Substitui NaN/Infinity por null antes de serializar
        def _fix(v):
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return None
            if isinstance(v, dict):  return {k: _fix(vv) for k, vv in v.items()}
            if isinstance(v, list):  return [_fix(vv) for vv in v]
            return v
        return super().iterencode(_fix(obj), _one_shot)

def _jdump(obj):
    return json.dumps(obj, cls=_NpEnc, ensure_ascii=False, separators=(",", ":"))


# =============================================================================
# ESTOQUE — processa Movimentação de Estoque.xlsx e gera payload JSON
# =============================================================================
def _processar_estoque(xlsx_path, jdump):
    import warnings
    from collections import defaultdict

    MESES_LABEL = ["","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

    def norm_emp(v):
        s = str(v or "").upper()
        if "ALFA" in s or "IAB" in s:    return "IAB"
        if "VERBUM" in s:                return "VERBUM"
        return "IAB"

    if not os.path.isfile(xlsx_path):
        print(f"      AVISO: arquivo de estoque não encontrado: {xlsx_path}")
        print("      Usando payload de estoque vazio.")
        empty = {"kpis":{}, "prods":[], "monthly":[]}
        return jdump({"IAB": empty, "VERBUM": empty,
                      "ultima_atualizacao": "—", "dias_historico": 0})

    print(f"      Lendo: {xlsx_path}")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(xlsx_path, sheet_name="Dados")

    # Renomear colunas para nomes padronizados
    df.columns = [str(c).strip() for c in df.columns]
    col_map = {
        "EMPRESA_ESTOQUE":   "empresa",
        "CODIGO_PRODUTO":    "cod",
        "DESCRICAO_PRODUTO": "desc",
        "DATA_MOVIMENTO":    "data",
        "TIPO_MOVIMENTO":    "tipo",
        "NATUREZA_TIPO_USO": "natureza",
        "QUANTIDADE":        "qty",
        "CUSTO_UNITARIO":    "cu",
        "CUSTO_TOTAL":       "custo_total",
        "SALDO_APOS_MOVIMENTO": "saldo_apos",
        "UNIDADE":           "unidade",
    }
    df.rename(columns=col_map, inplace=True)
    required = list(col_map.values())
    for c in required:
        if c not in df.columns:
            df[c] = None

    df["empresa"] = df["empresa"].apply(norm_emp)
    df["cod"]     = pd.to_numeric(df["cod"],     errors="coerce").fillna(0).astype(int).astype(str)
    df["qty"]     = pd.to_numeric(df["qty"],     errors="coerce").fillna(0)
    df["cu"]      = pd.to_numeric(df["cu"],      errors="coerce").fillna(0)
    df["custo_total"] = pd.to_numeric(df["custo_total"], errors="coerce").fillna(0)
    df["saldo_apos"]  = pd.to_numeric(df["saldo_apos"],  errors="coerce")
    df["data"]        = pd.to_datetime(df["data"], errors="coerce")
    df["unidade"]     = df["unidade"].fillna("UN").astype(str)
    df.dropna(subset=["data"], inplace=True)

    if df.empty:
        print("      AVISO: nenhum dado válido no arquivo de estoque.")
        empty = {"kpis":{}, "prods":[], "monthly":[]}
        return jdump({"IAB": empty, "VERBUM": empty,
                      "ultima_atualizacao": "—", "dias_historico": 0})

    datas = df["data"].dropna()
    dt_min, dt_max = datas.min(), datas.max()
    days_span = max((dt_max - dt_min).days, 1)
    ultima_atualizacao = dt_max.strftime("%d/%m/%Y")

    # ── saldo mais recente por produto por empresa ──
    df_sort = df.sort_values("data")
    latest_mask = df_sort.dropna(subset=["saldo_apos"]).groupby(["empresa","cod"])["data"].idxmax()
    df_latest = df_sort.loc[latest_mask.values][["empresa","cod","saldo_apos","cu","desc","unidade"]].copy()
    df_latest.rename(columns={"saldo_apos":"qty_atual"}, inplace=True)

    # ── movimentos por produto ──
    df_e = df[df["tipo"].str.upper().str.strip() == "ENTRADA"].copy()
    df_s = df[df["tipo"].str.upper().str.strip().str.startswith("SAÍ")].copy()
    if df_s.empty:
        df_s = df[df["tipo"].str.upper().str.strip().str.startswith("SAI")].copy()

    ent_qty = df_e.groupby(["empresa","cod"])["qty"].sum().rename("ent_qty")
    sai_qty = df_s.groupby(["empresa","cod"])["qty"].sum().rename("sai_qty")
    sai_val = df_s.groupby(["empresa","cod"])["custo_total"].sum().rename("sai_val")

    base = df_latest.set_index(["empresa","cod"])
    for s in [ent_qty, sai_qty, sai_val]:
        base = base.join(s, how="left")
    base = base.fillna(0).reset_index()

    # ── construir produto list por empresa ──
    def build_prods(emp):
        sub = base[base["empresa"] == emp].copy()
        sub["val"] = sub["qty_atual"].clip(lower=0) * sub["cu"]
        sub = sub[sub["val"] > 0].copy()
        sub = sub.sort_values("val", ascending=False).reset_index(drop=True)

        # DOH
        sub["avg_daily"] = sub["sai_qty"] / days_span
        sub["doh"] = sub.apply(
            lambda r: int(min(round(r["qty_atual"] / r["avg_daily"]), 9999))
            if r["avg_daily"] > 0 and r["qty_atual"] > 0 else 9999, axis=1)
        sub["giro"] = sub.apply(
            lambda r: round(r["sai_qty"] / r["qty_atual"], 2)
            if r["qty_atual"] > 0 and r["sai_qty"] > 0 else 0.0, axis=1)

        # ABC por custo
        total_val = sub["val"].sum()
        sub["acc"] = sub["val"].cumsum()
        sub["pct_acc"] = sub["acc"] / total_val * 100 if total_val else 0
        sub["abc"] = sub["pct_acc"].apply(lambda p: "A" if p <= 80 else ("B" if p <= 95 else "C"))

        prods = []
        for _, r in sub.iterrows():
            prods.append({
                "cod":  str(r["cod"]),
                "desc": str(r["desc"])[:60],
                "un":   str(r["unidade"]),
                "qty":  round(float(r["qty_atual"]), 2),
                "cu":   round(float(r["cu"]), 4),
                "val":  round(float(r["val"]), 2),
                "ent":  round(float(r["ent_qty"]), 2),
                "sai":  round(float(r["sai_qty"]), 2),
                "sai_v":round(float(r["sai_val"]), 2),
                "doh":  int(r["doh"]),
                "giro": round(float(r["giro"]), 2),
                "abc":  str(r["abc"]),
            })
        return prods, float(total_val)

    def build_kpis(prods, total_val):
        ativos  = [p for p in prods if p["qty"] > 0]
        parados = [p for p in prods if p["sai"] == 0 and p["qty"] > 0]
        abc_a   = [p for p in prods if p["abc"] == "A"]
        abc_b   = [p for p in prods if p["abc"] == "B"]
        abc_c   = [p for p in prods if p["abc"] == "C"]
        giro_it = [p for p in prods if p["giro"] > 0]
        doh_it  = [p for p in prods if 0 < p["doh"] < 9999]
        giro_med = (sum(p["giro"] * p["val"] for p in giro_it) / sum(p["val"] for p in giro_it)
                    if giro_it else 0.0)
        doh_med  = (sum(min(p["doh"],365) * p["val"] for p in doh_it) / sum(p["val"] for p in doh_it)
                    if doh_it else 0.0)
        return {
            "val_custo": round(total_val, 2),
            "n_skus":    len(ativos),
            "n_parados": len(parados),
            "tot_qty":   round(sum(p["qty"] for p in prods), 0),
            "tot_ent":   round(sum(p["ent"] for p in prods), 0),
            "tot_sai":   round(sum(p["sai"] for p in prods), 0),
            "giro_med":  round(giro_med, 2),
            "doh_med":   round(doh_med, 0),
            "abc_a":     len(abc_a),
            "abc_b":     len(abc_b),
            "abc_c":     len(abc_c),
            "val_abc_a": round(sum(p["val"] for p in abc_a), 2),
        }

    def build_monthly(emp):
        sub_e = df_e[df_e["empresa"] == emp].copy()
        sub_s = df_s[df_s["empresa"] == emp].copy()
        months = sorted(set(
            [(r.year, r.month) for r in sub_e["data"].dropna()] +
            [(r.year, r.month) for r in sub_s["data"].dropna()]
        ))
        result = []
        for yr, mo in months:
            e_q = float(sub_e[(sub_e["data"].dt.year==yr) & (sub_e["data"].dt.month==mo)]["qty"].sum())
            e_v = float(sub_e[(sub_e["data"].dt.year==yr) & (sub_e["data"].dt.month==mo)]["custo_total"].sum())
            s_q = float(sub_s[(sub_s["data"].dt.year==yr) & (sub_s["data"].dt.month==mo)]["qty"].sum())
            s_v = float(sub_s[(sub_s["data"].dt.year==yr) & (sub_s["data"].dt.month==mo)]["custo_total"].sum())
            result.append({
                "mes":     MESES_LABEL[mo] + "/" + str(yr)[2:],
                "ent_qty": round(e_q, 0),
                "ent_val": round(e_v, 2),
                "sai_qty": round(s_q, 0),
                "sai_val": round(s_v, 2),
            })
        return result

    payload = {"ultima_atualizacao": ultima_atualizacao, "dias_historico": days_span}
    for emp in ["IAB", "VERBUM"]:
        prods, total_val = build_prods(emp)
        payload[emp] = {
            "kpis":    build_kpis(prods, total_val),
            "prods":   prods,
            "monthly": build_monthly(emp),
        }
        print(f"      {emp}: {len(prods)} SKUs · R$ {total_val:,.0f}")

    return jdump(payload)


# ── FUNÇÃO PRINCIPAL ──────────────────────────────────────────────────────────
def gerar():
    print("=" * 60)
    print("  GERANDO DASHBOARD LOGÍSTICO")
    print("=" * 60)

    # 1. PARQUET
    print("[ 1 ] Carregando dados...")
    if not os.path.exists(PARQUET):
        print("ERRO: lancamentos.parquet não encontrado. Rode o ETL primeiro.")
        sys.exit(1)
    df   = pd.read_parquet(PARQUET)
    hoje = pd.Timestamp.now().normalize()
    print(f"      {len(df)} registros.")

    # 2. IMAGENS
    print("[ 2 ] Processando imagens...")
    imgs = {}
    for key, (path, max_w, q) in IMGS.items():
        nome = os.path.basename(path)
        candidatos = [path,
                      os.path.join(DASH_DIR, nome),
                      os.path.join(os.path.dirname(DASH_DIR), nome)]
        found = next((c for c in candidatos if os.path.exists(c)), None)
        if found:
            imgs[key] = webp(found, max_w, q)
        else:
            print(f"      AVISO: {nome} não encontrado.")
            imgs[key] = ""
    LI = imgs["LI"]; LV = imgs["LV"]
    CI = imgs["CI"]; CV = imgs["CV"]
    BG = imgs.get("BG", ""); LS = imgs.get("LS", "")

    # 3. PRÉ-PROCESSAMENTO
    print("[ 3 ] Pré-processando...")
    df["regiao"] = df["regiao"].fillna("").str.strip().str.upper()
    # Garantir coluna distancia_km — busca robusta por variações de nome
    if "distancia_km" not in df.columns:
        import unicodedata as _ud
        def _norm_col(s):
            return _ud.normalize("NFD", str(s)).encode("ascii","ignore").decode().lower()
        # Primeiro: nomes conhecidos (ordem de prioridade)
        _nomes_dist = ["Distancia (km)", "distancia (km)", "Distancia(km)",
                       "distancia km", "DISTANCIA (KM)"]
        _found = next((c for c in _nomes_dist if c in df.columns), None)
        # Fallback: qualquer coluna cujo nome normalizado contenha "dist" e "km"
        if _found is None:
            _found = next((c for c in df.columns
                           if "dist" in _norm_col(c) and "km" in _norm_col(c)), None)
        if _found:
            df["distancia_km"] = pd.to_numeric(df[_found], errors="coerce").fillna(0.0)
            print(f"      distancia_km mapeada de: '{_found}' | max={df['distancia_km'].max():.0f} km")
        else:
            df["distancia_km"] = 0.0
            print("      AVISO: coluna de distancia nao encontrada no parquet.")
    else:
        df["distancia_km"] = pd.to_numeric(df["distancia_km"], errors="coerce").fillna(0.0)
    if (df["distancia_km"] <= 0).any():
        df = _completar_distancias(df)
    # Ano Ref sempre como texto ("2027"), e lista de anos tirada dos próprios dados
    df["ano_ref"] = df["ano_ref"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df.loc[df["ano_ref"].isin(["", "nan", "None", "NaT"]), "ano_ref"] = None
    ANOS = sorted(df["ano_ref"].dropna().unique().tolist())
    ANO_ATUAL = ANOS[-1] if ANOS else str(hoje.year)
    ANO_ANT = str(int(ANO_ATUAL) - 1)
    print(f"      Anos Ref: {', '.join(ANOS)} (atual: {ANO_ATUAL})")
    df_pri = df[df["is_primeira_linha_entrega"] == 1].copy()
    df_pri["obs"] = df_pri["obs"].fillna("—")
    df_pri["regiao"] = df_pri["regiao"].fillna("").str.strip().str.upper()

    nfs_map = (df.groupby("id_entrega")["nf"]
                 .apply(lambda x: " / ".join(
                     str(int(v)) for v in sorted(x.dropna().unique())))
                 .to_dict())

    # 4. PAYLOAD GER (dados comprimidos para tabela Acompanhamento)
    print("[ 4 ] Payload GER...")
    emps = sorted(df_pri["empresa"].dropna().unique().tolist())
    anos = sorted(df_pri["ano_ref"].dropna().unique().tolist())
    tcs  = sorted(df_pri["tipo_cliente"].dropna().unique().tolist())
    nats = sorted(df_pri["natureza_operacao"].dropna().unique().tolist())
    trs  = sorted(df_pri["transportadora_norm"].dropna().unique().tolist())
    emp_i = {v: i for i, v in enumerate(emps)}
    ano_i = {v: i for i, v in enumerate(anos)}
    tc_i  = {v: i for i, v in enumerate(tcs)}
    nat_i = {v: i for i, v in enumerate(nats)}
    tr_i  = {v: i for i, v in enumerate(trs)}

    rows_comp = []
    for _, r in df.iterrows():
        rows_comp.append([
            emp_i.get(r["empresa"], 0),
            ano_i.get(str(r["ano_ref"]), 0),
            tc_i.get(r["tipo_cliente"], 0),
            nat_i.get(r["natureza_operacao"], 0),
            tr_i.get(r["transportadora_norm"], 0),
            round(float(r["vr_nf"]), 2),
            round(float(r["frete"]), 2),
            round(float(r["custo_embalagem"]), 2),
            round(float(r["total_impostos_calc"]), 2),
            int(r["is_primeira_linha_entrega"]),
            int(r["entrega_no_prazo"]),
            int(r["entrega_em_atraso"]),
        ])
    PAYLOAD_GER = _jdump({"E": emps, "A": anos, "T": tcs,
                           "N": nats, "R": trs, "D": rows_comp})

    # 5. PAYLOAD ENT (entregas pendentes e atrasos)
    print("[ 5 ] Payload ENT...")
    da_list, dv_list = [], []
    for _, row in df_pri.iterrows():
        prev = row["prev_entrega_cliente"]
        efet = row["data_efetiva_entrega"]
        try:
            if pd.isna(prev):
                da_list.append(0); dv_list.append(None); continue
            ref = efet if not pd.isna(efet) else hoje
            da_list.append(max(0, int((ref - prev).days)))
            dv_list.append(int((prev - hoje).days) if pd.isna(efet) else None)
        except Exception:
            da_list.append(0); dv_list.append(None)
    df_pri["_da"]   = da_list
    df_pri["_dv"]   = dv_list
    df_pri["nfs_ag"] = df_pri["id_entrega"].map(nfs_map)

    def rdict(row, c2=False):
        nf_str = sf(row.get("nfs_ag", ""))
        if nf_str == "—":
            nf_val = row["nf"]
            nf_str = str(int(nf_val)) if not pd.isna(nf_val) else "—"
        dv = row["_dv"]
        d = {
            "emp":  sf(row["empresa"]),
            "nf":   nf_str,
            "cli":  sf(row["cliente"]),
            "tr":   sf(row["transportadora_norm"]),
            "mun":  sf(row["municipio"]),
            "uf":   sf(row["uf"]),
            "reg":  sf(row["regiao"]),
            "col":  fmt_d(row["data_coleta"]),
            "prev": fmt_d(row["prev_entrega_cliente"]),
            "da":   int(row["_da"]),
            "dv":   int(dv) if (dv is not None and not (isinstance(dv, float) and math.isnan(dv))) else None,
            "obs":  sf(row["obs"]),
            "st":   sf(row["status_prazo"]),
        }
        if c2:
            d["ef"] = fmt_d(row["data_efetiva_entrega"])
        return d

    df_at = df_pri[df_pri["ano_ref"] == ANO_ATUAL].copy()
    df_pend = df_pri[df_pri["ano_ref"].isin([ANO_ATUAL, ANO_ANT])]   # pendência do ano anterior continua aparecendo
    q1   = df_pend[df_pend["data_efetiva_entrega"].isna()].sort_values("prev_entrega_cliente")
    q2   = df_pend[df_pend["entrega_em_atraso"] == 1].sort_values("data_efetiva_entrega", ascending=False)  # atrasos do ano atual e do anterior

    def agg(d):
        def g(col):
            return [
                {"k": str(r[col]), "total": int(r["total"]),
                 "prazo": int(r["prazo"]), "atraso": int(r["atraso"])}
                for _, r in d.groupby(col).agg(
                    total=("entrega_no_prazo", "count"),
                    prazo=("entrega_no_prazo", "sum"),
                    atraso=("entrega_em_atraso", "sum")
                ).reset_index().iterrows()
            ]
        mes_rows = []
        for _, r in (d.groupby(["ano", "mes"]).agg(
            total=("entrega_no_prazo", "count"),
            prazo=("entrega_no_prazo", "sum"),
            atraso=("entrega_em_atraso", "sum")
        ).reset_index().sort_values(["ano", "mes"]).iterrows()):
            mes_rows.append({
                "ano": int(r["ano"]), "mes": int(r["mes"]),
                "total": int(r["total"]), "prazo": int(r["prazo"]),
                "atraso": int(r["atraso"]),
            })
        return {
            "total":  int(d.shape[0]),
            "prazo":  int(d["entrega_no_prazo"].sum()),
            "atraso": int(d["entrega_em_atraso"].sum()),
            "reg": g("regiao"), "tc": g("tipo_cliente"),
            "tr":  g("transportadora_norm"), "mes": mes_rows,
        }

    # Helper: retorna None se o DataFrame estiver vazio (empresa sem dados no ano)
    def _safe(fn, d):
        return fn(d) if not d.empty else None

    G = {"TODAS": {"TODOS": agg(df_pri)}}
    for _e in ["IAB", "VERBUM"]:
        _d = df_pri[df_pri["empresa"] == _e]
        G[_e] = {"TODOS": _safe(agg, _d)}
        for _a in ANOS:
            G[_e][_a] = _safe(agg, _d[_d["ano_ref"] == _a])

    PAYLOAD_ENT = _jdump({
        "hoje": hoje.strftime("%Y-%m-%d"),
        "q1": [rdict(r) for _, r in q1.iterrows()],
        "q2": [rdict(r, c2=True) for _, r in q2.iterrows()],
        "G": G,
    })
    print(f"      Q1: {len(q1)} pendentes | Q2: {len(q2)} atrasos")

    anos_ref = sorted(df_pri["ano_ref"].dropna().unique().tolist())

    # 6. PAYLOAD VE (Visão Executiva)
    print("[ 6 ] Payload VE...")
    df_ve = df[df["is_primeira_linha_entrega"] == 1].copy()
    df_ve["regiao"] = df_ve["regiao"].fillna("").str.strip().str.upper()
    if "lead_time" not in df_ve.columns:
        df_ve["lead_time"] = (
            df_ve["data_efetiva_entrega"] - df_ve["data_coleta"]
        ).dt.days.clip(lower=0)

    def _kex(dp, dx):
        if len(dp) == 0:
            return None
        total  = len(dp)
        prazo  = int(dp["entrega_no_prazo"].sum())
        atraso = int(dp["entrega_em_atraso"].sum())
        ns     = round(prazo / total * 100, 1)
        lt_v   = dp["lead_time"].dropna() if "lead_time" in dp.columns else pd.Series(dtype=float)
        lt     = round(float(lt_v.mean()), 1) if len(lt_v) else 0.0
        frete  = float(dx["frete"].sum())
        fat_b  = float(dx[dx["natureza_operacao"] == "VENDA"]["vr_nf"].sum())
        dev    = float(dx[dx["natureza_operacao"] == "DEVOLUÇÃO"]["vr_nf"].sum())
        bon    = float(dx[dx["natureza_operacao"] == "BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l  = fat_b - dev - bon
        peso   = float(dx["peso_kg"].sum()) if "peso_kg" in dx.columns else 0.0
        c_kg   = round(frete / peso, 4) if peso else 0.0
        # km_dedup: distância contada 1× por entrega física (já deduplicado no parquet via is_primeira_linha_entrega)
        km     = float(dp["distancia_km"].sum()) if "distancia_km" in dp.columns else 0.0
        c_km   = round(frete / km, 4) if km else 0.0
        return {
            "total": total, "prazo": prazo, "atraso": atraso,
            "ns": ns, "taxa_at": round(atraso / total * 100, 1),
            "lt": lt, "lt_med": lt,
            "fat_b": round(fat_b), "fat_l": round(fat_l),
            "dev": round(dev), "bon": round(bon),
            "frete": round(frete),
            "fsf": round(frete / fat_b * 100, 2) if fat_b else 0.0,  # FSF% = Frete total ÷ Fat.Bruto
            "pct_dev": round(dev / fat_b * 100, 1) if fat_b else 0.0,
            "c_ent": round(frete / total, 2) if total else 0.0,
            "peso": round(peso, 1), "c_kg": c_kg,
            "km": round(km), "c_km": c_km,
        }

    _VE = {}
    for _emp in ["TODAS", "IAB", "VERBUM"]:
        _VE[_emp] = {}
        _de  = df    if _emp == "TODAS" else df[df["empresa"] == _emp]
        _dpe = df_ve if _emp == "TODAS" else df_ve[df_ve["empresa"] == _emp]
        for _ano in ["TODOS"] + ANOS:
            _da  = _de  if _ano == "TODOS" else _de[_de["ano_ref"].astype(str) == _ano]
            _dpa = _dpe if _ano == "TODOS" else _dpe[_dpe["ano_ref"].astype(str) == _ano]
            if _dpa.empty:
                _VE[_emp][_ano] = None; continue
            _regs, _transp, _mes = [], [], []
            for _r, _g in _dpa.groupby("regiao"):
                if not _r or _r == "NAN":
                    continue
                _k = _kex(_g, _da[_da["regiao"] == _r])
                if _k:
                    _k["k"] = str(_r); _regs.append(_k)
            for _t, _g in _dpa.groupby("transportadora_norm"):
                _k = _kex(_g, _da[_da["transportadora_norm"] == _t])
                if _k:
                    _k["k"] = str(_t); _transp.append(_k)
            for (_a, _m), _g in _dpa.groupby(["ano", "mes"]):
                _k = _kex(_g, _da[(_da["ano"] == _a) & (_da["mes"] == _m)])
                if _k:
                    _k["ano"] = int(_a); _k["mes"] = int(_m); _mes.append(_k)
            _VE[_emp][_ano] = {
                "total":  _kex(_dpa, _da),
                "regiao": sorted(_regs,   key=lambda x: -x["total"]),
                "transp": sorted(_transp, key=lambda x: -x["total"]),
                "mes":    sorted(_mes,    key=lambda x: (x["ano"], x["mes"])),
            }

    PAYLOAD_VE  = _jdump(_VE)
    MESES_VE_JS = _jdump(["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                           "Jul", "Ago", "Set", "Out", "Nov", "Dez"])

    # 7. PAYLOAD FIN (Financeiro) — SEM LIMITE DE CLIENTES
    print("[ 7 ] Payload FIN...")

    def _kfin(dx):
        fat_b  = float(dx[dx["natureza_operacao"] == "VENDA"]["vr_nf"].sum())
        dev    = float(dx[dx["natureza_operacao"] == "DEVOLUÇÃO"]["vr_nf"].sum())
        bon    = float(dx[dx["natureza_operacao"] == "BONIFICAÇÃO"]["vr_nf"].sum())
        fat_l  = fat_b - dev - bon
        frete  = float(dx["frete"].sum())
        emb    = float(dx["custo_embalagem"].sum())
        imp    = float(dx["total_impostos_calc"].sum())
        custo  = frete + emb
        margem = fat_l - custo - imp
        ents   = len(dx[dx["is_primeira_linha_entrega"] == 1])
        return {
            "fat_b":     round(fat_b),
            "dev":       round(dev),
            "bon":       round(bon),
            "fat_l":     round(fat_l),
            "frete":     round(frete),
            "emb":       round(emb),
            "imp":       round(imp),
            "custo":     round(custo),
            "margem_op": round(margem),
            "ents":      ents,
            "ticket_med": round(fat_l / ents) if ents else 0,
            "custo_med":  round(custo / ents) if ents else 0,
            "pct_dev":   round(dev   / fat_b * 100, 1) if fat_b else 0.0,
            "pct_bon":   round(bon   / fat_b * 100, 1) if fat_b else 0.0,
            "pct_frete": round(frete / fat_b * 100, 2) if fat_b else 0.0,  # FSF% = Frete total ÷ Fat.Bruto
            "pct_custo": round(custo / fat_l * 100, 2) if fat_l else 0.0,
            "pct_imp":   round(imp   / fat_l * 100, 2) if fat_l else 0.0,
            "pct_margem":round(margem/ fat_l * 100, 1) if fat_l else 0.0,
        }

    _FIN = {}
    for _emp in ["TODAS", "IAB", "VERBUM"]:
        _FIN[_emp] = {}
        _de = df if _emp == "TODAS" else df[df["empresa"] == _emp]
        for _ano in ["TODOS"] + ANOS:
            _da   = _de if _ano == "TODOS" else _de[_de["ano_ref"].astype(str) == _ano]
            if _da.empty:
                _FIN[_emp][_ano] = None; continue
            _regs, _trs, _mes, _clis = [], [], [], []
            for _r, _g in _da.groupby("regiao"):
                if not _r: continue
                _k = _kfin(_g); _k["k"] = str(_r); _regs.append(_k)
            for _t, _g in _da.groupby("transportadora_norm"):
                _k = _kfin(_g); _k["k"] = str(_t); _trs.append(_k)
            for (_a, _m), _g in _da.groupby(["ano", "mes"]):
                _k = _kfin(_g); _k["ano"] = int(_a); _k["mes"] = int(_m)
                _mes.append(_k)
            for _c, _g in _da.groupby("cliente"):
                _k = _kfin(_g); _k["k"] = str(_c)
                _r1 = _g.iloc[0] if len(_g) > 0 else None
                if _r1 is not None:
                    _k["mun"] = str(_r1.get("municipio", "")) if "municipio" in _g.columns else ""
                    _k["uf"]  = str(_r1.get("uf", ""))        if "uf"        in _g.columns else ""
                    _k["tc"]  = str(_r1.get("tipo_cliente", "")) if "tipo_cliente" in _g.columns else ""
                _clis.append(_k)
            # SEM LIMITE: todos os clientes
            _FIN[_emp][_ano] = {
                "total":    _kfin(_da),
                "regiao":   sorted(_regs, key=lambda x: -x["fat_l"]),
                "transp":   sorted(_trs,  key=lambda x: -x["frete"]),
                "mes":      sorted(_mes,  key=lambda x: (x["ano"], x["mes"])),
                "clientes": sorted(_clis, key=lambda x: -x["fat_l"]),
            }

    PAYLOAD_FIN = _jdump(_FIN)

    # 8. PAYLOAD TR (Transportadoras)
    print("[ 8 ] Payload TR...")
    df_tr = df[df["is_primeira_linha_entrega"] == 1].copy()
    df_tr["regiao"] = df_tr["regiao"].fillna("").str.strip().str.upper()
    if "lead_time" not in df_tr.columns:
        df_tr["lead_time"] = (
            df_tr["data_efetiva_entrega"] - df_tr["data_coleta"]
        ).dt.days.clip(lower=0)

    _TR = {}
    for _emp in ["TODAS", "IAB", "VERBUM"]:
        _TR[_emp] = {}
        _de  = df    if _emp == "TODAS" else df[df["empresa"] == _emp]
        _dpe = df_tr if _emp == "TODAS" else df_tr[df_tr["empresa"] == _emp]
        for _ano in ["TODOS"] + ANOS:
            _da  = _de  if _ano == "TODOS" else _de[_de["ano_ref"].astype(str) == _ano]
            _dpa = _dpe if _ano == "TODOS" else _dpe[_dpe["ano_ref"].astype(str) == _ano]
            if _dpa.empty:
                _TR[_emp][_ano] = None; continue
            _trs_l, _ufs, _mes_l, _regs_l = [], [], [], []
            for _t, _g in _dpa.groupby("transportadora_norm"):
                _k = _kex(_g, _da[_da["transportadora_norm"] == _t])
                if _k: _k["k"] = str(_t); _trs_l.append(_k)
            for _u, _g in _dpa.groupby("uf"):
                _dx_u = _da[_da["uf"] == _u]
                _k = _kex(_g, _dx_u)
                if _k:
                    _k["k"] = str(_u)
                    # Tipos de cliente e naturezas disponíveis
                    if "tipo_cliente" in _dx_u.columns:
                        _k["tcs"] = sorted([str(x) for x in _dx_u["tipo_cliente"].dropna().unique()])
                    else:
                        _k["tcs"] = []
                    _ufs.append(_k)
            for (_a, _m), _g in _dpa.groupby(["ano", "mes"]):
                _k = _kex(_g, _da[(_da["ano"] == _a) & (_da["mes"] == _m)])
                if _k: _k["ano"] = int(_a); _k["mes"] = int(_m); _mes_l.append(_k)
            _dpa_r = _dpa.copy()
            _dpa_r["regiao"] = _dpa_r["regiao"].fillna("").str.strip().str.upper()
            for _r, _g in _dpa_r.groupby("regiao"):
                if not _r or _r == "NAN": continue
                _k = _kex(_g, _da[_da["regiao"] == _r])
                if _k: _k["k"] = str(_r); _regs_l.append(_k)
            # Subtotais UF por natureza de operação
            _uf_por_nat = {}
            if "natureza_operacao" in _dpa.columns:
                for _nat, _gnat in _dpa.groupby("natureza_operacao"):
                    _da_nat = _da[_da["natureza_operacao"] == _nat] if "natureza_operacao" in _da.columns else _da
                    _ufs_nat = []
                    for _u, _gu in _gnat.groupby("uf"):
                        _kk = _kex(_gu, _da_nat[_da_nat["uf"] == _u])
                        if _kk: _kk["k"] = str(_u); _ufs_nat.append(_kk)
                    if _ufs_nat:
                        _uf_por_nat[str(_nat)] = sorted(_ufs_nat, key=lambda x: -x["total"])

            _TR[_emp][_ano] = {
                "total":      _kex(_dpa, _da),
                "transp":     sorted(_trs_l, key=lambda x: -x["total"]),
                "uf":         sorted(_ufs,   key=lambda x: -x["total"]),
                "uf_por_nat": _uf_por_nat,
                "mes":        sorted(_mes_l, key=lambda x: (x["ano"], x["mes"])),
                "regiao":     sorted(_regs_l,key=lambda x: -x["total"]),
            }

    PAYLOAD_TR = _jdump(_TR)

    # 9-EST. PAYLOAD ESTOQUE (Movimentação de Estoque.xlsx)
    print("[ 9-EST ] Processando dados de estoque...")
    PAYLOAD_EST = _processar_estoque(ESTOQUE_XLSX, _jdump)

    # 10. LER TEMPLATE
    print("[ 9 ] Lendo template...")
    tmpl_path = os.path.join(BASE_DIR, "dashboard_template.html")
    if not os.path.exists(tmpl_path):
        print("ERRO: dashboard_template.html não encontrado na pasta etl/")
        sys.exit(1)
    with open(tmpl_path, encoding="utf-8") as f:
        html = f.read()

    # 11. SUBSTITUIR PLACEHOLDERS
    print("[ 10 ] Injetando dados no template...")
    ano_btns = "".join(
        f'<button class="ano-btn" id="ab{a}" onclick="fAno(\'{a}\')">{a}</button>'
        for a in anos_ref
    )

    substitucoes = {
        "___LI___":           LI,
        "___LV___":           LV,
        "___CI___":           CI,
        "___CV___":           CV,
        "___BG___":           BG,
        "___PAYLOAD_GER___":  PAYLOAD_GER,
        "___PAYLOAD_ENT___":  PAYLOAD_ENT,
        "___PAYLOAD_VE___":   PAYLOAD_VE,
        "___MESES_VE___":     MESES_VE_JS,
        "___PAYLOAD_FIN___":  PAYLOAD_FIN,
        "___PAYLOAD_TR___":   PAYLOAD_TR,
        "___ANO_BTNS___":     ano_btns,
        "___ANOS_REF___":     _jdump(ANOS),
        "___ANO_ATUAL___":    ANO_ATUAL,
        "___ANO_ANT___":      ANO_ANT,
        "___ANO_BTNS_VE___":  "".join(f'<button class="ano-btn" id="ve-ab{a}" onclick="veAno(\'{a}\')">{a}</button>' for a in ANOS),
        "___ANO_BTNS_FIN___": "".join(f'<button class="ano-btn" id="fin-ab{a}" onclick="finAno(\'{a}\')">{a}</button>' for a in ANOS),
        "___ANO_BTNS_TR___":  "".join(f'<button class="ano-btn" id="tr-ab{a}" onclick="trAno(\'{a}\')">{a}</button>' for a in ANOS),
        "___GERADO_EM___":    datetime.now().strftime("%d/%m/%Y %H:%M"),
        "___PAYLOAD_EST___":  PAYLOAD_EST,
    }
    for ph, val in substitucoes.items():
        html = html.replace(ph, val)

    # Zerar qualquer placeholder residual
    import re as _re
    html = _re.sub(r'___[A-Z_]+___', 'null', html)

    # 12. SALVAR
    print("[ 11 ] Salvando...")
    os.makedirs(os.path.dirname(DASH_OUT), exist_ok=True)
    with open(DASH_OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"       → {DASH_OUT}")
    print()
    print("  ✅  DASHBOARD GERADO COM SUCESSO!")
    print("=" * 60)


if __name__ == "__main__":
    gerar()
