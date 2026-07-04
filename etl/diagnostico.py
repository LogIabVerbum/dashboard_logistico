import pandas as pd, re, sys, os

BASE = os.path.dirname(os.path.abspath(__file__))

def _encontrar_pasta():
    candidatos = [
        os.environ.get("OneDriveCommercial", ""),
        os.environ.get("OneDrive", ""),
        os.path.join(os.environ.get("USERPROFILE",""), "OneDrive - Instituto Alfa e Beto"),
    ]
    for base in candidatos:
        if not base: continue
        for sub in ["Logística-IAB\10 - DASHIBOARDs", "Logistica-IAB\10 - DASHIBOARDs", "10 - DASHIBOARDs"]:
            pasta = os.path.join(base, sub)
            if os.path.isdir(pasta): return pasta
    return os.path.normpath(os.path.join(BASE, "..", ".."))

PASTA = _encontrar_pasta()
QUADRO = os.path.join(PASTA, "Quadro de Envios.xlsx")
ARQUIVO_BASE = os.path.join(PASTA, "arquivo_base.xls")
print(f"Pasta: {PASTA}")
print(f"Quadro: {QUADRO}")

CNPJ_EMP = {
    "08.458.084/0001-13": "IAB",
    "57.638.518/0002-53": "VERBUM",
    "57.638.518/0001-72": "VERBUM",
}

def chave(data, nf, cnpj, emp=""):
    try: d = pd.Timestamp(data).strftime("%Y-%m-%d") if not pd.isna(data) else ""
    except: d = ""
    try: n = str(int(float(nf))) if not pd.isna(nf) else ""
    except: n = ""
    c = re.sub(r"\D","", str(cnpj)) if cnpj else ""
    return f"{d}|{n}|{c}|{str(emp).strip().upper()}"

print("Lendo Quadro de Envios...")
qe = pd.read_excel(QUADRO, sheet_name="Lançamentos", engine="openpyxl", dtype=str)
qe["Data Coleta"] = pd.to_datetime(qe["Data Coleta"], errors="coerce")
qe["NF"] = pd.to_numeric(qe["NF"], errors="coerce")

suely_qe = qe[qe["CNPJ / CPF"].str.contains("06.200.544", na=False)]
print(f"\nSUELY no Quadro: {len(suely_qe)} linhas")
print(suely_qe[["Data Coleta","NF","CNPJ / CPF","Cliente","Empresa","NATUREZA DA OPERAÇÃO"]].to_string())

print("\nChaves no Quadro (Suely):")
for _, r in suely_qe.iterrows():
    emp = str(r.get("Empresa","")).strip().upper()
    ch = chave(r["Data Coleta"], r["NF"], r["CNPJ / CPF"], emp)
    print(f"  {ch}")

print("\nLendo arquivo_base...")
ab = pd.read_excel(ARQUIVO_BASE, engine="xlrd", skiprows=2, header=0, dtype=str)
ab["DT_EMISSAO_NF"] = pd.to_datetime(ab["DT_EMISSAO_NF"], errors="coerce")
ab["NRO_NFE"] = pd.to_numeric(ab["NRO_NFE"], errors="coerce")

suely_ab = ab[ab["CPF_CNPJ_DESTINATARIO"].str.contains("06.200.544", na=False)]
print(f"\nSUELY no arquivo_base: {len(suely_ab)} linhas")

print("\nChaves no arquivo_base (Suely):")
chaves_qe = set()
for _, r in qe.iterrows():
    emp = str(r.get("Empresa","")).strip().upper()
    chaves_qe.add(chave(r["Data Coleta"], r["NF"], r["CNPJ / CPF"], emp))

for _, r in suely_ab.iterrows():
    emp = CNPJ_EMP.get(str(r.get("CNPJ_REMETENTE","")).strip(), "")
    ch = chave(r["DT_EMISSAO_NF"], r["NRO_NFE"], r["CPF_CNPJ_DESTINATARIO"], emp)
    existe = "JA EXISTE" if ch in chaves_qe else "NOVA (sera inserida)"
    print(f"  {ch}  →  {existe}")
