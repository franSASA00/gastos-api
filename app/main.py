from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import requests
import io
import time
import re

app = FastAPI(
    title="Gastos Mensuales API",
    description="API para leer el Excel de gastos desde Google Sheets y exponerlo a Power BI",
    version="2.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

SHEETS = {
    "Panel Principal":   "2058314116",
    "Tarjetas_Bancos":   "2069223113",
    "Gastos_Personales": "1789675783",
    "Hogar":             "1236276542",
    "Deudas_Varios":     "292959967",
    "Referencia":        "1655435187",
}

SHEET_ID = "2PACX-1vRGCLCZi3xe1djvQbl2L_ztbjIOq-iT-PYV0VAYyuzwpZbLTTOOuwWuYRgZpuFYxWxhSoQTF1Uqg52I"

_cache = {"data": None, "ts": 0}

def parse_ar_number(s) -> float | None:
    """Parsea números en formato argentino: 1.234.567,89 o 1234567,89"""
    if s is None:
        return None
    s = str(s).strip()
    if s in ('', '-', 'nan', 'NaN', '#N/A', '#VALUE!', '#REF!'):
        return None
    # Quitar símbolo $, espacios y paréntesis (formato negativo)
    negative = s.startswith('(') or s.startswith('-')
    s = re.sub(r'[\$\s\(\)]', '', s).replace('-', '')
    # Formato argentino: puntos = miles, coma = decimal
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        # Sin coma: solo puntos (pueden ser miles)
        parts = s.split('.')
        if len(parts) > 1 and len(parts[-1]) == 3:
            s = s.replace('.', '')  # son separadores de miles
        # Si no, es decimal normal
    try:
        val = float(s)
        return -val if negative else val
    except:
        return None

def fetch_sheet(gid: str) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/e/{SHEET_ID}/pub?output=csv&gid={gid}"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"No se pudo leer Google Sheets (HTTP {resp.status_code})")
    df = pd.read_csv(io.StringIO(resp.text), header=1, dtype=str)
    return df

def load_sheets() -> dict[str, pd.DataFrame]:
    now = time.time()
    if _cache["data"] and (now - _cache["ts"]) < 600:
        return _cache["data"]
    sheets = {name: fetch_sheet(gid) for name, gid in SHEETS.items()}
    _cache["data"] = sheets
    _cache["ts"] = now
    return sheets

def sheet_to_records(df: pd.DataFrame) -> list[dict]:
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df.to_dict(orient="records")

@app.get("/")
def root():
    return {"api": "Gastos Mensuales API v2.1 — Google Sheets",
            "endpoints": ["/panel","/tarjetas","/personales","/hogar","/deudas","/referencia","/resumen"]}

@app.get("/panel")
def get_panel():
    return {"sheet": "Panel Principal", "data": sheet_to_records(load_sheets()["Panel Principal"])}

@app.get("/tarjetas")
def get_tarjetas():
    return {"sheet": "Tarjetas_Bancos", "data": sheet_to_records(load_sheets()["Tarjetas_Bancos"])}

@app.get("/personales")
def get_personales():
    return {"sheet": "Gastos_Personales", "data": sheet_to_records(load_sheets()["Gastos_Personales"])}

@app.get("/hogar")
def get_hogar():
    return {"sheet": "Hogar", "data": sheet_to_records(load_sheets()["Hogar"])}

@app.get("/deudas")
def get_deudas():
    return {"sheet": "Deudas_Varios", "data": sheet_to_records(load_sheets()["Deudas_Varios"])}

@app.get("/referencia")
def get_referencia():
    return {"sheet": "Referencia", "data": sheet_to_records(load_sheets()["Referencia"])}

@app.get("/resumen")
def get_resumen():
    panel = load_sheets()["Panel Principal"]

    ROWS_OF_INTEREST = [
        "Tarjetas & Bancos", "Gastos Personales", "Hogar", "Deudas & Varios",
        "TOTAL GASTOS", "Ingreso Franco", "Ingreso Abi", "Pago de negocio",
        "Ingreso por rendimientos", "Nos deben", "TOTAL INGRESOS",
        "¿LLEGAMOS A PAGAR?", "% Franco", "% Abi",
    ]

    panel = panel.dropna(how="all").reset_index(drop=True)
    label_col = panel.columns[0]
    month_cols = panel.columns[1:]

    records = []
    for _, row in panel.iterrows():
        label = str(row[label_col]).strip()
        if label not in ROWS_OF_INTEREST:
            continue
        for mes in month_cols:
            v = parse_ar_number(row[mes])
            records.append({
                "categoria": label,
                "mes": str(mes).strip(),
                "valor": v
            })

    return {"data": records}

@app.get("/health")
def health():
    return {"status": "ok"}
