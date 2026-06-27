from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os
import time

app = FastAPI(
    title="Gastos Mensuales API",
    description="API para leer el Excel de gastos y exponerlo a Power BI",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "..", "Gastos_Mensuales_v3.xlsx")

_cache = {"data": None, "ts": 0}

def load_excel() -> dict[str, pd.DataFrame]:
    now = time.time()
    if _cache["data"] and (now - _cache["ts"]) < 600:
        return _cache["data"]
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(status_code=500, detail=f"Archivo Excel no encontrado en: {EXCEL_PATH}")
    sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, header=1)
    _cache["data"] = sheets
    _cache["ts"] = now
    return sheets

def sheet_to_records(df: pd.DataFrame) -> list[dict]:
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.where(pd.notna(df), other=None)
    return df.to_dict(orient="records")

@app.get("/")
def root():
    return {"api": "Gastos Mensuales API", "endpoints": ["/panel","/tarjetas","/personales","/hogar","/deudas","/referencia","/resumen"]}

@app.get("/panel")
def get_panel():
    sheet = load_excel().get("Panel Principal")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Panel Principal", "data": sheet_to_records(sheet)}

@app.get("/tarjetas")
def get_tarjetas():
    sheet = load_excel().get("Tarjetas_Bancos")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Tarjetas_Bancos", "data": sheet_to_records(sheet)}

@app.get("/personales")
def get_personales():
    sheet = load_excel().get("Gastos_Personales")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Gastos_Personales", "data": sheet_to_records(sheet)}

@app.get("/hogar")
def get_hogar():
    sheet = load_excel().get("Hogar")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Hogar", "data": sheet_to_records(sheet)}

@app.get("/deudas")
def get_deudas():
    sheet = load_excel().get("Deudas_Varios")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Deudas_Varios", "data": sheet_to_records(sheet)}

@app.get("/referencia")
def get_referencia():
    sheet = load_excel().get("Referencia")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")
    return {"sheet": "Referencia", "data": sheet_to_records(sheet)}

@app.get("/resumen")
def get_resumen():
    panel = load_excel().get("Panel Principal")
    if panel is None: raise HTTPException(status_code=404, detail="Pestaña no encontrada")

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
            v = row[mes]
            records.append({
                "categoria": label,
                "mes": str(mes).strip(),
                "valor": None if pd.isna(v) else v
            })

    return {"data": records}

@app.get("/health")
def health():
    return {"status": "ok"}
