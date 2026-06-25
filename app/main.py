from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import requests
import io
import os
from functools import lru_cache
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

ONEDRIVE_URL = os.getenv("ONEDRIVE_URL", "")  # URL pública del Excel en OneDrive

# Cache simple: recarga el Excel máximo cada 10 minutos
_cache = {"data": None, "ts": 0}

def load_excel() -> dict[str, pd.DataFrame]:
    now = time.time()
    if _cache["data"] and (now - _cache["ts"]) < 600:
        return _cache["data"]

    if not ONEDRIVE_URL:
        raise HTTPException(status_code=500, detail="ONEDRIVE_URL no configurada")

    # OneDrive: reemplazar el final de la URL para forzar descarga directa
    url = ONEDRIVE_URL.replace("embed", "download").replace("view.aspx", "download.aspx")
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        raise HTTPException(status_code=502, detail=f"No se pudo descargar el Excel (HTTP {resp.status_code})")

    sheets = pd.read_excel(io.BytesIO(resp.content), sheet_name=None, header=1)
    _cache["data"] = sheets
    _cache["ts"] = now
    return sheets


def sheet_to_records(df: pd.DataFrame) -> list[dict]:
    """Limpia el DataFrame y lo convierte a lista de dicts para JSON."""
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    # Rellenar NaN con None para JSON limpio
    df = df.where(pd.notna(df), other=None)
    return df.to_dict(orient="records")


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {
        "api": "Gastos Mensuales API",
        "endpoints": [
            "/panel",
            "/tarjetas",
            "/personales",
            "/hogar",
            "/deudas",
            "/referencia",
            "/resumen",
        ]
    }


@app.get("/panel")
def get_panel():
    sheets = load_excel()
    sheet = sheets.get("Panel Principal")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Panel Principal' no encontrada")
    return {"sheet": "Panel Principal", "data": sheet_to_records(sheet)}


@app.get("/tarjetas")
def get_tarjetas():
    sheets = load_excel()
    sheet = sheets.get("Tarjetas_Bancos")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Tarjetas_Bancos' no encontrada")
    return {"sheet": "Tarjetas_Bancos", "data": sheet_to_records(sheet)}


@app.get("/personales")
def get_personales():
    sheets = load_excel()
    sheet = sheets.get("Gastos_Personales")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Gastos_Personales' no encontrada")
    return {"sheet": "Gastos_Personales", "data": sheet_to_records(sheet)}


@app.get("/hogar")
def get_hogar():
    sheets = load_excel()
    sheet = sheets.get("Hogar")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Hogar' no encontrada")
    return {"sheet": "Hogar", "data": sheet_to_records(sheet)}


@app.get("/deudas")
def get_deudas():
    sheets = load_excel()
    sheet = sheets.get("Deudas_Varios")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Deudas_Varios' no encontrada")
    return {"sheet": "Deudas_Varios", "data": sheet_to_records(sheet)}


@app.get("/referencia")
def get_referencia():
    sheets = load_excel()
    sheet = sheets.get("Referencia")
    if sheet is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Referencia' no encontrada")
    return {"sheet": "Referencia", "data": sheet_to_records(sheet)}


@app.get("/resumen")
def get_resumen():
    """Endpoint especial: devuelve un resumen mensual plano, ideal para Power BI."""
    sheets = load_excel()
    panel = sheets.get("Panel Principal")
    if panel is None:
        raise HTTPException(status_code=404, detail="Pestaña 'Panel Principal' no encontrada")

    # Filas de interés (por label en col 0)
    ROWS_OF_INTEREST = [
        "Tarjetas & Bancos",
        "Gastos Personales",
        "Hogar",
        "Deudas & Varios",
        "TOTAL GASTOS",
        "Ingreso Franco",
        "Ingreso Abi",
        "TOTAL INGRESOS",
        "¿LLEGAMOS A PAGAR?",
        "% Franco",
        "% Abi",
    ]

    panel = panel.dropna(how="all").reset_index(drop=True)
    # Primera columna = etiquetas
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
