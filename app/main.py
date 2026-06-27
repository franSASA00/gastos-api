from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import requests
import io
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

# URL de OneDrive — se puede sobreescribir con variable de entorno
ONEDRIVE_SHARE_URL = os.getenv(
    "ONEDRIVE_URL",
    "https://1drv.ms/x/c/c5313725c1df88b8/IQCy5eYx8955Ro6BhqdCPnqTAUSRJ85oFRIm8yoatFYbbGY?e=1HzdeZ"
)

# Cache: recarga el Excel máximo cada 10 minutos
_cache = {"data": None, "ts": 0}

def onedrive_to_download_url(share_url: str) -> str:
    """Convierte un link de compartir de OneDrive a URL de descarga directa."""
    import base64
    encoded = base64.b64encode(share_url.encode()).decode()
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"

def load_excel() -> dict[str, pd.DataFrame]:
    now = time.time()
    if _cache["data"] and (now - _cache["ts"]) < 600:
        return _cache["data"]

    # Intentar primero desde archivo local (GitHub)
    local_path = os.path.join(os.path.dirname(__file__), "..", "Gastos_Mensuales_v3.xlsx")
    
    if os.path.exists(local_path):
        sheets = pd.read_excel(local_path, sheet_name=None, header=1)
    else:
        # Fallback: descargar desde OneDrive
        download_url = onedrive_to_download_url(ONEDRIVE_SHARE_URL)
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(download_url, headers=headers, timeout=30, allow_redirects=True)
        if resp.status_code != 200:
            raise HTTPException(status_code=502, detail=f"No se pudo descargar el Excel desde OneDrive (HTTP {resp.status_code})")
        sheets = pd.read_excel(io.BytesIO(resp.content), sheet_name=None, header=1)

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
    sheets = load_excel()
    sheet = sheets.get("Panel Principal")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Panel Principal' no encontrada")
    return {"sheet": "Panel Principal", "data": sheet_to_records(sheet)}

@app.get("/tarjetas")
def get_tarjetas():
    sheets = load_excel()
    sheet = sheets.get("Tarjetas_Bancos")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Tarjetas_Bancos' no encontrada")
    return {"sheet": "Tarjetas_Bancos", "data": sheet_to_records(sheet)}

@app.get("/personales")
def get_personales():
    sheets = load_excel()
    sheet = sheets.get("Gastos_Personales")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Gastos_Personales' no encontrada")
    return {"sheet": "Gastos_Personales", "data": sheet_to_records(sheet)}

@app.get("/hogar")
def get_hogar():
    sheets = load_excel()
    sheet = sheets.get("Hogar")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Hogar' no encontrada")
    return {"sheet": "Hogar", "data": sheet_to_records(sheet)}

@app.get("/deudas")
def get_deudas():
    sheets = load_excel()
    sheet = sheets.get("Deudas_Varios")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Deudas_Varios' no encontrada")
    return {"sheet": "Deudas_Varios", "data": sheet_to_records(sheet)}

@app.get("/referencia")
def get_referencia():
    sheets = load_excel()
    sheet = sheets.get("Referencia")
    if sheet is None: raise HTTPException(status_code=404, detail="Pestaña 'Referencia' no encontrada")
    return {"sheet": "Referencia", "data": sheet_to_records(sheet)}

@app.get("/resumen")
def get_resumen():
    sheets = load_excel()
    panel = sheets.get("Panel Principal")
    if panel is None: raise HTTPException(status_code=404, detail="Pestaña 'Panel Principal' no encontrada")

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
