# Gastos Mensuales API

API en FastAPI que lee el Excel de gastos desde OneDrive y expone los datos para Power BI.

---

## Endpoints disponibles

| Endpoint | Descripción |
|---|---|
| `GET /` | Lista de endpoints |
| `GET /panel` | Datos del Panel Principal |
| `GET /tarjetas` | Tarjetas & Bancos |
| `GET /personales` | Gastos Personales |
| `GET /hogar` | Hogar |
| `GET /deudas` | Deudas & Varios |
| `GET /referencia` | Tabla de referencia alquileres |
| `GET /resumen` | **Resumen mensual plano — usar este en Power BI** |
| `GET /health` | Health check |

---

## Pasos para deployar en Render

### 1. Subir a GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TU_USUARIO/gastos-api.git
git push -u origin main
```

### 2. Obtener la URL pública del Excel en OneDrive

1. Abrí el Excel en OneDrive
2. Hacé clic en **Compartir → Cualquier persona con el vínculo puede ver**
3. Copiá el link generado (formato: `https://1drv.ms/x/s!...`)

### 3. Crear el servicio en Render

1. Entrá a [render.com](https://render.com) y creá una cuenta gratuita
2. **New → Web Service → Connect a repository** → seleccioná tu repo
3. Configurá:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
4. En **Environment Variables** agregá:
   - Key: `ONEDRIVE_URL`
   - Value: la URL del Excel que copiaste en el paso 2
5. Click **Create Web Service**

Render te va a dar una URL pública tipo:
```
https://gastos-api.onrender.com
```

---

## Conectar en Power BI

### Endpoint recomendado: `/resumen`

Devuelve una tabla plana con columnas `categoria`, `mes`, `valor` — ideal para crear visualizaciones.

### Pasos en Power BI Desktop

1. **Obtener datos → Web**
2. URL: `https://gastos-api.onrender.com/resumen`
3. Power Query va a mostrar el JSON → expandí la columna `data`
4. Vas a tener 3 columnas: `categoria`, `mes`, `valor`
5. Desde ahí podés crear:
   - Gráfico de barras por mes
   - Filtro por categoría
   - Tarjeta de "¿Llegamos a pagar?" del mes actual
   - Tabla comparativa de ingresos vs gastos

### Programar actualización automática

En Power BI Service (online) podés configurar **Actualización programada** para que refresque los datos diariamente o cada semana.

---

## Notas

- El Excel se cachea por **10 minutos** para no sobrecargar OneDrive
- Si actualizás el Excel en OneDrive, los cambios se reflejan en la próxima llamada (máx. 10 min de delay)
- El plan gratuito de Render pone el servicio a dormir tras 15 min de inactividad — el primer request puede tardar ~30 seg en despertar
