
import pandas as pd
from dash import Dash, html, dcc, dash_table
import plotly.express as px
from pathlib import Path
import re
import unicodedata

# ---------- Ruta ----------
BASE_DIR = Path(__file__).resolve().parent
excel_path = BASE_DIR / "dashboard.xlsx"

# ---------- Lectura: usar la HOJA 'base' y encabezado en la fila 1 (header=0) ----------
df = pd.read_excel(
    excel_path,
    engine="openpyxl",
    sheet_name="base",  # <- hoja correcta según tu diagnóstico
    header=0,
    usecols="A:E"       # <- columnas A a E: ID, CATEGORIA, SUPERVISOR, GRAVEDAD, DESCRIPCION-DEL-RECLAMO
)

# ---------- Normalización de encabezados ----------
def normalize_header(s: str) -> str:
    s = str(s).strip().upper()
    # quitar acentos
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    # convertir guiones/guiones bajos en espacios
    s = s.replace("_", " ").replace("-", " ")
    # colapsar espacios
    s = re.sub(r"\s+", " ", s).strip()
    return s

df.columns = [normalize_header(c) for c in df.columns]

# ---------- Validación y mapeo (por si hay pequeñas variaciones) ----------
variantes = {
    "ID": ["ID"],
    "CATEGORIA": ["CATEGORIA"],
    "SUPERVISOR": ["SUPERVISOR"],
    "GRAVEDAD": ["GRAVEDAD"],
    "DESCRIPCION DEL RECLAMO": [
        "DESCRIPCION DEL RECLAMO", "DESCRIPCION DEL-RECLAMO", "DESCRIPCION-DEL-RECLAMO",
        "DESCRIPCION RECLAMO", "DESCRIPCION"
    ],
}

mapa = {}
for estandar, posibles in variantes.items():
    posibles_norm = {normalize_header(v) for v in posibles}
    encontrada = next((c for c in df.columns if c in posibles_norm), None)
    if encontrada:
        mapa[encontrada] = estandar

df = df.rename(columns=mapa)

esperadas = list(variantes.keys())
faltantes = [c for c in esperadas if c not in df.columns]
if faltantes:
    raise KeyError(f"Faltan columnas: {faltantes}. Detectadas: {list(df.columns)}")

# ---------- Limpieza y tipos ----------
for c in ["ID", "CATEGORIA", "SUPERVISOR", "DESCRIPCION DEL RECLAMO"]:
    df[c] = df[c].astype(str).str.strip()

# Intentar convertir GRAVEDAD a numérico si aplica
def to_numeric_safe(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(".", "", regex=False)  # miles
    s = s.str.replace(",", ".", regex=False)                  # coma decimal -> punto
    return pd.to_numeric(s, errors="coerce")

if not pd.api.types.is_numeric_dtype(df["GRAVEDAD"]):
    df["GRAVEDAD_NUM"] = to_numeric_safe(df["GRAVEDAD"])
else:
    df["GRAVEDAD_NUM"] = df["GRAVEDAD"]

# ---------- Gráficos ----------
cat_ct = df["CATEGORIA"].astype(str).value_counts(dropna=False).reset_index()
cat_ct.columns = ["CATEGORIA", "CANTIDAD"]
fig_categoria = px.bar(cat_ct, x="CATEGORIA", y="CANTIDAD", title="Reclamos por CATEGORIA")

sup_ct = df["SUPERVISOR"].astype(str).value_counts(dropna=False).nlargest(15).reset_index()
sup_ct.columns = ["SUPERVISOR", "CANTIDAD"]
fig_supervisor = px.bar(sup_ct, x="SUPERVISOR", y="CANTIDAD", title="Top 15 SUPERVISORES por cantidad")

if df["GRAVEDAD_NUM"].notna().sum() > 0:
    fig_gravedad = px.histogram(df, x="GRAVEDAD_NUM", nbins=20, title="Distribución de GRAVEDAD (numérica)")
    fig_gravedad.update_xaxes(title="GRAVEDAD")
else:
    grav_ct = df["GRAVEDAD"].astype(str).value_counts(dropna=False).reset_index()
    grav_ct.columns = ["GRAVEDAD", "CANTIDAD"]
    fig_gravedad = px.bar(grav_ct, x="GRAVEDAD", y="CANTIDAD", title="Distribución de GRAVEDAD (categorías)")

# ---------- Tabla ----------
cols_tabla = ["ID", "CATEGORIA", "SUPERVISOR", "GRAVEDAD", "DESCRIPCION DEL RECLAMO"]
tabla = dash_table.DataTable(
    columns=[{"name": c, "id": c} for c in cols_tabla],
    data=df[cols_tabla].tail(25).to_dict("records"),
    page_size=25,
    style_table={"overflowX": "auto"},
    style_cell={"textAlign": "left", "padding": "6px"},
)

# ---------- App ----------
app = Dash(__name__)
app.title = "Tablero de No Conformidades"

app.layout = html.Div(
    [
        html.H1("Tablero de No Conformidades", style={"textAlign": "center"}),
        dcc.Graph(figure=fig_categoria),
        dcc.Graph(figure=fig_supervisor),
        dcc.Graph(figure=fig_gravedad),
        html.Hr(),
        html.H3("Últimos 25 registros"),
        tabla,
    ],
    style={"maxWidth": "1100px", "margin": "0 auto", "padding": "20px"},
)

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=8050)


