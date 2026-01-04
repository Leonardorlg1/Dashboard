
import os
import pandas as pd
import dash
from dash import html, dcc
import plotly.express as px

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_XLSX = os.path.join(BASE_DIR, "dashboard.xlsx")

# Columnas que tu dashboard espera (ajústalas si cambian)
COLUMNAS_ESPERADAS = ["Categoría", "Estado", "Gravedad"]

def cargar_datos():
    if not os.path.exists(RUTA_XLSX):
        raise FileNotFoundError(f"No se encontró el archivo: {RUTA_XLSX}")

    # --- Lee sin encabezados para poder detectar la fila de header ---
    # (Si tu archivo tiene varias hojas, especifica sheet_name="Hoja1" u otro)
    raw = pd.read_excel(RUTA_XLSX, engine="openpyxl", header=None)

 # Busca la fila que contenga todas las columnas esperadas
    header_row_idx = None
    for idx in range(min(10, len(raw))):  # busca en las primeras 10 filas
        fila = raw.iloc[idx].astype(str).str.strip().tolist()
        if all(col in fila for col in COLUMNAS_ESPERADAS):
            header_row_idx = idx
            break

    if header_row_idx is None:
        # Si no se halló exactamente, intenta una detección parcial
        for idx in range(min(10, len(raw))):
            fila = raw.iloc[idx].astype(str).str.strip().tolist()
            # Si al menos 2 de las esperadas están en la fila, la tomamos como header
            match_count = sum(col in fila for col in COLUMNAS_ESPERADAS)
            if match_count >= 2:
                header_row_idx = idx
                break

if header_row_idx is None:
        # Último recurso: imprimir primeras filas para que veas qué hay
        preview = raw.head(5).to_string(index=False)
        raise ValueError(
            "No pude detectar la fila de encabezados.\n"
            f"Revisa las primeras filas del Excel:\n{preview}\n"
            "Opciones: mueve los encabezados a la primera fila, indica sheet_name, "
            "o dime los nombres exactos de las columnas para ajustar."
        )

    # Relee con el header correcto (saltando filas previas)
    df = pd.read_excel(
        RUTA_XLSX,
        engine="openpyxl",
        header=header_row_idx
    )

 # Limpieza básica
    df = df.dropna(how="all")

    # Validar columnas esperadas
    faltantes = set(COLUMNAS_ESPERADAS) - set(df.columns)
    if faltantes:
        # A veces vienen con espacios o mayúsculas distintas; normaliza
        # Intento de normalización simple:
        mapeo = {}
        for col in df.columns:
            norm = col.strip().lower()
            if norm == "categoria":
                mapeo[col] = "Categoría"
            elif norm == "estado":
                mapeo[col] = "Estado"
            elif norm == "gravedad":
                mapeo[col] = "Gravedad"

  if mapeo:
            df = df.rename(columns=mapeo)
            faltantes = set(COLUMNAS_ESPERADAS) - set(df.columns)

    if faltantes:
        raise KeyError(
            f"Faltan columnas: {faltantes}. Columnas detectadas: {list(df.columns)}.\n"
            "Ajusta los encabezados del Excel o actualiza COLUMNAS_ESPERADAS."
        )

    # Quita filas donde falten claves
    df = df.dropna(subset=["Categoría", "Estado", "Gravedad"])

    # Opcional: convierte tipos si hace falta
    # df["Gravedad"] = pd.to_numeric(df["Gravedad"], errors="coerce")

    return df

df = cargar_datos()

# --- Gráficos ---
fig_categoria = px.bar(
    df,
    x="Categoría",
    title="No conformidades por categoría",
)
fig_estado = px.pie(
    df,
    names="Estado",
    title="Distribución por estado",
)
fig_gravedad = px.histogram(
    df,
    x="Gravedad",
    title="Distribución por gravedad",
)

# --- App Dash ---
app = dash.Dash(__name__)
app.title = "Dashboard de No Conformidades"
app.layout = html.Div(
    [
        html.H1("Dashboard de No Conformidades", style={"textAlign": "center"}),
        dcc.Graph(figure=fig_categoria),
        dcc.Graph(figure=fig_estado),
        dcc.Graph(figure=fig_gravedad),
    ],
    style={"maxWidth": "1100px", "margin": "0 auto", "padding": "20px"},
)

if __name__ == "__main__":
    app.run_server(debug=True, host="127.0.0.1", port=8050)
