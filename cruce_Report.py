import streamlit as st
from pathlib import Path
import pandas as pd
import plotly.express as px

ruta = Path(r"/mnt/red")
archivos_csv = list(ruta.glob("*.csv"))
if not archivos_csv:
    raise FileNotFoundError(f"No se encontraron archivos CSV en: {ruta}")

# Leer y unir todos los archivos
df = pd.concat(
    [pd.read_csv(archivo) for archivo in archivos_csv],
    ignore_index=True
)
#st.title("Reporte")
#st.metric("Cantidad de filas", len(df))
#st.dataframe(df, use_container_width=True)

#Por ejemplo, si tu DataFrame df tiene las columnas region, fecha y ventas:

conteo = (
    df["resultado_final"]
    .astype("string")
    .fillna("Sin dato")
    .value_counts()
    .rename_axis("resultado_final")
    .reset_index(name="cantidad")
)

fig = px.bar(
    conteo,
    x="resultado_final",
    y="cantidad",
    text="cantidad",
    title="Cantidad por resultado final"
)

st.plotly_chart(fig, use_container_width=True)
