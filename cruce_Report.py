import streamlit as st
from pathlib import Path
import pandas as pd
import plotly.express as px

ruta = Path(r"/mnt/red")

st.title("Reporte")
if st.button("↻ Actualizar"):
    st.rerun()

archivos_csv = list(ruta.glob("*.csv"))
if not archivos_csv:
    st.warning(f"No se encontraron archivos CSV en: {ruta}")
    st.stop()

# Leer y unir todos los archivos
df = pd.concat(
    [
        pd.read_csv(archivo).assign(**{"FUID origen": archivo.name})
        for archivo in archivos_csv
    ],
    ignore_index=True
)
st.caption(f"{len(archivos_csv)} archivo(s) CSV · {len(df):,} fila(s)")

# Filtro por archivo de origen
fuid_origenes = sorted(df["FUID origen"].dropna().unique())
fuid_origen_seleccionados = st.multiselect(
    "Filtrar por FUID origen",
    options=fuid_origenes,
    default=fuid_origenes,
)

df_filtrado = df[df["FUID origen"].isin(fuid_origen_seleccionados)]

valores_grafica = {
    "titulo_grafica": "Cantidad por resultado final",
    "etiqueta_eje_x": "Resultado final",
    "etiqueta_eje_y": "Cantidad",
}
for clave, valor in valores_grafica.items():
    st.session_state.setdefault(clave, valor)

conteo = (
    df_filtrado["resultado_final"]
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
    title=st.session_state["titulo_grafica"],
    labels={
        "resultado_final": st.session_state["etiqueta_eje_x"],
        "cantidad": st.session_state["etiqueta_eje_y"],
    },
)
fig.update_layout(title={"x": 0.5, "xanchor": "center"})

st.plotly_chart(fig, use_container_width=True)

st.subheader("Personalizar gráfico")
st.text_input(
    "Título",
    key="titulo_grafica",
)
st.text_input(
    "Texto del eje X",
    key="etiqueta_eje_x",
)
st.text_input(
    "Texto del eje Y",
    key="etiqueta_eje_y",
)
