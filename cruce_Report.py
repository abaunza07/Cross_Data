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

columnas_graficas = ["archivo_difusion", "ciclo_vida"]

# Leer y unir todos los archivos, conservando explícitamente las columnas de análisis.
def leer_csv(archivo: Path) -> pd.DataFrame:
    datos = pd.read_csv(archivo)
    datos.columns = datos.columns.str.strip().str.replace("\ufeff", "", regex=False)
    columnas_dataframe = list(dict.fromkeys([*datos.columns, *columnas_graficas]))
    return datos.reindex(columns=columnas_dataframe)


df = pd.concat(
    [
        leer_csv(archivo).assign(**{"FUID origen": archivo.name})
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
    default=[],
)

df_filtrado = df[df["FUID origen"].isin(fuid_origen_seleccionados)]

colores_resultado = {
    "Si": "#1976D2",
    "Sí": "#1976D2",
    "No": "#64B5F6",
    "HUERFANO": "#FF2B2B",
    "Huerfano": "#FF2B2B",
    "huérfano": "#FF2B2B",
}

TITULO_GENERAL = "Reporte Cruce FUID servidor lenovo"
ETIQUETAS_RESULTADO = {
    "resultado_final": "Resultado final",
    "cantidad": "Cantidad",
}


def preparar_conteo(
    df_base: pd.DataFrame,
    columnas: list[str],
) -> pd.DataFrame:
    datos = df_base[columnas].copy()
    for columna in columnas:
        datos[columna] = datos[columna].astype("string").fillna(
            "Sin dato" if columna != "resultado_final" else "Sin respuesta"
        )
    return (
        datos.value_counts()
        .rename("cantidad")
        .reset_index()
    )


def mostrar_graficas(df_graficas: pd.DataFrame, titulo: str) -> None:
    if df_graficas.empty:
        st.info("No hay datos para los filtros seleccionados.")
        return

    conteo = preparar_conteo(df_graficas, ["resultado_final"])
    fig = px.bar(
        conteo,
        x="resultado_final",
        y="cantidad",
        text="cantidad",
        title=titulo,
        labels=ETIQUETAS_RESULTADO,
        color="resultado_final",
        color_discrete_map=colores_resultado,
    )
    fig.update_layout(
        title={"x": 0.5, "xanchor": "center"},
        legend={
            "orientation": "h",
            "x": 0.5,
            "xanchor": "center",
            "y": -0.15,
            "yanchor": "top",
            "font": {"size": 10},
            "entrywidth": 90,
            "entrywidthmode": "pixels",
        },
        margin={"l": 40, "r": 40, "t": 80, "b": 80},
    )
    st.plotly_chart(fig, use_container_width=True)

    fig_torta = px.pie(
        conteo,
        names="resultado_final",
        values="cantidad",
        title=titulo,
        labels=ETIQUETAS_RESULTADO,
        color="resultado_final",
        color_discrete_map=colores_resultado,
    )
    fig_torta.update_traces(
        textposition="outside",
        textfont={"size": 11, "color": "black"},
        texttemplate="%{percent:.1%}",
        hovertemplate="%{label}: %{percent:.1%}<extra></extra>",
    )
    fig_torta.update_layout(
        title={"x": 0.5, "xanchor": "center"},
        uniformtext_minsize=10,
        uniformtext_mode="show",
        margin={"l": 60, "r": 60, "t": 80, "b": 80},
        showlegend=True,
        legend={
            "orientation": "h",
            "x": 0.5,
            "xanchor": "center",
            "y": -0.08,
            "yanchor": "top",
            "font": {"size": 10},
            "entrywidth": 90,
            "entrywidthmode": "pixels",
        },
    )
    st.plotly_chart(fig_torta, use_container_width=True)


def crear_graficas_archivo(
    df_base: pd.DataFrame,
    titulo: str,
) -> tuple[object, object] | None:
    columnas_requeridas = ["archivo_difusion", "resultado_final"]
    if any(columna not in df_base.columns for columna in columnas_requeridas):
        return None

    conteo = preparar_conteo(df_base, columnas_requeridas)
    conteo_barra = conteo.copy()
    conteo_barra.loc[
        conteo_barra["archivo_difusion"].eq("Sin dato"),
        "archivo_difusion",
    ] = ""
    conteo["segmento"] = conteo["archivo_difusion"].where(
        conteo["archivo_difusion"].ne("Sin dato"),
        conteo["resultado_final"],
    )
    conteo.loc[
        conteo["archivo_difusion"].ne("Sin dato"),
        "segmento",
    ] = (
        conteo["archivo_difusion"]
        + " · "
        + conteo["resultado_final"]
    )
    etiquetas = {
        "archivo_difusion": "resultado_final",
        "resultado_final": "Tipo de respuesta",
        "cantidad": "Cantidad",
    }

    fig_barra = px.bar(
        conteo_barra,
        x="archivo_difusion",
        y="cantidad",
        color="resultado_final",
        barmode="group",
        text="cantidad",
        title=titulo,
        labels=etiquetas,
        color_discrete_map=colores_resultado,
    )
    fig_barra.update_layout(
        title={"x": 0.5, "xanchor": "center"},
        margin={"l": 60, "r": 60, "t": 80, "b": 100},
        legend={"orientation": "h", "x": 0.5, "xanchor": "center", "y": -0.2},
    )

    fig_torta = px.pie(
        conteo,
        names="segmento",
        values="cantidad",
        color="resultado_final",
        title=titulo,
        labels=etiquetas,
        color_discrete_map=colores_resultado,
    )
    fig_torta.update_traces(
        textposition="outside",
        textfont={"size": 11, "color": "black"},
        texttemplate="%{percent:.1%}",
        hovertemplate="%{label}: %{value}<extra></extra>",
    )
    fig_torta.update_layout(
        title={"x": 0.5, "xanchor": "center"},
        uniformtext_minsize=10,
        uniformtext_mode="show",
        margin={"l": 70, "r": 70, "t": 80, "b": 100},
        legend={"orientation": "h", "x": 0.5, "xanchor": "center", "y": -0.12},
    )
    return fig_barra, fig_torta


tab_general, tab_difusion, tab_ciclo_vida = st.tabs(
    ["Reporte general", "Archivo de difusión", "Ciclo de vida"]
)

with tab_general:
    mostrar_graficas(df_filtrado, TITULO_GENERAL)

    st.subheader("Gráficas por archivo CSV y consolidado")
    archivos_graficas = [
        (archivo.stem, leer_csv(archivo))
        for archivo in sorted(archivos_csv)
    ]
    archivos_graficas.append(("Total", df.drop(columns=["FUID origen"])))

    for nombre_archivo, datos_archivo in archivos_graficas:
        graficas_archivo = crear_graficas_archivo(
            datos_archivo,
            f"Reporte Cruce FUID servidor lenovo - {nombre_archivo}",
        )
        if graficas_archivo is None:
            st.warning(
                f"No se puede generar la gráfica de {nombre_archivo}: "
                "faltan archivo_difusion o resultado_final."
            )
            continue
        with st.expander(nombre_archivo, expanded=nombre_archivo == "Total"):
            st.plotly_chart(graficas_archivo[0], use_container_width=True)
            st.plotly_chart(graficas_archivo[1], use_container_width=True)