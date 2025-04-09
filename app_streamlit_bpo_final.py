
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(layout="wide")
st.title("Distribución BPO con Ajuste a Melissa y Agente Incontactable")

# Simulación: cargar datos desde archivo subido por el usuario
archivo = st.file_uploader("📂 Cargar archivo Excel", type=["xlsx"])
if archivo:
    df = pd.read_excel(archivo)

    # Vista previa
    st.subheader("👀 Vista previa")
    st.dataframe(df.head(15), height=500, use_container_width=True)

    # Agentes base
    agentes_bpo = [
        "Alysson Garcia",
        "Ana Paniagua",
        "Julio de Leon",
        "Melissa Florian",
        "Nancy Zet"
    ]

    # Agregar a Abigail los sábados
    fecha_actual = datetime.now()
    if fecha_actual.weekday() == 5:  # sábado
        agentes_bpo.append("Abigail Vasquez")

    # Identificar incontactables
    cantidad_incontactable = df[df["Agente BPO"] == "Agente Incontactable"].shape[0]
    indices_sin_asignar = df[df["Agente BPO"].isna()].index.tolist()
    total_distribuibles = len(indices_sin_asignar)

    # Ponderaciones
    ponderaciones = {a: 0.75 if a == "Melissa Florian" else 1.0 for a in agentes_bpo}
    total_ponderado = sum(ponderaciones.values())

    # Calcular cantidades exactas y asegurar que la suma cuadre
    cantidades = {a: int((ponderaciones[a] / total_ponderado) * total_distribuibles) for a in agentes_bpo}
    asignados = sum(cantidades.values())
    faltan = total_distribuibles - asignados

    # Asignar faltantes (por redondeo) a los primeros agentes en orden (excepto Melissa si ya tuvo menos)
    agentes_ordenados = sorted(agentes_bpo, key=lambda x: -ponderaciones[x])
    for i in range(faltan):
        cantidades[agentes_ordenados[i % len(agentes_ordenados)]] += 1

    # Asignar registros a cada agente
    for agente in agentes_bpo:
        for _ in range(cantidades[agente]):
            if indices_sin_asignar:
                df.at[indices_sin_asignar.pop(0), "Agente BPO"] = agente

    # Tabla resumen
    resumen = df["Agente BPO"].value_counts().sort_index().reset_index()
    resumen.columns = ["Agente", "Cantidad"]
    resumen.loc[len(resumen.index)] = ["Total general", resumen["Cantidad"].sum()]

    st.subheader("📊 Resumen de distribución")
    st.dataframe(resumen, use_container_width=True)

    # Descarga del resumen
    resumen_csv = resumen.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Descargar resumen CSV", resumen_csv, file_name="resumen_distribucion.csv", mime="text/csv")
