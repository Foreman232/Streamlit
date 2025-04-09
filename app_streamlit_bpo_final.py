import streamlit as st
import pandas as pd

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

    # Identificar incontactables
    cantidad_incontactable = df[df["Agente BPO"] == "Agente Incontactable"].shape[0]
    indices_sin_asignar = df[df["Agente BPO"].isna()].index.tolist()
    total_distribuibles = len(indices_sin_asignar)

    # Agentes activos
    agentes_dia = [a for a in agentes_bpo if a != "Agente Incontactable"]
    otros_agentes = [a for a in agentes_dia if a != "Melissa Florian"]
    n_otros = len(otros_agentes)

    # Cálculo base
    x = total_distribuibles // (n_otros + 0.75)
    distribucion = {a: x for a in otros_agentes}
    distribucion["Melissa Florian"] = int(x * 0.75)

    # Redondeo y ajuste
    asignados = sum(distribucion.values())
    faltan = total_distribuibles - asignados

    # Asignar faltantes entre los agentes normales (excluyendo Melissa)
    agentes_ordenados = sorted(otros_agentes, key=lambda a: -distribucion[a])
    for i in range(faltan):
        distribucion[agentes_ordenados[i % len(agentes_ordenados)]] += 1

    # Asignar registros a cada agente
    for agente in distribucion:
        for _ in range(distribucion[agente]):
            if indices_sin_asignar:
                df.at[indices_sin_asignar.pop(0), "Agente BPO"] = agente

    # Tabla resumen
    resumen = df["Agente BPO"].value_counts().sort_index().reset_index()
    resumen.columns = ["Agente", "Cantidad"]
    resumen.loc[len(resumen.index)] = ["Total general", resumen["Cantidad"].sum()]

    st.subheader("📊 Resumen de distribución")
    st.dataframe(resumen, use_container_width=True)
    # Guardar cambios