import streamlit as st
import pandas as pd

# Datos de entrada simulados
data = {
    "Agente BPO": ["Agente Incontactable", "Alysson Garcia", "Ana Paniagua", "Julio de Leon", 
                   "Melissa Florian", "Nancy Zet"],
    "Cuenta de Agente BPO": [101, 194, 194, 194, 97, 194]
}
df = pd.DataFrame(data)

# Identificar registros de agentes incontactables
cantidad_incontactable = df[df["Agente BPO"] == "Agente Incontactable"]["Cuenta de Agente BPO"].sum()
indices_sin_asignar = df[df["Agente BPO"].isna()].index.tolist()
total_distribuibles = len(indices_sin_asignar)

# Agentes activos (sin el incontactable)
agentes_dia = [a for a in df["Agente BPO"].unique() if a != "Agente Incontactable"]
otros_agentes = [a for a in agenes_dia if a != "Melissa Florian"]
n_otros = len(otros_agentes)

# Calcular cantidad base para otros agentes
x = total_distribuibles // (n_otros + 0.75)  # 0.75 para Melissa
distribucion = {a: x for a in otros_agentes}
distribucion["Melissa Florian"] = int(x * 0.75)  # Melissa recibe 25% menos

# Redondeo y ajuste
asignados = sum(distribucion.values())
faltan = int(round(total_distribuibles - asignados))

# Repartir sobrantes entre los agentes normales (sin modificar Melissa)
agentes_ordenados = sorted(otros_agentes, key=lambda a: -distribucion[a])
for i in range(faltan):
    distribucion[agentes_ordenados[i % len(agentes_ordenados)]] += 1

# Asignar al DataFrame
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