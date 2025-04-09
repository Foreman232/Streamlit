# Distribución automática de registros (Melissa con 25% menos)
cantidad_incontactable = df[df["Agente BPO"] == "Agente Incontactable"].shape[0]
indices_sin_asignar = df[df["Agente BPO"].isna()].index.tolist()
total_distribuibles = len(indices_sin_asignar)

# Agentes activos (sin el incontactable)
agentes_dia = [a for a in agentes_bpo if a != "Agente Incontactable"]
otros_agentes = [a for a in agentes_dia if a != "Melissa Florian"]
n_otros = len(otros_agentes)

# Calcular cantidad base para otros agentes
x = total_distribuibles // (n_otros + 0.75)
distribucion = {a: x for a in otros_agentes}
distribucion["Melissa Florian"] = int(x * 0.75)  # Melissa recibe 25% menos

# Redondeo y ajuste
asignados = sum(distribucion.values())
faltan = total_distribuibles - asignados

# Asignar faltantes entre los agentes normales (excluyendo Melissa)
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