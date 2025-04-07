
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta

# === CONFIGURACIÓN ===
archivo_entrada = "Programa.xlsx"
archivo_salida = f"Programa_Modificado_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

# === FUNCIONES ===

def remove_accents(text):
    if isinstance(text, str):
        return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

def asignar_fecha(row, fecha_actual, fecha_siguiente):
    if isinstance(row, str):
        valor = row.strip().lower()
        if valor == "ad":
            return fecha_actual.strftime("%d/%m/%Y")
        elif valor in ["od", "on demand", "bamx"]:
            return fecha_siguiente.strftime("%d/%m/%Y")
    try:
        fecha = pd.to_datetime(row)
        return fecha.strftime("%d/%m/%Y")
    except:
        return row

# === STREAMLIT INTERFAZ ===
st.set_page_config(layout="wide", page_title="Procesador de Archivos BPO", page_icon="📁")
st.image("images/bpo_character.png", width=220)

st.markdown("### 📁 Procesador de Archivos BPO")
st.write("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

uploaded_file = st.file_uploader("📤 Sube tu archivo Excel (.xlsx)", type=["xlsx"])
fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
meses_es = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun", 7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}
fecha_oportunidad = f"{fecha_actual.day}-{meses_es[fecha_actual.month]}-{fecha_actual.year}"
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
    df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
    df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
    df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
    df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

    df["Día de recolección"] = df["Día de recolección"].apply(lambda r: asignar_fecha(r, fecha_actual, fecha_siguiente))
    df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)
    df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
    df["Fecha de cierre"] = fecha_cierre
    df["Etapa"] = "Pendiente de Contacto"
    df["Agente BPO"] = ""

    # Reglas de asignación
    clientes_melissa_exclusivos = ["OXXO", "Axionlog"]
    df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_melissa_exclusivos), case=False, na=False), "Agente BPO"] = "Melissa Florian"

    clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
    df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False)].copy()

    agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
    asignaciones = df["Agente BPO"].value_counts().to_dict()
    for agente in agentes_bpo:
        if agente not in asignaciones:
            asignaciones[agente] = 0

    indices_repartir = df_repartir.index.tolist()
    for i, idx in enumerate(indices_repartir):
        agente = agentes_bpo[i % len(agentes_bpo)]
        df.at[idx, "Agente BPO"] = agente
        asignaciones[agente] += 1

    df_sin_asignar = df[df["Agente BPO"] == ""].copy()
    indices_sin_asignar = df_sin_asignar.index.tolist()
    registros_por_agente = len(df) // len(agentes_bpo)
    faltantes = {agente: max(0, registros_por_agente - asignaciones[agente]) for agente in agentes_bpo}

    for agente in agentes_bpo:
        for _ in range(faltantes[agente]):
            if indices_sin_asignar:
                df.at[indices_sin_asignar.pop(0), "Agente BPO"] = agente

    i = 0
    while indices_sin_asignar:
        idx = indices_sin_asignar.pop(0)
        agente = agentes_bpo[i % len(agentes_bpo)]
        df.at[idx, "Agente BPO"] = agente
        i += 1

    # Vista previa amigable
    st.subheader("👀 Vista previa")
    columnas_clave = [
        "Delv Ship-To Name", "Order Quantity", "Esquema", "Coordinador LT", 
        "Shpt Haulier Name", "Ejecutivo RBO", "Motivo", "Fecha de recolección"
    ]
    st.dataframe(df[columnas_clave], use_container_width=True)

    # Descarga
    df.to_excel(archivo_salida, index=False)
    with open(archivo_salida, "rb") as f:
        st.download_button("📥 Descargar archivo procesado", f, file_name=archivo_salida)

    st.markdown("---")
    st.markdown("❤️ Hecho con amor por el equipo de **BPO Innovations**")

