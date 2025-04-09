
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import time
import os

st.set_page_config(layout="wide", page_title="📁 Procesador BPO", page_icon="📊")

col1, col2 = st.columns([1, 5])
with col1:
    st.image("images/bpo_character.png", width=100)
with col2:
    st.title("📁 Procesador BPO")
    st.caption("Automatiza limpieza de datos y asignación de agentes BPO para tu archivo Excel.")

with st.expander("ℹ️ ¿Qué hace esta herramienta?"):
    st.markdown("""
    - Corrige campos vacíos o incorrectos.
    - Asigna automáticamente agentes BPO.
    - Detecta y etiqueta como 'Incontactables' según lista externa.
    - Asigna a Ana Paniagua todos los registros con motivo 'Adicionales'.
    - Agrega a Abigail Vasquez a la distribución si es sábado.
    - Descarga un archivo limpio, listo para usar.
    """)

uploaded_file = st.file_uploader("📤 Sube tu archivo Excel para procesar", type=["xlsx"])

fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_oportunidad = f"{fecha_actual.day}-{fecha_actual.strftime('%b').lower()}-{fecha_actual.year}"
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

def remove_accents(text):
    if isinstance(text, str):
        return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

def asignar_fecha(row):
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

if uploaded_file:
    with st.spinner("⏳ Procesando archivo..."):
        time.sleep(1)
        df = pd.read_excel(uploaded_file)
        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
        df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
        df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
        df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
        df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

        df["Día de recolección"] = df["Día de recolección"].apply(asignar_fecha)
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)
        df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
        df["Fecha de cierre"] = fecha_cierre
        df["Etapa"] = "Pendiente de Contacto"
        df["Agente BPO"] = ""

        if os.path.exists("Incontactables.xlsx"):
            try:
                df_incontactables = pd.read_excel("Incontactables.xlsx", sheet_name=0)
                df["Delv Ship-To Party"] = df["Delv Ship-To Party"].astype(str)
                df_incontactables["Delv Ship-To Party"] = df_incontactables["Delv Ship-To Party"].astype(str)
                df.loc[df["Delv Ship-To Party"].isin(df_incontactables["Delv Ship-To Party"]), "Agente BPO"] = "Agente Incontactable"
            except Exception as e:
                st.warning(f"No se pudo procesar 'Incontactables.xlsx'. Error: {e}")
        else:
            st.info("Puedes subir manualmente 'Incontactables.xlsx' a la raíz del proyecto en Streamlit Cloud si deseas usarlo.")

        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:  # sábado
            agentes_bpo.append("Abigail Vasquez")

        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), "Agente BPO"] = "Melissa Florian"

        df.loc[df["Motivo"].str.contains("adicionales", case=False, na=False), "Agente BPO"] = "Ana Paniagua"

        # Distribución personalizada después de asignar "Agente Incontactable"
total_registros = len(df)
incontactables = df[df["Agente BPO"] == "Agente Incontactable"]
total_incontactables = len(incontactables)
total_para_repartir = total_registros - total_incontactables

# Lista base de agentes
agentes_base = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
if fecha_actual.weekday() == 5:  # sábado
    agentes_base.append("Abigail Vasquez")

# Aplicar 25% menos a Melissa
if "Melissa Florian" in agentes_base:
    n_agentes = len(agentes_base)
    peso_normal = 1
    peso_melissa = 0.75
    pesos = {agente: peso_normal for agente in agentes_base}
    pesos["Melissa Florian"] = peso_melissa

    suma_pesos = sum(pesos.values())
    asignaciones_personalizadas = {agente: int((peso / suma_pesos) * total_para_repartir) for agente, peso in pesos.items()}
else:
    n_agentes = len(agentes_base)
    asignaciones_personalizadas = {agente: total_para_repartir // n_agentes for agente in agentes_base}

# Rellenar distribución en el dataframe
df.loc[df["Agente BPO"] != "Agente Incontactable", "Agente BPO"] = ""  # Reinicia BPO asignados

sin_asignar = df[df["Agente BPO"] == ""].copy()
indices_sin_asignar = sin_asignar.index.tolist()

for agente in agentes_base:
    cantidad = asignaciones_personalizadas[agente]
    for _ in range(cantidad):
        if indices_sin_asignar:
            idx = indices_sin_asignar.pop(0)
            df.at[idx, "Agente BPO"] = agente

# Si quedaron sin asignar por redondeo, asignamos en orden cíclico
i = 0
while indices_sin_asignar:
    agente = agentes_base[i % len(agentes_base)]
    idx = indices_sin_asignar.pop(0)
    df.at[idx, "Agente BPO"] = agente
    i += 1

        st.success("✅ Archivo procesado con éxito")
        st.markdown("### 👀 Vista previa")
        st.dataframe(df.head(15), height=500, use_container_width=True)

        output_file = "Programa_Modificado.xlsx"
        columnas_finales = [
            'Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr',
            'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo',
            'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO'
        ]
        df = df[[col for col in columnas_finales if col in df.columns]]
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Descargar Programa_Modificado.xlsx",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.markdown("---")
st.caption("🚀 Creado por el equipo de BPO Innovations")
