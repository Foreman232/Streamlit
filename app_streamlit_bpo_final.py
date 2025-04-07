
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import base64

st.set_page_config(layout="wide", page_title="Procesador BPO")

# Estilos personalizados
st.markdown(
    "<style>body { background-color: #1e1e1e; color: white; }</style>",
    unsafe_allow_html=True
)

# Layout con imágenes
col1, col2 = st.columns([1, 2])
with col1:
    st.image("images/bpo_character.png", width=250)
with col2:
    st.image("images/bpo_innovations_logo.jpg", width=100)
    st.markdown("## 📂 Procesador de Archivos BPO")
    st.markdown("Sube tu archivo Excel y descarga uno limpio con fechas correctas y agentes BPO asignados automáticamente.")

# Funciones auxiliares
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

# Subida de archivo
uploaded_file = st.file_uploader("📥 Sube tu archivo Excel", type=["xlsx"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Archivo cargado correctamente")

        st.markdown("### 👀 Vista previa")
        st.dataframe(df.head())

        # Procesamiento
        fecha_actual = datetime.today()
        fecha_siguiente = fecha_actual + timedelta(days=1)
        fecha_cierre = fecha_actual.strftime("%d/%m/%Y")
        meses_es = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun", 7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}
        mes_in_spanish = meses_es[fecha_actual.month]
        fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
        df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
        df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
        df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
        df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

        df["Día de recolección"] = df["Día de recolección"].apply(lambda row: asignar_fecha(row, fecha_actual, fecha_siguiente))
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)
        df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
        df["Fecha de cierre"] = fecha_cierre
        df["Etapa"] = "Pendiente de Contacto"
        df["Agente BPO"] = ""

        clientes_melissa_exclusivos = ["OXXO", "Axionlog"]
        df.loc[
            df["Nombre de oportunidad1"].str.contains('|'.join(clientes_melissa_exclusivos), case=False, na=False),
            "Agente BPO"
        ] = "Melissa Florian"

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
        total_registros = df.shape[0]
        registros_por_agente = total_registros // len(agentes_bpo)
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

        column_order = [
            "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
            "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
            "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
        ]
        df_final = df[column_order]
        output_file = "Programa_Modificado.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{output_file}">📥 Descargar archivo procesado</a>'
            st.markdown(href, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")

# Footer
st.markdown("---")
st.markdown("📍 Hecho por el equipo de **BPO Innovations**")
