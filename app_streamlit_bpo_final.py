
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import base64

st.set_page_config(page_title="Procesador BPO", layout="wide", page_icon="📁")

# Fondo oscuro y estilos personalizados
page_style = """
<style>
body {
    background-color: #1e1e1e;
    color: white;
}
h1, h2, h3, p, .stTextInput > label, .stFileUploader > label {
    color: white;
}
</style>
"""
st.markdown(page_style, unsafe_allow_html=True)

# Imagen
col1, col2 = st.columns([1, 2])
with col1:
    st.image("images/trayectoria.png", use_container_width=True)
with col2:
    st.markdown("### 📁 Procesador de Archivos BPO")
    st.write("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

# Subir archivo
uploaded_file = st.file_uploader("Carga el archivo Excel", type=["xlsx"])

# === FUNCIONES ===
def remove_accents(text):
    if isinstance(text, str):
        return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

def asignar_fecha(row):
    if isinstance(row, str):
        valor = row.strip().lower()
        if valor == "ad":
            return f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
        elif valor in ["od", "on demand", "bamx"]:
            return f"{fecha_siguiente.day}/{fecha_siguiente.month}/{fecha_siguiente.year}"
    try:
        fecha = pd.to_datetime(row)
        return f"{fecha.day}/{fecha.month}/{fecha.year}"
    except:
        return row

# Procesamiento
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Hoja1")

    # Fechas base
    fecha_actual = datetime.today()
    fecha_siguiente = fecha_actual + timedelta(days=1)
    fecha_cierre = f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
    meses_es = {
        1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
    }
    mes_in_spanish = meses_es[fecha_actual.month]
    fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

    # Limpieza básica
    df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
    df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
    df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
    df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
    df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

    # Fechas y campos adicionales
    df["Día de recolección"] = df["Día de recolección"].apply(asignar_fecha)
    df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)
    df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
    df["Fecha de cierre"] = fecha_cierre
    df["Etapa"] = "Pendiente de Contacto"
    df["Agente BPO"] = ""

    # Asignaciones
    agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
    asignaciones = df["Agente BPO"].value_counts().to_dict()
    for agente in agentes_bpo:
        if agente not in asignaciones:
            asignaciones[agente] = 0

    # Asignación específica
    df.loc[df["Nombre de oportunidad1"].str.contains("OXXO|Axionlog", case=False, na=False), "Agente BPO"] = "Melissa Florian"

    # Asignación proporcional
    indices_sin_asignar = df[df["Agente BPO"] == ""].index.tolist()
    for i, idx in enumerate(indices_sin_asignar):
        agente = agentes_bpo[i % len(agentes_bpo)]
        df.at[idx, "Agente BPO"] = agente

    # Reordenar columnas
    column_order = [
        "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
        "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
        "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
    ]
    df_final = df[column_order]

    # Exportar archivo
    output_filename = f"Programa_Modificado_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    df_final.to_excel(output_filename, index=False)
    with open(output_filename, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{output_filename}">📥 Descargar archivo procesado</a>'
        st.markdown(href, unsafe_allow_html=True)
