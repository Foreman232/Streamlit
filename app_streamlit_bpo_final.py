
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import unicodedata
import io

st.set_page_config(
    page_title="Procesador de Archivos BPO",
    layout="wide",
    page_icon="📊"
)

# Estilo personalizado
st.markdown('''
<style>
    body, .stApp {
        background-color: #1c1f26;
        color: white;
    }
    .css-18e3th9 {
        background-color: #1c1f26 !important;
    }
    h1, h2, h3, h4 {
        color: #ffffff;
    }
    .stButton>button {
        background-color: #00c7b7;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 1rem;
    }
    .stDownloadButton>button {
        background-color: #00c7b7;
        color: white;
        font-weight: bold;
        border-radius: 8px;
    }
</style>
''', unsafe_allow_html=True)

st.title("📊 Procesador de Archivos BPO")

# Confirmamos que se ve la imagen decorativa
st.image("images/trayectoria.png", use_container_width=True)

uploaded_file = st.file_uploader("📁 Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("✅ Archivo cargado correctamente")
    df = pd.read_excel(uploaded_file)

    def remove_accents(text):
        if isinstance(text, str):
            return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
        return text

    fecha_actual = datetime.today()
    fecha_siguiente = fecha_actual + timedelta(days=1)
    fecha_cierre = f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
    meses_es = {
        1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
    }
    mes_in_spanish = meses_es[fecha_actual.month]
    fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

    if "Esquema" in df:
        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
    if "Shpt Haulier Name" in df:
        df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
    if "Motivo" in df:
        df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")
    if "Día de recolección" in df:
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
        df["Día de recolección"] = df["Día de recolección"].apply(asignar_fecha)
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)

    if "Delv Ship-To Name" in df:
        df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
    df["Fecha de cierre"] = fecha_cierre
    df["Etapa"] = "Pendiente de Contacto"
    df["Agente BPO"] = ""

    agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
    clientes_melissa_exclusivos = ["OXXO", "Axionlog"]
    df.loc[
        df["Nombre de oportunidad1"].str.contains('|'.join(clientes_melissa_exclusivos), case=False, na=False),
        "Agente BPO"
    ] = "Melissa Florian"

    df_sin_agente = df[df["Agente BPO"] == ""].copy()
    indices = df_sin_agente.index.tolist()
    for i, idx in enumerate(indices):
        agente = agentes_bpo[i % len(agentes_bpo)]
        df.at[idx, "Agente BPO"] = agente

    columnas_orden = [col for col in [
        "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
        "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
        "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
    ] if col in df.columns]
    df_final = df[columnas_orden]

    output = io.BytesIO()
    df_final.to_excel(output, index=False)
    st.download_button(
        label="📥 Descargar archivo procesado",
        data=output.getvalue(),
        file_name="archivo_procesado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
