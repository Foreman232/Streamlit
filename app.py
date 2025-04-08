
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import unicodedata

st.set_page_config(page_title="Procesador BPO", layout="centered")
st.title("📁 Procesador BPO")

st.image("images/bpo_character.png", width=150)
st.caption("Automatiza la limpieza de datos y asignación de agentes BPO")

fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")
fecha_oportunidad = f"{fecha_actual.day}-{fecha_actual.strftime('%b').lower()}-{fecha_actual.year}"

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

archivo = st.file_uploader("📤 Sube tu archivo Excel", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo)
    df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR")
    df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR")
    df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
    df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR")
    df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents)

    if "Día de recolección" in df.columns:
        df["Día de recolección"] = df["Día de recolección"].apply(asignar_fecha)
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)

    df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
    df["Fecha de cierre"] = fecha_cierre
    df["Etapa"] = "Pendiente de Contacto"
    df["Agente BPO"] = ""

    # Distribución de agentes
    agentes = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
    exclusivas = ["OXXO", "Axionlog"]
    df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas), case=False, na=False), "Agente BPO"] = "Melissa Florian"

    a_repartir = df[df["Agente BPO"] == ""].copy()
    indices = a_repartir.index.tolist()
    for i, idx in enumerate(indices):
        df.at[idx, "Agente BPO"] = agentes[i % len(agentes)]

    # Eliminar columnas extras
    eliminar = [
        "Código de llamada", "Cantidad Confirmada", "Persona", "Puesto", "Comentarios",
        "Del Block", "Diferencia de Pallets", "Porcentaje de variación", "Variación",
        "Coordinador LT.1", "Respuesta", "Comentarios adicoinales", "Seguimiento"
    ]
    for col in eliminar:
        if col in df.columns:
            df.drop(columns=col, inplace=True)

    st.success("✅ Procesado con éxito")
    st.dataframe(df.head(10))
    df.to_excel("Programa_Procesado.xlsx", index=False)
    with open("Programa_Procesado.xlsx", "rb") as f:
        st.download_button("📥 Descargar archivo procesado", f, file_name="Programa_Procesado.xlsx")
