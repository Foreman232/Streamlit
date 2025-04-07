
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Procesador BPO", layout="wide", page_icon="📊")

st.markdown(
    """
    <style>
    body {
        background-color: #0e1117;
        color: white;
    }
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    """,
    unsafe_allow_html=True,
)

st.image("images/trayectoria.png", use_column_width=True)

st.title("📊 Procesador de Archivos BPO")
st.markdown("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

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

    # === EJECUCIÓN PRINCIPAL ===
    fecha_actual = datetime.today()
    fecha_siguiente = fecha_actual + timedelta(days=1)
    fecha_cierre = f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
    meses_es = {
        1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
    }
    mes_in_spanish = meses_es[fecha_actual.month]
    fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

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

    agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
    df_shuffled = df.sample(frac=1, random_state=42).reset_index(drop=True)
    total = len(df_shuffled)
    por_agente = total // len(agentes_bpo)
    extras = total % len(agentes_bpo)

    idx = 0
    for agente in agentes_bpo:
        asignaciones = por_agente + (1 if extras > 0 else 0)
        extras -= 1 if extras > 0 else 0
        for _ in range(asignaciones):
            if idx < total:
                df_shuffled.at[idx, "Agente BPO"] = agente
                idx += 1

    columnas_final = [
        "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
        "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
        "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
    ]
    df_final = df_shuffled[columnas_final]
    
    output = BytesIO()
    df_final.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ Archivo procesado con éxito.")
    st.download_button(
        label="📥 Descargar archivo procesado",
        data=output,
        file_name=f"Programa_Modificado_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
