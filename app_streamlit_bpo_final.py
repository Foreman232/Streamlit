
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO

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

# Interfaz
st.set_page_config(page_title="Procesador de Archivos BPO", layout="wide", page_icon="📊")

col1, col2 = st.columns([1, 2])

with col1:
    st.image("images/trayectoria.png", width=300)

with col2:
    st.image("images/bpo_logo.png", width=100)
    st.markdown("## 📁 Procesador de Archivos BPO")
    st.markdown("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

    uploaded_file = st.file_uploader("📥 Sube tu archivo Excel (.xlsx)", type=["xlsx"])

    if uploaded_file:
        fecha_actual = datetime.today()
        fecha_siguiente = fecha_actual + timedelta(days=1)
        fecha_cierre = fecha_actual.strftime("%d/%m/%Y")
        meses_es = {
            1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
            7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
        }
        mes_in_spanish = meses_es[fecha_actual.month]
        fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

        df = pd.read_excel(uploaded_file, sheet_name="Hoja1")
        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
        df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
        df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
        df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
        df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

        df["Día de recolección"] = df["Día de recolección"].apply(lambda x: asignar_fecha(x, fecha_actual, fecha_siguiente))
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

        output = BytesIO()
        df_final.to_excel(output, index=False)
        st.success("✅ Archivo procesado con éxito")
        st.download_button(label="📥 Descargar archivo procesado", data=output.getvalue(), file_name="Programa_Procesado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    