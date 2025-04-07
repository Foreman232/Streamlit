import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(layout="wide", page_title="Procesador de Archivos BPO", page_icon="📊")

st.markdown("<style>body {background-color: #1e1e1e; color: white; font-family: 'Segoe UI', sans-serif;}</style>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 2])
with col1:
    st.image("images/trayectoria.png", width=350)

with col2:
    st.image("images/bpo_logo.png", width=100)
    st.markdown("<h1 style='color: white;'>📊 Procesador de Archivos BPO</h1>", unsafe_allow_html=True)
    st.write("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

    uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        fecha_actual = datetime.today()
        fecha_siguiente = fecha_actual + timedelta(days=1)
        fecha_cierre = f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
        meses_es = {1:"ene",2:"feb",3:"mar",4:"abr",5:"may",6:"jun",7:"jul",8:"ago",9:"sep",10:"oct",11:"nov",12:"dic"}
        fecha_oportunidad = f"{fecha_actual.day}-{meses_es[fecha_actual.month]}-{fecha_actual.year}"

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

        exclusivos = ["OXXO", "Axionlog"]
        df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(exclusivos), case=False, na=False), "Agente BPO"] = "Melissa Florian"

        a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(a_repartir), case=False, na=False)].copy()

        agentes = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        asignaciones = df["Agente BPO"].value_counts().to_dict()
        for a in agentes:
            asignaciones[a] = asignaciones.get(a, 0)

        indices_repartir = df_repartir.index.tolist()
        for i, idx in enumerate(indices_repartir):
            agente = agentes[i % len(agentes)]
            df.at[idx, "Agente BPO"] = agente
            asignaciones[agente] += 1

        df_sin_asignar = df[df["Agente BPO"] == ""].copy()
        indices_sin_asignar = df_sin_asignar.index.tolist()
        total = df.shape[0]
        por_agente = total // len(agentes)
        faltantes = {a: max(0, por_agente - asignaciones[a]) for a in agentes}

        for a in agentes:
            for _ in range(faltantes[a]):
                if indices_sin_asignar:
                    df.at[indices_sin_asignar.pop(0), "Agente BPO"] = a

        i = 0
        while indices_sin_asignar:
            idx = indices_sin_asignar.pop(0)
            df.at[idx, "Agente BPO"] = agentes[i % len(agentes)]
            i += 1

        columnas = [
            "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
            "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
            "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
        ]
        df_final = df[columnas]
        output = BytesIO()
        df_final.to_excel(output, index=False)
        st.success("✅ Archivo procesado exitosamente")
        st.download_button("📥 Descargar archivo procesado", output.getvalue(), file_name="Archivo_Procesado.xlsx")

st.markdown("<br><center style='color:gray;'>Hecho con 💚 por el equipo de BPO Innovations</center>", unsafe_allow_html=True)
