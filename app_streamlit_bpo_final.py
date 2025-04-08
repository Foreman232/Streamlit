
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import time
import os

# ----------------------- CONFIGURACIÓN INICIAL ----------------------- #
st.set_page_config(layout="wide", page_title="📁 Procesador BPO", page_icon="📊")

# Tema claro forzado (evita fondo negro en algunos despliegues)
st.markdown("""
    <style>
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #f5f5f5;
        color: #1c1c1c;
    }
    </style>
""", unsafe_allow_html=True)

# ----------------------- ENCABEZADO ----------------------- #
col1, col2 = st.columns([1, 5])
with col1:
    st.image("images/bpo_character.png", width=90)
with col2:
    st.markdown("""
    <div style='text-align: left;'>
        <h1 style='font-size: 2.5em; margin-bottom: 0;'>📁 Procesador BPO</h1>
        <p style='font-size: 1.1em; color: gray;'>Automatiza limpieza de datos y asignación de agentes BPO para tu archivo Excel</p>
    </div>
    """, unsafe_allow_html=True)

# ----------------------- MENÚ LATERAL ----------------------- #
st.sidebar.image("images/bpo_logo.png", width=150)
opcion = st.sidebar.radio("Navegación", ["Subir archivo", "Vista previa", "Resumen"])

# ----------------------- FUNCIÓN AUXILIAR ----------------------- #
def remove_accents(text):
    if isinstance(text, str):
        return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text

# ----------------------- BLOQUE PRINCIPAL ----------------------- #
fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_oportunidad = f"{fecha_actual.day}-{fecha_actual.strftime('%b').lower()}-{fecha_actual.year}"
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

uploaded_file = st.file_uploader("📥 Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("🎯 Analizando datos, por favor espera..."):
        time.sleep(1)
        df = pd.read_excel(uploaded_file)

        # Validación de columnas obligatorias
        columnas_requeridas = ["Delv Ship-To Name", "Motivo", "Delv Ship-To Party"]
        for col in columnas_requeridas:
            if col not in df.columns:
                st.error(f"❌ Falta la columna obligatoria: {col}")
                st.stop()

        # Limpieza y estandarización
        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
        df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
        df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
        df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
        df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

        df["Día de recolección"] = df["Día de recolección"].apply(lambda row: fecha_siguiente.strftime("%d/%m/%Y") if str(row).strip().lower() in ["od", "on demand", "bamx"] else fecha_actual.strftime("%d/%m/%Y") if str(row).strip().lower() == "ad" else row)
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)

        df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
        df["Fecha de cierre"] = fecha_cierre
        df["Etapa"] = "Pendiente de Contacto"
        df["Agente BPO"] = ""

        # Incontactables
        if os.path.exists("Incontactables.xlsx"):
            try:
                df_incontactables = pd.read_excel("Incontactables.xlsx", sheet_name=0)
                df["Delv Ship-To Party"] = df["Delv Ship-To Party"].astype(str).str.strip()
                df_incontactables["Delv Ship-To Party"] = df_incontactables["Delv Ship-To Party"].astype(str).str.strip()
                df.loc[df["Delv Ship-To Party"].isin(df_incontactables["Delv Ship-To Party"]), "Agente BPO"] = "Incontactables"
            except Exception as e:
                st.warning(f"No se pudo procesar 'Incontactables.xlsx'. Error: {e}")

        # Ana Paniagua si motivo = adicionales
        df["Motivo"] = df["Motivo"].astype(str).str.lower().str.strip()
        df.loc[df["Motivo"] == "adicionales", "Agente BPO"] = "Ana Paniagua"

        # Lista de agentes
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:  # sábado
            agentes_bpo.append("Abigail Vasquez")

        # Melissa exclusiva
        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), "Agente BPO"] = "Melissa Florian"

        # Reparto por clientes especiales
        clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False) & (df["Agente BPO"] == "")].copy()
        asignaciones = df["Agente BPO"].value_counts().to_dict()
        for agente in agentes_bpo:
            if agente not in asignaciones:
                asignaciones[agente] = 0
        indices_repartir = df_repartir.index.tolist()
        for i, idx in enumerate(indices_repartir):
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente
            asignaciones[agente] += 1

        # Resto de distribución
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

        # Mostrar vista previa
        if opcion == "Vista previa":
            st.markdown("### 👀 Vista previa")
            st.dataframe(df.head(10), use_container_width=True, height=400)

        # Mostrar resumen
        if opcion == "Resumen":
            st.markdown("## 📊 Resumen")
            st.metric("Registros totales", len(df))
            st.metric("Incontactables", df[df["Agente BPO"] == "Incontactables"].shape[0])
            st.metric("Asignados a Ana (Adicionales)", df[df["Agente BPO"] == "Ana Paniagua"].shape[0])

        # Descargar archivo
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
                label="📥 Descargar archivo procesado",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("✅ Archivo procesado con éxito")
        st.image("images/bpo_logo.png", width=100)

