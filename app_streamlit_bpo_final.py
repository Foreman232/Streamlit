import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import time
import os

# Configuración de la página
st.set_page_config(layout="wide", page_title="🚀 Procesador Chep", page_icon="📊")

# Cabecera con imagen y títulos
col1, col2 = st.columns([1, 5])
with col1:
    st.image("images/bpo_character.png", width=100)
with col2:
    st.title("🚀 Procesador de Datos")
    st.caption("Automatiza limpieza de datos y asignación de agentes BPO para tu archivo Excel.")

with st.expander("ℹ️ ¿Qué hace esta herramienta?"):
    st.markdown("""
    - Descarga un archivo limpio, listo para usar.
    """)

# Selector de archivo
uploaded_file = st.file_uploader("📤 Sube tu archivo Excel para procesar", type=["xlsx"])

# Fechas
fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_oportunidad = f"{fecha_actual.day}-{fecha_actual.strftime('%b').lower()}-{fecha_actual.year}"
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

# Funciones de utilidad
def remove_accents(text):
    if isinstance(text, str):
        return ''.join(
            c for c in unicodedata.normalize('NFD', text) 
            if unicodedata.category(c) != 'Mn'
        )
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
        dfs = pd.read_excel(uploaded_file, sheet_name=None)
        columnas_obligatorias = ['Delv Ship-To Name', 'Esquema']
        hoja_correcta = None
        nombre_hoja_seleccionada = None
        for nombre, df_candidate in dfs.items():
            if set(columnas_obligatorias).issubset(df_candidate.columns):
                hoja_correcta = df_candidate
                nombre_hoja_seleccionada = nombre
                st.write(f"Se ha seleccionado la hoja: {nombre}")
                break
        if hoja_correcta is None:
            st.error("No se encontró una hoja que contenga las columnas necesarias.")
        else:
            df = hoja_correcta

        ############################################
        # 1. Limpieza y ajustes básicos
        ############################################
        df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(
            lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x
        )
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

        ######################################################
        # 2. Reemplazo manual (opcional)
        ######################################################
        st.markdown("### 👤 Reemplazo Manual de Agente (opcional)")
        col_excluir, col_reemplazo = st.columns(2)
        with col_excluir:
            agente_excluir = st.selectbox("Agente a excluir", ["Ninguno", "Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"])
        with col_reemplazo:
            agente_reemplazo = st.text_input("Reemplazo (nombre completo)", placeholder="Ej. Laura Martínez")

        # 2.2 Definir la lista de agentes BPO
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:
            agentes_bpo.append("Abigail Vasquez")

        if agente_excluir != "Ninguno" and agente_reemplazo.strip() != "":
            agentes_bpo = [agente_reemplazo if ag == agente_excluir else ag for ag in agentes_bpo]

        ######################################################
        # (Todo el resto de tu código sigue igual)
        ######################################################

        # Resto del procesamiento...
        # [Aquí va toda tu lógica de asignaciones y generación de archivos tal como ya está en tu código]

        # 🔁 Puedes seguir pegando el resto del código desde la sección:
        # # 2.3 Reglas especiales:
        # ... hasta el final ...

