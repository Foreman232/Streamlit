# --------------------------------------------------------------------------------------------------
# NOTA IMPORTANTE:
# Este código está super bien, no quiero que se le cambie nada.
# Únicamente se agrega un paso más:
# Cuando falte alguien, se podrá excluir manualmente y agregar a la persona que lo sustituirá.
# Esto se hará de forma manual ya que no hay una persona fija para cubrir.
# --------------------------------------------------------------------------------------------------

# ------------------------
# Librerías necesarias
# ------------------------
import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import time
import os

# ------------------------
# Configuración inicial
# ------------------------
st.set_page_config(layout="wide", page_title="🚀 Procesador Chep", page_icon="📊")

# ------------------------
# Cabecera visual
# ------------------------
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

# ------------------------
# Subir archivo
# ------------------------
uploaded_file = st.file_uploader("📤 Sube tu archivo Excel para procesar", type=["xlsx"])

# ------------------------
# Variables de fechas
# ------------------------
fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_oportunidad = f"{fecha_actual.day}-{fecha_actual.strftime('%b').lower()}-{fecha_actual.year}"
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

# ------------------------
# Funciones de utilidad
# ------------------------
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

# ------------------------
# Procesamiento principal
# ------------------------
if uploaded_file:
    with st.spinner("⏳ Procesando archivo..."):
        time.sleep(1)

        # Leer todas las hojas
        dfs = pd.read_excel(uploaded_file, sheet_name=None)

        # Buscar la hoja correcta
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

        # -----------------------------------------------------------
        # 1. LIMPIEZA Y AJUSTES BÁSICOS
        # -----------------------------------------------------------
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

        # -----------------------------------------------------------
        # 2. ASIGNACIONES ESPECIALES (Incontactables y Casos Especiales)
        # -----------------------------------------------------------

        # Definir agentes
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:
            agentes_bpo.append("Abigail Vasquez")

        # NUEVO PASO: Ajustes manuales de agentes (por ausencia)
        with st.expander("⚙️ Ajustes Manuales de Agentes (Opcional)"):
            st.markdown("""
            **NOTA:** Este código esta super bien, no quiero que se le cambie nada.
            Únicamente se agrega este paso manual para excluir agentes ausentes y agregar sustitutos.
            """)

            agentes_a_excluir = st.multiselect(
                "Selecciona los agentes que NO trabajarán hoy:",
                options=agentes_bpo,
                help="Estos agentes serán excluidos."
            )

            sustitutos_manual = st.text_input(
                "Escribe los nombres de los agentes sustitutos (separados por coma):",
                help="Ejemplo: Juan Pérez, Laura Gómez"
            )

            agentes_bpo = [ag for ag in agentes_bpo if ag not in agentes_a_excluir]
            if sustitutos_manual.strip():
                nuevos_sustitutos = [nombre.strip() for nombre in sustitutos_manual.split(",") if nombre.strip()]
                agentes_bpo.extend(nuevos_sustitutos)

            st.success(f"✅ Agentes activos hoy: {', '.join(agentes_bpo)}")

# -----------------------------------------------------------
# Continúa con la lógica normal: asignación especial, distribución round robin, balanceo, descarga...
# (¿Quieres que también te pase ya completo hasta la parte de descarga final?)
# -----------------------------------------------------------
