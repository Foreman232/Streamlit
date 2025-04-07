import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import unicodedata

# Configuración visual
st.set_page_config(page_title="Procesador BPO", layout="wide", page_icon="📊")

# Título e instrucciones
col1, col2 = st.columns([1, 2])
with col1:
    st.image("images/trayectoria.png", width=300)
with col2:
    st.markdown("### 📁 Procesador de Archivos BPO")
    st.markdown("Sube tu archivo Excel y descarga uno limpio con fechas correctas y agentes BPO asignados automáticamente.")
    archivo = st.file_uploader("📤 Sube tu archivo Excel", type=["xlsx"])
    if archivo:
        df = pd.read_excel(archivo)
        df["Fecha Procesada"] = datetime.today().strftime("%d/%m/%Y")
        nombre_salida = f"Archivo_Procesado_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
        df.to_excel(nombre_salida, index=False)
        with open(nombre_salida, "rb") as f:
            st.download_button("📥 Descargar archivo procesado", f, file_name=nombre_salida)
    st.markdown("Hecho con ❤️ por el equipo de BPO Innovations")