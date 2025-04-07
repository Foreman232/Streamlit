
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# ==== CONFIGURACIÓN DE PÁGINA ====
st.set_page_config(page_title="Procesador BPO", layout="wide")

# ==== ESTILOS PERSONALIZADOS ====
st.markdown("""
<style>
    .css-18e3th9 {
        padding: 2rem 1rem 2rem 1rem;
    }
    .stButton>button {
        background-color: #00c7b7;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 0.6em 1.5em;
    }
    footer {
        visibility: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ==== ENCABEZADO ====
col1, col2, col3 = st.columns([1,1,2])
with col1:
    st.image("images/trayectoria.png", use_column_width=True)
with col2:
    st.image("images/bpo_innovations_logo.jpg", width=160)
with col3:
    st.markdown("### 📊 Procesador de Archivos BPO")
    st.markdown("Sube tu archivo Excel y descarga uno limpio con fechas corregidas y agentes BPO asignados automáticamente.")

st.markdown("---")

# ==== SUBIDA DE ARCHIVO ====
archivo = st.file_uploader("📁 Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if archivo:
    try:
        df = pd.read_excel(archivo)
        fecha_actual = datetime.today()
        fecha_siguiente = fecha_actual + timedelta(days=1)
        meses_es = {1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
                    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"}
        fecha_cierre = f"{fecha_actual.day}/{fecha_actual.month}/{fecha_actual.year}"
        fecha_oportunidad = f"{fecha_actual.day}-{meses_es[fecha_actual.month]}-{fecha_actual.year}"

        def asignar_fecha(valor):
            if isinstance(valor, str):
                valor = valor.strip().lower()
                if valor == "ad":
                    return fecha_actual.strftime("%-d/%-m/%Y")
                elif valor in ["od", "on demand", "bamx"]:
                    return fecha_siguiente.strftime("%-d/%-m/%Y")
            try:
                fecha = pd.to_datetime(valor)
                return fecha.strftime("%-d/%-m/%Y")
            except:
                return valor

        df["Día de recolección"] = df["Día de recolección"].apply(asignar_fecha)
        df.rename(columns={"Día de recolección": "Fecha de recolección"}, inplace=True)
        df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
        df["Fecha de cierre"] = fecha_cierre
        df["Etapa"] = "Pendiente de Contacto"
        df["Agente BPO"] = ""

        # Agentes y lógica de asignación
        clientes_melissa_exclusivos = ["OXXO", "Axionlog"]
        df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_melissa_exclusivos), case=False, na=False),
               "Agente BPO"] = "Melissa Florian"

        agentes = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        sin_asignar = df[df["Agente BPO"] == ""].copy()
        repartidos = sin_asignar.index.tolist()
        for i, idx in enumerate(repartidos):
            df.at[idx, "Agente BPO"] = agentes[i % len(agentes)]

        columnas_final = [
            "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
            "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
            "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
        ]
        df_final = df[columnas_final]

        st.success("✅ Archivo procesado con éxito")

        output = BytesIO()
        df_final.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button("📥 Descargar archivo procesado", data=output,
                           file_name=f"Archivo_Procesado_{datetime.today().strftime('%d-%m-%Y')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
