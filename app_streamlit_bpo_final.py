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
        # Leer todas las hojas del Excel
        dfs = pd.read_excel(uploaded_file, sheet_name=None)
        columnas_obligatorias = ['Delv Ship-To Name', 'Esquema']
        hoja_correcta = None
        for nombre, df_candidate in dfs.items():
            if set(columnas_obligatorias).issubset(df_candidate.columns):
                hoja_correcta = df_candidate
                st.write(f"Se ha seleccionado la hoja: {nombre}")
                break
        if hoja_correcta is None:
            st.error("No se encontró una hoja que contenga las columnas necesarias.")
        else:
            df = hoja_correcta

        # Limpieza
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

        # Incontactables
        if os.path.exists("Incontactables.xlsx"):
            try:
                df_incontactables = pd.read_excel("Incontactables.xlsx", sheet_name=0)
                df["Delv Ship-To Party"] = df["Delv Ship-To Party"].astype(str)
                df_incontactables["Delv Ship-To Party"] = df_incontactables["Delv Ship-To Party"].astype(str)
                df.loc[
                    df["Delv Ship-To Party"].isin(df_incontactables["Delv Ship-To Party"]), 
                    "Agente BPO"
                ] = "Agente Incontactable"
            except Exception as e:
                st.warning(f"No se pudo procesar 'Incontactables.xlsx'. Error: {e}")

        # Lista inicial de agentes
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:
            agentes_bpo.append("Abigail Vasquez")

        # Selección de reemplazo si alguien no está
        st.subheader("👥 Agentes disponibles hoy")
        agente_faltante = st.selectbox("¿Hay algún agente ausente hoy?", ["Ninguno"] + agentes_bpo)
        reemplazo_manual = None
        if agente_faltante != "Ninguno":
            agentes_bpo.remove(agente_faltante)
            reemplazo_manual = st.selectbox(f"Selecciona el reemplazo para {agente_faltante}:", agentes_bpo)

        # Reglas especiales
        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[
            df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), 
            "Agente BPO"
        ] = "Melissa Florian"
        df.loc[
            df["Motivo"].str.contains("adicionales", case=False, na=False), 
            "Agente BPO"
        ] = "Ana Paniagua"

        clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False)].copy()
        indices_repartir = df_repartir[df_repartir["Agente BPO"] == ""].index.tolist()
        for i, idx in enumerate(indices_repartir):
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente

        # Cupo teórico
        forzadas_por_agente = df[df["Agente BPO"] != ""].groupby("Agente BPO").size().to_dict()
        total_general = df.shape[0]
        incontactables = forzadas_por_agente.get("Agente Incontactable", 0)
        remainder = total_general - incontactables
        n_agentes = len(agentes_bpo)
        x = remainder / (n_agentes - 0.25)
        cupo_teorico = {}
        for agente in agentes_bpo:
            if agente == "Melissa Florian":
                cupo_teorico[agente] = int(0.75 * x)
            else:
                cupo_teorico[agente] = int(x)
        if reemplazo_manual:
            cupo_teorico[reemplazo_manual] += cupo_teorico.get(agente_faltante, 0)

        filas_adicionales_para = {}
        for agente in agentes_bpo:
            ya_forzadas = forzadas_por_agente.get(agente, 0)
            disponible = cupo_teorico[agente] - ya_forzadas
            filas_adicionales_para[agente] = max(disponible, 0)

        df_sin_asignar = df[df["Agente BPO"] == ""].copy()
        indices_sin_asignar = list(df_sin_asignar.index)
        while indices_sin_asignar:
            for agente in agentes_bpo:
                if not indices_sin_asignar or filas_adicionales_para[agente] <= 0:
                    continue
                idx = indices_sin_asignar.pop(0)
                df.at[idx, "Agente BPO"] = agente
                filas_adicionales_para[agente] -= 1
            else:
                if indices_sin_asignar:
                    for idx in indices_sin_asignar:
                        agente_max = max(filas_adicionales_para, key=filas_adicionales_para.get)
                        df.at[idx, "Agente BPO"] = agente_max
                    break

        # Balanceo
        agentes_balancear = [a for a in agentes_bpo if a != "Melissa Florian"]
        conteo_final = df["Agente BPO"].value_counts().to_dict()
        promedio_aprox = sum(conteo_final.get(a, 0) for a in agentes_balancear) // len(agentes_balancear)
        agentes_por_encima = [(a, conteo_final.get(a, 0) - promedio_aprox) for a in agentes_balancear if conteo_final.get(a, 0) > promedio_aprox + 1]
        agentes_por_debajo = [(a, promedio_aprox - conteo_final.get(a, 0)) for a in agentes_balancear if conteo_final.get(a, 0) < promedio_aprox]

        for ag_enc, sobrante in agentes_por_encima:
            for ag_fal, faltante in agentes_por_debajo:
                while sobrante > 0 and faltante > 0:
                    idx_mover = df[df["Agente BPO"] == ag_enc].index[0]
                    df.at[idx_mover, "Agente BPO"] = ag_fal
                    sobrante -= 1
                    faltante -= 1

        # Mostrar resultados
        conteo_final = df["Agente BPO"].value_counts().to_dict()
        resumen_html = "<div class='resumen-container'><div class='resumen-title'>📊 Resumen de Distribución Final</div>"
        for agente in agentes_bpo:
            resumen_html += f"<div class='resumen-item'><strong>{agente}:</strong> {conteo_final.get(agente, 0)} (máx teórico: {cupo_teorico.get(agente, 'N/A')})</div>"
        resumen_html += f"<div class='resumen-item'><strong>Agente Incontactable:</strong> {conteo_final.get('Agente Incontactable', 0)}</div>"
        resumen_html += f"<div class='resumen-item'><strong>Total:</strong> {df.shape[0]}</div></div>"

        st.markdown("""
        <style>
        .resumen-container {
            background: #f7f9fc;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .resumen-title {
            font-size: 1.25rem;
            font-weight: bold;
        }
        .resumen-item {
            margin: 5px 0;
        }
        </style>
        """, unsafe_allow_html=True)
        st.markdown(resumen_html, unsafe_allow_html=True)

        columnas_finales = [
            'Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr',
            'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo',
            'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO'
        ]
        df_final = df[[col for col in columnas_finales if col in df.columns]]

        st.markdown("### 👀 Vista previa")
        st.dataframe(df_final.head(15), height=500, use_container_width=True)

        now_str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        excel_filename = f"Programa_Modificado_{now_str}.xlsx"
        csv_filename = f"Programa_Modificado_{now_str}.csv"
        df_final.to_excel(excel_filename, index=False)
        df_final.to_csv(csv_filename, index=False)

        col1, col2 = st.columns(2)
        with col1:
            with open(excel_filename, "rb") as f:
                st.download_button("📥 Descargar Excel", f, file_name=excel_filename)
        with col2:
            with open(csv_filename, "rb") as f:
                st.download_button("📥 Descargar CSV", f, file_name=csv_filename)

st.caption("🚀 Creado por el equipo de BPO Innovations")


