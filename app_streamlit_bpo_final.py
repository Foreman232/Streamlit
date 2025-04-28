import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime, timedelta
import time
import os

# Configuración de la página
st.set_page_config(layout="wide", page_title="🚀 Procesador Chep", page_icon="📈")

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
uploaded_file = st.file_uploader("📄 Sube tu archivo Excel para procesar", type=["xlsx"])

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

        # Agentes
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:
            agentes_bpo.append("Abigail Vasquez")

        # Ajustes manuales de agentes
        with st.expander("⚙️ Ajustes de Agentes Manuales (Opcional)"):
            st.markdown("""
            - **Agentes ausentes**: selecciona quienes no trabajarán hoy.
            - **Sustitutos**: escribe los nombres de quienes los cubrirán.
            """)
            agentes_a_excluir = st.multiselect(
                "Selecciona los agentes AUSENTES hoy:",
                options=agentes_bpo,
                help="Estos agentes no recibirán registros hoy."
            )
            sustitutos_manual = st.text_input(
                "Escribe el nombre de los AGENTES SUSTITUTOS (separados por comas):",
                help="Ejemplo: Juan Pérez, Laura Gómez"
            )
            agentes_bpo = [ag for ag in agentes_bpo if ag not in agentes_a_excluir]
            if sustitutos_manual.strip():
                nuevos_sustitutos = [nombre.strip() for nombre in sustitutos_manual.split(",") if nombre.strip()]
                agentes_bpo.extend(nuevos_sustitutos)
            st.success(f"✅ Agentes activos para la asignación: {', '.join(agentes_bpo)}")

        # Reglas especiales
        if os.path.exists("Incontactables.xlsx"):
            try:
                df_incontactables = pd.read_excel("Incontactables.xlsx", sheet_name=0)
                df["Delv Ship-To Party"] = df["Delv Ship-To Party"].astype(str)
                df_incontactables["Delv Ship-To Party"] = df_incontactables["Delv Ship-To Party"].astype(str)
                df.loc[df["Delv Ship-To Party"].isin(df_incontactables["Delv Ship-To Party"]), "Agente BPO"] = "Agente Incontactable"
            except Exception as e:
                st.warning(f"No se pudo procesar 'Incontactables.xlsx'. Error: {e}")

        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), "Agente BPO"] = "Melissa Florian"
        df.loc[df["Motivo"].str.contains("adicionales", case=False, na=False), "Agente BPO"] = "Ana Paniagua"

        clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False)].copy()
        indices_repartir = df_repartir[df_repartir["Agente BPO"] == ""].index.tolist()
        for i, idx in enumerate(indices_repartir):
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente

        # Asignación normal
        forzadas_por_agente = df[df["Agente BPO"] != ""].groupby("Agente BPO").size().to_dict()
        total_general = df.shape[0]
        incontactables = forzadas_por_agente.get("Agente Incontactable", 0)
        remainder = total_general - incontactables
        n_agentes = len(agentes_bpo)
        x = remainder / (n_agentes - 0.25)

        cupo_teorico = {}
        for agente in agentes_bpo:
            cupo_teorico[agente] = int(0.75 * x) if agente == "Melissa Florian" else int(x)

        filas_adicionales_para = {}
        for agente in agentes_bpo:
            ya_forzadas = forzadas_por_agente.get(agente, 0)
            disponible = cupo_teorico[agente] - ya_forzadas
            filas_adicionales_para[agente] = max(disponible, 0)

        df_sin_asignar = df[df["Agente BPO"] == ""].copy()
        indices_sin_asignar = list(df_sin_asignar.index)

        while indices_sin_asignar:
            asignado_en_ronda = False
            for agente in agentes_bpo:
                if not indices_sin_asignar:
                    break
                if filas_adicionales_para.get(agente, 0) > 0:
                    idx = indices_sin_asignar.pop(0)
                    df.at[idx, "Agente BPO"] = agente
                    filas_adicionales_para[agente] -= 1
                    asignado_en_ronda = True
            if not asignado_en_ronda and indices_sin_asignar:
                for idx in indices_sin_asignar:
                    agente_max = max(filas_adicionales_para, key=filas_adicionales_para.get)
                    df.at[idx, "Agente BPO"] = agente_max
                break

        # Balanceo
        agentes_balancear = [ag for ag in agentes_bpo if ag != "Melissa Florian"]
        conteo_final = df["Agente BPO"].value_counts().to_dict()
        total_balancear = sum(conteo_final.get(ag, 0) for ag in agentes_balancear)
        promedio_aprox = total_balancear // len(agentes_balancear)

        agentes_por_encima = [(ag, conteo_final.get(ag, 0) - promedio_aprox) for ag in agentes_balancear if conteo_final.get(ag, 0) > promedio_aprox + 1]
        agentes_por_debajo = [(ag, promedio_aprox - conteo_final.get(ag, 0)) for ag in agentes_balancear if conteo_final.get(ag, 0) < promedio_aprox]
        agentes_por_encima.sort(key=lambda x: x[1], reverse=True)
        agentes_por_debajo.sort(key=lambda x: x[1], reverse=True)

        for ag_enc, sobrante in agentes_por_encima:
            for ag_fal, faltante in agentes_por_debajo:
                while sobrante > 0 and faltante > 0:
                    filas_ag_enc = df[df["Agente BPO"] == ag_enc].index.tolist()
                    if not filas_ag_enc:
                        break
                    idx_mover = filas_ag_enc[0]
                    df.at[idx_mover, "Agente BPO"] = ag_fal
                    sobrante -= 1
                    faltante -= 1
                agentes_por_debajo = [(ag, faltante) if ag == ag_fal else (ag, f) for ag, f in agentes_por_debajo]

        conteo_final = df["Agente BPO"].value_counts().to_dict()

        resumen_html = "<div class='resumen-container'>"
        resumen_html += "<div class='resumen-title'>📊 Resumen de Distribución Final</div>"
        for agente in agentes_bpo:
            resumen_html += f"<div class='resumen-item'><strong>{agente}:</strong> {conteo_final.get(agente, 0)}</div>"
        resumen_html += f"<div class='resumen-item'><strong>Agente Incontactable:</strong> {conteo_final.get('Agente Incontactable', 0)}</div>"
        resumen_html += f"<div class='resumen-item'><strong>Total general:</strong> {df.shape[0]}</div>"
        resumen_html += "</div>"

        st.markdown("""
        <style>
        .resumen-container {background: #f7f9fc; padding: 20px; border-radius: 8px; margin-top: 20px;}
        .resumen-title {font-size: 1.25rem; font-weight: bold; color: #333; margin-bottom: 10px;}
        .resumen-item {font-size: 1rem; margin: 5px 0; color: #555;}
        </style>
        """, unsafe_allow_html=True)
        st.markdown(resumen_html, unsafe_allow_html=True)

        columnas_finales = ['Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr', 'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo', 'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO']
        df_final = df[[col for col in columnas_finales if col in df.columns]]

        now_str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        excel_filename = f"Programa_Modificado_{now_str}.xlsx"
        csv_filename = f"Programa_Modificado_{now_str}.csv"

        df_final.to_excel(excel_filename, index=False)
        df_final.to_csv(csv_filename, index=False)

        col1, col2 = st.columns(2)
        with col1:
            with open(excel_filename, "rb") as f:
                st.download_button("\ud83d\udcc5 Descargar Excel", data=f, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            with open(csv_filename, "rb") as f:
                st.download_button("\ud83d\udcc5 Descargar CSV", data=f, file_name=csv_filename, mime="text/csv")

st.markdown("---")
st.caption("🚀 Creado por el equipo de BPO Innovations")
