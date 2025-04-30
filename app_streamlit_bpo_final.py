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

# Lista de agentes BPO
agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Christian Tocay", "Nancy Zet", "Melissa Florian"]
if fecha_actual.weekday() == 5:  # sábado
    agentes_bpo.append("Abigail Vasquez")

# 🔁 Reemplazo manual de agente (opcional)
st.subheader("🔁 Reemplazo manual de un agente BPO (opcional)")
agente_ausente = st.selectbox("Selecciona al agente que está ausente", ["Ninguno"] + agentes_bpo, key="ausente")

agente_reemplazo = None
reemplazo_realizado = False
reemplazo_info = ""

if agente_ausente != "Ninguno":
    agente_reemplazo = st.text_input("Nombre del agente que lo va a sustituir (nuevo)", key="reemplazo")
    if agente_reemplazo:
        if agente_reemplazo in agentes_bpo:
            st.warning("⚠️ El agente de reemplazo ya está en la lista. Escribe un nuevo nombre diferente.")
        else:
            agentes_bpo = [agente_reemplazo if ag == agente_ausente else ag for ag in agentes_bpo]
            reemplazo_realizado = True
            reemplazo_info = f"ℹ️ {agente_ausente} fue reemplazado manualmente por {agente_reemplazo}."
            st.success(f"✅ {agente_ausente} ha sido reemplazado por {agente_reemplazo}")

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
            df = hoja_correcta.copy()

        # Limpieza y ajustes
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

                 # 2. Asignaciones forzadas
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

        # Casos especiales
        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[
            df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), 
            "Agente BPO"
        ] = "Melissa Florian"

        df.loc[
            df["Motivo"].str.contains("adicionales", case=False, na=False), 
            "Agente BPO"
        ] = "Ana Paniagua"

        # Distribución forzada clientes especiales
        clientes_especiales = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_especial = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_especiales), case=False, na=False)].copy()
        indices_a_repartir = df_especial[df_especial["Agente BPO"] == ""].index.tolist()
        for i, idx in enumerate(indices_a_repartir):
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente

        # Calcular cupo
        forzadas = df[df["Agente BPO"] != ""].groupby("Agente BPO").size().to_dict()
        total = df.shape[0]
        incontactables = forzadas.get("Agente Incontactable", 0)
        remainder = total - incontactables
        x = remainder / (len(agentes_bpo) - 0.25)

        cupo_teorico = {
            agente: int(0.75 * x) if agente == "Melissa Florian" else int(x)
            for agente in agentes_bpo
        }

        disponibles = {
            agente: max(cupo_teorico[agente] - forzadas.get(agente, 0), 0)
            for agente in agentes_bpo
        }

        indices_sin_asignar = df[df["Agente BPO"] == ""].index.tolist()

        while indices_sin_asignar:
            asignado = False
            for agente in agentes_bpo:
                if disponibles[agente] > 0 and indices_sin_asignar:
                    idx = indices_sin_asignar.pop(0)
                    df.at[idx, "Agente BPO"] = agente
                    disponibles[agente] -= 1
                    asignado = True
            if not asignado:
                for idx in indices_sin_asignar:
                    mayor = max(disponibles, key=disponibles.get)
                    df.at[idx, "Agente BPO"] = mayor
                break
        # Balanceo final
        agentes_normales = [ag for ag in agentes_bpo if ag != "Melissa Florian"]
        if "Agente Incontactable" in agentes_normales:
            agentes_normales.remove("Agente Incontactable")

        conteo_final = df["Agente BPO"].value_counts().to_dict()
        total_balancear = sum(conteo_final.get(ag, 0) for ag in agentes_normales)
        promedio = total_balancear // len(agentes_normales)

        agentes_sobra = [(ag, conteo_final.get(ag, 0) - promedio) for ag in agentes_normales if conteo_final.get(ag, 0) > promedio + 1]
        agentes_falta = [(ag, promedio - conteo_final.get(ag, 0)) for ag in agentes_normales if conteo_final.get(ag, 0) < promedio]

        for ag_sobra, sobra in agentes_sobra:
            for ag_falta, falta in agentes_falta:
                while sobra > 0 and falta > 0:
                    idx_mover = df[df["Agente BPO"] == ag_sobra].index[0]
                    df.at[idx_mover, "Agente BPO"] = ag_falta
                    sobra -= 1
                    falta -= 1
                conteo_final[ag_sobra] -= sobra
                conteo_final[ag_falta] = conteo_final.get(ag_falta, 0) + falta

        # Resumen final
        conteo_final = df["Agente BPO"].value_counts().to_dict()
        resumen_html = "<div class='resumen-container'>"
        resumen_html += "<div class='resumen-title'>📊 Resumen de Distribución Final</div>"
        for agente in agentes_bpo:
            if agente == "Agente Incontactable":
                continue
            asignado = conteo_final.get(agente, 0)
            resumen_html += f"<div class='resumen-item'><strong>{agente}:</strong> {asignado} (máximo teórico: {cupo_teorico.get(agente, 'N/A')})</div>"
        resumen_html += f"<div class='resumen-item'><strong>Agente Incontactable:</strong> {conteo_final.get('Agente Incontactable', 0)}</div>"
        if reemplazo_realizado:
            resumen_html += f"<div class='resumen-item'><em>{reemplazo_info}</em></div>"
        resumen_html += "</div>"

        st.markdown(
            """
            <style>
            .resumen-container {
                background: #f7f9fc;
                padding: 20px;
                border-radius: 8px;
                margin-top: 20px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                max-width: 600px;
            }
            .resumen-title {
                font-size: 1.25rem;
                font-weight: bold;
                color: #333;
                margin-bottom: 10px;
            }
            .resumen-item {
                font-size: 1rem;
                margin: 5px 0;
                color: #555;
            }
            </style>
            """, unsafe_allow_html=True
        )
        st.markdown(resumen_html, unsafe_allow_html=True)
        st.success("✅ Archivo procesado con éxito")

        # Exportar resultado
        columnas_finales = [
            'Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr',
            'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo',
            'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO'
        ]
        df_final = df[[col for col in columnas_finales if col in df.columns]]

        st.markdown("### 👀 Vista previa de los primeros registros")
        st.dataframe(df_final.head(15), height=400, use_container_width=True)

        now_str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        excel_filename = f"Programa_Modificado_{now_str}.xlsx"
        csv_filename = f"Programa_Modificado_{now_str}.csv"
        df_final.to_excel(excel_filename, index=False)
        df_final.to_csv(csv_filename, index=False)

        col1, col2 = st.columns(2)
        with col1:
            with open(excel_filename, "rb") as f:
                st.download_button("📥 Descargar Excel", data=f, file_name=excel_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            with open(csv_filename, "rb") as f:
                st.download_button("📥 Descargar CSV", data=f, file_name=csv_filename, mime="text/csv")

st.markdown("---")
st.caption("🚀 Creado por el equipo de BPO Innovations")




