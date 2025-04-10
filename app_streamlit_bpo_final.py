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
    - Corrige campos vacíos o incorrectos.
    - Asigna automáticamente agentes BPO.
    - Detecta y etiqueta como 'Incontactables' según lista externa.
    - Asigna a Ana Paniagua todos los registros con motivo 'Adicionales'.
    - Agrega a Abigail Vasquez a la distribución si es sábado.
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
        df = pd.read_excel(uploaded_file)
        
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

        # Procesar incontactables
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
        else:
            st.info("Puedes subir manualmente 'Incontactables.xlsx' a la raíz del proyecto en Streamlit Cloud si deseas usarlo.")

        # Lista de agentes BPO
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:  # sábado
            agentes_bpo.append("Abigail Vasquez")
        
        # Reglas especiales de asignación
        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[
            df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), 
            "Agente BPO"
        ] = "Melissa Florian"
        df.loc[
            df["Motivo"].str.contains("adicionales", case=False, na=False), 
            "Agente BPO"
        ] = "Ana Paniagua"

        # Repartir para clientes específicos
        clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False)].copy()
        asignaciones = df["Agente BPO"].value_counts().to_dict()
        # Inicializar conteo para cada agente si no hay aún asignaciones
        for agente in agentes_bpo:
            if agente not in asignaciones:
                asignaciones[agente] = 0
        indices_repartir = df_repartir[df_repartir["Agente BPO"] == ""].index.tolist()
        for i, idx in enumerate(indices_repartir):
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente
            asignaciones[agente] += 1

        # Asignar los que aún están sin agente (round-robin)
        df_sin_asignar = df[df["Agente BPO"] == ""].copy()
        indices_sin_asignar = df_sin_asignar.index.tolist()
        i = 0
        while indices_sin_asignar:
            idx = indices_sin_asignar.pop(0)
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente
            i += 1

        # Aquí se muestra la distribución REAL según la asignación en el DataFrame (si se quisiera solo el conteo real usarías value_counts)
        # Pero a continuación calculamos la distribución según la regla solicitada:
        # Se toma el total general y se resta los "Agente Incontactable" para repartir el resto entre agentes
        total_general = df.shape[0]
        incontactables = df[df["Agente BPO"] == "Agente Incontactable"].shape[0]
        remainder = total_general - incontactables

        # Número de agentes a repartir (sin incluir "Agente Incontactable")
        n_agentes = len(agentes_bpo)
        # Fórmula: (n_agentes - 1)*x + 0.75*x = remainder  =>  x = remainder / (n_agentes - 0.25)
        x = remainder / (n_agentes - 0.25)

        # Crear el diccionario de distribución
        distribucion = {}
        for agente in agentes_bpo:
            if agente == "Melissa Florian":
                distribucion[agente] = int(0.75 * x)
            else:
                distribucion[agente] = int(x)
        distribucion["Agente Incontactable"] = incontactables

        # Ajuste para cuadrar el total (por si la conversión a enteros genera diferencia)
        suma_asignada = sum(distribucion[ag] for ag in distribucion if ag != "Agente Incontactable")
        diferencia = remainder - suma_asignada
        if diferencia != 0:
            # Se ajusta sumando la diferencia al primer agente que no sea Melissa Florian
            for agente in agentes_bpo:
                if agente != "Melissa Florian":
                    distribucion[agente] += diferencia
                    break

        # Mostrar el resumen de distribución final
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
        resumen_html = "<div class='resumen-container'>"
        resumen_html += "<div class='resumen-title'>📊 Resumen de Distribución Final</div>"
        for agente, monto in distribucion.items():
            resumen_html += f"<div class='resumen-item'><strong>{agente}:</strong> {monto}</div>"
        resumen_html += "</div>"
        st.markdown(resumen_html, unsafe_allow_html=True)

        st.success("✅ Archivo procesado con éxito")

        # Solo las 14 columnas finales para la vista previa y descarga
        columnas_finales = [
            'Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr',
            'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo',
            'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO'
        ]
        df_final = df[[col for col in columnas_finales if col in df.columns]]

        st.markdown("### 👀 Vista previa de los primeros registros (14 columnas finales)")
        st.dataframe(df_final.head(15), height=500, use_container_width=True)

        # Preparar archivos para descarga
        now_str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        excel_filename = f"Programa_Modificado_{now_str}.xlsx"
        csv_filename = f"Programa_Modificado_{now_str}.csv"
        
        # Exportar a Excel
        df_final.to_excel(excel_filename, index=False)
        # Exportar a CSV
        df_final.to_csv(csv_filename, index=False)
        
        # Mostrar botones de descarga para Excel y CSV con la fecha y mensaje CHEP
        st.markdown(f"**CHEP**: Archivo generado el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        col1, col2 = st.columns(2)
        with col1:
            with open(excel_filename, "rb") as f:
                st.download_button(
                    label="📥 Descargar Excel",
                    data=f,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with col2:
            with open(csv_filename, "rb") as f:
                st.download_button(
                    label="📥 Descargar CSV",
                    data=f,
                    file_name=csv_filename,
                    mime="text/csv"
                )

st.markdown("---")
st.caption("🚀 Creado por el equipo de BPO Innovations")
