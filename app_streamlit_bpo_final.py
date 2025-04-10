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
        
        ### Limpieza y ajustes ###
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
        df["Agente BPO"] = ""  # Inicialmente vacío

        ### Asignaciones forzadas ###
        # 1. Incontactables (si existe el archivo)
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

        # 2. Definir la lista de agentes BPO  
        agentes_bpo = ["Ana Paniagua", "Alysson Garcia", "Julio de Leon", "Nancy Zet", "Melissa Florian"]
        if fecha_actual.weekday() == 5:  # sábado
            agentes_bpo.append("Abigail Vasquez")
        
        # 3. Reglas especiales:
        #   - Casos OXXO y Axionlog asignan a Melissa Florian
        exclusivas_melissa = ["OXXO", "Axionlog"]
        df.loc[
            df["Nombre de oportunidad1"].str.contains('|'.join(exclusivas_melissa), case=False, na=False), 
            "Agente BPO"
        ] = "Melissa Florian"
        #   - Casos "adicionales" asignan a Ana Paniagua
        df.loc[
            df["Motivo"].str.contains("adicionales", case=False, na=False), 
            "Agente BPO"
        ] = "Ana Paniagua"
        
        # 4. Repartir para clientes específicos (forzado)
        clientes_a_repartir = ["La Comer", "Fresko", "Sumesa", "City Market"]
        df_repartir = df[df["Nombre de oportunidad1"].str.contains('|'.join(clientes_a_repartir), case=False, na=False)].copy()
        indices_repartir = df_repartir[df_repartir["Agente BPO"] == ""].index.tolist()
        for i, idx in enumerate(indices_repartir):
            # Se asignan en orden de la lista
            agente = agentes_bpo[i % len(agentes_bpo)]
            df.at[idx, "Agente BPO"] = agente

        ### Asignar los registros restantes respetando el cupo definido para cada agente ###
        # Primero, se calculan cuántos registros forzados (asignados mediante reglas especiales) tiene cada agente.
        # Se ignoran los que siguen vacíos.
        forzadas_por_agente = df[df["Agente BPO"] != ""].groupby("Agente BPO").size().to_dict()

        # Calcular el total general y el total de incontactables ya asignados.
        total_general = df.shape[0]
        incontactables = forzadas_por_agente.get("Agente Incontactable", 0)
        remainder = total_general - incontactables  # Total a repartir entre los agentes BPO

        # Número de agentes (en la lista de agentes_bpo) a los que se asigna el resto.
        n_agentes = len(agentes_bpo)
        # Fórmula: (n_agentes - 1)*x + 0.75*x = remainder  =>  x = remainder / (n_agentes - 0.25)
        x = remainder / (n_agentes - 0.25)

        # Definir cupo teórico para cada agente:
        # (siendo Melissa 25% menos)
        cupo_teorico = {}
        for agente in agentes_bpo:
            if agente == "Melissa Florian":
                cupo_teorico[agente] = int(0.75 * x)
            else:
                cupo_teorico[agente] = int(x)
        # Nota: "Agente Incontactable" ya se asignó de forma forzada

        # Calcular cuántas filas adicionales puede recibir cada agente:
        # Disponible = cupo_teorico - filas ya forzadas (si tiene más, se limita a 0)
        filas_adicionales_para = {}
        for agente in agentes_bpo:
            ya_forzadas = forzadas_por_agente.get(agente, 0)
            disponible = cupo_teorico[agente] - ya_forzadas
            if disponible < 0:
                disponible = 0
            filas_adicionales_para[agente] = disponible

        # Obtener los índices de filas sin asignar (columna vacía)
        df_sin_asignar = df[df["Agente BPO"] == ""].copy()
        indices_sin_asignar = list(df_sin_asignar.index)

        # Asignación round-robin respetando la capacidad adicional de cada agente.
        # Se recorre en orden a los agentes y se asigna solo si tiene capacidad disponible.
        i = 0
        # Para evitar un loop infinito, se verifica que en cada ronda se asigne al menos una fila.
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
            # Si en una ronda completa no se asigna ninguno, significa que se agotó el cupo de todos.
            # En ese caso, para no dejar filas sin asignar, se asignan las restantes arbitrariamente.
            if not asignado_en_ronda and indices_sin_asignar:
                for idx in indices_sin_asignar:
                    # Se asigna al agente que tenga el mayor cupo adicional (aunque sea negativo)
                    agente_max = max(filas_adicionales_para, key=filas_adicionales_para.get)
                    df.at[idx, "Agente BPO"] = agente_max
                break  # Se asignaron todas

        ### Mostrar resumen de distribución ###
        # Se arma un diccionario final con los totales por agente según el DataFrame final.
        conteo_final = df["Agente BPO"].value_counts().to_dict()
        # Se muestra junto con el cupo teórico calculado para comparar.
        resumen_html = "<div class='resumen-container'>"
        resumen_html += "<div class='resumen-title'>📊 Resumen de Distribución Final</div>"
        for agente in agentes_bpo:
            # Se muestra: forzadas + adicionales asignadas vs. cupo teórico
            asignado = conteo_final.get(agente, 0)
            resumen_html += f"<div class='resumen-item'><strong>{agente}:</strong> {asignado} (máximo: {cupo_teorico[agente]})</div>"
        # Agregar a los incontactables
        incontactable_total = conteo_final.get("Agente Incontactable", 0)
        resumen_html += f"<div class='resumen-item'><strong>Agente Incontactable:</strong> {incontactable_total}</div>"
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

        ### Vista previa y descarga ###
        columnas_finales = [
            'Delv Ship-To Party', 'Delv Ship-To Name', 'Order Quantity', 'Delivery Nbr',
            'Esquema', 'Coordinador LT', 'Shpt Haulier Name', 'Ejecutivo RBO', 'Motivo',
            'Fecha de recolección', 'Nombre de oportunidad1', 'Fecha de cierre', 'Etapa', 'Agente BPO'
        ]
        df_final = df[[col for col in columnas_finales if col in df.columns]]

        st.markdown("### 👀 Vista previa de los primeros registros (14 columnas finales)")
        st.dataframe(df_final.head(15), height=500, use_container_width=True)

        now_str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        excel_filename = f"Programa_Modificado_{now_str}.xlsx"
        csv_filename = f"Programa_Modificado_{now_str}.csv"
        
        # Exportar a Excel y CSV
        df_final.to_excel(excel_filename, index=False)
        df_final.to_csv(csv_filename, index=False)
        
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

