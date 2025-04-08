
import pandas as pd
import unicodedata
from datetime import datetime, timedelta

# === CONFIGURACIÓN ===
archivo_entrada = "Programa  llamadas CAM - BPO del 07 al 12 de abril del 2025.xlsx"
archivo_incontactables = "Incontactables.xlsx"
nombre_hoja = "Hoja1"
archivo_salida = f"Programa_CAM_Modificado_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

# === FUNCIONES ===
def remove_accents(text):
    if isinstance(text, str):
        return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
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

# === EJECUCIÓN PRINCIPAL ===

# Fechas base
fecha_actual = datetime.today()
fecha_siguiente = fecha_actual + timedelta(days=1)
fecha_cierre = fecha_actual.strftime("%d/%m/%Y")

# Mes en español
meses_es = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
}
mes_in_spanish = meses_es[fecha_actual.month]
fecha_oportunidad = f"{fecha_actual.day}-{mes_in_spanish}-{fecha_actual.year}"

# Leer archivo principal y de incontactables
df = pd.read_excel(archivo_entrada, sheet_name=nombre_hoja)
df_incontactables = pd.read_excel(archivo_incontactables)
delv_incontactables = df_incontactables['Delv Ship-To Party'].astype(str).tolist()

# Limpieza básica
df["Esquema"] = df["Esquema"].fillna("SIN ASIGNAR").apply(lambda x: "SIN ASIGNAR" if x not in ["Dedicado", "Regular"] else x)
df["Coordinador LT"] = df["Coordinador LT"].fillna("SIN ASIGNAR").replace("#N/A", "SIN ASIGNAR")
df["Shpt Haulier Name"] = df["Shpt Haulier Name"].fillna("Sin Asignar").apply(remove_accents)
df["Ejecutivo RBO"] = df["Ejecutivo RBO"].fillna("SIN ASIGNAR").replace(["#N/A", "N/A"], "SIN ASIGNAR")
df["Motivo"] = df["Motivo"].fillna("#N/A").apply(remove_accents).replace("N/A", "#N/A")

# Procesar fecha de recolección
if "Día de recolección" in df.columns:
    df["Fecha de recolección"] = df["Día de recolección"].apply(asignar_fecha)
else:
    df["Fecha de recolección"] = ""

# Campos adicionales
df["Nombre de oportunidad1"] = df["Delv Ship-To Name"] + " " + fecha_oportunidad
df["Fecha de cierre"] = fecha_cierre
df["Etapa"] = "Pendiente de Contacto"
df["Agente BPO"] = ""

# Asignar "Incontactables"
df.loc[df["Delv Ship-To Party"].astype(str).isin(delv_incontactables), "Agente BPO"] = "Incontactables"

# Reordenar columnas
column_order = [
    "Delv Ship-To Party", "Delv Ship-To Name", "Order Quantity", "Delivery Nbr",
    "Esquema", "Coordinador LT", "Shpt Haulier Name", "Ejecutivo RBO", "Motivo",
    "Fecha de recolección", "Nombre de oportunidad1", "Fecha de cierre", "Etapa", "Agente BPO"
]
df_final = df[column_order]

# Guardar archivo final
df_final.to_excel(archivo_salida, index=False)
print(f"✅ Archivo generado exitosamente: {archivo_salida}")
