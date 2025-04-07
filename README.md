
# 📊 Procesador de Archivos BPO

Esta app permite subir un archivo Excel (.xlsx), limpiar sus datos y asignar automáticamente agentes BPO según reglas predefinidas. Está construida con **Streamlit** y se ejecuta 100% en la nube.

## 🚀 ¿Qué hace esta app?

- Limpia columnas como “Esquema”, “Motivo”, “Shpt Haulier Name”, entre otras
- Asigna fechas según reglas de “AD”, “OD”, “On Demand”
- Asigna automáticamente el Agente BPO de manera equilibrada
- Devuelve un archivo limpio y listo para subir a Salesforce

## 🧪 Cómo usar

1. Abre la app 👉 [https://bpo-procesador.streamlit.app](https://bpo-procesador.streamlit.app)
2. Sube tu archivo Excel
3. Haz clic en “Procesar”
4. Descarga tu archivo limpio

## 🛠 Requisitos del entorno

Este repositorio incluye un archivo `requirements.txt` para instalar dependencias:

```
streamlit
pandas
openpyxl
```

---

Creado con ❤️ por el equipo de **BPO Innovations**.
