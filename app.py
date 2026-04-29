# Streamlit cloud: marco de trabajo donde crear la app del móvil
import streamlit as st
# Pandas: libreria para manejo y tratamiento de datos. Permite trabajo con excel
import pandas as pd
# Calendario
from datetime import datetime
#os: módulo para interaccionar con Windows, Lynux o macOS
import os
# Importar imágenes para la app
from PIL import Image

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Seguimiento de Obra")
st.title("Control de Seguimiento de Obra")

# --- CARGA DE LOGO ---
# Asegúrate de tener un archivo 'logo.png' en tu carpeta
imagen1 = Image.open('Señal Electricidad.jpg')
st.image(imagen1, caption='Obra eléctrica', use_container_width=True)
imagen2 = Image.open('Logo Masaveu.png')
st.image(imagen2, caption='Logo Masaveu', use_container_width=True)

# --- FORMULARIO DE ENTRADA ---
with st.form("formulario_obra", clear_on_submit=True):
    st.subheader("Nuevo Registro")
    
    Nombre_Trabajador = st.text_input("Nombre del Trabajador")
    
    Tarea = st.selectbox("Seleccione la Tarea", [
        "Trazado y marcado de cajas, tubos y cuadros", 
        "Ejecución rozas en paredes y techos", 
        "Montaje de soportes", 
        "Colocación tubos y conductos", 
        "Tendido de cables",
        "Identificación y etiquetado",
        "Conexionado cables en bornas o regletas",
        "Instalación y conexionado de mecanismos",
        "Fijación de carril DIN y mecanismos en cuadro eléctrico",
        "Cableado interno del cuadro eléctrico",
        "Configuración de equipos domóticos y/o automáticos",
        "Conexionado de sensores/actuadores de equipos domóticos/automáticos",
        "Pruebas de continuidad",
        "Pruebas de aislamiento",
        "Verificación de tierras",
        "Programación del automatismo",
        "Pruebas de funcionamiento"
    ])
    
    Estado = st.selectbox("Estado de la Tarea", [
        "OK, finalizado sin errores", 
        "Finalizado, pero con errores pendientes de corregir", 
        "Finalizado y corregidos los errores"
    ])
    
    Comentarios = st.text_input("Comentarios")
    
    Fecha = st.date_input("Fecha", datetime.now())
    
    enviar_datos = st.form_submit_button("Guardar")

# --- LÓGICA DE ALMACENAMIENTO ---
archivo_datos = "registros_obra.csv"

if enviar_datos:
    nuevo_registro = {
        "Fecha": Fecha.strftime("%d/%m/%Y"),
        "Trabajador": Nombre_Trabajador,
        "Tarea": Tarea,
        "Estado": Estado
    }
    
    df = pd.DataFrame([nuevo_registro])
    
    # Guardar en CSV (modo append)
    if not os.path.isfile(archivo_datos):
        df.to_csv(archivo_datos, index=False)
    else:
        df.to_csv(archivo_datos, mode='a', header=False, index=False)
    
    st.success("✅ Registro guardado localmente.")

# --- SECCIÓN DE EXPORTACIÓN Y ENVÍO ---
st.divider()
st.subheader("Gestión de Reportes")

if os.path.exists(archivo_datos):
    df_total = pd.read_csv(archivo_datos)
    st.write("Vista previa de los últimos registros:")
    st.dataframe(df_total.tail())

    # Botón para descargar el Excel
    # Nota: Usamos ExcelWriter para generar el archivo .xlsx
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_total.to_excel(writer, index=False, sheet_name='Registros')
    
    st.download_button(
        label="📥 Descargar Excel",
        data=buffer.getvalue(),
        file_name=f"reporte_obra_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Botón para simular envío por correo
    if st.button("📧 Enviar Excel a la empresa"):
        # Aquí iría la lógica de SMTP (Gmail/Outlook)
        # Por seguridad, Streamlit Cloud requiere configurar 'secrets' para las contraseñas
        st.info("Función de envío activada. (Requiere configuración de servidor SMTP en Secrets)")
else:
    st.info("Aún no hay registros guardados.")
