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
#Envío de mails
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

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

def enviar_correo(archivo_adjunto):
    try:
        # Extraer datos de los Secrets
        remitente = st.secrets["email"]["user"]
        password = st.secrets["email"]["password"]
        destinatario = st.secrets["email"]["receiver"]

        # Crear el mensaje
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = f"Reporte de Obra - {datetime.now().strftime('%d/%m/%Y')}"

        # Adjuntar el archivo Excel
        part = MIMEBase('application', 'octet-stream')
        with open(archivo_adjunto, "rb") as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={archivo_adjunto}")
        msg.attach(part)

        # Conexión con el servidor SMTP de Gmail
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(remitente, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error técnico: {e}")
        return False

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

    # Botón para envío por correo
   if st.button("📧 Enviar Excel al Jefe de Turno"):
    if os.path.exists("registros_obra.csv"):
        # Convertimos el CSV actual a Excel para que el jefe lo vea bien
        nombre_excel = "reporte_obra.xlsx"
        pd.read_csv("registros_obra.csv").to_excel(nombre_excel, index=False)
        
        with st.spinner("Enviando correo..."):
            if enviar_correo(nombre_excel):
                st.success(f"✅ ¡Correo enviado con éxito a {st.secrets['email']['receiver']}!")
    else:
        st.error("No hay datos guardados para enviar.")
