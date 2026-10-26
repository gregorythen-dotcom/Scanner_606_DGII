import json
import os
from google import genai
from google.genai import types
import openpyxl
from PIL import Image
import streamlit as st

# Configuración de la interfaz adaptada a pantalla de teléfono
st.set_page_config(
    page_title="Scanner 606 DGII",
    page_icon="🧾",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.title("🧾 Scanner Facturas 606 DGII")
st.caption("Extracción automática de comprobantes fiscales con Gemini AI")

# Inicializar API Key de Gemini
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key:
    api_key = st.text_input(
        "Ingresa tu GEMINI_API_KEY:", type="password", key="api_key_input"
    )

if api_key:
    client = genai.Client(api_key=api_key)

    # 1. Opción de Captura
    opcion = st.radio("Origen de la factura:", ["📸 Tomar Foto", "📁 Subir Imagen/PDF"], horizontal=True)

    archivo = None
    if opcion == "📸 Tomar Foto":
        archivo = st.camera_input("Captura la factura")
    else:
        archivo = st.file_uploader("Selecciona imagen", type=["jpg", "jpeg", "png"])

    # 2. Procesamiento con Gemini
    if archivo is not None:
        imagen = Image.open(archivo)
        st.image(imagen, caption="Comprobante cargado", use_container_width=True)

        if st.button("⚡ Extraer Datos con Gemini", type="primary", use_container_width=True):
            with st.spinner("Leyendo comprobante fiscal..."):
                prompt = """
                Analiza este comprobante fiscal dominicano y extrae la información en este JSON:
                {
                  "rnc_emisor": "Solo números del RNC o Cédula sin guiones",
                  "ncf": "Código del NCF completo (ej. B0100015748)",
                  "fecha": "Fecha en formato YYYYMMDD (ej. 20260905)",
                  "monto_bienes": 0.00,
                  "monto_servicios": 0.00,
                  "itbis": 0.00,
                  "total": 0.00
                }
                Si un monto no aplica, coloca 0.00.
                """
                
                try:
                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=[imagen, prompt],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.1
                        )
                    )
                    st.session_state["datos"] = json.loads(response.text)
                    st.success("¡Lectura completada!")
                except Exception as e:
                    st.error(f"Error al procesar con Gemini: {e}")

    # 3. Formulario de Verificación y Guardado
    if "datos" in st.session_state:
        datos = st.session_state["datos"]
        st.markdown("---")
        st.subheader("📋 Confirmación de Datos")

        with st.form("form_606"):
            rnc = st.text_input("RNC / Cédula Emisor", value=str(datos.get("rnc_emisor", "")))
            ncf = st.text_input("NCF Comprobante", value=str(datos.get("ncf", "")))
            fecha = st.text_input("Fecha (YYYYMMDD)", value=str(datos.get("fecha", "")))
            bienes = st.number_input("Monto Bienes (RD$)", value=float(datos.get("monto_bienes", 0.0)))
            servicios = st.number_input("Monto Servicios (RD$)", value=float(datos.get("monto_servicios", 0.0)))
            itbis = st.number_input("ITBIS Facturado (RD$)", value=float(datos.get("itbis", 0.0)))
            total = st.number_input("Total Facturado (RD$)", value=float(datos.get("total", 0.0)))

            guardar = st.form_submit_button("💾 Guardar en Formato 606", use_container_width=True)

            if guardar:
                ruta_excel = "Formato_606.xlsm"
                if os.path.exists(ruta_excel):
                    wb = openpyxl.load_workbook(ruta_excel, keep_vba=True)
                    ws = wb.active
                    
                    # Buscar primera fila libre desde la fila 12
                    fila = 12
                    while ws.cell(row=fila, column=2).value is not None:
                        fila += 1

                    tipo_id = 1 if len(rnc.replace("-", "").strip()) == 9 else 2

                    ws.cell(row=fila, column=1, value=tipo_id)
                    ws.cell(row=fila, column=2, value=rnc.replace("-", "").strip())
                    ws.cell(row=fila, column=3, value="02")
                    ws.cell(row=fila, column=4, value=ncf)
                    ws.cell(row=fila, column=6, value=fecha)
                    ws.cell(row=fila, column=8, value=servicios)
                    ws.cell(row=fila, column=9, value=bienes)
                    ws.cell(row=fila, column=10, value=total)
                    ws.cell(row=fila, column=11, value=itbis)

                    wb.save(ruta_excel)
                    st.balloons()
                    st.success(f"✅ Comprobante NCF {ncf} guardado exitosamente en la fila {fila} del 606.")
                    del st.session_state["datos"]
                else:
                    st.error("No se encontró el archivo 'Formato_606.xlsm' en el directorio.")

import time

# Dentro de la función o bloque donde presionas "Extraer Datos con Gemini"
if st.button("⚡ Extraer Datos con Gemini"):
    with st.spinner("Procesando imagen con Gemini..."):
        max_retries = 3
        success = False
        
        for attempt in range(max_retries):
            try:
                # Llamada a la API con el modelo actualizado
                response = model.generate_content([prompt, image])
                success = True
                break  # Si tiene éxito, sale del bucle
            except Exception as e:
                if "503" in str(e) and attempt < max_retries - 1:
                    time.sleep(2)  # Espera 2 segundos si el servidor de Google está saturado
                    continue
                else:
                    st.error(f"Error al procesar con Gemini: {e}")
                    break
        
        if success:
            st.success("¡Datos extraídos con éxito!")
            # Aquí continúa tu código normal para mostrar y guardar los datos en el Formato 606
else:
    st.warning("Ingresa tu clave de API para comenzar.")
