import streamlit as st
import datetime
import pandas as pd
import requests
import base64
import io
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

# Configuración de página
st.set_page_config(page_title="Control de Aseo - Taller", page_icon="🛠️", layout="centered")

st.title("📸 Control Diario de Estado y Aseo")
st.write("Registra el estado del taller al ingresar o salir para respaldar tu labor.")

# Formulario Streamlit
with st.form("form_aseo"):
    fecha = st.date_input("Fecha de Inspección", datetime.date.today())
    jornada = st.selectbox("Jornada", ["Mañana", "Tarde", "Noche"])
    momento = st.radio("Momento", ["Ingreso (Inicio de jornada)", "Entrega (Fin de jornada)"])
    instructor = st.text_input("Instructor a Cargo", value="William Santiago Hurtado Carmona")
    
    st.subheader("Evidencias Fotográficas Obligatorias")
    foto_canecas = st.file_uploader("1. Estado de Canecas y Basura", type=["jpg", "jpeg", "png"])
    foto_maquinas = st.file_uploader("2. Estado de Máquinas (Tornos / Fresadoras / CNC)", type=["jpg", "jpeg", "png"])
    foto_salon = st.file_uploader("3. Estado del Salón, Sillas y Suelo", type=["jpg", "jpeg", "png"])
    
    observaciones = st.text_area("Observaciones o Novedades (Ej: Sillas desordenadas, residuos encontrados, etc.)")
    
    submitted = st.form_submit_button("Guardar Registro y Generar Informe Word")

if submitted:
    if not (foto_canecas and foto_maquinas and foto_salon):
        st.error("⚠️ Debes subir obligatoriamente las 3 fotografías.")
    else:
        try:
            # Credenciales desde los Secrets de Streamlit
            GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
            REPO_NAME = st.secrets["REPO_NAME"] # Formato: "usuario/repositorio"
            
            headers = {
                "Authorization": f"token {GITHUB_TOKEN}",
                "Accept": "application/vnd.github.v3+json"
            }
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            prefix = f"{fecha}_{jornada}_{momento.split()[0]}".lower()
            
            fotos = {
                "canecas": foto_canecas,
                "maquinas": foto_maquinas,
                "salon": foto_salon
            }
            
            urls_fotos = {}
            for clave, archivo_foto in fotos.items():
                ruta_github = f"evidencias/{fecha}/{prefix}_{clave}_{timestamp}.png"
                contenido_bytes = archivo_foto.read()
                archivo_foto.seek(0)
                
                content_encoded = base64.b64encode(contenido_bytes).decode("utf-8")
                
                url_api = f"https://api.github.com/repos/{REPO_NAME}/contents/{ruta_github}"
                payload = {
                    "message": f"Subir evidencia {clave} - {fecha}",
                    "content": content_encoded
                }
                
                resp = requests.put(url_api, json=payload, headers=headers)
                urls_fotos[clave] = f"https://raw.githubusercontent.com/{REPO_NAME}/main/{ruta_github}"

            # Descargar la plantilla Word directamente desde GitHub
            url_plantilla = f"https://raw.githubusercontent.com/{REPO_NAME}/main/Formato_Entrada_y_salida_Taller.docx"
            resp_plantilla = requests.get(url_plantilla, headers=headers)
            
            if resp_plantilla.status_code == 200:
                plantilla_bytes = io.BytesIO(resp_plantilla.content)
                doc = docx.Document(plantilla_bytes)
            else:
                doc = docx.Document() # Plan B si no la encuentra

            # Rellenar datos en el documento
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        texto = cell.text
                        if "DD / MM / AAAA" in texto:
                            cell.text = f"Fecha: {str(fecha)}"
                        elif "Jornada" in texto:
                            cell.text = f"Jornada: {jornada} ({momento})"
                        elif "Instructor" in texto:
                            cell.text = f"Instructor Responsable: {instructor}"

            # Añadir observaciones y fotos al informe Word
            doc.add_paragraph().paragraph_format.space_after = Pt(6)
            p_obs = doc.add_paragraph()
            r_ot = p_obs.add_run("Observaciones / Novedades Registradas:\n")
            r_ot.bold = True
            r_ot.font.name = "Calibri"
            r_ot.font.size = Pt(11)
            
            r_ox = p_obs.add_run(observaciones if observaciones else "Ninguna novedad reportada.")
            r_ox.font.name = "Calibri"
            r_ox.font.size = Pt(10)

            doc.add_paragraph().paragraph_format.space_after = Pt(10)

            p_f = doc.add_paragraph()
            r_f = p_f.add_run("REGISTRO FOTOGRÁFICO DE EVIDENCIAS")
            r_f.bold = True
            r_f.font.name = "Calibri"
            r_f.font.size = Pt(12)

            for label, uploaded_file in [("1. Estado de Canecas y Basura", foto_canecas), 
                                         ("2. Estado de Máquinas", foto_maquinas), 
                                         ("3. Estado del Salón y Sillas", foto_salon)]:
                
                doc.add_paragraph().add_run(label).bold = True
                img = Image.open(uploaded_file)
                temp_path = f"temp_{label[0]}.png"
                img.save(temp_path)
                doc.add_picture(temp_path, width=Inches(4.3))
                doc.add_paragraph().paragraph_format.space_after = Pt(6)

            # Preparar descarga del documento Word
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            st.success("✅ ¡Evidencias subidas a GitHub correctamente!")
            st.download_button(
                label="📥 Descargar Informe en Word con Fotos",
                data=doc_io,
                file_name=f"Informe_Aseo_{fecha}_{jornada}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"❌ Error al procesar: {e}")
