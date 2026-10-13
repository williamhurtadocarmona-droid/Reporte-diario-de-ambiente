from datetime import datetime
from io import BytesIO
import os
import streamlit as st
from docx import Document
from docx.shared import Inches

st.set_page_config(
    page_title="Reporte Diario de Ambiente - Taller", layout="centered"
)

st.title("🛠️ Generador de Reporte Diario de Ambiente")
st.subheader("Taller de Mecanizado - SENA")

# Formulario de entrada
with st.form("reporte_form"):
  col1, col2 = st.columns(2)
  with col1:
    fecha_inspeccion = st.date_input(
        "Fecha de Inspección", value=datetime.now()
    )
    jornada = st.selectbox(
        "Jornada", ["Mañana", "Tarde", "24 H", "Mixta"]
    )
  with col2:
    momento = st.selectbox(
        "Momento de Inspección",
        ["Ingreso (Inicio)", "Entrega (Fin de jornada)"],
    )
    instructor = st.text_input(
        "Instructor Responsable", value="William Santiago Hurtado Carmona"
    )

  observaciones = st.text_area(
      "Observaciones / Novedades del Ambiente",
      placeholder=(
          "Detalle aquí si se encontraron residuos de comida, sillas"
          " desordenadas, máquinas sucias, etc."
      ),
  )

  st.markdown("### 📷 Registro Fotográfico de Evidencias")
  foto_canecas = st.file_uploader(
      "1. Estado de Canecas y Basura General",
      type=["png", "jpg", "jpeg"],
      key="canecas",
  )
  foto_maquinas = st.file_uploader(
      "2. Estado de Máquinas (Tornos / Fresadoras / CNC)",
      type=["png", "jpg", "jpeg"],
      key="maquinas",
  )
  foto_salon = st.file_uploader(
      "3. Estado del Salón, Sillas y Suelo",
      type=["png", "jpg", "jpeg"],
      key="salon",
  )

  submitted = st.form_submit_button("Generar y Descargar Reporte")

if submitted:
  try:
    # Verificar si existe el archivo base en el repositorio
    template_path = "Formato_Entrada y salida_Taller.docx"
    if os.path.exists(template_path):
      doc = Document(template_path)
    else:
      # Si no lo encuentra localmente, crea una estructura base limpia
      doc = Document()
      doc.add_heading("INFORME DIARIO DE CONTROL Y ASEO DE AMBIENTE", 0)

    # Añadir o reemplazar datos en el documento (Ejemplo básico de volcado estructurado)
    doc.add_paragraph(
        f"Fecha de Inspección: {fecha_inspeccion.strftime('%d / %m / %Y')}"
    )
    doc.add_paragraph(f"Jornada: {jornada}")
    doc.add_paragraph(f"Momento de Inspección: {momento}")
    doc.add_paragraph(f"Instructor Responsable: {instructor}")
    doc.add_paragraph(f"Observaciones / Novedades: {observaciones}")

    doc.add_heading("REGISTRO FOTOGRÁFICO DE EVIDENCIAS", level=1)

    # Función auxiliar para agregar fotos de forma segura
    def agregar_foto_al_doc(titulo_foto, archivo_subido):
      doc.add_heading(titulo_foto, level=2)
      if archivo_subido is not None:
        # Guardar temporalmente la imagen subida por Streamlit
        img_bytes = BytesIO(archivo_subido.getvalue())
        doc.add_picture(img_bytes, width=Inches(4.0))
      else:
        doc.add_paragraph("[ Sin registro fotográfico proporcionado ]")

    agregar_foto_al_doc("1. Estado de Canecas y Basura General", foto_canecas)
    agregar_foto_al_doc(
        "2. Estado de Máquinas (Tornos / Fresadoras / CNC)", foto_maquinas
    )
    agregar_foto_al_doc("3. Estado del Salón, Sillas y Suelo", foto_salon)

    doc.add_paragraph(f"\nInstructor Responsable: {instructor}")
    doc.add_paragraph("Firma de Conformidad / Registro: ______________________")

    # Guardar en memoria para descarga directa en Streamlit
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("¡Reporte generado con éxito!")
    st.download_button(
        label="📥 Descargar Documento Rellenado (.docx)",
        data=buffer,
        file_name=(
            f"Reporte_Ambiente_{fecha_inspeccion.strftime('%Y-%m-%d')}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )

  except Exception as e:
    st.error(f"Ocurrió un error al generar el archivo: {e}")
