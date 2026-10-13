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
    jornada_opcion = st.radio(
        "Jornada", ["Mañana", "Tarde", "24 H"], horizontal=True
    )
  with col2:
    momento_opcion = st.radio(
        "Momento de Inspección",
        ["Ingreso (Inicio)", "Entrega (Fin de jornada)"],
        horizontal=True,
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
    template_path = "Formato_Entrada y salida_Taller.docx"
    if not os.path.exists(template_path):
      st.error(
          "No se encontró el archivo 'Formato_Entrada y salida_Taller.docx' en"
          " el repositorio."
      )
    else:
      doc = Document(template_path)

      # 1. Reemplazar datos de texto básicos en los párrafos
      fecha_str = fecha_inspeccion.strftime("%d / %m / %Y")

      for p in doc.paragraphs:
        if "Fecha de Inspección:" in p.text:
          p.text = f"Fecha de Inspección: {fecha_str}"

        if "Jornada:" in p.text:
          # Marcar la opción seleccionada visualmente
          m = "[X]" if jornada_opcion == "Mañana" else "[ ]"
          t = "[X]" if jornada_opcion == "Tarde" else "[ ]"
          h = "[X]" if jornada_opcion == "24 H" else "[ ]"
          p.text = (
              f"Jornada:   [  ] Mañana   [  ] Tarde   [  ] 24 H".replace(
                  "[  ] Mañana", f"{m} Mañana"
              )
              .replace("[  ] Tarde", f"{t} Tarde")
              .replace("[  ] 24 H", f"{h} 24 H")
          )

        if "Momento de Inspección:" in p.text:
          ing = (
              "[X]" if momento_opcion == "Ingreso (Inicio)" else "[  ]"
          )
          ent = (
              "[X]" if momento_opcion == "Entrega (Fin de jornada)" else "[  ]"
          )
          p.text = (
              "Momento de Inspección:   [  ] Ingreso (Inicio)   [  ] Entrega"
              " (Fin de jornada)"
          ).replace("[  ] Ingreso (Inicio)", f"{ing} Ingreso (Inicio)").replace(
              "[  ] Entrega (Fin de jornada)",
              f"{ent} Entrega (Fin de jornada)",
          )

        if "Instructor Responsable:" in p.text:
          p.text = f"Instructor Responsable: {instructor}"

        if (
            "Observaciones / Novedades del Ambiente:" in p.text
            or "(Detalle aquí" in p.text
        ):
          p.text = (
              "Observaciones / Novedades del Ambiente:\n"
              f"{observaciones if observaciones else 'Sin novedades reportadas.'}"
          )

      # 2. Reemplazar las etiquetas de fotos en los párrafos por las imágenes reales
      # Recorremos los párrafos buscando los textos indicadores de imágenes
      for i, p in enumerate(doc.paragraphs):
        texto_p = p.text.strip()

        # Identificamos el bloque de la caneca
        if "1. Estado de Canecas y Basura General" in texto_p:
          if foto_canecas is not None:
            p.text = "1. Estado de Canecas y Basura General"
            run = p.add_run()
            run.add_picture(
                BytesIO(foto_canecas.getvalue()), width=Inches(3.5)
            )

        # Identificamos el bloque de las máquinas
        elif "2. Estado de Máquinas (Tornos / Fresadoras / CNC)" in texto_p:
          if foto_maquinas is not None:
            p.text = "2. Estado de Máquinas (Tornos / Fresadoras / CNC)"
            run = p.add_run()
            run.add_picture(
                BytesIO(foto_maquinas.getvalue()), width=Inches(3.5)
            )

        # Identificamos el bloque del salón
        elif "3. Estado del Salón, Sillas y Suelo" in texto_p:
          if foto_salon is not None:
            p.text = "3. Estado del Salón, Sillas y Suelo"
            run = p.add_run()
            run.add_picture(BytesIO(foto_salon.getvalue()), width=Inches(3.5))

      # Guardar en memoria para descarga
      buffer = BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.success("¡Reporte integrado y generado correctamente!")
      st.download_button(
          label="📥 Descargar Formato Rellenado (.docx)",
          data=buffer,
          file_name=(
              f"Reporte_Ambiente_{fecha_inspeccion.strftime('%Y-%m-%d')}.docx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )

  except Exception as e:
    st.error(f"Ocurrió un error al procesar el documento: {e}")
