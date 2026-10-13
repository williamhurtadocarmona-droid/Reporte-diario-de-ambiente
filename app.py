from datetime import datetime
from io import BytesIO
import os
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

st.set_page_config(
    page_title="Reporte Diario de Ambiente - Taller", layout="centered"
)

st.title("🛠️ Generador de Reporte Diario de Ambiente")
st.subheader("Taller de Mecanizado - SENA")

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
          f"No se encontró el archivo '{template_path}' en el repositorio."
      )
    else:
      doc = DocxTemplate(template_path)

      # Validación de casillas
      m_chk = "X" if jornada_opcion == "Mañana" else " "
      t_chk = "X" if jornada_opcion == "Tarde" else " "
      h_chk = "X" if jornada_opcion == "24 H" else " "
      jornada_texto = f"[{m_chk}] Mañana   [{t_chk}] Tarde   [{h_chk}] 24 H"

      ing_chk = "X" if momento_opcion == "Ingreso (Inicio)" else " "
      ent_chk = "X" if momento_opcion == "Entrega (Fin de jornada)" else " "
      momento_texto = f"[{ing_chk}] Ingreso (Inicio)   [{ent_chk}] Entrega (Fin de jornada)"

      # Diccionario con etiquetas simplificadas (sin espacios)
      context = {
          "fecha": fecha_inspeccion.strftime("%d / %m / %Y"),
          "jornada": jornada_texto,
          "momento": momento_texto,
          "observaciones": (
              observaciones if observaciones else "Sin novedades reportadas."
          ),
      }

      # Manejo de imágenes con etiquetas limpias
      if foto_canecas is not None:
        context["foto_canecas"] = InlineImage(
            doc, BytesIO(foto_canecas.getvalue()), width=Inches(3.5)
        )
      else:
        context["foto_canecas"] = "[ Sin fotografía adjunta ]"

      if foto_maquinas is not None:
        context["foto_maquinas"] = InlineImage(
            doc, BytesIO(foto_maquinas.getvalue()), width=Inches(3.5)
        )
      else:
        context["foto_maquinas"] = "[ Sin fotografía adjunta ]"

      if foto_salon is not None:
        context["foto_salon"] = InlineImage(
            doc, BytesIO(foto_salon.getvalue()), width=Inches(3.5)
        )
      else:
        context["foto_salon"] = "[ Sin fotografía adjunta ]"

      # Renderizar la plantilla
      doc.render(context)

      # Guardar en memoria para descarga
      buffer = BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.success("¡Reporte generado perfectamente sin errores!")
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
    st.error(f"Ocurrió un error al procesar la plantilla: {e}")
