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
  st.info(
      "💡 Tip: Mantén presionada la tecla **Ctrl** (o **Cmd** en Mac) para"
      " seleccionar varias imágenes en cada apartado."
  )

  fotos_canecas = st.file_uploader(
      "1. Estado de Canecas y Basura General",
      type=["png", "jpg", "jpeg"],
      accept_multiple_files=True,
      key="canecas",
  )
  fotos_maquinas = st.file_uploader(
      "2. Estado de Máquinas (Tornos / Fresadoras / CNC)",
      type=["png", "jpg", "jpeg"],
      accept_multiple_files=True,
      key="maquinas",
  )
  fotos_salon = st.file_uploader(
      "3. Estado del Salón, Sillas y Suelo",
      type=["png", "jpg", "jpeg"],
      accept_multiple_files=True,
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

      # Diccionario base con los textos
      context = {
          "fecha": fecha_inspeccion.strftime("%d / %m / %Y"),
          "jornada": jornada_texto,
          "momento": momento_texto,
          "observaciones": (
              observaciones if observaciones else "Sin novedades reportadas."
          ),
      }

      # Función segura para insertar múltiples imágenes pasando por cada párrafo del documento
      def insertar_multiples_fotos(tag_name, lista_archivos):
        if not lista_archivos:
          doc.replace_string(tag_name, "[ Sin fotografía adjunta ]")
          return

        # Buscamos el párrafo donde está la etiqueta y reemplazamos con la primera imagen,
        # y añadimos las demás justo debajo de manera limpia.
        for p in doc.paragraphs:
          if tag_name in p.text:
            p.text = p.text.replace(
                tag_name, ""
            )  # Limpiamos el texto de la etiqueta
            run = p.add_run()
            for i, archivo in enumerate(lista_archivos):
              if i > 0:
                # Añadir un salto de línea para las siguientes imágenes
                p_nuevo = doc.add_paragraph()
                run_nuevo = p_nuevo.add_run()
                InlineImage(
                    doc, BytesIO(archivo.getvalue()), width=Inches(3.5)
                )._add_to_run(run_nuevo)
              else:
                InlineImage(
                    doc, BytesIO(archivo.getvalue()), width=Inches(3.5)
                )._add_to_run(run)
            break

      # Renderizamos textos primero
      doc.render(context)

      # Insertamos las imágenes múltiples de forma segura por etiqueta
      insertar_multiples_fotos("{{foto_canecas}}", fotos_canecas)
      insertar_multiples_fotos("{{foto_maquinas}}", fotos_maquinas)
      insertar_multiples_fotos("{{foto_salon}}", fotos_salon)

      # Guardar en memoria para descarga
      buffer = BytesIO()
      doc.save(buffer)
      buffer.seek(0)

      st.success("¡Reporte generado con éxito y sin errores de formato!")
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
    st.error(f"Ocurrió un error al procesar el archivo: {e}")
