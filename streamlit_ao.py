import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from datetime import datetime
import pandas as pd

st.set_page_config(layout="wide")

st.title("Generador AO")

# Información básica del informe
col1, col2 = st.columns(2)
with col1:
    nombre_ao = st.text_input("Nombre (AO)", "")
    fecha_operacion = st.date_input("Fecha", datetime.now())
with col2:
    zona_operaciones = st.text_input("Zona", "")

# Observaciones clave
st.header("📋 Observaciones clave")
observaciones = st.text_area("Añade las observaciones principales", height=150)

# Timeline de movimientos
st.header("⏱️ Cronología de movimientos")

# Inicializar la lista de eventos y el key para el uploader en la sesión si no existe
if 'eventos' not in st.session_state:
    st.session_state.eventos = []
if 'upload_key' not in st.session_state:
    st.session_state.upload_key = 0

# Formulario para añadir eventos
with st.expander("Añadir nuevo evento al timeline"):
    col1, col2 = st.columns([1, 3])
    with col1:
        hora = st.time_input("Hora")
    with col2:
        descripcion = st.text_area("Descripción del evento", height=100)
    
    # Subida de imagen con key dinámica
    imagen = st.file_uploader("Añadir imagen al evento (opcional)", 
                            type=['png', 'jpg', 'jpeg'],
                            key=f"uploader_{st.session_state.upload_key}")
    
    if st.button("Añadir evento"):
        st.session_state.eventos.append({
            'hora': hora,
            'descripcion': descripcion,
            'imagen': imagen.getvalue() if imagen else None
        })
        # Incrementar el key para forzar un nuevo uploader limpio
        st.session_state.upload_key += 1
        st.success("Evento añadido correctamente")

# Mostrar timeline
if st.session_state.eventos:
    for idx, evento in enumerate(st.session_state.eventos):
        col1, col2 = st.columns([1, 4])
        with col1:
            st.write(f"**{evento['hora'].strftime('%H:%M')}**")
        with col2:
            st.write(evento['descripcion'])
            if evento['imagen']:
                st.image(evento['imagen'], width=300)
    
    # Botón para limpiar timeline
    if st.button("Limpiar timeline"):
        st.session_state.eventos = []
        st.success("Timeline limpiado correctamente")

# Sección de interacciones
st.header("👥 Interacciones")
interacciones = st.text_area("Describe las interacciones relevantes", height=150)

# Generar DOCX
if st.button("📤 Exportar a Word"):
    doc = Document()
    
    # Título y encabezado
    doc.add_heading(f"🕵️ AO - \"{nombre_ao}\"", 0)
    doc.add_paragraph(f"Fecha: {fecha_operacion.strftime('%d de %B de %Y')}")
    doc.add_paragraph(f"Zona de operaciones: {zona_operaciones}")
    
    # Observaciones clave
    doc.add_heading("Observaciones clave", level=1)
    doc.add_paragraph(observaciones)
    
    # Timeline
    doc.add_heading("Cronología de movimientos", level=1)
    for evento in st.session_state.eventos:
        p = doc.add_paragraph()
        p.add_run(f"{evento['hora'].strftime('%H:%M')} – ").bold = True
        p.add_run(evento['descripcion'])
        
        if evento['imagen']:
            doc.add_picture(BytesIO(evento['imagen']), width=Inches(3))
    
    # Interacciones
    doc.add_heading("Interacciones", level=1)
    doc.add_paragraph(interacciones)
    
    # Guardar y descargar
    buffer = BytesIO()
    doc.save(buffer)
    st.download_button(
        label="📥 Descargar informe",
        data=buffer.getvalue(),
        file_name=f"AO_{nombre_ao}_{fecha_operacion.strftime('%Y%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
