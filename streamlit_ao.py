import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
from datetime import datetime
import pandas as pd
from urllib.parse import quote_plus

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
# Inicializar valores para los inputs del formulario para poder resetearlos
if 'hora_str' not in st.session_state:
    st.session_state.hora_str = "00:00"
if 'descripcion' not in st.session_state:
    st.session_state.descripcion = ""

# Formulario para añadir eventos
with st.expander("Añadir nuevo evento al timeline"):
    col1, col2 = st.columns([1, 3])
    with col1:
        # Usar solo key; el valor inicial se definió en st.session_state arriba
        hora_str = st.text_input("Hora (HH:MM)", key="hora_str")
        try:
            # Convertir el texto a objeto time
            hora = datetime.strptime(hora_str, "%H:%M").time()
        except ValueError:
            st.error("Por favor, ingrese la hora en formato HH:MM")
            hora = None
    with col2:
        # Añadir key para poder limpiar la descripción (no pasar value=)
        descripcion = st.text_area("Descripción del evento", height=100, key="descripcion")
    
    # Subida de imagen con key dinámica
    imagen = st.file_uploader("Añadir imagen al evento (opcional)", 
                            type=['png', 'jpg', 'jpeg'],
                            key=f"uploader_{st.session_state.upload_key}")
    
    # Callback para añadir evento (modifica st.session_state desde el callback)
    def add_event_callback():
        try:
            h = datetime.strptime(st.session_state.hora_str, "%H:%M").time()
        except Exception:
            st.session_state.add_error = "Hora inválida. Introduce HH:MM antes de añadir el evento."
            return

        uploader_key = f"uploader_{st.session_state.upload_key}"
        imagen_widget = st.session_state.get(uploader_key)
        imagen_bytes = imagen_widget.getvalue() if imagen_widget else None

        st.session_state.eventos.append({
            'hora': h,
            'descripcion': st.session_state.descripcion,
            'imagen': imagen_bytes
        })
        st.session_state.upload_key += 1
        # Reset de los campos (se hace dentro del callback)
        st.session_state.hora_str = "00:00"
        st.session_state.descripcion = ""
        st.session_state.add_error = ""
        st.session_state.add_success = "Evento añadido correctamente"

    st.button("Añadir evento", on_click=add_event_callback)

    # Mostrar mensajes producidos por el callback
    if st.session_state.get('add_error'):
        st.error(st.session_state.get('add_error'))
    if st.session_state.get('add_success'):
        st.success(st.session_state.get('add_success'))
        # opcional: mantener o limpiar el mensaje en otro momento

# Mostrar timeline
if st.session_state.eventos:
    for idx, evento in enumerate(st.session_state.eventos):
        col1, col2, col3 = st.columns([1, 3, 1])
        with col1:
            hora_display = evento['hora'].strftime('%H:%M') if evento.get('hora') else "—"
            st.write(f"**{hora_display}**")
        with col2:
            st.write(evento['descripcion'])
            if evento['imagen']:
                st.image(evento['imagen'], width=300)
        with col3:
            # Botones de edición y eliminación
            if st.button("✏️ Editar", key=f"edit_{idx}"):
                st.session_state.editing_idx = idx
            if st.button("🗑️ Eliminar", key=f"delete_{idx}"):
                st.session_state.eventos.pop(idx)
                st.rerun()

        # Formulario de edición si este evento está siendo editado
        if 'editing_idx' in st.session_state and st.session_state.editing_idx == idx:
            with st.expander("Editar evento", expanded=True):
                hora_edit = st.text_input(
                    "Hora (HH:MM)", 
                    value=evento['hora'].strftime("%H:%M"),
                    key=f"hora_edit_{idx}"
                )
                desc_edit = st.text_area(
                    "Descripción", 
                    value=evento['descripcion'],
                    key=f"desc_edit_{idx}"
                )
                imagen_edit = st.file_uploader(
                    "Actualizar imagen (opcional)",
                    type=['png', 'jpg', 'jpeg'],
                    key=f"img_edit_{idx}"
                )
                
                col1, col2 = st.columns([1,3])
                with col1:
                    if st.button("💾 Guardar", key=f"save_{idx}"):
                        try:
                            nueva_hora = datetime.strptime(hora_edit, "%H:%M").time()
                            st.session_state.eventos[idx]['hora'] = nueva_hora
                            st.session_state.eventos[idx]['descripcion'] = desc_edit
                            if imagen_edit:
                                st.session_state.eventos[idx]['imagen'] = imagen_edit.getvalue()
                            del st.session_state.editing_idx
                            st.rerun()
                        except ValueError:
                            st.error("Formato de hora inválido (HH:MM)")
                with col2:
                    if st.button("❌ Cancelar", key=f"cancel_{idx}"):
                        del st.session_state.editing_idx
                        st.rerun()
    
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
    doc.add_heading(f"🕵️‍♂️ AO - \"{nombre_ao}\"", 0)
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

# Generar resumen para WhatsApp
def build_wa_summary(max_obs=300, max_desc=120, max_items=10):
    lines = []
    lines.append(f"AO: {nombre_ao}")
    lines.append(f"Fecha: {fecha_operacion.strftime('%d/%m/%Y')}")
    lines.append(f"Zona: {zona_operaciones}")

    if observaciones:
        obs = observaciones.strip()
        if len(obs) > max_obs:
            obs = obs[:max_obs].rsplit(' ', 1)[0] + "..."
        lines.append("")
        lines.append("Observaciones:")
        lines.append(obs)

    if st.session_state.eventos:
        lines.append("")
        lines.append("Cronología:")
        eventos_sorted = sorted(
            st.session_state.eventos,
            key=lambda e: e.get('hora') or datetime.strptime("00:00", "%H:%M").time()
        )
        for e in eventos_sorted[:max_items]:
            hora_text = e.get('hora').strftime('%H:%M') if e.get('hora') else "--:--"
            desc = (e.get('descripcion') or "").strip()
            if len(desc) > max_desc:
                desc = desc[:max_desc].rsplit(' ', 1)[0] + "..."
            lines.append(f"{hora_text} - {desc}")
        if len(st.session_state.eventos) > max_items:
            lines.append(f"...(+{len(st.session_state.eventos) - max_items} más)")

    if interacciones:
        inter = interacciones.strip()
        if len(inter) > max_obs:
            inter = inter[:max_obs].rsplit(' ', 1)[0] + "..."
        lines.append("")
        lines.append("Interacciones:")
        lines.append(inter)

    return "\n".join(lines)

if 'wa_summary' not in st.session_state:
    st.session_state.wa_summary = ""

if st.button("✉️ Generar resumen WhatsApp"):
    st.session_state.wa_summary = build_wa_summary()

if st.session_state.wa_summary:
    st.text_area("Resumen para WhatsApp (copiar/pegar)", value=st.session_state.wa_summary, height=260, key="wa_show")
    st.download_button(
        label="Descargar resumen (.txt)",
        data=st.session_state.wa_summary,
        file_name=f"AO_{nombre_ao}_resumen.txt",
        mime="text/plain"
    )
    #wa_url = f"https://api.whatsapp.com/send?text={quote_plus(st.session_state.wa_summary)}"
    #st.markdown(f"[Abrir WhatsApp Web con el mensaje]({wa_url})")
