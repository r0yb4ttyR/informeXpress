# 🕵️ Generador de Informes Operativos

Esta aplicación Streamlit permite crear y gestionar informes operativos de manera interactiva, con soporte para timeline de eventos, imágenes y exportación a formato DOCX.

## 📋 Características

- Creación de informes operativos completos
- Timeline interactivo con soporte para imágenes
- Gestión de observaciones clave
- Registro de interacciones
- Exportación a formato DOCX y PDF
- Niveles de clasificación del informe
- Guardar y cargar informes en formato JSON
- Interfaz con tema corporativo, panel lateral y pestañas

## 🛠️ Instalación

1. Clona el repositorio:
```bash
git clone https://github.com/elp0t4s-prog/informeXpress
cd STREAMLIT_AO
```

2. Crea un entorno virtual (opcional pero recomendado):
```bash
python -m venv venv
source venv/bin/activate  # En Linux/Mac
# o
venv\Scripts\activate     # En Windows
```

3. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## 🚀 Uso

1. Inicia la aplicación:
```bash
streamlit run streamlit_ao.py
```

2. Accede a la aplicación en tu navegador (por defecto en `http://localhost:8501`)

## 📝 Funcionalidades

### Información Básica
- Nombre de la Operación (AO)
- Fecha de la operación
- Zona de operaciones

### Observaciones Clave
- Registro de observaciones principales
- Formato de texto libre para máxima flexibilidad

### Timeline de Movimientos
- Añadir eventos con hora específica
- Incluir descripciones detalladas
- Adjuntar imágenes a los eventos (opcional)
- Vista previa de imágenes en tiempo real
- Limpieza automática del selector de imágenes

### Interacciones
- Documentación de interacciones relevantes
- Formato libre para detalles específicos

### Exportación
- Generación de documentos DOCX y PDF
- Portada con metadatos y banner de clasificación
- Cronología estructurada con imágenes (tabla en DOCX)
- Numeración de páginas en el pie
- Nombre de archivo personalizado con fecha

## 🤝 Contribuir

Si deseas contribuir al proyecto:

1. Haz un Fork del repositorio
2. Crea una rama para tu funcionalidad (`git checkout -b feature/AmazingFeature`)
3. Realiza tus cambios
4. Commit de tus cambios (`git commit -m 'Add some AmazingFeature'`)
5. Push a la rama (`git push origin feature/AmazingFeature`)
6. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.