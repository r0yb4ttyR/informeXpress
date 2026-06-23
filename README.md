# 🕵️ informeXpress — Herramienta para la confección de actividades operativas

## 🧭 Descripción general
informeXpress es una herramienta diseñada para facilitar la creación, gestión y documentación de actividades operativas, tanto en entornos físicos como digitales.
Permite generar informes profesionales de manera rápida, estructurada y con soporte para geolocalización, cronología de eventos, análisis y exportación final.

Está pensada para usuarios que necesitan documentar:

- Seguimientos operativos

- Observaciones en campo

- Monitoreo de perfiles o publicaciones en redes sociales

- Actividades de análisis OSINT/SOCMINT

- Cualquier situación que requiera registrar eventos con contexto, ubicación y evidencias

- Su interfaz es sencilla, intuitiva y accesible para usuarios sin conocimientos técnicos.

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

## 🚀 Características principales

- Creación de actividades operativas con estructura modular

- Geolocalización de eventos mediante coordenadas o búsqueda manual

- Cronología editable para ordenar y contextualizar hechos

- Adjuntar imágenes o evidencias a cada evento

- Guardado de sesiones en formato .json

- Carga de sesiones previas para continuar el trabajo

- Exportación del informe final en formato profesional

- Interfaz limpia y moderna

- Configuración persistente mediante .streamlit/config.toml

## 🧭 Flujo de uso

1. Introducir los datos generales de la actividad

2. Añadir eventos a la cronología

3. Geolocalizar cada evento (opcional pero recomendado)

4. Adjuntar imágenes o capturas

5. Guardar la sesión en .json o continuar trabajando

6. Exportar el informe final

## 💾 Guardado y carga de sesiones

La aplicación permite:

- Guardar el estado completo del informe en un archivo .json

- Cargar un archivo .json previamente guardado

- Continuar la edición sin perder información

- Ideal para informes largos o actividades que se desarrollan en varios días.

## 🗺️ Geolocalización de eventos

Cada evento puede incluir coordenadas GPS, lo que permite reconstruir la actividad en el espacio y el tiempo.

## 📦 Instalación
Requisitos:

- Python 3.9+

- pip

- Streamlit

Instalación:
```
pip install -r requirements.txt
```

Ejecución en local:
```
streamlit run streamlit_ao.py
```

## 🗂️ Estructura del proyecto

informeXpress/
│── .devcontainer/
│── .streamlit/
│     └── config.toml
│── README.md
│── requirements.txt
└── streamlit_ao.py

## 🧪 Ejemplos de uso

A continuación se muestran dos casos prácticos para entender cómo aplicar informeXpress en situaciones reales.

### 🕵️‍♂️ Ejemplo 1 — Seguimiento operativo físico
Situación:  
Un usuario necesita documentar un seguimiento discreto a una persona o vehículo.

**Flujo**:

1. Crear una nueva actividad y registrar los datos generales.

2. Añadir un evento inicial:

`“Inicio del seguimiento en Calle X a las 09:15.”`

3. Continuar insertando eventos desde la pestaña correspondiente. En cada evento tenemos la posibilidad de geolocalizar cada punto relevante, así como de adjuntar fotografías de las personas, vehículos o lugares visitados.

4. Añadir observaciones clave e interacciones, para una lectura más clara de los datos contenidos. Para ello tenemos dos pestañas con campos que permiten la inclusión de carácteres sin límite.

5. Exportar el informe final para archivarlo o compartirlo.

**Resultado**:  
Un informe claro, cronológico y geolocalizado del seguimiento.

### 🌐 Ejemplo 2 — Monitoreo de un perfil (SOCMINT)

**Situación**:  
Un usuario detecta una publicación relevante en un perfil de redes sociales y necesita documentarla.

**Flujo**:

1. Crear una nueva actividad con el nombre del perfil o caso.

Añadir un evento:

```
“Publicación detectada a las 14:32 con contenido potencialmente relevante.”
```

2. Registrar el texto o resumen de la publicación, así como captura de pantalla de la misma, así como usar los campos observaciones claves e interacciones para el análisis de la misma.

3. Exportar el informe final.

**Resultado**:  
Un informe SOCMINT estructurado, con análisis y ubicación contextual.

## 🆕 Novedades respecto a versiones anteriores
- Geolocalización integrada

- Interfaz más profesional

- Guardado/carga en JSON

- Mejor gestión de imágenes

- Exportación más limpia

- Código reorganizado

- Mejor rendimiento

## 📄 Licencia
(Se utilizará la licencia definida en el repositorio GitHub.)


## 🛠️ Instalación

1. Clona el repositorio:
```bash
git clone https://github.com/r0yb4ttyR/informeXpress
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


## 🤝 Contribuir

Si deseas contribuir al proyecto:

1. Haz un Fork del repositorio
2. Crea una rama para tu funcionalidad (`git checkout -b feature/AmazingFeature`)
3. Realiza tus cambios
4. Commit de tus cambios (`git commit -m 'Add some AmazingFeature'`)
5. Push a la rama (`git push origin feature/AmazingFeature`)
6. Abre un Pull Request