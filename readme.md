# Sistema Generador de Exámenes de Estadística

Un sistema web completo para la generación de exámenes de estadística básica con múltiples variantes, hojas de respuestas y plantillas de calificación. Desarrollado para la Facultad de Humanidades de la Universidad Panamericana.

## Características

- **Múltiples variantes**: Genera hasta 10 versiones diferentes de un mismo examen.
- **Estructura estándar**: Mantiene el formato institucional requerido.
- **Secciones completas**: Preguntas de selección múltiple, identificación de gráficos estadísticos y ejercicios prácticos.
- **Sistema de calificación**: Incluye hojas de respuesta para estudiantes y plantillas de corrección para docentes.
- **Exportación flexible**: Genera archivos en formatos Word y PDF.

## Requisitos

- Python 3.8 o superior
- Pip (gestor de paquetes de Python)
- Navegador web moderno

## Instalación

1. Clone este repositorio:
   ```bash
   git clone https://github.com/username/generador-examenes-estadistica.git
   cd generador-examenes-estadistica
   ```

2. Cree un entorno virtual (recomendado):
   ```bash
   python -m venv venv
   ```

3. Active el entorno virtual:
   
   En Windows:
   ```bash
   venv\Scripts\activate
   ```
   
   En MacOS/Linux:
   ```bash
   source venv/bin/activate
   ```

4. Instale las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

## Uso

1. Inicie la aplicación:
   ```bash
   python run.py
   ```

2. Acceda a la aplicación web desde su navegador:
   ```
   http://localhost:5000
   ```

3. Desde la interfaz principal puede:
   - Generar nuevas variantes de exámenes
   - Ver una vista previa de los exámenes generados
   - Descargar los exámenes en formato Word
   - Descargar las hojas de respuesta en PDF
   - Descargar las plantillas de calificación en PDF

## Estructura del Examen

Cada examen generado consta de tres series:

1. **Primera Serie (40 puntos)**: 10 preguntas de selección múltiple sobre conceptos básicos de estadística.
2. **Segunda Serie (20 puntos)**: 6 escenarios donde debe identificarse el tipo de gráfico estadístico más apropiado.
3. **Tercera Serie (40 puntos)**: 4 ejercicios prácticos que incluyen:
   - Cálculo del coeficiente de Gini
   - Tabla de distribución de frecuencias usando método Sturgers
   - Diagrama de tallo y hoja
   - Medidas de tendencia central

## Estructura del Proyecto

```
generador-examenes-estadistica/
├── app.py                  # Aplicación principal Flask
├── run.py                  # Script para iniciar la aplicación
├── requirements.txt        # Dependencias del proyecto
├── static/                 # Archivos estáticos (CSS, JS, imágenes)
├── templates/              # Plantillas HTML
│   ├── base.html
│   ├── index.html
│   └── previsualizar.html
├── examenes/               # Exámenes generados (Word)
├── hojas_respuesta/        # Hojas de respuesta (PDF)
├── plantillas/             # Plantillas de calificación (PDF)
└── variantes/              # Datos JSON de las variantes
```

## Solución de Problemas

**Error con Pillow/PIL**:
- Asegúrese de tener instaladas las dependencias del sistema para Pillow.
- En Ubuntu/Debian: `sudo apt-get install python3-pil.imagetk`
- En Windows: asegúrese de tener instalada la versión correcta de Pillow para su versión de Python.

**Error con python-docx**:
- Si tiene problemas con la instalación o uso de python-docx, intente: `pip install --upgrade python-docx`

## Personalización

El sistema está diseñado para ser fácilmente personalizable. Puede modificar:

1. El banco de preguntas en `app.py`:
   - `preguntas_base_primera` para la primera serie
   - `preguntas_base_segunda` para la segunda serie

2. Los datos base para los ejercicios prácticos:
   - `gini_exercises` para el problema del coeficiente de Gini
   - `sturgers_exercises` para el problema de distribución de frecuencias
   - `stem_leaf_exercises` para el problema de tallo y hoja
   - `central_tendency_exercises` para el problema de medidas de tendencia central

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Consulte el archivo LICENSE para más detalles.

## Contacto

Para preguntas, soporte o personalización, contacte a:
- Email: soporte@universidad-panamericana.edu.gt
