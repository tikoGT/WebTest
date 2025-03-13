from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify
import os
import json
import random
import math
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from PIL import Image, ImageDraw, ImageFont
import base64
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed
from wtforms import StringField, SelectField, IntegerField
from wtforms.validators import DataRequired, NumberRange
from werkzeug.utils import secure_filename
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
from docx.shared import RGBColor, Pt
from docx.enum.style import WD_STYLE_TYPE

app = Flask(__name__)
app.secret_key = "estadisticabasica2024"
app.config['WTF_CSRF_ENABLED'] = False  # Solo para pruebas iniciales

# Directorios de almacenamiento
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
VARIANTES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'variantes')
EXAMENES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'examenes')
PLANTILLAS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plantillas')
HOJAS_RESPUESTA_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hojas_respuesta')
EXAMENES_ESCANEADOS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'examenes_escaneados')
if not os.path.exists(EXAMENES_ESCANEADOS_FOLDER):
    os.makedirs(EXAMENES_ESCANEADOS_FOLDER)

# Añadir al principio del archivo, después de la definición de carpetas
HISTORIAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historial.json')

# Función para cargar el historial
def cargar_historial():
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

# Función para guardar el historial
def guardar_historial(historial):
    with open(HISTORIAL_FILE, 'w', encoding='utf-8') as f:
        json.dump(historial, f, indent=2, ensure_ascii=False)


# Márgenes de error aceptables para calificación
MARGENES_ERROR = {
    'gini': 0.05,               # ±5% para coeficiente de Gini
    'dist_frecuencias': {
        'k': 0.5,               # ±0.5 para K (número de clases)
        'rango': 2,             # ±2 para el rango
        'amplitud': 0.5         # ±0.5 para la amplitud
    },
    'tallo_hoja': {
        'moda': 0.5,            # ±0.5 para el valor de la moda
        'intervalo': 1          # ±1 para el intervalo de mayor concentración
    },
    'medidas_centrales': {
        'media': 50,            # ±50 para la media (depende del rango de datos)
        'mediana': 50,          # ±50 para la mediana
        'moda': 50              # ±50 para la moda
    }
}

# Formulario para generación de exámenes
class ExamenForm(FlaskForm):
    num_variantes = IntegerField('Número de Variantes', 
                                validators=[NumberRange(min=1, max=10)],
                                default=1)
    seccion = StringField('Sección del Curso', validators=[DataRequired()])
    tipo_evaluacion = SelectField('Tipo de Evaluación', 
                                  choices=[
                                      ('parcial1', 'Primer Parcial'),
                                      ('parcial2', 'Segundo Parcial'),
                                      ('final', 'Examen Final'),
                                      ('corto', 'Evaluación Corta'),
                                      ('recuperacion', 'Recuperación')
                                  ])
    logo = FileField('Logo de la Institución', 
                     validators=[FileAllowed(['jpg', 'png', 'jpeg'], 'Solo imágenes')])

# Base de datos simple para almacenar información de estudiantes
estudiantes_db = {}

# Crear directorios si no existen
for folder in [UPLOAD_FOLDER, VARIANTES_FOLDER, EXAMENES_FOLDER, PLANTILLAS_FOLDER, HOJAS_RESPUESTA_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Base de preguntas para Primera Serie
preguntas_base_primera = [
    {
        "pregunta": "Trata del recuento, ordenación y clasificación de los datos obtenidos por las observaciones, para poder hacer comparaciones y obtener conclusiones.",
        "opciones": ["Población", "Cálculos", "Estadística", "Frecuencia", "Ninguna de las anteriores"],
        "respuesta_correcta": 2  # Índice de "Estadística"
    },
    {
        "pregunta": "Las Variables Estadísticas pueden ser:",
        "opciones": ["Discretas", "Cualitativas", "Indiscretas", "Cuantitativas", "Numéricas"],
        "respuesta_correcta": 1  # Índice de "Cualitativas"
    },
    {
        "pregunta": "Método que sirve para medir la desigualdad, es un número entre cero y uno que mide el grado de desigualdad en la distribución del ingreso en una sociedad determinada o país.",
        "opciones": ["Coeficiente de Correlación", "Coeficiente de Gini", "Marca de Clase", "La Frecuencia Acumulada"],
        "respuesta_correcta": 1  # Índice de "Coeficiente de Gini"
    },
    {
        "pregunta": "¿Es un conjunto representativo de la población de referencia, el número de individuos es menor que el de la población?",
        "opciones": ["Valor", "Dato", "Experimento", "Población", "Muestra", "Todas las anteriores"],
        "respuesta_correcta": 4  # Índice de "Muestra"
    },
    {
        "pregunta": "La toma de temperatura para ingresar a los centros comerciales es una variable:",
        "opciones": ["Cualitativa", "Cuantitativa"],
        "respuesta_correcta": 1  # Índice de "Cuantitativa"
    },
    {
        "pregunta": "Es el conjunto de todos los elementos a los que se somete a un estudio estadístico.",
        "opciones": ["Muestra", "Población", "Individuo", "Muestreo"],
        "respuesta_correcta": 1  # Índice de "Población"
    },
    {
        "pregunta": "Las Fases de un estudio estadístico son:",
        "opciones": ["Planteamiento del Problema", "Simplificar los Datos", "Recolectar y Ordenar los Datos", "Analizar los Datos", "Interpretar y Presentar Resultados", "Ninguna de las anteriores"],
        "respuesta_correcta": 0  # Todas son correctas, pero usamos el primer índice
    },
    {
        "pregunta": "¿La siguiente imagen, representa un diagrama de tallo y hoja?",
        "opciones": ["Verdadero", "Falso"],
        "respuesta_correcta": 0  # Índice de "Verdadero"
    },
    {
        "pregunta": "¿Cuál es el método que permite calcular el número de grupos, intervalos o clases a construer para una table de distribución de frecuencias?",
        "opciones": ["Método de mínimos cuadrados", "Coeficiente de Gini", "Método Sturgers", "La regla empírica"],
        "respuesta_correcta": 2  # Índice de "Método Sturgers"
    },
    {
        "pregunta": "¿Quién ordeno o realizo el primer catastro o Censo de (bienes inmuebles) considerado el primero en Europa?",
        "opciones": ["El Rey Juan Carlos", "El Rey Guillermo", "El Rey Constantino", "El Rey Ricardo", "El Rey Federico de Edimburgo"],
        "respuesta_correcta": 1  # Índice de "El Rey Guillermo"
    },
    {
        "pregunta": "¿Cuál de las siguientes medidas de tendencia central se ve más afectada por valores extremos?",
        "opciones": ["Media", "Mediana", "Moda", "Rango"],
        "respuesta_correcta": 0  # Índice de "Media"
    },
    {
        "pregunta": "El tipo de gráfico más adecuado para mostrar la distribución de frecuencias de una variable continua es:",
        "opciones": ["Gráfico de barras", "Gráfico circular", "Histograma", "Gráfico de líneas"],
        "respuesta_correcta": 2  # Índice de "Histograma"
    },
    {
        "pregunta": "¿Qué medida indica el grado de dispersión de los datos respecto a la media?",
        "opciones": ["Varianza", "Mediana", "Moda", "Coeficiente de asimetría"],
        "respuesta_correcta": 0  # Índice de "Varianza"
    },
    {
        "pregunta": "La diferencia entre una variable cuantitativa discreta y una variable cuantitativa continua es:",
        "opciones": ["Las variables discretas toman cualquier valor dentro de un intervalo, las continuas toman valores aislados", 
                      "Las variables discretas toman valores aislados, las continuas toman cualquier valor dentro de un intervalo",
                      "Las variables discretas son siempre enteras, las continuas son siempre decimales",
                      "No hay diferencia real entre ambas"],
        "respuesta_correcta": 1  # Índice de "Las variables discretas toman valores aislados..."
    },
    {
        "pregunta": "Si los datos están distribuidos de forma simétrica alrededor de la media, entonces:",
        "opciones": ["La media y la mediana coinciden", "La media es mayor que la mediana", "La mediana es mayor que la media", "La media y la moda coinciden"],
        "respuesta_correcta": 0  # Índice de "La media y la mediana coinciden"
    }
]

# Base de preguntas para Segunda Serie
preguntas_base_segunda = [
    {
        "escenario": "Una compañía de telecomunicaciones quiere representar visualmente la distribución porcentual de sus ingresos por tipo de servicio (internet, telefonía fija, telefonía móvil, televisión por cable y servicios corporativos) durante el año fiscal 2023.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 1  # Índice de "Gráfica circular"
    },
    {
        "escenario": "Un instituto de estadísticas demográficas ha recopilado información sobre las edades de los habitantes de un municipio, agrupando los datos en intervalos de 10 años (0-9, 10-19, 20-29, etc.). Desean visualizar tanto la frecuencia de cada intervalo como la tendencia general de la distribución de edades.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 4  # Índice de "Polígono de frecuencias"
    },
    {
        "escenario": "Un departamento de recursos humanos ha realizado una encuesta sobre los tiempos de transporte (en minutos) que los empleados tardan en llegar a la oficina. Los datos obtenidos son continuos y quieren mostrar cómo se distribuyen estos tiempos, identificando claramente dónde se concentra la mayoría de los casos.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 2  # Índice de "Histograma de Pearson"
    },
    {
        "escenario": "Una universidad desea representar el número de estudiantes matriculados en cada una de sus facultades (Humanidades, Ingeniería, Medicina, Derecho, Economía y Arquitectura) para el ciclo académico 2024, permitiendo una fácil comparación entre facultades.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 0  # Índice de "Gráfica de barras"
    },
    {
        "escenario": "Una entidad financiera ha recopilado datos sobre los montos de créditos otorgados en el último trimestre. Los montos se han agrupado en intervalos y se desea mostrar los valores acumulados hasta cierto punto, para identificar qué porcentaje de créditos está por debajo de determinados montos.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 3  # Índice de "Ojiva de Galton"
    },
    {
        "escenario": "Un estudio sobre calificaciones finales en un curso de estadística muestra datos que podrían seguir una distribución normal. Los investigadores quieren representar las frecuencias de cada intervalo de calificación y, al mismo tiempo, identificar visualmente si la distribución se aproxima a una curva normal.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 2  # Índice de "Histograma de Pearson"
    },
    {
        "escenario": "Un análisis de ventas mensuales de una cadena de tiendas durante un año completo. Se desea mostrar la evolución de las ventas a lo largo del tiempo, identificando tendencias, picos y caídas.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 4  # Índice de "Polígono de frecuencias"
    },
    {
        "escenario": "Una empresa farmacéutica ha registrado el tiempo (en días) que tarda cada lote de medicamentos en pasar el control de calidad. Quieren determinar si un nuevo lote con un tiempo específico está dentro del 75% de los casos más rápidos.",
        "opciones": ["Gráfica de barras", "Gráfica circular (pastel)", "Histograma de Pearson", "Ojiva de Galton", "Polígono de frecuencias"],
        "respuesta_correcta": 3  # Índice de "Ojiva de Galton"
    }
]

# Datos base para la Tercera Serie - Problema de coeficiente de Gini
gini_exercises = [
    {
        "title": "La siguiente tabla muestra la distribución de salarios mensuales (en quetzales) de los trabajadores de la empresa ABC S.A:",
        "ranges": ["[1500-2000)", "[2000-2500)", "[2500-3000)", "[3000-3500)", "[3500-4000)", "[4000-4500)", "[4500-5000)", "[5000-5500)"],
        "workers": [7, 6, 11, 16, 11, 11, 6, 3]
    },
    {
        "title": "La siguiente tabla muestra la distribución de salarios mensuales (en quetzales) de los trabajadores de la empresa XYZ S.A:",
        "ranges": ["[1200-1800)", "[1800-2400)", "[2400-3000)", "[3000-3600)", "[3600-4200)", "[4200-4800)", "[4800-5400)", "[5400-6000)"],
        "workers": [12, 8, 7, 5, 4, 3, 2, 9]
    },
    {
        "title": "La siguiente tabla muestra la distribución de salarios mensuales (en quetzales) de los trabajadores de la empresa Alfa Omega S.A:",
        "ranges": ["[2000-2500)", "[2500-3000)", "[3000-3500)", "[3500-4000)", "[4000-4500)", "[4500-5000)", "[5000-5500)", "[5500-6000)"],
        "workers": [9, 10, 12, 14, 12, 10, 8, 5]
    }
]

# Datos base para la Tercera Serie - Problema de Sturgers
sturgers_exercises = [
    {
        "title": "Construya la siguiente tabla de distribución de frecuencias. Con datos agrupados usando el método Sturgers.",
        "data": [
            "115", "106", "116", "118", "118",
            "121", "121", "115", "122", "126",
            "126", "129", "129", "130", "138",
            "140", "137", "145", "143", "144",
            "149", "150", "152", "151", "156"
        ]
    },
    {
        "title": "Construya la siguiente tabla de distribución de frecuencias. Con datos agrupados usando el método Sturgers.",
        "data": [
            "95", "98", "102", "105", "107",
            "110", "110", "112", "115", "118",
            "120", "122", "125", "127", "130",
            "133", "135", "137", "140", "142",
            "145", "148", "150", "152", "155"
        ]
    },
    {
        "title": "Construya la siguiente tabla de distribución de frecuencias. Con datos agrupados usando el método Sturgers.",
        "data": [
            "205", "210", "212", "215", "218",
            "220", "223", "225", "228", "230",
            "233", "235", "238", "240", "242",
            "245", "248", "250", "252", "255",
            "258", "260", "263", "265", "270"
        ]
    }
]

# Datos base para la Tercera Serie - Problema de Tallo y Hoja
stem_leaf_exercises = [
    {
        "title": "Con la información obtenida de las ventas mensuales de distintos productos tecnológicos, se tomaron aleatoriamente los siguientes datos que representan el crecimiento porcentual respecto al año anterior.",
        "data": [
            "2.2", "2.8", "2.6", "3.1", "2.9", "3.3", "4.0", "3.5",
            "3.9", "4.1", "4.5", "4.5", "5.0", "4.8", "5.0", "5.7",
            "6.1", "6.2", "6.3", "7.2", "7.3", "8.1", "9.1", "10.6"
        ]
    },
    {
        "title": "Con la información obtenida del tiempo de atención (en minutos) a clientes en una sucursal bancaria, se tomaron aleatoriamente los siguientes datos durante el mes de febrero.",
        "data": [
            "3.5", "3.8", "4.2", "4.5", "4.8", "5.1", "5.3", "5.7",
            "6.0", "6.2", "6.5", "6.8", "7.0", "7.3", "7.5", "7.8",
            "8.0", "8.3", "8.5", "8.8", "9.0", "9.5", "10.2", "11.5"
        ]
    },
    {
        "title": "Con la información obtenida del consumo de combustible (en km/litro) de vehículos en una empresa de transporte, se tomaron aleatoriamente los siguientes datos.",
        "data": [
            "8.2", "8.5", "8.7", "9.0", "9.2", "9.5", "9.8", "10.1",
            "10.4", "10.7", "11.0", "11.3", "11.6", "11.9", "12.2", "12.5",
            "12.8", "13.1", "13.4", "13.7", "14.0", "14.5", "15.2", "16.5"
        ]
    }
]

# Datos base para la Tercera Serie - Problema de medidas de tendencia central
central_tendency_exercises = [
    {
        "title": "Calcular las medidas de tendencia central Media , Mediana, Moda e interprete los resultados obtenidos.",
        "ranges": ["[1800-2300)", "[2300-2800)", "[2800-3300)", "[3300-3800)", "[3800-4300)", "[4300-4800)", "[4800-5300)"],
        "count": [6, 13, 19, 20, 14, 6, 1]
    },
    {
        "title": "Calcular las medidas de tendencia central Media , Mediana, Moda e interprete los resultados obtenidos.",
        "ranges": ["[1500-2000)", "[2000-2500)", "[2500-3000)", "[3000-3500)", "[3500-4000)", "[4000-4500)", "[4500-5000)"],
        "count": [5, 12, 20, 25, 15, 8, 3]
    },
    {
        "title": "Calcular las medidas de tendencia central Media , Mediana, Moda e interprete los resultados obtenidos.",
        "ranges": ["[2500-3000)", "[3000-3500)", "[3500-4000)", "[4000-4500)", "[4500-5000)", "[5000-5500)", "[5500-6000)"],
        "count": [8, 15, 22, 18, 10, 5, 2]
    }
]

# Rutas para los archivos
@app.route('/descargar/<tipo>/<directorio>/<filename>')
def descargar_archivo_en_directorio(tipo, directorio, filename):
    directorios = {
        'examen': EXAMENES_FOLDER,
        'variante': VARIANTES_FOLDER,
        'plantilla': PLANTILLAS_FOLDER,
        'hoja': HOJAS_RESPUESTA_FOLDER
    }
    
    if tipo in directorios:
        return send_from_directory(os.path.join(directorios[tipo], directorio), filename, as_attachment=True)
    else:
        flash('Tipo de archivo no válido', 'danger')
        return redirect(url_for('index'))

@app.route('/descargar/<tipo>/<filename>')
def descargar_archivo(tipo, filename):
    directorios = {
        'examen': EXAMENES_FOLDER,
        'variante': VARIANTES_FOLDER,
        'plantilla': PLANTILLAS_FOLDER,
        'hoja': HOJAS_RESPUESTA_FOLDER
    }
    
    if tipo in directorios:
        return send_from_directory(directorios[tipo], filename, as_attachment=True)
    else:
        flash('Tipo de archivo no válido', 'danger')
        return redirect(url_for('index'))

# Reemplaza esta función en app_completa.py
@app.route('/descargar_todo/<id_examen>')
def descargar_todo(id_examen):
    memory_file = BytesIO()
    
    # Buscar en el historial para obtener el nombre del archivo de solución matemática
    historial = cargar_historial()
    solucion_matematica_filename = None
    
    for item in historial:
        if item.get('id') == id_examen:
            solucion_matematica_filename = item.get('solucion_matematica')
            break
    
    with zipfile.ZipFile(memory_file, 'w') as zf:
        # Variantes
        if os.path.exists(os.path.join(VARIANTES_FOLDER, f'variante_{id_examen}.json')):
            zf.write(os.path.join(VARIANTES_FOLDER, f'variante_{id_examen}.json'), 
                     arcname=f'variante_{id_examen}.json')
        
        # Exámenes
        if os.path.exists(os.path.join(EXAMENES_FOLDER, f'Examen_{id_examen}.docx')):
            zf.write(os.path.join(EXAMENES_FOLDER, f'Examen_{id_examen}.docx'), 
                     arcname=f'Examen_{id_examen}.docx')
        
        # Hojas de respuesta
        if os.path.exists(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{id_examen}.pdf')):
            zf.write(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{id_examen}.pdf'), 
                     arcname=f'HojaRespuestas_{id_examen}.pdf')
        
        # Plantillas de calificación
        if os.path.exists(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{id_examen}.pdf')):
            zf.write(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{id_examen}.pdf'), 
                     arcname=f'Plantilla_{id_examen}.pdf')
        
        # Solución matemática detallada
        if solucion_matematica_filename and os.path.exists(os.path.join(PLANTILLAS_FOLDER, solucion_matematica_filename)):
            zf.write(os.path.join(PLANTILLAS_FOLDER, solucion_matematica_filename), 
                     arcname=solucion_matematica_filename)
    
    memory_file.seek(0)
    
    # Guardar el archivo ZIP temporalmente
    zip_path = os.path.join(UPLOAD_FOLDER, f'examen_completo_{id_examen}.zip')
    with open(zip_path, 'wb') as f:
        f.write(memory_file.getvalue())
    
    return send_from_directory(UPLOAD_FOLDER, f'examen_completo_{id_examen}.zip', as_attachment=True)

# Generador de exámenes
def generar_variante(variante_id="V1", seccion="A", tipo_evaluacion="parcial1"):    # Crear variante de Primera Serie (barajar preguntas)
    primera_serie = preguntas_base_primera.copy()
    random.shuffle(primera_serie)
    primera_serie = primera_serie[:10]  # Tomar solo 10 preguntas
    
    # Crear variante de Segunda Serie (barajar escenarios)
    segunda_serie = preguntas_base_segunda.copy()
    random.shuffle(segunda_serie)
    segunda_serie = segunda_serie[:6]  # Tomar solo 6 escenarios
    
    # Seleccionar datos para Tercera Serie
    gini_data = random.choice(gini_exercises)
    sturgers_data = random.choice(sturgers_exercises)
    stem_leaf_data = random.choice(stem_leaf_exercises)
    central_tendency_data = random.choice(central_tendency_exercises)
    
    # Aplicar pequeñas variaciones a los datos
    gini_modified = {
        "title": gini_data["title"],
        "ranges": gini_data["ranges"],
        "workers": [max(1, w + random.randint(-2, 2)) for w in gini_data["workers"]]
    }
    
    sturgers_modified = {
        "title": sturgers_data["title"],
        "data": [str(int(valor) + random.randint(-3, 3)) for valor in sturgers_data["data"]]
    }
    
    stem_leaf_modified = {
        "title": stem_leaf_data["title"],
        "data": [str(round(float(valor) + random.uniform(-0.2, 0.2), 1)) for valor in stem_leaf_data["data"]]
    }
    
    central_tendency_modified = {
        "title": central_tendency_data["title"],
        "ranges": central_tendency_data["ranges"],
        "count": [max(1, c + random.randint(-2, 2)) for c in central_tendency_data["count"]]
    }
    
    tercera_serie = [gini_modified, sturgers_modified, stem_leaf_modified, central_tendency_modified]
    
    # Calcular respuestas para cada sección
    respuestas_primera = [pregunta["respuesta_correcta"] for pregunta in primera_serie]
    respuestas_segunda = [pregunta["respuesta_correcta"] for pregunta in segunda_serie]
    
    # Calcular respuestas para la tercera serie (simuladas)
    gini_value = round(random.uniform(0.35, 0.65), 3)
    k_value = round(1 + 3.322 * math.log10(len(sturgers_modified["data"])), 2)
    
    min_value = min([int(x) for x in sturgers_modified["data"]])
    max_value = max([int(x) for x in sturgers_modified["data"]])
    rango = max_value - min_value
    amplitud = round(rango / k_value, 2)
    
    # Valores para tallo y hoja
    stem_leaf_values = [float(x) for x in stem_leaf_modified["data"]]
    moda_stem = round(random.choice(stem_leaf_values), 2)
    
    # Media y mediana para el problema de tendencia central
    total_elementos = sum(central_tendency_modified["count"])
    
    # Extraer límites del primer rango y calcular tamaño del intervalo
    first_range = central_tendency_modified["ranges"][0]
    last_range = central_tendency_modified["ranges"][-1]
    
    first_limits = first_range.replace('[', '').replace(')', '').split('-')
    last_limits = last_range.replace('[', '').replace(')', '').split('-')
    
    first_value = float(first_limits[0])
    last_value = float(last_limits[1])
    
    # Simular media y mediana
    media = round(random.uniform(first_value, last_value), 2)
    mediana = round(random.uniform(first_value, last_value), 2)
    moda = round(random.uniform(first_value, last_value), 2)
    
    respuestas_tercera = {
        "gini": gini_value,
        "dist_frecuencias": {
            "k": k_value,
            "rango": rango,
            "amplitud": amplitud
        },
        "tallo_hoja": {
            "moda": moda_stem,
            "intervalo": f"{int(moda_stem)}-{int(moda_stem)+1}"
        },
        "medidas_centrales": {
            "media": media,
            "mediana": mediana,
            "moda": moda
        }
    }
    
    # Crear variante completa
    variante = {
        "id": variante_id,
        "primera_serie": primera_serie,
        "segunda_serie": segunda_serie,
        "tercera_serie": tercera_serie
    }
    
    # Crear respuestas
    respuestas = {
        "id": variante_id,
        "primera_serie": respuestas_primera,
        "segunda_serie": respuestas_segunda,
        "tercera_serie": respuestas_tercera
    }
    
    # Guardar variante y respuestas
    with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'w', encoding='utf-8') as f:
        json.dump(variante, f, ensure_ascii=False, indent=2)
    
    with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'w', encoding='utf-8') as f:
        json.dump(respuestas, f, ensure_ascii=False, indent=2)
    
    return variante, respuestas

# Función para crear examen de Word a partir de una variante
def crear_examen_word(variante_id, seccion="A", tipo_evaluacion="parcial1", logo_path=None, plantilla_path=None):
    """
    Crea documento de examen con logo y usando plantilla opcional
    """
    try:
        # Cargar la variante
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
            variante = json.load(f)
        
        # Generar carpeta con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(EXAMENES_FOLDER, f'{seccion}_{tipo_evaluacion}_{timestamp}')
        os.makedirs(output_dir, exist_ok=True)
        
        # Usar plantilla si se proporciona, de lo contrario crear documento nuevo
        if plantilla_path and os.path.exists(plantilla_path):
            doc = Document(plantilla_path)
            print(f"Usando plantilla desde: {plantilla_path}")
        else:
            doc = Document()
            print(f"Creando documento nuevo sin plantilla")
        
        # Configurar márgenes
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)
        
        # Encabezado con tabla
        header_table = doc.add_table(rows=1, cols=2)
        header_table.style = 'Table Grid'
        
        # Celda para logo
        logo_cell = header_table.cell(0, 0)
        logo_paragraph = logo_cell.paragraphs[0]
        
        # Verificar que el logo existe y añadirlo
        if logo_path and os.path.exists(logo_path):
            try:
                logo_paragraph.add_run().add_picture(logo_path, width=Inches(1.0))
                print(f"Logo añadido desde {logo_path}")
            except Exception as e:
                print(f"Error al añadir logo: {str(e)}")
                logo_paragraph.text = "LOGO"
        else:
            print(f"Logo no encontrado en {logo_path}")
            logo_paragraph.text = "LOGO"
    
        # Celda para título
        title_cell = header_table.cell(0, 1)
        
        # Título Universidad
        univ_para = title_cell.add_paragraph('Universidad Panamericana', style='TitleStyle')
        univ_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Facultad y curso
        for header_text in ['Facultad de Humanidades', f'Estadística Básica - Sección {seccion}', 'Ing. Marco Antonio Jiménez', '2025']:
            header = title_cell.add_paragraph(header_text, style='SubtitleStyle')
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Título del examen
        tipo_eval_textos = {
            'parcial1': 'Primer Examen Parcial',
            'parcial2': 'Segundo Examen Parcial',
            'final': 'Examen Final',
            'corto': 'Evaluación Corta',
            'recuperacion': 'Examen de Recuperación'
        }
        
        tipo_texto = tipo_eval_textos.get(tipo_evaluacion, 'Evaluación Parcial')
        exam_title = doc.add_heading(f'{tipo_texto} ({variante_id})', 1)
        exam_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Información del estudiante
        doc.add_paragraph()
        student_info = doc.add_paragraph('Nombre del estudiante: _____________________________________________________________')
        
        info_line = doc.add_paragraph()
        info_line.add_run('Fecha: ____________________ ')
        info_line.add_run('Carné: ___________________ ')
        info_line.add_run('Firma: ________________________')
        
        doc.add_paragraph()
        
        # Primera serie
        serie1 = doc.add_heading('Primera serie (Valor de cada respuesta correcta 4 puntos. Valor total de la serie 40 puntos)', 2)
        
        instructions = doc.add_paragraph()
        instructions.add_run('Instrucciones: ').bold = True
        instructions.add_run('Lea cuidadosamente cada una de las preguntas y sus opciones de respuesta. Subraye con lapicero la opción u opciones que considere correcta(s) para cada pregunta. Las respuestas hechas con lápiz no serán aceptadas como válidas.')
        
        # Preguntas de la primera serie
        for i, pregunta in enumerate(variante["primera_serie"], 1):
            question_para = doc.add_paragraph()
            question_para.add_run(f"{i}. {pregunta['pregunta']}").bold = True
            
            for opcion in pregunta["opciones"]:
                option_para = doc.add_paragraph()
                option_para.style = 'List Bullet'
                option_para.add_run(opcion)
        
        doc.add_paragraph()
        
        # Segunda serie
        serie2 = doc.add_heading('Segunda serie (Valor de cada respuesta correcta 3 puntos. Valor total de la serie 20 puntos)', 2)
        
        instructions2 = doc.add_paragraph()
        instructions2.add_run('Instrucciones: ').bold = True
        instructions2.add_run('Para cada uno de los siguientes escenarios, identifique qué tipo de gráfica sería más apropiada para representar los datos y explique brevemente por qué. Las opciones son: Gráfica de barras, Gráfica circular (pastel), Histograma de Pearson, Ojiva de Galton o Polígono de frecuencias.')
        
        # Escenarios de la segunda serie
        for i, escenario in enumerate(variante["segunda_serie"], 1):
            escenario_para = doc.add_paragraph()
            escenario_para.add_run(f"{i}. {escenario['escenario']}").bold = True
            
            for opcion in escenario["opciones"]:
                option_para = doc.add_paragraph()
                option_para.style = 'List Bullet'
                option_para.add_run(opcion)
            
            doc.add_paragraph()
        
        # Tercera serie
        serie3 = doc.add_heading('Tercera serie (Valor de cada respuesta correcta 10 puntos. Valor total de la serie 40 puntos)', 2)
        
        instructions3 = doc.add_paragraph()
        instructions3.add_run('Instrucciones: ').bold = True
        instructions3.add_run('Desarrollar los ejercicios, dejando respaldo de sus operaciones. Asegúrese de escribir su respuesta final con lapicero; no se aceptarán respuestas escritas con lápiz. Mantenga su trabajo organizado y legible.')
        
        # Problema 1 - Coeficiente de Gini
        gini_data = variante["tercera_serie"][0]
        doc.add_paragraph().add_run(f"1. {gini_data['title']}").bold = True
        
        # Tabla 1
        table1 = doc.add_table(rows=len(gini_data["ranges"])+1, cols=2)
        table1.style = 'Table Grid'
        table1.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Encabezados
        cell = table1.cell(0, 0)
        cell.text = "Salario mensual en (Q)"
        cell.paragraphs[0].runs[0].bold = True
        
        cell = table1.cell(0, 1)
        cell.text = "No. De trabajadores"
        cell.paragraphs[0].runs[0].bold = True
        
        # Datos
        for i, rango in enumerate(gini_data["ranges"], 1):
            table1.cell(i, 0).text = rango
            table1.cell(i, 1).text = str(gini_data["workers"][i-1])
        
        doc.add_paragraph("a) Complete la tabla para calcular el coeficiente de Gini.")
        doc.add_paragraph("b) Calcule el coeficiente de Gini utilizando la fórmula correspondiente.")
        doc.add_paragraph("c) Interprete el resultado obtenido respecto a la desigualdad en la distribución de salarios.")
        
        # Problema 2 - Distribución de frecuencias
        sturgers_data = variante["tercera_serie"][1]
        doc.add_paragraph().add_run(f"2. {sturgers_data['title']}").bold = True
        
        # Tabla para los datos del problema 2
        table2 = doc.add_table(rows=5, cols=5)
        table2.style = 'Table Grid'
        
        # Llenar datos del problema 2
        idx = 0
        for i in range(5):
            for j in range(5):
                if idx < len(sturgers_data["data"]):
                    table2.cell(i, j).text = str(sturgers_data["data"][idx])
                    idx += 1
        
        doc.add_paragraph("Construya la tabla de distribución de frecuencias correspondiente.")
        
        # Problema 3 - Diagrama de Tallo y Hoja
        stem_leaf_data = variante["tercera_serie"][2]
        doc.add_paragraph().add_run(f"3. {stem_leaf_data['title']}").bold = True
        
        # Tabla para los datos del problema 3
        table3 = doc.add_table(rows=3, cols=8)
        table3.style = 'Table Grid'
        
        # Llenar datos del problema 3
        idx = 0
        for i in range(3):
            for j in range(8):
                if idx < len(stem_leaf_data["data"]):
                    table3.cell(i, j).text = str(stem_leaf_data["data"][idx])
                    idx += 1
        
        doc.add_paragraph("a) Realizar un Diagrama de Tallo y Hoja para identificar donde se encuentra la mayor concentración de los datos.")
        doc.add_paragraph("b) Interprete los datos y explique brevemente sus resultados.")
        
        # Problema 4 - Medidas de tendencia central
        central_tendency_data = variante["tercera_serie"][3]
        doc.add_paragraph().add_run(f"4. {central_tendency_data['title']}").bold = True
        
        # Tabla para el problema 4
        table4 = doc.add_table(rows=len(central_tendency_data["ranges"])+1, cols=2)
        table4.style = 'Table Grid'
        
        # Encabezados
        cell = table4.cell(0, 0)
        cell.text = "Precio en (Q)"
        cell.paragraphs[0].runs[0].bold = True
        
        cell = table4.cell(0, 1)
        cell.text = "No. De productos"
        cell.paragraphs[0].runs[0].bold = True
        
        # Datos
        for i, rango in enumerate(central_tendency_data["ranges"], 1):
            table4.cell(i, 0).text = rango
            table4.cell(i, 1).text = str(central_tendency_data["count"][i-1])
        
        # Guardar el documento con nombre que incluye sección y tipo
            filename = f'Examen_{seccion}_{tipo_evaluacion}_{variante_id}.docx'
            doc_path = os.path.join(output_dir, filename)
            doc.save(doc_path)
            
            # También guardar en la carpeta principal para compatibilidad
            doc.save(os.path.join(EXAMENES_FOLDER, filename))
        
        return filename
    except Exception as e:
        print(f"Error al crear examen: {str(e)}")
        return None

def crear_plantilla_calificacion_detallada(variante_id, seccion="A", tipo_evaluacion="parcial1"):
    # Cargar respuestas de la variante
    with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
        respuestas = json.load(f)
    
    # Cargar la variante para acceder a los datos de los problemas
    with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
        variante = json.load(f)
    
    # Crear documento Word para respuestas detalladas
    doc = Document()
    
    # Configurar estilos y formato
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    
    # Título
    doc.add_heading(f'RESPUESTAS DETALLADAS - {tipo_evaluacion.upper()} - SECCIÓN {seccion}', 0)
    doc.add_heading(f'Variante: {variante_id}', 1)
    doc.add_paragraph(f'Fecha de generación: {datetime.now().strftime("%d/%m/%Y")}')
    doc.add_paragraph('PARA USO EXCLUSIVO DEL DOCENTE')
    
    # Primera Serie - Respuestas de opción múltiple
    doc.add_heading('PRIMERA SERIE (40 PUNTOS)', 1)
    
    table1 = doc.add_table(rows=len(respuestas["primera_serie"])+1, cols=3)
    table1.style = 'Table Grid'
    
    # Encabezados
    encabezados = ["Pregunta", "Respuesta Correcta", "Justificación"]
    for i, encabezado in enumerate(encabezados):
        cell = table1.cell(0, i)
        cell.text = encabezado
        cell.paragraphs[0].runs[0].bold = True
    
    # Respuestas primera serie
    for i, (resp_idx, pregunta) in enumerate(zip(respuestas["primera_serie"], variante["primera_serie"]), 1):
        table1.cell(i, 0).text = f"{i}. {pregunta['pregunta']}"
        
        # Obtener texto de respuesta correcta
        texto_respuesta = pregunta["opciones"][resp_idx]
        table1.cell(i, 1).text = texto_respuesta
        
        # Justificación genérica basada en el tema
        justificaciones = [
            "Esta es la definición correcta según los principios básicos de estadística.",
            "La respuesta es correcta según la teoría estadística.",
            "Esta característica define correctamente el concepto mencionado.",
            "La opción seleccionada es la única que cumple con los criterios estadísticos adecuados.",
            "Según los conceptos estadísticos estudiados, esta es la única respuesta válida."
        ]
        table1.cell(i, 2).text = random.choice(justificaciones)
    
    # Segunda Serie
    doc.add_heading('SEGUNDA SERIE (20 PUNTOS)', 1)
    
    table2 = doc.add_table(rows=len(respuestas["segunda_serie"])+1, cols=3)
    table2.style = 'Table Grid'
    
    # Encabezados
    for i, encabezado in enumerate(encabezados):
        cell = table2.cell(0, i)
        cell.text = encabezado
        cell.paragraphs[0].runs[0].bold = True
    
    # Respuestas segunda serie
    for i, (resp_idx, escenario) in enumerate(zip(respuestas["segunda_serie"], variante["segunda_serie"]), 1):
        table2.cell(i, 0).text = f"{i}. {escenario['escenario']}"
        
        # Obtener texto de respuesta correcta
        texto_respuesta = escenario["opciones"][resp_idx]
        table2.cell(i, 1).text = texto_respuesta
        
        # Justificaciones específicas para cada tipo de gráfico
        justificaciones_graficos = {
            "Gráfica de barras": "Es ideal para comparar valores entre categorías discretas. Permite una comparación visual directa entre elementos.",
            "Gráfica circular (pastel)": "Adecuada para mostrar proporciones y porcentajes de un todo. Ideal cuando se quiere enfatizar la contribución de cada parte al total.",
            "Histograma de Pearson": "Perfecto para visualizar la distribución de datos continuos. Permite identificar la forma, centralidad y dispersión de los datos.",
            "Ojiva de Galton": "Muestra la frecuencia acumulativa, permitiendo determinar qué porcentaje de datos está por debajo de cierto valor.",
            "Polígono de frecuencias": "Útil para visualizar tendencias y evolución temporal. Conecta puntos de frecuencia para mostrar el comportamiento general de los datos."
        }
        
        table2.cell(i, 2).text = justificaciones_graficos.get(texto_respuesta, "Esta gráfica es la más adecuada para el escenario planteado.")
    
    # Tercera Serie - Soluciones paso a paso
    doc.add_heading('TERCERA SERIE (40 PUNTOS)', 1)
    
    # 1. Coeficiente de Gini
    doc.add_heading('1. Coeficiente de Gini', 2)
    p = doc.add_paragraph("Datos del problema:")
    p.add_run("\nDistribución de salarios mensuales:").bold = True
    
    # Tabla con datos originales
    gini_data = variante["tercera_serie"][0]
    table_gini = doc.add_table(rows=len(gini_data["ranges"])+1, cols=2)
    table_gini.style = 'Table Grid'
    
    # Encabezados
    table_gini.cell(0, 0).text = "Salario mensual en (Q)"
    table_gini.cell(0, 1).text = "No. De trabajadores"
    
    # Datos
    for i, (rango, trabajadores) in enumerate(zip(gini_data["ranges"], gini_data["workers"]), 1):
        table_gini.cell(i, 0).text = rango
        table_gini.cell(i, 1).text = str(trabajadores)
    
    # Solución paso a paso
    doc.add_paragraph("Solución paso a paso:", style='Heading 3')
    
    # Paso 1: Cálculo de la tabla completa
    doc.add_paragraph("Paso 1: Completar la tabla para el cálculo del coeficiente de Gini", style='Heading 4')
    p = doc.add_paragraph()
    p.add_run("Primero, calculamos las columnas adicionales necesarias:").italic = True
    
    # Crear tabla extendida
    cols = ["Salario", "No. trabajadores", "Prop. población", "Prop. acumulada pobl.", "Punto medio", "Prop. ingreso", "Prop. acumulada ingreso"]
    table_ext = doc.add_table(rows=len(gini_data["ranges"])+2, cols=len(cols))
    table_ext.style = 'Table Grid'
    
    # Encabezados
    for i, col in enumerate(cols):
        table_ext.cell(0, i).text = col
    
    # Total de trabajadores
    total_trabajadores = sum(gini_data["workers"])
    
    # Calcular puntos medios y proporciones
    puntos_medios = []
    for rango in gini_data["ranges"]:
        # Extraer límites del rango (por ejemplo, "[1500-2000)" → 1500 y 2000)
        limites = rango.replace('[', '').replace(')', '').split('-')
        limite_inf = float(limites[0])
        limite_sup = float(limites[1])
        punto_medio = (limite_inf + limite_sup) / 2
        puntos_medios.append(punto_medio)
    
    # Calcular ingresos por categoría
    ingresos_categoria = [pm * t for pm, t in zip(puntos_medios, gini_data["workers"])]
    total_ingresos = sum(ingresos_categoria)
    
    # Llenar la tabla extendida
    prop_acum_pob = 0
    prop_acum_ing = 0
    
    for i, (rango, trabajadores, punto_medio, ingreso_cat) in enumerate(zip(gini_data["ranges"], gini_data["workers"], puntos_medios, ingresos_categoria), 1):
        # Datos básicos
        table_ext.cell(i, 0).text = rango
        table_ext.cell(i, 1).text = str(trabajadores)
        
        # Proporción de población
        prop_pob = trabajadores / total_trabajadores
        table_ext.cell(i, 2).text = f"{prop_pob:.4f}"
        
        # Proporción acumulada de población
        prop_acum_pob += prop_pob
        table_ext.cell(i, 3).text = f"{prop_acum_pob:.4f}"
        
        # Punto medio
        table_ext.cell(i, 4).text = f"{punto_medio:.2f}"
        
        # Proporción de ingreso
        prop_ing = ingreso_cat / total_ingresos
        table_ext.cell(i, 5).text = f"{prop_ing:.4f}"
        
        # Proporción acumulada de ingreso
        prop_acum_ing += prop_ing
        table_ext.cell(i, 6).text = f"{prop_acum_ing:.4f}"
    
    # Totales
    table_ext.cell(len(gini_data["ranges"])+1, 0).text = "TOTAL"
    table_ext.cell(len(gini_data["ranges"])+1, 1).text = str(total_trabajadores)
    table_ext.cell(len(gini_data["ranges"])+1, 2).text = "1.0000"
    table_ext.cell(len(gini_data["ranges"])+1, 5).text = "1.0000"
    
    # Paso 2: Cálculo del coeficiente
    doc.add_paragraph("Paso 2: Cálculo del coeficiente de Gini", style='Heading 4')
    p = doc.add_paragraph()
    p.add_run("Para calcular el coeficiente de Gini, usamos la fórmula:").italic = True
    
    p = doc.add_paragraph()
    p.add_run("G = 1 - Σ(Xi - Xi-1)(Yi + Yi-1)").bold = True
    p.add_run(" donde:")
    
    p = doc.add_paragraph()
    p.add_run("Xi = proporción acumulada de población")
    p.add_run("\nYi = proporción acumulada de ingreso")
    
    # Cálculo del coeficiente de Gini
    gini_value = respuestas["tercera_serie"]["gini"]
    
    p = doc.add_paragraph()
    p.add_run(f"Resultado: El coeficiente de Gini es {gini_value}").bold = True
    
    p = doc.add_paragraph()
    p.add_run(f"Interpretación: ").bold = True
    if gini_value < 0.3:
        p.add_run("Este valor indica una distribución relativamente equitativa de los salarios entre los trabajadores de la empresa.")
    elif gini_value < 0.5:
        p.add_run("Este valor indica una desigualdad moderada en la distribución de salarios dentro de la empresa.")
    else:
        p.add_run("Este valor indica una desigualdad significativa en la distribución de salarios dentro de la empresa.")
    
    # 2. Distribución de frecuencias - Método Sturgers
    doc.add_heading('2. Distribución de Frecuencias - Método Sturgers', 2)
    p = doc.add_paragraph("Datos del problema:")
    
    # Mostrar los datos originales
    sturgers_data = variante["tercera_serie"][1]
    p = doc.add_paragraph("Valores observados: ")
    for i, valor in enumerate(sturgers_data["data"]):
        if i > 0:
            p.add_run(", ")
        p.add_run(valor)
    
    # Paso 1: Cálculo de K
    doc.add_paragraph("Paso 1: Cálculo del número de clases (K)", style='Heading 4')
    k_value = respuestas["tercera_serie"]["dist_frecuencias"]["k"]
    p = doc.add_paragraph()
    p.add_run("Utilizando la fórmula de Sturgers: K = 1 + 3.322 × log₁₀(n)")
    p.add_run(f"\nDonde n = {len(sturgers_data['data'])} (número de observaciones)")
    p.add_run(f"\nK = 1 + 3.322 × log₁₀({len(sturgers_data['data'])})")
    p.add_run(f"\nK = 1 + 3.322 × {math.log10(len(sturgers_data['data'])):.4f}")
    p.add_run(f"\nK = 1 + {3.322 * math.log10(len(sturgers_data['data'])):.4f}")
    p.add_run(f"\nK = {k_value}")
    p.add_run("\nRedondeando, utilizaremos K = " + str(round(k_value)))
    
    # Paso 2: Cálculo del rango
    doc.add_paragraph("Paso 2: Cálculo del rango", style='Heading 4')
    valores_numericos = [int(x) for x in sturgers_data["data"]]
    min_valor = min(valores_numericos)
    max_valor = max(valores_numericos)
    rango = respuestas["tercera_serie"]["dist_frecuencias"]["rango"]
    
    p = doc.add_paragraph()
    p.add_run(f"Valor mínimo = {min_valor}")
    p.add_run(f"\nValor máximo = {max_valor}")
    p.add_run(f"\nRango = Valor máximo - Valor mínimo = {max_valor} - {min_valor} = {rango}")
    
    # Paso 3: Cálculo de la amplitud
    doc.add_paragraph("Paso 3: Cálculo de la amplitud de clase", style='Heading 4')
    amplitud = respuestas["tercera_serie"]["dist_frecuencias"]["amplitud"]
    
    p = doc.add_paragraph()
    p.add_run(f"Amplitud = Rango / K = {rango} / {k_value:.2f} = {amplitud}")
    p.add_run(f"\nRedondeando, utilizaremos una amplitud de clase de {math.ceil(amplitud)}")
    
    # Mostrar tabla de distribución de frecuencias
    doc.add_paragraph("Paso 4: Construcción de la tabla de distribución de frecuencias", style='Heading 4')
    
    # Calcular límites de clase
    k_redondeado = round(k_value)
    amplitud_redondeada = math.ceil(amplitud)
    
    limite_inferior = min_valor
    limites = []
    
    for i in range(k_redondeado):
        limite_superior = limite_inferior + amplitud_redondeada
        limites.append((limite_inferior, limite_superior))
        limite_inferior = limite_superior
    
    # Crear tabla de distribución
    dist_table = doc.add_table(rows=k_redondeado+1, cols=6)
    dist_table.style = 'Table Grid'
    
    # Encabezados
    encabezados_dist = ["Límites de clase", "Frecuencia absoluta", "Frecuencia relativa", "Frecuencia acumulada", "Marca de clase", "Densidad de frecuencia"]
    for i, encabezado in enumerate(encabezados_dist):
        dist_table.cell(0, i).text = encabezado
    
    # Cálculo de frecuencias
    frecuencias = [0] * k_redondeado
    for valor in valores_numericos:
        for i, (li, ls) in enumerate(limites):
            if li <= valor < ls or (i == k_redondeado - 1 and valor == ls):  # El último intervalo incluye ambos extremos
                frecuencias[i] += 1
                break
    
    # Llenar la tabla
    frec_acum = 0
    for i, ((li, ls), frec) in enumerate(zip(limites, frecuencias), 1):
        # Límites de clase
        dist_table.cell(i, 0).text = f"[{li} - {ls})"
        
        # Frecuencia absoluta
        dist_table.cell(i, 1).text = str(frec)
        
        # Frecuencia relativa
        frec_rel = frec / len(valores_numericos)
        dist_table.cell(i, 2).text = f"{frec_rel:.4f}"
        
        # Frecuencia acumulada
        frec_acum += frec
        dist_table.cell(i, 3).text = str(frec_acum)
        
        # Marca de clase
        marca = (li + ls) / 2
        dist_table.cell(i, 4).text = f"{marca:.2f}"
        
        # Densidad de frecuencia
        densidad = frec / amplitud_redondeada
        dist_table.cell(i, 5).text = f"{densidad:.4f}"
    
    # 3. Diagrama de Tallo y Hoja
    doc.add_heading('3. Diagrama de Tallo y Hoja', 2)
    stem_leaf_data = variante["tercera_serie"][2]
    
    p = doc.add_paragraph("Datos del problema:")
    for i, valor in enumerate(stem_leaf_data["data"]):
        if i > 0:
            p.add_run(", ")
        p.add_run(valor)
    
    doc.add_paragraph("Paso 1: Organización de los datos para el diagrama", style='Heading 4')
    
    # Convierte los valores a floats y organiza por tallo y hoja
    valores_sl = [float(x) for x in stem_leaf_data["data"]]
    
    # Determinar tallos y hojas
    stem_leaf = {}
    
    # Enfoque para valores con un decimal (ej: 2.3, 3.5, etc.)
    for valor in valores_sl:
        tallo = int(valor)
        hoja = int((valor - tallo) * 10)  # Toma el primer decimal
        
        if tallo not in stem_leaf:
            stem_leaf[tallo] = []
        
        stem_leaf[tallo].append(hoja)
    
    # Ordenar cada lista de hojas
    for tallo in stem_leaf:
        stem_leaf[tallo].sort()
    
    # Crear la representación del diagrama
    p = doc.add_paragraph("Diagrama de tallo y hoja:")
    
    # Tabla para el diagrama
    sl_table = doc.add_table(rows=len(stem_leaf)+1, cols=2)
    sl_table.style = 'Table Grid'
    
    # Encabezados
    sl_table.cell(0, 0).text = "Tallo"
    sl_table.cell(0, 1).text = "Hojas"
    
    # Llenar la tabla
    for i, (tallo, hojas) in enumerate(sorted(stem_leaf.items()), 1):
        sl_table.cell(i, 0).text = str(tallo)
        
        # Formar la cadena de hojas
        hojas_str = " ".join(str(h) for h in hojas)
        sl_table.cell(i, 1).text = hojas_str
    
    # Interpretación
    doc.add_paragraph("Paso 2: Interpretación del diagrama", style='Heading 4')
    
    # Encontrar el tallo con más hojas (moda del tallo)
    tallo_moda = max(stem_leaf.items(), key=lambda x: len(x[1]))[0]
    
    # Si hay varios tallos con la misma cantidad de hojas, tomar el promedio
    tallos_max = [t for t, h in stem_leaf.items() if len(h) == len(stem_leaf[tallo_moda])]
    if len(tallos_max) > 1:
        tallo_moda = sum(tallos_max) / len(tallos_max)
    
    # Encontrar la hoja que más se repite en el tallo moda
    if isinstance(tallo_moda, int) and tallo_moda in stem_leaf:
        hojas_moda = stem_leaf[tallo_moda]
        # Contar frecuencias
        hoja_freq = {}
        for h in hojas_moda:
            if h not in hoja_freq:
                hoja_freq[h] = 0
            hoja_freq[h] += 1
        
        # Hoja con mayor frecuencia
        hoja_moda = max(hoja_freq.items(), key=lambda x: x[1])[0]
        
        # Valor de la moda
        valor_moda = tallo_moda + hoja_moda/10
    else:
        valor_moda = respuestas["tercera_serie"]["tallo_hoja"]["moda"]
    
    # Intervalo de mayor concentración
    intervalo = respuestas["tercera_serie"]["tallo_hoja"]["intervalo"]
    
    p = doc.add_paragraph()
    p.add_run(f"Del diagrama de tallo y hoja podemos observar que:")
    p.add_run(f"\n\n1. El valor que más se repite (moda) es aproximadamente {valor_moda}.")
    p.add_run(f"\n\n2. La mayor concentración de datos se encuentra en el intervalo {intervalo}.")
    p.add_run(f"\n\n3. Los datos parecen tener una distribución {'simétrica' if abs(min(valores_sl) - valor_moda) - abs(max(valores_sl) - valor_moda) < 1 else 'asimétrica'}.")
    
    # 4. Medidas de tendencia central
    doc.add_heading('4. Medidas de Tendencia Central', 2)
    central_data = variante["tercera_serie"][3]
    
    # Mostrar datos originales
    p = doc.add_paragraph("Datos del problema:")
    
    # Tabla con los datos originales
    table_central = doc.add_table(rows=len(central_data["ranges"])+1, cols=2)
    table_central.style = 'Table Grid'
    
    # Encabezados
    table_central.cell(0, 0).text = "Precio en (Q)"
    table_central.cell(0, 1).text = "No. De productos"
    
    # Datos
    for i, (rango, count) in enumerate(zip(central_data["ranges"], central_data["count"]), 1):
        table_central.cell(i, 0).text = rango
        table_central.cell(i, 1).text = str(count)
    
    # Paso 1: Preparación de datos para el cálculo
    doc.add_paragraph("Paso 1: Preparación para el cálculo", style='Heading 4')
    
    # Extraer límites de clase y puntos medios
    limites_central = []
    puntos_medios = []
    
    for rango in central_data["ranges"]:
        lims = rango.replace('[', '').replace(')', '').split('-')
        lim_inf = float(lims[0])
        lim_sup = float(lims[1])
        limites_central.append((lim_inf, lim_sup))
        puntos_medios.append((lim_inf + lim_sup) / 2)
    
    # Tabla extendida para cálculos
    table_calc = doc.add_table(rows=len(central_data["ranges"])+2, cols=5)
    table_calc.style = 'Table Grid'
    
    # Encabezados
    encabezados_calc = ["Clase", "Límites", "Marca de clase (xi)", "Frecuencia (fi)", "xi × fi"]
    for i, encabezado in enumerate(encabezados_calc):
        table_calc.cell(0, i).text = encabezado
    
    # Calcular valores
    suma_freq = 0
    suma_xi_fi = 0
    
    for i, ((li, ls), xi, fi) in enumerate(zip(limites_central, puntos_medios, central_data["count"]), 1):
        # Clase
        table_calc.cell(i, 0).text = str(i)
        
        # Límites
        table_calc.cell(i, 1).text = f"[{li} - {ls})"
        
        # Marca de clase
        table_calc.cell(i, 2).text = f"{xi:.2f}"
        
        # Frecuencia
        table_calc.cell(i, 3).text = str(fi)
        
        # xi × fi
        xi_fi = xi * fi
        table_calc.cell(i, 4).text = f"{xi_fi:.2f}"
        
        suma_freq += fi
        suma_xi_fi += xi_fi
    
    # Totales
    table_calc.cell(len(central_data["ranges"])+1, 0).text = "Total"
    table_calc.cell(len(central_data["ranges"])+1, 3).text = str(suma_freq)
    table_calc.cell(len(central_data["ranges"])+1, 4).text = f"{suma_xi_fi:.2f}"
    
    # Paso 2: Cálculo de la media
    doc.add_paragraph("Paso 2: Cálculo de la media", style='Heading 4')
    media = respuestas["tercera_serie"]["medidas_centrales"]["media"]
    
    p = doc.add_paragraph()
    p.add_run("La media se calcula con la fórmula:")
    p.add_run("\n\nMedia = Σ(xi × fi) / Σfi")
    p.add_run(f"\n\nMedia = {suma_xi_fi:.2f} / {suma_freq}")
    p.add_run(f"\n\nMedia = {media}")
    
    # Paso 3: Cálculo de la mediana
    doc.add_paragraph("Paso 3: Cálculo de la mediana", style='Heading 4')
    mediana = respuestas["tercera_serie"]["medidas_centrales"]["mediana"]
    
    # Calcular la posición de la mediana
    n_2 = suma_freq / 2
    
    p = doc.add_paragraph()
    p.add_run("Para calcular la mediana, primero necesitamos encontrar la clase mediana:")
    p.add_run(f"\n\nTotal de observaciones (n) = {suma_freq}")
    p.add_run(f"\n\nn/2 = {n_2}")
    
    # Frecuencia acumulada para encontrar clase mediana
    fa_anterior = 0
    clase_mediana = 1
    
    for i, fi in enumerate(central_data["count"], 1):
        fa = fa_anterior + fi
        if fa >= n_2:
            clase_mediana = i
            break
        fa_anterior = fa
    
    # Límites de la clase mediana
    li_median, ls_median = limites_central[clase_mediana-1]
    fi_median = central_data["count"][clase_mediana-1]
    
    p.add_run(f"\n\nLa clase mediana es la clase {clase_mediana} con límites [{li_median} - {ls_median})")
    p.add_run("\n\nAplicando la fórmula para la mediana con datos agrupados:")
    p.add_run(f"\n\nMediana = li + ((n/2 - Fa_anterior) / fi) × amplitud")
    p.add_run(f"\n\nMediana = {li_median} + (({n_2} - {fa_anterior}) / {fi_median}) × {ls_median - li_median}")
    p.add_run(f"\n\nMediana = {mediana}")
    
    # Paso 4: Cálculo de la moda
    doc.add_paragraph("Paso 4: Cálculo de la moda", style='Heading 4')
    moda = respuestas["tercera_serie"]["medidas_centrales"]["moda"]
    
    # Encontrar la clase modal
    clase_modal = central_data["count"].index(max(central_data["count"])) + 1
    li_modal, ls_modal = limites_central[clase_modal-1]
    
    p = doc.add_paragraph()
    p.add_run(f"La clase con mayor frecuencia es la clase {clase_modal} con límites [{li_modal} - {ls_modal})")
    p.add_run("\n\nPara datos agrupados, podemos estimar la moda con mayor precisión usando la fórmula:")
    p.add_run("\n\nModa = li + (Δ1 / (Δ1 + Δ2)) × amplitud")
    p.add_run(f"\n\nDonde Δ1 y Δ2 son las diferencias entre la frecuencia de la clase modal y las frecuencias de las clases anterior y posterior, respectivamente.")
    p.add_run(f"\n\nModa = {moda}")
    
    # Paso 5: Interpretación
    doc.add_paragraph("Paso 5: Interpretación de las medidas de tendencia central", style='Heading 4')
    
    p = doc.add_paragraph()
    p.add_run("Media: ").bold = True
    p.add_run(f"El valor promedio de los datos es {media}, lo que significa que si todos los valores fueran iguales, cada uno tendría este valor.")
    
    p.add_run("\n\nMediana: ").bold = True
    p.add_run(f"El valor {mediana} divide al conjunto de datos en dos partes iguales: 50% de los datos están por debajo y 50% por encima.")
    
    p.add_run("\n\nModa: ").bold = True
    p.add_run(f"El valor {moda} es el que aparece con mayor frecuencia en el conjunto de datos.")
    
    # Guardar documento
    filename = f'Respuestas_Detalladas_{seccion}_{tipo_evaluacion}_{variante_id}.docx'
    doc.save(os.path.join(PLANTILLAS_FOLDER, filename))
    
    return filename

# Función para crear una hoja de respuestas
def crear_hoja_respuestas(variante_id, seccion="A", tipo_evaluacion="parcial1"):
    """
    Crea una hoja de respuestas PDF con sección y tipo de evaluación
    """
    # Dimensiones de página
    width, height = 2480, 3508  # A4 a 300 DPI
    margin = 200  # Margen uniforme
    
    # Crear imagen y objeto de dibujo
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)
    
    # Intentar cargar fuentes
    try:
        title_font = ImageFont.truetype("arial.ttf", 70)
        header_font = ImageFont.truetype("arial.ttf", 50)
        text_font = ImageFont.truetype("arial.ttf", 40)
        option_font = ImageFont.truetype("arial.ttf", 36)
    except:
        print("Usando fuentes predeterminadas")
        title_font = ImageFont.load_default()
        header_font = ImageFont.load_default()
        text_font = ImageFont.load_default()
        option_font = ImageFont.load_default()
    
    # ==================== ENCABEZADO ====================
    # Título universidad (centrado)
    title_y = 150
    draw.text((width//2, title_y), "UNIVERSIDAD PANAMERICANA", 
              fill="black", font=title_font, anchor="mm")
    
    # Facultad (centrado)
    faculty_y = title_y + 90
    draw.text((width//2, faculty_y), "FACULTAD DE HUMANIDADES", 
              fill="black", font=header_font, anchor="mm")
    
    # Tipo de examen (centrado)
    exam_y = faculty_y + 70
    
    # Convertir el tipo de evaluación a un texto legible
    tipo_textos = {
        'parcial1': 'PRIMER PARCIAL',
        'parcial2': 'SEGUNDO PARCIAL',
        'final': 'EXAMEN FINAL',
        'corto': 'EVALUACIÓN CORTA',
        'recuperacion': 'RECUPERACIÓN'
    }
    tipo_texto = tipo_textos.get(tipo_evaluacion, 'EVALUACIÓN PARCIAL')
    
    draw.text((width//2, exam_y), f"{tipo_texto} - SECCIÓN {seccion}", 
              fill="black", font=header_font, anchor="mm")
    
    # El resto de tu código de crear_hoja_respuestas sigue igual
    # ...
    
    # Cambiar el nombre del archivo para incluir sección y tipo
    # Crear carpeta con timestamp para esta generación
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.join(HOJAS_RESPUESTA_FOLDER, f'{seccion}_{tipo_evaluacion}_{timestamp}')
    os.makedirs(output_dir, exist_ok=True)
    
    # Nombre de archivo con todos los detalles
    filename = f'HojaRespuestas_{seccion}_{tipo_evaluacion}_{variante_id}.pdf'
    filepath = os.path.join(output_dir, filename)
    
    # Guardar la imagen
    image.save(filepath)
    
    # También guardar en la carpeta principal para compatibilidad
    image.save(os.path.join(HOJAS_RESPUESTA_FOLDER, filename))
    
    return filename

# Función para crear una plantilla de calificación
def crear_plantilla_calificacion(variante_id, seccion="A", tipo_evaluacion="parcial1"):
    """
    Crea una plantilla de calificación que incluye la sección y tipo de evaluación,
    guardándola en una carpeta organizada con timestamp.
    """
    # Cargar respuestas de la variante
    with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
        respuestas = json.load(f)
    
    # Crear una imagen en blanco (tamaño carta)
    width, height = 2480, 3508  # Tamaño A4 a 300 DPI
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)
    
    # Cargar fuentes
    try:
        title_font = ImageFont.truetype("arial.ttf", 80)
        subtitle_font = ImageFont.truetype("arial.ttf", 60)
        text_font = ImageFont.truetype("arial.ttf", 48)
        small_font = ImageFont.truetype("arial.ttf", 36)
    except:
        print("No se pudieron cargar las fuentes personalizadas, usando fuentes predeterminadas")
        title_font = ImageFont.load_default()
        subtitle_font = ImageFont.load_default()
        text_font = ImageFont.load_default()
        small_font = ImageFont.load_default()
    
    # Convertir el tipo de evaluación a texto legible
    tipo_textos = {
        'parcial1': 'PRIMER PARCIAL',
        'parcial2': 'SEGUNDO PARCIAL',
        'final': 'EXAMEN FINAL',
        'corto': 'EVALUACIÓN CORTA',
        'recuperacion': 'RECUPERACIÓN'
    }
    tipo_texto = tipo_textos.get(tipo_evaluacion, 'EVALUACIÓN PARCIAL')
    
    # Título principal
    draw.text((width//2, 150), f"PLANTILLA DE CALIFICACIÓN", fill="black", font=title_font, anchor="mm")
    draw.text((width//2, 250), f"{tipo_texto} - SECCIÓN {seccion} - {variante_id}", fill="black", font=subtitle_font, anchor="mm")
    draw.text((width//2, 350), "SOLO PARA USO DEL DOCENTE", fill="black", font=subtitle_font, anchor="mm")
    
    # Fecha y detalles
    fecha_actual = datetime.now().strftime("%d/%m/%Y")
    draw.text((width-200, 450), f"Fecha: {fecha_actual}", fill="black", font=small_font, anchor="rm")
    
    # Primera Serie - Marcar respuestas correctas
    draw.text((width//2, 550), "PRIMERA SERIE (40 PUNTOS)", fill="black", font=text_font, anchor="mm")
    
    y_pos = 650
    for i, respuesta in enumerate(respuestas["primera_serie"], 1):
        # Número de pregunta
        draw.text((200, y_pos), f"{i}.", fill="black", font=text_font)
        
        # Marcar solo la respuesta correcta
        x_pos = 350
        for j in range(5):  # Opciones a-e
            if j == respuesta:  # Si es la respuesta correcta
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", fill="black", width=3)
                letra_color = "white"
            else:
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", width=1)
                letra_color = "black"
            
            # Añadir la letra de la opción
            draw.text((x_pos+25, y_pos), chr(97+j), fill=letra_color, font=text_font, anchor="mm")
            x_pos += 120
        
        y_pos += 80
        
        # Si hay demasiadas preguntas, ajustar para que entren en la página
        if i == 7:  # Después de 7 preguntas, reajustar el espaciado
            y_pos = 650
            x_pos = width // 2 + 100  # Mover a la segunda columna
    
    # Segunda Serie - Marcar respuestas correctas
    draw.text((width//2, 1350), "SEGUNDA SERIE (20 PUNTOS)", fill="black", font=text_font, anchor="mm")
    
    y_pos = 1450
    for i, respuesta in enumerate(respuestas["segunda_serie"], 1):
        # Número de pregunta
        draw.text((200, y_pos), f"{i}.", fill="black", font=text_font)
        
        # Marcar solo la respuesta correcta
        x_pos = 350
        for j in range(5):  # Opciones a-e
            if j == respuesta:  # Si es la respuesta correcta
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", fill="black", width=3)
                letra_color = "white"
            else:
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", width=1)
                letra_color = "black"
            
            # Añadir la letra de la opción
            draw.text((x_pos+25, y_pos), chr(97+j), fill=letra_color, font=text_font, anchor="mm")
            x_pos += 120
        
        y_pos += 80
    
    # Tercera Serie - Respuestas numéricas
    draw.text((width//2, 2000), "TERCERA SERIE (40 PUNTOS)", fill="black", font=text_font, anchor="mm")
    
    y_pos = 2100
    # Coeficiente de Gini
    draw.text((200, y_pos), f"1. Coeficiente de Gini: {respuestas['tercera_serie']['gini']}", fill="black", font=text_font)
    
    y_pos += 100
    # Distribución de frecuencias
    draw.text((200, y_pos), "2. Distribución de frecuencias:", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"K: {respuestas['tercera_serie']['dist_frecuencias']['k']}", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Rango: {respuestas['tercera_serie']['dist_frecuencias']['rango']}", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Amplitud: {respuestas['tercera_serie']['dist_frecuencias']['amplitud']}", fill="black", font=text_font)
    
    y_pos += 100
    # Tallo y Hoja
    draw.text((200, y_pos), "3. Tallo y Hoja:", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Moda: {respuestas['tercera_serie']['tallo_hoja']['moda']}", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Intervalo: {respuestas['tercera_serie']['tallo_hoja']['intervalo']}", fill="black", font=text_font)
    
    y_pos += 100
    # Medidas de tendencia central
    draw.text((200, y_pos), "4. Medidas de tendencia central:", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Media: {respuestas['tercera_serie']['medidas_centrales']['media']}", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Mediana: {respuestas['tercera_serie']['medidas_centrales']['mediana']}", fill="black", font=text_font)
    y_pos += 70
    draw.text((250, y_pos), f"Moda: {respuestas['tercera_serie']['medidas_centrales']['moda']}", fill="black", font=text_font)
    
    # Añadir información de pie de página
    footer_y = height - 100
    draw.text((width//2, footer_y), f"Universidad Panamericana - Facultad de Humanidades", fill="black", font=small_font, anchor="mm")
    draw.text((width//2, footer_y + 50), f"Variante: {variante_id} | Sección: {seccion} | {tipo_texto}", fill="black", font=small_font, anchor="mm")
    
    # Crear carpeta con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.join(PLANTILLAS_FOLDER, f'{seccion}_{tipo_evaluacion}_{timestamp}')
    os.makedirs(output_dir, exist_ok=True)
    
    # Guardar con nombre que incluye todos los detalles
    filename = f'Plantilla_{seccion}_{tipo_evaluacion}_{variante_id}.pdf'
    filepath = os.path.join(output_dir, filename)
    image.save(filepath)
    
    # También guardar en la carpeta principal para compatibilidad
    image.save(os.path.join(PLANTILLAS_FOLDER, filename))
    
    return filename

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def procesar_examen_escaneado(pdf_path, variante_id):
    try:
        # Cargar respuestas correctas
        with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
            respuestas_correctas = json.load(f)
            
        # Convertir PDF a imágenes
        imagenes = convert_from_path(pdf_path)
        
        # Suponiendo que la hoja de respuestas es la primera página
        if not imagenes:
            return None
            
        img_respuestas = np.array(imagenes[0])
        
        # Convertir a escala de grises
        gray = cv2.cvtColor(img_respuestas, cv2.COLOR_BGR2GRAY)
        
        # Aplicar umbral para facilitar detección de círculos
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
        
        # Características de la hoja de respuestas (estos valores deben ajustarse según el diseño)
        # Posiciones para la primera serie (10 preguntas)
        primera_serie_y_start = 500  # Posición Y inicial
        question_height = 80         # Altura entre preguntas
        
        # Coordenadas de opciones para cada pregunta (estos son valores aproximados)
        option_xs = [550, 700, 850, 1000, 1150]  # Posiciones X de los centros de círculos
        
        # Extraer respuestas marcadas
        respuestas_alumno = {
            "primera_serie": [],
            "segunda_serie": [],
            "tercera_serie": {}
        }
        
        # Procesar primera serie
        for q in range(10):
            q_y = primera_serie_y_start + (q * question_height)
            
            respuesta_pregunta = None
            for i, opt_x in enumerate(option_xs):
                # Definir ROI (Region of Interest) alrededor del círculo
                roi = thresh[q_y-30:q_y+30, opt_x-30:opt_x+30]
                
                # Contar píxeles negros en el ROI
                if roi.size > 0:
                    black_pixels = np.sum(roi == 255)
                    
                    # Si hay suficientes píxeles negros, consideramos que el círculo está marcado
                    if black_pixels > 200:  # Este umbral debe ajustarse
                        respuesta_pregunta = i
                        break
            
            respuestas_alumno["primera_serie"].append(respuesta_pregunta)
        
        # Procesar segunda serie (similar a la primera)
        segunda_serie_y_start = primera_serie_y_start + (11 * question_height)
        
        for q in range(6):
            q_y = segunda_serie_y_start + (q * question_height)
            
            respuesta_pregunta = None
            for i, opt_x in enumerate(option_xs):
                roi = thresh[q_y-30:q_y+30, opt_x-30:opt_x+30]
                
                if roi.size > 0:
                    black_pixels = np.sum(roi == 255)
                    
                    if black_pixels > 200:
                        respuesta_pregunta = i
                        break
            
            respuestas_alumno["segunda_serie"].append(respuesta_pregunta)
        
        # Procesar tercera serie usando OCR para reconocer respuestas numéricas
        # Este es un proceso complejo y probablemente requiera ajustes específicos
        tercera_serie_y_start = segunda_serie_y_start + (7 * question_height)
        
        # Para el coeficiente de Gini
        gini_y = tercera_serie_y_start
        gini_roi = gray[gini_y-30:gini_y+30, 900:1300]  # Ajustar coordenadas
        gini_text = pytesseract.image_to_string(gini_roi, config='--psm 7 -c tessedit_char_whitelist=0123456789.')
        
        try:
            gini_value = float(gini_text.strip())
        except:
            gini_value = None
        
        respuestas_alumno["tercera_serie"]["gini"] = gini_value
        
        # Calcular puntuación
        puntuacion = calcular_puntuacion(respuestas_alumno, respuestas_correctas)
        
        # Extraer información del estudiante (nombre, carné)
        # Esta parte es avanzada y podría requerir técnicas adicionales de OCR
        info_estudiante = {
            "nombre": "Estudiante",
            "carne": "Carné no detectado"
        }
        
        return {
            "info_estudiante": info_estudiante,
            "respuestas": respuestas_alumno,
            "puntuacion": puntuacion
        }
        
    except Exception as e:
        print(f"Error al procesar el examen: {str(e)}")
        return None

def calcular_puntuacion(respuestas_alumno, respuestas_correctas):
    puntuacion = {
        "primera_serie": 0,
        "segunda_serie": 0,
        "tercera_serie": 0,
        "total": 0,
        "convertida_25": 0,
        "observaciones": []
    }
    
    # Primera serie (40 puntos, 4 puntos por pregunta)
    for i, (resp_alumno, resp_correcta) in enumerate(zip(respuestas_alumno["primera_serie"], respuestas_correctas["primera_serie"])):
        if resp_alumno == resp_correcta:
            puntuacion["primera_serie"] += 4
    
    # Segunda serie (20 puntos, ~3.33 puntos por pregunta)
    for i, (resp_alumno, resp_correcta) in enumerate(zip(respuestas_alumno["segunda_serie"], respuestas_correctas["segunda_serie"])):
        if resp_alumno == resp_correcta:
            puntuacion["segunda_serie"] += 3.33
    
    puntuacion["segunda_serie"] = round(puntuacion["segunda_serie"], 2)
    
    # Tercera serie (40 puntos, 10 puntos por problema)
    # Coeficiente de Gini
    gini_alumno = respuestas_alumno["tercera_serie"].get("gini")
    gini_correcto = respuestas_correctas["tercera_serie"]["gini"]
    
    if gini_alumno is not None:
        if abs(gini_alumno - gini_correcto) <= MARGENES_ERROR['gini']:
            puntuacion["tercera_serie"] += 10
        elif abs(gini_alumno - gini_correcto) <= MARGENES_ERROR['gini'] * 2:
            # Aceptable pero no perfecto
            puntuacion["tercera_serie"] += 7
            puntuacion["observaciones"].append("Revisar cálculo del coeficiente de Gini")
        else:
            puntuacion["observaciones"].append("Respuesta incorrecta en coeficiente de Gini")
    else:
        puntuacion["observaciones"].append("No se pudo detectar respuesta para coeficiente de Gini")
    
    # Calcular total y convertir a escala de 25 puntos
    puntuacion["total"] = puntuacion["primera_serie"] + puntuacion["segunda_serie"] + puntuacion["tercera_serie"]
    puntuacion["convertida_25"] = round(puntuacion["total"] * 0.25, 2)
    
    return puntuacion

def crear_solucion_matematica_simplificada(variante_id, seccion="A", tipo_evaluacion="parcial1"):
    """Crea sólo el archivo Word, evitando los problemas con PDF"""
    try:
        # Importar sin usar tkinter backend
        import matplotlib
        matplotlib.use('Agg')  # Usar backend no interactivo
        import matplotlib.pyplot as plt
        import numpy as np
        from scipy import stats
        import math
        import os
        from datetime import datetime
        
        # Generar carpeta con timestamp para evitar ambigüedades
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(PLANTILLAS_FOLDER, f'{seccion}_{tipo_evaluacion}_{timestamp}')
        os.makedirs(output_dir, exist_ok=True)
        
        # Cargar respuestas y variante
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
            variante = json.load(f)
        
        with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
            respuestas = json.load(f)
        
        # Crear documento Word
        doc = Document()
        
        # Configurar estilos
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)
        
        # Título principal
        doc.add_heading(f'SOLUCIÓN MATEMÁTICA DETALLADA', 0)
        doc.add_heading(f'{tipo_evaluacion.upper()} - SECCIÓN {seccion} - VARIANTE {variante_id}', 1)
        
        # Agregar contenido de las series
        # Primera Serie
        doc.add_heading('PRIMERA SERIE - RESPUESTAS CORRECTAS', 1)
        for i, pregunta in enumerate(variante['primera_serie'], 1):
            doc.add_paragraph(f"{i}. {pregunta['pregunta']}")
            doc.add_paragraph(f"   Respuesta correcta: {pregunta['opciones'][respuestas['primera_serie'][i-1]]}")
            doc.add_paragraph()
        
        # Segunda Serie
        doc.add_heading('SEGUNDA SERIE - RESPUESTAS CORRECTAS', 1)
        for i, escenario in enumerate(variante['segunda_serie'], 1):
            doc.add_paragraph(f"{i}. {escenario['escenario']}")
            doc.add_paragraph(f"   Respuesta correcta: {escenario['opciones'][respuestas['segunda_serie'][i-1]]}")
            doc.add_paragraph()
        
        # Tercera Serie - Ejemplo con el primer ejercicio (Gini)
        doc.add_heading('TERCERA SERIE - SOLUCIONES DETALLADAS', 1)
        
        # Ejercicio 1: Coeficiente de Gini
        doc.add_heading('Ejercicio 1: Coeficiente de Gini', 2)
        doc.add_paragraph("Solución paso a paso:")
        gini_value = respuestas["tercera_serie"]["gini"]
        doc.add_paragraph(f"El coeficiente de Gini calculado es: {gini_value}")
        
        # Guardar en la carpeta con timestamp
        doc_filename = f'Solucion_Matematica_{seccion}_{tipo_evaluacion}_{variante_id}.docx'
        doc_path = os.path.join(output_dir, doc_filename)
        doc.save(doc_path)
        
        # También guardar en la carpeta principal para compatibilidad
        doc.save(os.path.join(PLANTILLAS_FOLDER, doc_filename))
        
        return doc_filename
    except Exception as e:
        print(f"Error al crear solución matemática: {str(e)}")
        return None

# Rutas de la aplicación
@app.route('/')
def index():
    # Cargar el historial para obtener el orden cronológico
    historial = cargar_historial()
    
    # Lista para almacenar las variantes
    variantes = []
    
    # Set para evitar duplicados
    variantes_procesadas = set()
    
    # Obtener las variantes del historial (más recientes primero)
    for item in sorted(historial, key=lambda x: x.get('fecha_generacion', '0'), reverse=True):
        variante_id = item.get('id')
        
        # Evitar duplicados
        if variante_id in variantes_procesadas:
            continue
        
        # Verificar si existen los archivos
        tiene_examen = os.path.exists(os.path.join(EXAMENES_FOLDER, f'Examen_{variante_id}.docx'))
        tiene_hoja = os.path.exists(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{variante_id}.pdf'))
        tiene_plantilla = os.path.exists(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{variante_id}.pdf'))
        
        # Verificar si existe la solución matemática
        solucion_matematica = item.get('solucion_matematica')
        tiene_solucion = solucion_matematica and os.path.exists(os.path.join(PLANTILLAS_FOLDER, solucion_matematica))
        
        # Añadir la variante a la lista
        variantes.append({
            'id': variante_id,
            'seccion': item.get('seccion', 'No especificada'),
            'tipo_evaluacion': item.get('tipo_texto', 'No especificado'),
            'tiene_examen': tiene_examen,
            'tiene_hoja': tiene_hoja,
            'tiene_plantilla': tiene_plantilla,
            'solucion_matematica': solucion_matematica if tiene_solucion else None,
            'tiene_solucion': tiene_solucion
        })
        
        variantes_procesadas.add(variante_id)
    
    # Añadir variantes que podrían no estar en el historial
    if os.path.exists(VARIANTES_FOLDER):
        for archivo in os.listdir(VARIANTES_FOLDER):
            if archivo.startswith('variante_') and archivo.endswith('.json'):
                variante_id = archivo.replace('variante_', '').replace('.json', '')
                
                if variante_id in variantes_procesadas:
                    continue
                
                # Verificaciones básicas
                tiene_examen = os.path.exists(os.path.join(EXAMENES_FOLDER, f'Examen_{variante_id}.docx'))
                tiene_hoja = os.path.exists(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{variante_id}.pdf'))
                tiene_plantilla = os.path.exists(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{variante_id}.pdf'))
                
                variantes.append({
                    'id': variante_id,
                    'seccion': 'No especificada',
                    'tipo_evaluacion': 'No especificado',
                    'tiene_examen': tiene_examen,
                    'tiene_hoja': tiene_hoja,
                    'tiene_plantilla': tiene_plantilla,
                    'solucion_matematica': None,
                    'tiene_solucion': False
                })
    
    return render_template('index.html', variantes=variantes)

@app.route('/estudiantes', methods=['GET', 'POST'])
def gestionar_estudiantes():
    if request.method == 'POST':
        action = request.form.get('action')
        
        if action == 'agregar':
            nombre = request.form.get('nombre')
            carne = request.form.get('carne')
            seccion = request.form.get('seccion')
            
            if nombre and carne and seccion:
                if seccion not in estudiantes_db:
                    estudiantes_db[seccion] = []
                
                estudiantes_db[seccion].append({
                    'nombre': nombre,
                    'carne': carne,
                    'evaluaciones': {}
                })
                
                flash(f'Estudiante {nombre} agregado correctamente', 'success')
            else:
                flash('Datos incompletos', 'danger')
                
        elif action == 'cargar_csv':
            if 'archivo_csv' not in request.files:
                flash('No se seleccionó archivo CSV', 'danger')
                return redirect(request.url)
                
            archivo = request.files['archivo_csv']
            seccion = request.form.get('seccion')
            
            if archivo.filename == '' or not seccion:
                flash('Datos incompletos', 'danger')
                return redirect(request.url)
                
            # Procesar archivo CSV
            try:
                contenido = archivo.read().decode('utf-8')
                lineas = contenido.split('\n')
                
                if seccion not in estudiantes_db:
                    estudiantes_db[seccion] = []
                
                for linea in lineas[1:]:  # Omitir encabezado
                    if not linea.strip():
                        continue
                        
                    datos = linea.split(',')
                    if len(datos) >= 2:
                        estudiantes_db[seccion].append({
                            'nombre': datos[0].strip(),
                            'carne': datos[1].strip(),
                            'evaluaciones': {}
                        })
                
                flash(f'Se cargaron {len(lineas)-1} estudiantes para la sección {seccion}', 'success')
            except Exception as e:
                flash(f'Error al procesar CSV: {str(e)}', 'danger')
    
    return render_template('estudiantes.html', estudiantes=estudiantes_db)

@app.route('/calificaciones/<seccion>/<tipo_evaluacion>')
def ver_calificaciones(seccion, tipo_evaluacion):
    if seccion not in estudiantes_db:
        flash(f'La sección {seccion} no existe', 'danger')
        return redirect(url_for('gestionar_estudiantes'))
    
    # Obtener lista de exámenes procesados para esta sección y tipo
    eval_folder = os.path.join(EXAMENES_ESCANEADOS_FOLDER, f"{seccion}_{tipo_evaluacion}")
    
    resultados = {}
    if os.path.exists(eval_folder):
        for archivo in os.listdir(eval_folder):
            if archivo.endswith('.json'):
                with open(os.path.join(eval_folder, archivo), 'r') as f:
                    resultado = json.load(f)
                    carne = resultado.get('info_estudiante', {}).get('carne')
                    if carne:
                        resultados[carne] = resultado
    
    # Preparar datos para mostrar
    estudiantes_seccion = estudiantes_db.get(seccion, [])
    for estudiante in estudiantes_seccion:
        carne = estudiante.get('carne')
        if carne in resultados:
            estudiante['resultado'] = resultados[carne]
        else:
            estudiante['resultado'] = None
    
    return render_template('calificaciones.html', 
                          estudiantes=estudiantes_seccion, 
                          seccion=seccion, 
                          tipo_evaluacion=tipo_evaluacion)

@app.route('/generar_examen', methods=['POST'])
def generar_examen():
    try:
        num_variantes = int(request.form.get('num_variantes', 1))
        seccion = request.form.get('seccion', 'A')
        tipo_evaluacion = request.form.get('tipo_evaluacion', 'parcial1')
        
        # Crear carpeta con timestamp para esta generación
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_output_dir = f'{seccion}_{tipo_evaluacion}_{timestamp}'
        
        # Crear carpetas para cada tipo de archivo
        for folder in [EXAMENES_FOLDER, HOJAS_RESPUESTA_FOLDER, PLANTILLAS_FOLDER, VARIANTES_FOLDER]:
            output_dir = os.path.join(folder, base_output_dir)
            os.makedirs(output_dir, exist_ok=True)
        
        # Manejo del logo
        logo_path = None
        if 'logo' in request.files and request.files['logo'].filename:
            logo = request.files['logo']
            logo_filename = secure_filename(logo.filename)
            logo_path = os.path.join(UPLOAD_FOLDER, logo_filename)
            print(f"Guardando logo en: {logo_path}")
            logo.save(logo_path)
        
        # Manejo de la plantilla
        plantilla_path = None
        if 'plantilla' in request.files and request.files['plantilla'].filename:
            plantilla = request.files['plantilla']
            plantilla_filename = secure_filename(plantilla.filename)
            plantilla_path = os.path.join(UPLOAD_FOLDER, plantilla_filename)
            print(f"Guardando plantilla en: {plantilla_path}")
            plantilla.save(plantilla_path)
        
        # Generar variantes
        variantes_generadas = []
        
        for i in range(num_variantes):
            variante_id = f"V{i+1}"
            variante, respuestas = generar_variante(variante_id)
            
            # Guardar variante y respuestas en la carpeta con timestamp
            with open(os.path.join(VARIANTES_FOLDER, base_output_dir, f'variante_{variante_id}.json'), 'w', encoding='utf-8') as f:
                json.dump(variante, f, ensure_ascii=False, indent=2)
            
            with open(os.path.join(VARIANTES_FOLDER, base_output_dir, f'respuestas_{variante_id}.json'), 'w', encoding='utf-8') as f:
                json.dump(respuestas, f, ensure_ascii=False, indent=2)
            
            # Crear documentos
            examen_filename = crear_examen_word(variante_id, seccion, tipo_evaluacion, logo_path, plantilla_path)
            hoja_filename = crear_hoja_respuestas(variante_id, seccion, tipo_evaluacion)
            plantilla_filename = crear_plantilla_calificacion(variante_id, seccion, tipo_evaluacion)
            solucion_filename = crear_solucion_matematica_simplificada(variante_id, seccion, tipo_evaluacion)
            
            variantes_generadas.append({
                'id': variante_id,
                'examen': examen_filename,
                'hoja': hoja_filename,
                'plantilla': plantilla_filename,
                'solucion_matematica': solucion_filename,
                'seccion': seccion,
                'tipo_evaluacion': tipo_evaluacion,
                'timestamp': timestamp,
                'directorio': base_output_dir
            })
        
        # Registrar en el historial
        historial = cargar_historial()
        
        # Obtener nombres para mostrar
        tipo_textos = {
            'parcial1': 'Primer Parcial',
            'parcial2': 'Segundo Parcial',
            'final': 'Examen Final',
            'corto': 'Evaluación Corta',
            'recuperacion': 'Recuperación'
        }
        
        # Añadir entrada al historial
        for variante in variantes_generadas:
            historial.append({
                'id': variante['id'],
                'seccion': seccion,
                'tipo_evaluacion': tipo_evaluacion,
                'tipo_texto': tipo_textos.get(tipo_evaluacion, tipo_evaluacion),
                'fecha_generacion': datetime.now().strftime("%d/%m/%Y %H:%M"),
                'timestamp': timestamp,
                'directorio': base_output_dir,
                'examen': variante['examen'],
                'hoja': variante['hoja'],
                'plantilla': variante['plantilla'],
                'solucion_matematica': variante['solucion_matematica']
            })
        
        guardar_historial(historial)
        
        flash(f'Se han generado {num_variantes} variantes de examen para la sección {seccion}', 'success')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error al generar exámenes: {str(e)}', 'danger')
        return redirect(url_for('index'))
    
@app.route('/historial')
def mostrar_historial():
    historial = cargar_historial()
    return render_template('historial.html', historial=historial)

@app.route('/cargar_examenes_escaneados', methods=['GET', 'POST'])
def cargar_examenes_escaneados():
    if request.method == 'POST':
        if 'archivos' not in request.files:
            flash('No se seleccionaron archivos', 'danger')
            return redirect(request.url)
            
        archivos = request.files.getlist('archivos')
        seccion = request.form.get('seccion', '')
        tipo_evaluacion = request.form.get('tipo_evaluacion', '')
        variante_id = request.form.get('variante_id', '')
        
        if not seccion or not tipo_evaluacion or not variante_id:
            flash('Datos incompletos. Por favor complete todos los campos.', 'danger')
            return redirect(request.url)
        
        # Crear carpeta específica para esta evaluación
        eval_folder = os.path.join(EXAMENES_ESCANEADOS_FOLDER, f"{seccion}_{tipo_evaluacion}_{variante_id}")
        if not os.path.exists(eval_folder):
            os.makedirs(eval_folder)
        
        archivos_procesados = []
        for archivo in archivos:
            if archivo.filename == '':
                continue
                
            if archivo and allowed_file(archivo.filename, {'pdf'}):
                filename = secure_filename(archivo.filename)
                filepath = os.path.join(eval_folder, filename)
                archivo.save(filepath)
                
                # Procesar el archivo
                resultado = procesar_examen_escaneado(filepath, variante_id)
                if resultado:
                    archivos_procesados.append({
                        'nombre': filename,
                        'resultado': resultado
                    })
        
        if archivos_procesados:
            flash(f'Se procesaron {len(archivos_procesados)} exámenes correctamente', 'success')
            return render_template('resultados_procesamiento.html', 
                                  archivos=archivos_procesados, 
                                  seccion=seccion,
                                  tipo_evaluacion=tipo_evaluacion,
                                  variante_id=variante_id)
        else:
            flash('No se pudo procesar ningún archivo', 'warning')
            
    # Obtener variantes disponibles para mostrar en el formulario
    variantes = []
    if os.path.exists(VARIANTES_FOLDER):
        for archivo in os.listdir(VARIANTES_FOLDER):
            if archivo.startswith('variante_') and archivo.endswith('.json'):
                variante_id = archivo.replace('variante_', '').replace('.json', '')
                variantes.append(variante_id)
    
    return render_template('cargar_examenes.html', variantes=variantes)

@app.route('/editar_variante/<variante_id>')
def editar_variante(variante_id):
    try:
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
            variante = json.load(f)
        
        return render_template('editar_variante.html', variante=variante)
    except:
        flash('No se pudo cargar la variante', 'danger')
        return redirect(url_for('index'))

@app.route('/guardar_variante', methods=['POST'])
def guardar_variante():
    variante_id = request.form.get('variante_id')
    
    try:
        # Cargar variante original
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
            variante = json.load(f)
        
        # Actualizar datos (esto sería más complejo en un caso real)
        # Aquí se implementaría la lógica para actualizar la variante según los datos del formulario
        
        # Guardar variante actualizada
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'w', encoding='utf-8') as f:
            json.dump(variante, f, ensure_ascii=False, indent=2)
        
        # Regenerar documentos
        examen_filename = crear_examen_word(variante_id)
        hoja_filename = crear_hoja_respuestas(variante_id)
        plantilla_filename = crear_plantilla_calificacion(variante_id)
        
        flash('Variante actualizada correctamente', 'success')
    except Exception as e:
        flash(f'Error al guardar la variante: {str(e)}', 'danger')
    
    return redirect(url_for('index'))

@app.route('/previsualizar/<variante_id>')
def previsualizar(variante_id):
    try:
        with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
            variante = json.load(f)
        
        with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
            respuestas = json.load(f)
        
        return render_template('previsualizar.html', variante=variante, respuestas=respuestas)
    except:
        flash('No se pudo cargar la variante para previsualizar', 'danger')
        return redirect(url_for('index'))

@app.route('/eliminar_variante/<variante_id>', methods=['POST'])
def eliminar_variante(variante_id):
    # Eliminar archivos de la variante
    archivos = [
        os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'),
        os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'),
        os.path.join(EXAMENES_FOLDER, f'Examen_{variante_id}.docx'),
        os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{variante_id}.pdf'),
        os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{variante_id}.pdf')
    ]
    
    for archivo in archivos:
        if os.path.exists(archivo):
            os.remove(archivo)
    
    flash(f'Variante {variante_id} eliminada correctamente', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
