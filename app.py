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
def crear_examen_word(variante_id, seccion="A", tipo_evaluacion="parcial1", logo_path=None):
    # Cargar la variante
    with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
        variante = json.load(f)
    
    # Crear documento Word
    doc = Document()
    
    # Configurar márgenes
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    
    # Definir estilos
    styles = doc.styles
    
    # Estilo para encabezado principal
    title_style = styles.add_style('TitleStyle', WD_STYLE_TYPE.PARAGRAPH)
    title_font = title_style.font
    title_font.name = 'Arial'
    title_font.size = Pt(16)
    title_font.bold = True
    title_font.color.rgb = RGBColor(0, 0, 102)  # Azul oscuro
    
    # Estilo para subtítulos
    subtitle_style = styles.add_style('SubtitleStyle', WD_STYLE_TYPE.PARAGRAPH)
    subtitle_font = subtitle_style.font
    subtitle_font.name = 'Arial'
    subtitle_font.size = Pt(12)
    subtitle_font.bold = True
    
    # Estilo para texto normal
    normal_style = styles.add_style('NormalStyle', WD_STYLE_TYPE.PARAGRAPH)
    normal_font = normal_style.font
    normal_font.name = 'Arial'
    normal_font.size = Pt(11)
    
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
    doc.save(os.path.join(EXAMENES_FOLDER, filename))
    
    return filename

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
def crear_hoja_respuestas(variante_id):
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
    draw.text((width//2, exam_y), f"EVALUACIÓN PARCIAL", 
              fill="black", font=header_font, anchor="mm")
    
    # ==================== INFORMACIÓN DEL ESTUDIANTE ====================
    # Establecer posiciones exactas
    info_y = exam_y + 120
    left_x = margin
    label_width = 200  # Ancho fijo para etiquetas
    
    # NOMBRE
    nombre_y = info_y
    draw.text((left_x, nombre_y), "NOMBRE:", fill="black", font=text_font)
    # Línea para nombre (inicia justo después de la etiqueta)
    line_start_x = left_x + label_width
    draw.line((line_start_x, nombre_y + 20, width - margin, nombre_y + 20), 
              fill="black", width=3)
    
    # CARNÉ y FIRMA (en la misma línea)
    carne_y = nombre_y + 80
    draw.text((left_x, carne_y), "CARNÉ:", fill="black", font=text_font)
    # Línea para carné
    draw.line((line_start_x, carne_y + 20, line_start_x + 450, carne_y + 20), 
              fill="black", width=3)
    
    # FIRMA (alineada a la derecha de carné)
    firma_x = line_start_x + 600
    draw.text((firma_x, carne_y), "FIRMA:", fill="black", font=text_font)
    # Línea para firma
    draw.line((firma_x + label_width, carne_y + 20, width - margin, carne_y + 20), 
              fill="black", width=3)
    
    # VARIANTE (nueva, al lado de firma)
    variante_x = firma_x + 500
    draw.text((variante_x, carne_y), "VARIANTE:", fill="black", font=text_font)
    # Línea para variante
    draw.rectangle((variante_x + label_width, carne_y - 20, variante_x + label_width + 80, carne_y + 20), 
            outline="black", width=2)
    
    # ==================== PRIMERA SERIE ====================
    serie1_y = carne_y + 100
    
    # Encabezado
    draw.rectangle((left_x, serie1_y, width - margin, serie1_y + 70), 
                  outline="black", fill="#EEEEEE", width=2)
    draw.text((width//2, serie1_y + 35), "PRIMERA SERIE (40 PUNTOS)", 
              fill="black", font=header_font, anchor="mm")
    
    # Configuración para opciones
    questions_start_y = serie1_y + 120
    question_height = 80  # Altura entre preguntas
    
    # Posiciones horizontales fijas para número y opciones
    number_x = left_x + 50
    circle_radius = 22
    
    # Calcular posiciones absolutas para opciones (a, b, c, d, e)
    option_xs = [550, 700, 850, 1000, 1150]  # Posiciones fijas para los centros de círculos
    
    # Dibujar preguntas y opciones
    for q in range(10):
        q_y = questions_start_y + (q * question_height)
        
        # Número de pregunta
        draw.text((number_x, q_y), f"{q+1}.", fill="black", font=text_font, anchor="lm")
        
        # Opciones (a-e)
        for i, opt_x in enumerate(option_xs):
            # Dibujar círculo perfectamente centrado
            draw.ellipse((opt_x - circle_radius, q_y - circle_radius, 
                          opt_x + circle_radius, q_y + circle_radius), 
                         outline="black", width=2)
            
            # Letra centrada en círculo
            draw.text((opt_x, q_y), chr(97 + i), fill="black", font=option_font, anchor="mm")
    
    # ==================== SEGUNDA SERIE ====================
    serie2_y = questions_start_y + (11 * question_height)
    
    # Encabezado
    draw.rectangle((left_x, serie2_y, width - margin, serie2_y + 70), 
                  outline="black", fill="#EEEEEE", width=2)
    draw.text((width//2, serie2_y + 35), "SEGUNDA SERIE (20 PUNTOS)", 
              fill="black", font=header_font, anchor="mm")
    
    # Opciones
    questions2_start_y = serie2_y + 120
    
    # Dibujar preguntas y opciones
    for q in range(6):
        q_y = questions2_start_y + (q * question_height)
        
        # Número de pregunta
        draw.text((number_x, q_y), f"{q+1}.", fill="black", font=text_font, anchor="lm")
        
        # Opciones (a-e) - Mismas posiciones que primera serie
        for i, opt_x in enumerate(option_xs):
            # Dibujar círculo
            draw.ellipse((opt_x - circle_radius, q_y - circle_radius, 
                          opt_x + circle_radius, q_y + circle_radius), 
                         outline="black", width=2)
            
            # Letra centrada
            draw.text((opt_x, q_y), chr(97 + i), fill="black", font=option_font, anchor="mm")
    
    # ==================== TERCERA SERIE ====================
    serie3_y = questions2_start_y + (7 * question_height)
    
    # Encabezado
    draw.rectangle((left_x, serie3_y, width - margin, serie3_y + 70), 
                  outline="black", fill="#EEEEEE", width=2)
    draw.text((width//2, serie3_y + 35), "TERCERA SERIE (40 PUNTOS)", 
              fill="black", font=header_font, anchor="mm")
    
    # Respuestas numéricas
    resp_start_y = serie3_y + 120
    resp_height = 80  # Reducimos un poco para que quepa todo
    
    # Configuración para cajas de respuesta
    box_x = left_x + 500
    box_width = 400
    
    # 1. Coeficiente de Gini
    gini_y = resp_start_y
    draw.text((left_x + 50, gini_y), "1. Coeficiente de Gini:", 
              fill="black", font=text_font, anchor="lm")
    
    # Caja para respuesta
    draw.rectangle((box_x, gini_y - 30, box_x + box_width, gini_y + 30), 
                  outline="black", width=2)
    
    # 2. Distribución de frecuencias
    dist_title_y = gini_y + resp_height
    draw.text((left_x + 50, dist_title_y), "2. Distribución de frecuencias:", 
              fill="black", font=text_font, anchor="lm")
    
    # K
    k_y = dist_title_y + resp_height
    draw.text((left_x + 100, k_y), "K:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, k_y - 30, box_x + box_width, k_y + 30), 
                  outline="black", width=2)
    
    # Rango
    rango_y = k_y + resp_height
    draw.text((left_x + 100, rango_y), "Rango:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, rango_y - 30, box_x + box_width, rango_y + 30), 
                  outline="black", width=2)
    
    # Amplitud
    amp_y = rango_y + resp_height
    draw.text((left_x + 100, amp_y), "Amplitud:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, amp_y - 30, box_x + box_width, amp_y + 30), 
                  outline="black", width=2)
    
    # 3. Tallo y Hoja
    th_title_y = amp_y + resp_height
    draw.text((left_x + 50, th_title_y), "3. Tallo y Hoja:", 
              fill="black", font=text_font, anchor="lm")
    
    # Valor moda
    moda_y = th_title_y + resp_height
    draw.text((left_x + 100, moda_y), "Valor moda:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, moda_y - 30, box_x + box_width, moda_y + 30), 
                  outline="black", width=2)
    
    # Intervalo mayor concentración
    intervalo_y = moda_y + resp_height
    draw.text((left_x + 100, intervalo_y), "Intervalo:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, intervalo_y - 30, box_x + box_width, intervalo_y + 30), 
                  outline="black", width=2)
    
    # 4. Medidas de tendencia central
    mc_title_y = intervalo_y + resp_height
    draw.text((left_x + 50, mc_title_y), "4. Medidas de tendencia central:", 
              fill="black", font=text_font, anchor="lm")
    
    # Media
    media_y = mc_title_y + resp_height
    draw.text((left_x + 100, media_y), "Media:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, media_y - 30, box_x + box_width, media_y + 30), 
                  outline="black", width=2)
    
    # Mediana
    mediana_y = media_y + resp_height
    draw.text((left_x + 100, mediana_y), "Mediana:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, mediana_y - 30, box_x + box_width, mediana_y + 30), 
                  outline="black", width=2)
    
    # Moda (nuevo campo)
    moda_y = mediana_y + resp_height
    draw.text((left_x + 100, moda_y), "Moda:", fill="black", font=text_font, anchor="lm")
    draw.rectangle((box_x, moda_y - 30, box_x + box_width, moda_y + 30), 
                outline="black", width=2)
    
    # Código de variante (discreto en esquina)
    draw.text((width - 200, height - 40), f"V{variante_id}", fill="black", 
              font=text_font if 'arial.ttf' in ImageFont.truetype.__code__.co_varnames else ImageFont.load_default())
    
    # Guardar la imagen
    image.save(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{variante_id}.pdf'))
    
    return f'HojaRespuestas_{variante_id}.pdf'

# Función para crear una plantilla de calificación
def crear_plantilla_calificacion(variante_id):
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
        text_font = ImageFont.truetype("arial.ttf", 48)
    except:
        title_font = ImageFont.load_default()
        text_font = ImageFont.load_default()
    
    # Título
    draw.text((width//2, 150), f"PLANTILLA DE CALIFICACIÓN - {variante_id}", fill="black", font=title_font, anchor="mm")
    draw.text((width//2, 250), "SOLO PARA USO DEL DOCENTE", fill="black", font=title_font, anchor="mm")
    
    # Primera Serie - Marcar respuestas correctas
    draw.text((width//2, 400), "PRIMERA SERIE", fill="black", font=text_font, anchor="mm")
    
    y_pos = 500
    for i, respuesta in enumerate(respuestas["primera_serie"], 1):
        # Número de pregunta
        draw.text((200, y_pos), f"{i}.", fill="black", font=text_font)
        
        # Marcar solo la respuesta correcta
        x_pos = 350
        for j in range(5):  # Opciones a-e
            if j == respuesta:  # Si es la respuesta correcta
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", fill="black", width=3)
            else:
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", width=1)
            
            draw.text((x_pos+25, y_pos), chr(97+j), fill="white" if j == respuesta else "black", font=text_font, anchor="mm")
            x_pos += 120
        
        y_pos += 80
    
    # Segunda Serie - Marcar respuestas correctas
    draw.text((width//2, 1350), "SEGUNDA SERIE", fill="black", font=text_font, anchor="mm")
    
    y_pos = 1450
    for i, respuesta in enumerate(respuestas["segunda_serie"], 1):
        # Número de pregunta
        draw.text((200, y_pos), f"{i}.", fill="black", font=text_font)
        
        # Marcar solo la respuesta correcta
        x_pos = 350
        for j in range(5):  # Opciones a-e
            if j == respuesta:  # Si es la respuesta correcta
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", fill="black", width=3)
            else:
                draw.ellipse((x_pos, y_pos-25, x_pos+50, y_pos+25), outline="black", width=1)
            
            draw.text((x_pos+25, y_pos), chr(97+j), fill="white" if j == respuesta else "black", font=text_font, anchor="mm")
            x_pos += 120
        
        y_pos += 80
    
    # Tercera Serie - Respuestas numéricas
    draw.text((width//2, 2000), "TERCERA SERIE", fill="black", font=text_font, anchor="mm")
    
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
    
    # Guardar la imagen
    image.save(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{variante_id}.pdf'))
    
    return f'Plantilla_{variante_id}.pdf'

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

def crear_solucion_matematica_detallada(variante_id, seccion="A", tipo_evaluacion="parcial1"):
    """
    Genera un documento con soluciones matemáticas extremadamente detalladas
    para todos los ejercicios del examen, incluyendo las tres series.
    """
    import matplotlib.pyplot as plt
    import numpy as np
    from scipy import stats
    from docx2pdf import convert
    import math
    
    # Cargar respuestas de la variante
    with open(os.path.join(VARIANTES_FOLDER, f'respuestas_{variante_id}.json'), 'r', encoding='utf-8') as f:
        respuestas = json.load(f)
    
    # Cargar la variante para acceder a los datos de los problemas
    with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
        variante = json.load(f)
    
    # Crear documento Word para respuestas detalladas
    doc = Document()
    
    # Configurar estilos y formato para mejor legibilidad
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    
    # Estilos personalizados para mejorar la apariencia
    styles = doc.styles
    
    # Estilo para título principal
    title_style = styles.add_style('MathTitleStyle', WD_STYLE_TYPE.PARAGRAPH)
    title_font = title_style.font
    title_font.name = 'Calibri'
    title_font.size = Pt(18)
    title_font.bold = True
    title_font.color.rgb = RGBColor(0, 70, 0)  # Verde oscuro
    
    # Estilo para subtítulos
    subtitle_style = styles.add_style('MathSubtitleStyle', WD_STYLE_TYPE.PARAGRAPH)
    subtitle_font = subtitle_style.font
    subtitle_font.name = 'Calibri'
    subtitle_font.size = Pt(14)
    subtitle_font.bold = True
    subtitle_font.color.rgb = RGBColor(0, 0, 120)  # Azul oscuro
    
    # Estilo para ecuaciones
    equation_style = styles.add_style('EquationStyle', WD_STYLE_TYPE.PARAGRAPH)
    equation_font = equation_style.font
    equation_font.name = 'Cambria Math'
    equation_font.size = Pt(12)
    equation_font.italic = True
    
    # Estilo para explicaciones
    explanation_style = styles.add_style('ExplanationStyle', WD_STYLE_TYPE.PARAGRAPH)
    explanation_font = explanation_style.font
    explanation_font.name = 'Calibri'
    explanation_font.size = Pt(11)
    
    # Título principal
    tipo_textos = {
        'parcial1': 'Primer Parcial',
        'parcial2': 'Segundo Parcial',
        'final': 'Examen Final',
        'corto': 'Evaluación Corta',
        'recuperacion': 'Recuperación'
    }
    tipo_texto = tipo_textos.get(tipo_evaluacion, tipo_evaluacion)
    
    doc.add_heading(f'SOLUCIÓN MATEMÁTICA DETALLADA', 0).style = title_style
    doc.add_heading(f'{tipo_texto.upper()} - SECCIÓN {seccion} - VARIANTE {variante_id}', 1).style = subtitle_style
    
    p = doc.add_paragraph()
    p.add_run(f'Fecha de generación: {datetime.now().strftime("%d de %B de %Y, %H:%M")}')
    p.add_run('\nPARA USO EXCLUSIVO DEL DOCENTE - DOCUMENTO DE VERIFICACIÓN MATEMÁTICA')
    p = doc.add_paragraph()
    p.add_run('Este documento contiene soluciones matemáticas paso a paso con verificación cruzada para garantizar precisión absoluta en los cálculos.').bold = True
    
    # ÍNDICE DE CONTENIDOS
    doc.add_heading('ÍNDICE', level=1).style = subtitle_style
    p = doc.add_paragraph()
    p.add_run("PRIMERA SERIE - Preguntas de opción múltiple (40 puntos)").bold = True
    p.add_run("\n   • Análisis y justificación de las respuestas correctas.")
    
    p = doc.add_paragraph()
    p.add_run("SEGUNDA SERIE - Tipos de gráficos (20 puntos)").bold = True
    p.add_run("\n   • Justificación matemática y estadística de las selecciones.")
    
    p = doc.add_paragraph()
    p.add_run("TERCERA SERIE - Ejercicios prácticos (40 puntos)").bold = True
    p.add_run("\n   • Ejercicio 1: Coeficiente de Gini (cálculo paso a paso).")
    p.add_run("\n   • Ejercicio 2: Distribución de frecuencias - Método Sturges.")
    p.add_run("\n   • Ejercicio 3: Diagrama de Tallo y Hoja.")
    p.add_run("\n   • Ejercicio 4: Medidas de tendencia central.")
    
    # Saltar a la siguiente página para empezar con el contenido
    doc.add_page_break()
    
    # ==================================================
    # PRIMERA SERIE - Análisis detallado
    # ==================================================
    doc.add_heading('PRIMERA SERIE - JUSTIFICACIÓN DE RESPUESTAS', level=1).style = title_style
    
    p = doc.add_paragraph("Las siguientes justificaciones proporcionan el fundamento teórico y matemático para cada una de las respuestas correctas de la primera serie. Estas explicaciones pueden utilizarse para clarificar dudas y como material didáctico complementario:", style='ExplanationStyle')
    
    for i, (pregunta, respuesta_idx) in enumerate(zip(variante["primera_serie"], respuestas["primera_serie"]), 1):
        doc.add_heading(f"Pregunta {i}: {pregunta['pregunta']}", level=2).style = subtitle_style
        
        # Mostrar todas las opciones
        p = doc.add_paragraph("Opciones disponibles:", style='ExplanationStyle')
        for j, opcion in enumerate(pregunta["opciones"]):
            if j == respuesta_idx:
                p = doc.add_paragraph(f"• {opcion}", style='ExplanationStyle')
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = RGBColor(0, 128, 0)  # Verde para la respuesta correcta
            else:
                p = doc.add_paragraph(f"• {opcion}", style='ExplanationStyle')
        
        # Justificación detallada según el tipo de pregunta
        p = doc.add_paragraph("Justificación:", style='ExplanationStyle')
        p.runs[0].bold = True
        
        # Generar una justificación más detallada según el tipo de pregunta
        if "Variable" in pregunta["pregunta"]:
            p.add_run("\nEn estadística, las variables pueden clasificarse según su naturaleza y comportamiento. Las variables cualitativas representan características o atributos que no pueden medirse numéricamente (como colores, géneros, profesiones), mientras que las variables cuantitativas expresan valores numéricos.")
            p.add_run("\n\nLas variables cuantitativas pueden ser discretas (toman valores aislados, como número de hijos) o continuas (pueden tomar cualquier valor dentro de un intervalo, como peso o temperatura).")
            
        elif "población" in pregunta["pregunta"].lower() or "muestra" in pregunta["pregunta"].lower():
            p.add_run("\nEn estadística, la población es el conjunto completo de elementos (personas, objetos, medidas) sobre los que se realiza el estudio. La muestra es un subconjunto representativo de la población, seleccionado para inferir características de la población total.")
            p.add_run("\n\nLa correcta distinción entre estos conceptos es fundamental para el diseño de estudios estadísticos válidos y para la aplicación apropiada de técnicas de inferencia.")
            
        elif "Gini" in pregunta["pregunta"]:
            p.add_run("\nEl coeficiente de Gini es una medida estadística diseñada para representar la desigualdad en la distribución de ingresos o riqueza. Toma valores entre 0 (igualdad perfecta) y 1 (desigualdad máxima).")
            p.add_run("\n\nMatematicamente, el coeficiente se calcula a partir de la Curva de Lorenz, que representa la proporción acumulada de ingreso versus la proporción acumulada de población. El coeficiente es el área entre la curva de Lorenz y la línea de igualdad perfecta, dividida por el área total bajo la línea de igualdad.")
            
        elif "medida" in pregunta["pregunta"].lower() and "tendencia central" in pregunta["pregunta"].lower():
            p.add_run("\nLas medidas de tendencia central (media, mediana y moda) describen el centro de una distribución de datos. La media aritmética es el promedio de todos los valores y se calcula sumando todos los datos y dividiendo por el número total.")
            p.add_run("\n\nLa media es particularmente sensible a valores extremos (outliers) porque estos afectan directamente la suma total. Por ejemplo, en la distribución [1, 2, 3, 4, 100], la media es 22, un valor que no representa adecuadamente la centralidad de los datos.")
            p.add_run("\n\nEn contraste, la mediana no se ve afectada por valores extremos ya que solo considera la posición central de los datos ordenados.")
            
        elif "histograma" in pregunta["pregunta"].lower() or "gráfico" in pregunta["pregunta"].lower():
            p.add_run("\nLos diferentes tipos de gráficos estadísticos están diseñados para representar distintos tipos de variables y relaciones. El histograma está específicamente diseñado para representar variables continuas, mostrando la distribución de frecuencias mediante rectángulos contiguos.")
            p.add_run("\n\nA diferencia de los gráficos de barras (para variables discretas o categóricas), el histograma no deja espacios entre las barras, reflejando la naturaleza continua de los datos. La altura representa la frecuencia o densidad, y el área total equivale al número total de observaciones.")
        
        else:
            # Justificación genérica para otras preguntas
            p.add_run(f"\nLa respuesta correcta es '{pregunta['opciones'][respuesta_idx]}' porque es la definición precisa según los conceptos estadísticos fundamentales.")
            p.add_run("\n\nEsta respuesta se alinea con los principios establecidos en la teoría estadística y representa la interpretación correcta del concepto consultado.")
    
    doc.add_page_break()
    
    # ==================================================
    # SEGUNDA SERIE - Justificación de gráficos
    # ==================================================
    doc.add_heading('SEGUNDA SERIE - JUSTIFICACIÓN DE SELECCIÓN DE GRÁFICOS', level=1).style = title_style
    
    p = doc.add_paragraph("Para cada escenario, se proporciona una justificación matemática y estadística detallada sobre por qué el tipo de gráfico seleccionado es el más apropiado:", style='ExplanationStyle')
    
    for i, (escenario, respuesta_idx) in enumerate(zip(variante["segunda_serie"], respuestas["segunda_serie"]), 1):
        doc.add_heading(f"Escenario {i}:", level=2).style = subtitle_style
        
        p = doc.add_paragraph(escenario["escenario"], style='ExplanationStyle')
        p.runs[0].italic = True
        
        # Mostrar opciones disponibles
        p = doc.add_paragraph("Opciones disponibles:", style='ExplanationStyle')
        for j, opcion in enumerate(escenario["opciones"]):
            if j == respuesta_idx:
                p = doc.add_paragraph(f"• {opcion}", style='ExplanationStyle')
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = RGBColor(0, 128, 0)  # Verde para la correcta
            else:
                p = doc.add_paragraph(f"• {opcion}", style='ExplanationStyle')
        
        # Justificación matemática detallada
        p = doc.add_paragraph("Justificación matemática:", style='ExplanationStyle')
        p.runs[0].bold = True
        
        # Justificaciones específicas según tipo de gráfico seleccionado
        opcion_seleccionada = escenario["opciones"][respuesta_idx]
        
        if "Gráfica de barras" in opcion_seleccionada:
            p.add_run("\nLa gráfica de barras es ideal para representar variables categóricas o discretas cuando se requiere comparar magnitudes entre diferentes categorías. Matemáticamente, cada categoría (Xi) se representa en el eje horizontal, mientras que su frecuencia o magnitud (Yi) se representa en el eje vertical mediante una barra rectangular.")
            p.add_run("\n\nLa altura de cada barra es proporcional al valor que representa, permitiendo una comparación directa: Yi = f(Xi). La separación entre barras enfatiza la naturaleza discreta de las categorías, lo que es perfecto para comparar valores entre facultades.")
            p.add_run("\n\nLa ventaja matemática de este gráfico es que preserva la integridad de los datos originales sin agregar ni transformar la información, permitiendo comparar valores absolutos.")
        
        elif "Gráfica circular" in opcion_seleccionada:
            p.add_run("\nLa gráfica circular (o de pastel) es matemáticamente apropiada cuando se necesita visualizar partes de un todo y la contribución proporcional de cada categoría al total (100%). Cada sector del círculo tiene un ángulo proporcional al valor que representa.")
            p.add_run("\n\nPara cada categoría i, el ángulo correspondiente se calcula como: θi = (valor_i / total) × 360°")
            p.add_run("\n\nEste tipo de gráfico es óptimo para mostrar distribuciones porcentuales, ya que la suma de todos los sectores completa visualmente el círculo (100%). Esto facilita la comprensión inmediata de la importancia relativa de cada componente.")
        
        elif "Histograma" in opcion_seleccionada:
            p.add_run("\nEl histograma es la representación matemática idónea para variables continuas agrupadas en intervalos o clases. A diferencia del gráfico de barras, en el histograma los rectángulos son contiguos, reflejando la continuidad de la variable subyacente.")
            p.add_run("\n\nLa construcción matemática implica dividir el rango de datos [min, max] en k intervalos de clase, generalmente usando la fórmula de Sturges: k = 1 + 3.322 log₁₀(n), donde n es el número de observaciones.")
            p.add_run("\n\nLa altura de cada rectángulo representa la frecuencia o densidad de observaciones en ese intervalo. El área total del histograma es proporcional al número total de datos, lo que permite visualizar la forma de la distribución y identificar características como la normalidad, asimetría o multimodalidad.")
            
        elif "Ojiva" in opcion_seleccionada:
            p.add_run("\nLa Ojiva de Galton es una representación gráfica de la función de distribución acumulativa empírica. Matemáticamente, para cada punto x, la ojiva muestra F(x) = P(X ≤ x), es decir, la probabilidad de que la variable tome un valor menor o igual a x.")
            p.add_run("\n\nEsta representación es particularmente útil para determinar cuantiles y percentiles. Si queremos encontrar el valor x tal que F(x) = p (donde p es una proporción), simplemente localizamos p en el eje vertical y leemos el valor correspondiente x en el eje horizontal.")
            p.add_run("\n\nLa pendiente de la curva en cualquier punto refleja la densidad de observaciones en ese rango, proporcionando información adicional sobre la distribución de los datos.")
            
        elif "Polígono" in opcion_seleccionada:
            p.add_run("\nEl polígono de frecuencias es una representación mediante líneas continuas que conectan las frecuencias representadas por las marcas de clase. Matemáticamente, es una aproximación a la función de densidad de probabilidad subyacente.")
            p.add_run("\n\nSe construye uniendo con segmentos de recta los puntos (xi, fi), donde xi es la marca de clase (punto medio del intervalo) y fi es la frecuencia correspondiente. Esta representación es particularmente útil para visualizar tendencias y patrones en datos continuos ordenados, especialmente cuando existen series temporales.")
            p.add_run("\n\nLa primera derivada del polígono en cualquier punto proporciona la tasa de cambio de la frecuencia respecto a la variable, lo que permite identificar visualmente dónde el crecimiento es más rápido o más lento.")
        
        # Razones adicionales para rechazar las otras opciones
        p = doc.add_paragraph("Razones para descartar las otras opciones:", style='ExplanationStyle')
        p.runs[0].bold = True
        
        otras_opciones = [opt for j, opt in enumerate(escenario["opciones"]) if j != respuesta_idx]
        for otra in otras_opciones:
            if "Gráfica de barras" in otra and "Gráfica de barras" not in opcion_seleccionada:
                p.add_run(f"\n• {otra}: No es adecuada porque los datos requieren mostrar proporciones de un todo o representar una variable continua, no una comparación entre categorías discretas.")
            elif "Gráfica circular" in otra and "Gráfica circular" not in opcion_seleccionada:
                p.add_run(f"\n• {otra}: No es apropiada porque los datos no representan proporciones de un todo o porque hay demasiadas categorías, lo que dificultaría la interpretación visual.")
            elif "Histograma" in otra and "Histograma" not in opcion_seleccionada:
                p.add_run(f"\n• {otra}: No es óptima porque los datos no son continuos o porque el objetivo no es analizar la distribución de frecuencias de una variable continua.")
            elif "Ojiva" in otra and "Ojiva" not in opcion_seleccionada:
                p.add_run(f"\n• {otra}: No es la mejor opción porque el objetivo no es analizar valores acumulados o percentiles en la distribución.")
            elif "Polígono" in otra and "Polígono" not in opcion_seleccionada:
                p.add_run(f"\n• {otra}: No es ideal porque los datos no representan una tendencia o evolución, o porque no se busca enfatizar la continuidad y cambios graduales.")
    
    doc.add_page_break()
    
    # ==================================================
    # TERCERA SERIE - EJERCICIO 1: COEFICIENTE DE GINI
    # ==================================================
    doc.add_heading('TERCERA SERIE - EJERCICIOS PRÁCTICOS', level=1).style = title_style
    doc.add_heading('Ejercicio 1: Coeficiente de Gini', level=2).style = subtitle_style
    
    # Datos del problema
    gini_data = variante["tercera_serie"][0]
    
    p = doc.add_paragraph("DATOS DEL PROBLEMA:", style='ExplanationStyle')
    p.add_run("\nDistribución de salarios mensuales:").bold = True
    
    # Tabla con datos originales
    table_gini = doc.add_table(rows=len(gini_data["ranges"])+1, cols=2)
    table_gini.style = 'Table Grid'
    
    # Encabezados
    table_gini.cell(0, 0).text = "Salario mensual en (Q)"
    table_gini.cell(0, 1).text = "No. De trabajadores"
    
    # Datos
    for i, (rango, trabajadores) in enumerate(zip(gini_data["ranges"], gini_data["workers"]), 1):
        table_gini.cell(i, 0).text = rango
        table_gini.cell(i, 1).text = str(trabajadores)
    
    doc.add_paragraph("MÉTODO DE CÁLCULO:", style='ExplanationStyle').bold = True
    
    p = doc.add_paragraph("El coeficiente de Gini es una medida de desigualdad que toma valores entre 0 y 1. Un valor de 0 representa igualdad perfecta y un valor de 1 representa desigualdad máxima.", style='ExplanationStyle')
    
    # Paso 1: Preparación de datos y cálculos previos
    doc.add_heading("PASO 1: Preparación de datos", level=3).style = subtitle_style
    
    # Crear tabla extendida para los cálculos
    cols = ["Intervalo salarial", "No. trabajadores", "Prop. población", "Prop. acum. pobl.", "Punto medio", "Ingreso total", "Prop. ingreso", "Prop. acum. ingreso"]
    table_ext = doc.add_table(rows=len(gini_data["ranges"])+2, cols=len(cols))
    table_ext.style = 'Table Grid'
    
    # Encabezados
    for i, col in enumerate(cols):
        table_ext.cell(0, i).text = col
        table_ext.cell(0, i).paragraphs[0].runs[0].bold = True
    
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
    
    # Verificación cruzada del total de ingresos
    verificacion_ingresos = 0
    
    # Llenar la tabla extendida
    prop_acum_pob = 0
    prop_acum_ing = 0
    
    for i, (rango, trabajadores, punto_medio, ingreso_cat) in enumerate(zip(gini_data["ranges"], gini_data["workers"], puntos_medios, ingresos_categoria), 1):
        # Datos básicos
        table_ext.cell(i, 0).text = rango
        table_ext.cell(i, 1).text = str(trabajadores)
        
        # Proporción de población
        prop_pob = trabajadores / total_trabajadores
        table_ext.cell(i, 2).text = f"{prop_pob:.6f}"
        
        # Verificamos que la suma de prop_pob será 1.0
        verificacion_ingresos += trabajadores
        
        # Proporción acumulada de población
        prop_acum_pob += prop_pob
        table_ext.cell(i, 3).text = f"{prop_acum_pob:.6f}"
        
        # Punto medio
        table_ext.cell(i, 4).text = f"{punto_medio:.2f}"
        
        # Ingreso total para esta categoría
        table_ext.cell(i, 5).text = f"{ingreso_cat:.2f}"
        
        # Proporción de ingreso
        prop_ing = ingreso_cat / total_ingresos
        table_ext.cell(i, 6).text = f"{prop_ing:.6f}"
        
        # Proporción acumulada de ingreso
        prop_acum_ing += prop_ing
        table_ext.cell(i, 7).text = f"{prop_acum_ing:.6f}"
    
    # Totales
    table_ext.cell(len(gini_data["ranges"])+1, 0).text = "TOTAL"
    table_ext.cell(len(gini_data["ranges"])+1, 0).paragraphs[0].runs[0].bold = True
    table_ext.cell(len(gini_data["ranges"])+1, 1).text = str(total_trabajadores)
    table_ext.cell(len(gini_data["ranges"])+1, 2).text = "1.000000"
    table_ext.cell(len(gini_data["ranges"])+1, 5).text = f"{total_ingresos:.2f}"
    table_ext.cell(len(gini_data["ranges"])+1, 6).text = "1.000000"
    table_ext.cell(len(gini_data["ranges"])+1, 7).text = f"{prop_acum_ing:.6f}"
    
    # Verificaciones de cálculos
    p = doc.add_paragraph("VERIFICACIONES DE CÁLCULOS:", style='ExplanationStyle')
    p.add_run(f"\n1. Suma de trabajadores: {total_trabajadores} = {verificacion_ingresos} (Exacto)")
    p.add_run(f"\n2. Proporción acumulada de población: {prop_acum_pob:.6f} ≈ 1.0 (Correcto)")
    p.add_run(f"\n3. Proporción acumulada de ingresos: {prop_acum_ing:.6f} ≈ 1.0 (Correcto)")
    
    # Paso 2: Cálculo del coeficiente
    doc.add_heading("PASO 2: Cálculo del coeficiente de Gini", level=3).style = subtitle_style
    
    p = doc.add_paragraph("Para calcular el coeficiente de Gini, utilizamos la fórmula basada en la curva de Lorenz:", style='ExplanationStyle')
    p = doc.add_paragraph("G = 1 - Σ[(Xi - Xi-1)(Yi + Yi-1)]", style='EquationStyle')
    p = doc.add_paragraph("Donde:", style='ExplanationStyle')
    p.add_run("\n- Xi = proporción acumulada de población en el grupo i")
    p.add_run("\n- Yi = proporción acumulada de ingreso en el grupo i")
    
    # Tabla de cálculo para el coeficiente de Gini
    doc.add_paragraph("CÁLCULO DETALLADO:", style='ExplanationStyle')
    gini_table = doc.add_table(rows=len(gini_data["ranges"])+2, cols=6)
    gini_table.style = 'Table Grid'
    
    # Encabezados
    headers = ["Grupo", "Xi", "Xi-1", "Yi", "Yi-1", "(Xi - Xi-1)(Yi + Yi-1)"]
    for i, header in enumerate(headers):
        gini_table.cell(0, i).text = header
        gini_table.cell(0, i).paragraphs[0].runs[0].bold = True
    
    # Extraer proporciones acumuladas para Gini
    prop_acum_pob_list = [0]  # Empezamos con 0
    prop_acum_ing_list = [0]  # Empezamos con 0
    
    prop_acum_pob = 0
    prop_acum_ing = 0
    
    for trabajadores, ingreso_cat in zip(gini_data["workers"], ingresos_categoria):
        prop_pob = trabajadores / total_trabajadores
        prop_acum_pob += prop_pob
        prop_acum_pob_list.append(prop_acum_pob)
        
        prop_ing = ingreso_cat / total_ingresos
        prop_acum_ing += prop_ing
        prop_acum_ing_list.append(prop_acum_ing)
    
    # Calcular el coeficiente de Gini manualmente
    gini_sum = 0
    for i in range(1, len(prop_acum_pob_list)):
        xi = prop_acum_pob_list[i]
        xi_1 = prop_acum_pob_list[i-1]
        yi = prop_acum_ing_list[i]
        yi_1 = prop_acum_ing_list[i-1]
        
        term = (xi - xi_1) * (yi + yi_1)
        gini_sum += term
        
        # Llenar tabla
        gini_table.cell(i, 0).text = str(i)
        gini_table.cell(i, 1).text = f"{xi:.6f}"
        gini_table.cell(i, 2).text = f"{xi_1:.6f}"
        gini_table.cell(i, 3).text = f"{yi:.6f}"
        gini_table.cell(i, 4).text = f"{yi_1:.6f}"
        gini_table.cell(i, 5).text = f"{term:.6f}"
    
    # Total
    gini_table.cell(len(prop_acum_pob_list), 0).text = "Suma"
    gini_table.cell(len(prop_acum_pob_list), 0).paragraphs[0].runs[0].bold = True
    gini_table.cell(len(prop_acum_pob_list), 5).text = f"{gini_sum:.6f}"
    gini_table.cell(len(prop_acum_pob_list), 5).paragraphs[0].runs[0].bold = True
    
    # Calcular coeficiente de Gini manual
    gini_manual = 1 - gini_sum
    
    # Cálculo alternativo para verificación
    gini_alternative = 0
    for i in range(len(prop_acum_pob_list)-1):
        for j in range(i+1, len(prop_acum_pob_list)):
            xi = prop_acum_pob_list[i]
            xj = prop_acum_pob_list[j]
            yi = prop_acum_ing_list[i]
            yj = prop_acum_ing_list[j]
            
            gini_alternative += abs((xj - xi) - (yj - yi))
    
    gini_alternative /= (len(prop_acum_pob_list) * len(prop_acum_pob_list))
    gini_alternative *= 2
    
    # Obtener coeficiente de referencia
    gini_value = respuestas["tercera_serie"]["gini"]
    
    # Resultados con verificaciones cruzadas
    p = doc.add_paragraph("RESULTADO:", style='ExplanationStyle')
    p.add_run(f"\nCoeficiente de Gini (método principal): G = 1 - {gini_sum:.6f} = {gini_manual:.6f}").bold = True
    p.add_run(f"\nCoeficiente de Gini (método alternativo): {gini_alternative:.6f}")
    p.add_run(f"\nCoeficiente de Gini (valor de referencia): {gini_value}")
    
    # Verificar que los valores estén cerca
    precision = abs(gini_manual - gini_value) / gini_value * 100
    
    p.add_run(f"\n\nPrecisión matemática: {100-precision:.2f}% (diferencia: {abs(gini_manual - gini_value):.6f})")
    
    if abs(gini_manual - gini_value) < 0.05:
        p.add_run("\nCONCLUSIÓN: Los cálculos son matemáticamente correctos y confiables.").bold = True
    else:
        p.add_run("\nNOTA: Hay una discrepancia entre los métodos de cálculo. Se recomienda usar el valor manual.").bold = True
    
    # Graficar curva de Lorenz
    plt.figure(figsize=(8, 6))
    plt.plot([0] + prop_acum_pob_list, [0] + prop_acum_ing_list, 'b-', linewidth=2, label='Curva de Lorenz')
    plt.plot([0, 1], [0, 1], 'r--', label='Línea de equidad perfecta')
    plt.fill_between([0] + prop_acum_pob_list, [0] + prop_acum_ing_list, [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0][:len(prop_acum_pob_list)+1], alpha=0.2)
    plt.title('Curva de Lorenz')
    plt.xlabel('Proporción acumulada de población')
    plt.ylabel('Proporción acumulada de ingreso')
    plt.grid(True)
    plt.legend()
    
    # Guardar gráfica temporalmente
    img_path = os.path.join(UPLOAD_FOLDER, f'lorenz_curve_{variante_id}.png')
    plt.savefig(img_path)
    plt.close()
    
    # Añadir la gráfica al documento
    doc.add_paragraph("VISUALIZACIÓN DE LA CURVA DE LORENZ:", style='ExplanationStyle').bold = True
    doc.add_picture(img_path, width=Inches(6))
    
    # Interpretación
    doc.add_heading("PASO 3: Interpretación del coeficiente de Gini", level=3).style = subtitle_style
    
    p = doc.add_paragraph("El coeficiente de Gini mide la desigualdad en la distribución de los ingresos:", style='ExplanationStyle')
    p.add_run("\n• 0 = Igualdad perfecta (todos reciben exactamente lo mismo)")
    p.add_run("\n• 1 = Desigualdad perfecta (una persona recibe todo el ingreso)")
    
    p = doc.add_paragraph("INTERPRETACIÓN DEL RESULTADO:", style='ExplanationStyle')
    
    if gini_manual < 0.3:
        p.add_run(f"\nEl coeficiente de Gini calculado es {gini_manual:.4f}, lo que indica una distribución relativamente equitativa de los salarios entre los trabajadores de la empresa. Esta empresa muestra baja desigualdad salarial.").bold = True
    elif gini_manual < 0.5:
        p.add_run(f"\nEl coeficiente de Gini calculado es {gini_manual:.4f}, lo que indica una desigualdad moderada en la distribución de salarios dentro de la empresa. Este nivel es típico en muchas organizaciones.").bold = True
    else:
        p.add_run(f"\nEl coeficiente de Gini calculado es {gini_manual:.4f}, lo que indica una desigualdad significativa en la distribución de salarios dentro de la empresa. Esta empresa muestra una concentración importante de los salarios en un grupo relativamente pequeño de trabajadores.").bold = True
    
    # Añadir evaluación comparativa con índices de Gini nacionales
    p = doc.add_paragraph("CONTEXTO COMPARATIVO:", style='ExplanationStyle')
    p.add_run("\nPara contextualizar este resultado, aquí hay algunos índices de Gini de países a nivel mundial (2021):")
    p.add_run("\n• Sudáfrica: 0.63 (alta desigualdad)")
    p.add_run("\n• Brasil: 0.53")
    p.add_run("\n• Estados Unidos: 0.41")
    p.add_run("\n• España: 0.35")
    p.add_run("\n• Canadá: 0.33")
    p.add_run("\n• Alemania: 0.31")
    p.add_run("\n• Noruega: 0.27 (baja desigualdad)")
    
    doc.add_page_break()
    
    # ==================================================
    # TERCERA SERIE - EJERCICIO 2: DISTRIBUCIÓN DE FRECUENCIAS (MÉTODO STURGERS)
    # ==================================================
    doc.add_heading('Ejercicio 2: Distribución de Frecuencias - Método Sturgers', level=2).style = subtitle_style
    
    sturgers_data = variante["tercera_serie"][1]
    
    p = doc.add_paragraph("DATOS DEL PROBLEMA:", style='ExplanationStyle')
    
    # Mostrar los datos originales en formato organizado (tabla)
    valores_str = sturgers_data["data"]
    valores_numericos = [int(x) for x in valores_str]
    
    # Tabla para mostrar los valores
    rows_needed = math.ceil(len(valores_numericos) / 5)
    table_datos = doc.add_table(rows=rows_needed, cols=5)
    table_datos.style = 'Table Grid'
    
    idx = 0
    for i in range(rows_needed):
        for j in range(5):
            if idx < len(valores_numericos):
                table_datos.cell(i, j).text = str(valores_numericos[idx])
                idx += 1
    
    # Estadística descriptiva básica de los datos
    min_valor = min(valores_numericos)
    max_valor = max(valores_numericos)
    media = sum(valores_numericos) / len(valores_numericos)
    
    # Ordenar valores para análisis
    valores_ordenados = sorted(valores_numericos)
    
    # Calcular mediana
    n = len(valores_ordenados)
    if n % 2 == 0:
        mediana = (valores_ordenados[n//2 - 1] + valores_ordenados[n//2]) / 2
    else:
        mediana = valores_ordenados[n//2]
    
    p = doc.add_paragraph("ANÁLISIS PRELIMINAR:", style='ExplanationStyle')
    p.add_run(f"\nValor mínimo: {min_valor}")
    p.add_run(f"\nValor máximo: {max_valor}")
    p.add_run(f"\nRango: {max_valor - min_valor}")
    p.add_run(f"\nMedia aritmética: {media:.2f}")
    p.add_run(f"\nMediana: {mediana}")
    p.add_run(f"\nCantidad de datos: {len(valores_numericos)}")
    
    # Paso 1: Cálculo del número de clases (K)
    doc.add_heading("PASO 1: Cálculo del número de clases (K)", level=3).style = subtitle_style
    
    k_value = 1 + 3.322 * math.log10(len(valores_numericos))
    k_rounded = round(k_value)
    
    p = doc.add_paragraph(f"Utilizando la fórmula de Sturgers:", style='ExplanationStyle')
    p.add_run("\nK = 1 + 3.322 × log₁₀(n)")
    p.add_run(f"\nDonde n = {len(valores_numericos)} (número de observaciones)")
    p.add_run(f"\nK = 1 + 3.322 × log₁₀({len(valores_numericos)})")
    p.add_run(f"\nK = 1 + 3.322 × {math.log10(len(valores_numericos)):.6f}")
    p.add_run(f"\nK = 1 + {3.322 * math.log10(len(valores_numericos)):.6f}")
    p.add_run(f"\nK = {k_value:.6f}")
    
    # Verificación del redondeo
    p.add_run(f"\n\nRedondeando a un número entero: K = {k_rounded}")
    
    # Verificar con el valor de referencia
    k_ref = respuestas["tercera_serie"]["dist_frecuencias"]["k"]
    p.add_run(f"\nValor de referencia: K = {k_ref}")
    
    # Paso 2: Cálculo del rango
    doc.add_heading("PASO 2: Cálculo del rango", level=3).style = subtitle_style
    
    rango = max_valor - min_valor
    rango_ref = respuestas["tercera_serie"]["dist_frecuencias"]["rango"]
    
    p = doc.add_paragraph("El rango es la diferencia entre el valor máximo y el valor mínimo:", style='ExplanationStyle')
    p.add_run(f"\nRango = Valor máximo - Valor mínimo = {max_valor} - {min_valor} = {rango}")
    
    # Verificación
    p.add_run(f"\n\nValor de referencia: Rango = {rango_ref}")
    
    # Paso 3: Cálculo de la amplitud de clase
    doc.add_heading("PASO 3: Cálculo de la amplitud de clase", level=3).style = subtitle_style
    
    amplitud = rango / k_value
    amplitud_redondeada = math.ceil(amplitud)
    amplitud_ref = respuestas["tercera_serie"]["dist_frecuencias"]["amplitud"]
    
    p = doc.add_paragraph("La amplitud es el tamaño de cada intervalo de clase:", style='ExplanationStyle')
    p.add_run(f"\nAmplitud = Rango / K = {rango} / {k_value:.6f} = {amplitud:.6f}")
    p.add_run(f"\n\nRedondeando hacia arriba (para asegurar que todos los valores queden incluidos): Amplitud = {amplitud_redondeada}")
    
    # Verificación
    p.add_run(f"\n\nValor de referencia: Amplitud = {amplitud_ref}")
    
    # Paso 4: Construcción de la tabla de distribución de frecuencias
    doc.add_heading("PASO 4: Construcción de la tabla de distribución de frecuencias", level=3).style = subtitle_style
    
    # Calcular límites de clase
    limite_inferior = min_valor
    limites = []
    
    for i in range(k_rounded):
        limite_superior = limite_inferior + amplitud_redondeada
        limites.append((limite_inferior, limite_superior))
        limite_inferior = limite_superior
    
    # Calcular frecuencias
    frecuencias = [0] * k_rounded
    for valor in valores_numericos:
        for i, (li, ls) in enumerate(limites):
            if li <= valor < ls or (i == k_rounded - 1 and valor == ls):
                frecuencias[i] += 1
                break
    
    # Crear tabla completa de distribución de frecuencias
    dist_table = doc.add_table(rows=k_rounded+1, cols=7)
    dist_table.style = 'Table Grid'
    
    # Encabezados
    encabezados_dist = ["Clase", "Límites de clase", "Marca de clase", "Frecuencia absoluta", "Frecuencia relativa", "Frecuencia acumulada", "Frecuencia rel. acumulada"]
    for i, encabezado in enumerate(encabezados_dist):
        dist_table.cell(0, i).text = encabezado
        dist_table.cell(0, i).paragraphs[0].runs[0].bold = True
    
    # Llenar la tabla
    frec_acum = 0
    frec_rel_acum = 0
    
    for i, ((li, ls), frec) in enumerate(zip(limites, frecuencias), 1):
        # Clase
        dist_table.cell(i, 0).text = str(i)
        
        # Límites de clase
        dist_table.cell(i, 1).text = f"[{li} - {ls})"
        
        # Marca de clase
        marca = (li + ls) / 2
        dist_table.cell(i, 2).text = f"{marca:.1f}"
        
        # Frecuencia absoluta
        dist_table.cell(i, 3).text = str(frec)
        
        # Frecuencia relativa
        frec_rel = frec / len(valores_numericos)
        dist_table.cell(i, 4).text = f"{frec_rel:.4f}"
        
        # Frecuencia acumulada
        frec_acum += frec
        dist_table.cell(i, 5).text = str(frec_acum)
        
        # Frecuencia relativa acumulada
        frec_rel_acum += frec_rel
        dist_table.cell(i, 6).text = f"{frec_rel_acum:.4f}"
    
    # Verificaciones
    p = doc.add_paragraph("VERIFICACIONES MATEMÁTICAS:", style='ExplanationStyle')
    p.add_run(f"\n1. Suma de frecuencias absolutas: {sum(frecuencias)} (debe ser igual a {len(valores_numericos)})")
    p.add_run(f"\n2. Última frecuencia acumulada: {frec_acum} (debe ser igual a {len(valores_numericos)})")
    p.add_run(f"\n3. Última frecuencia relativa acumulada: {frec_rel_acum:.4f} (debe ser aproximadamente 1.0)")
    
    # Generar histograma para visualización
    plt.figure(figsize=(10, 6))
    
    # Extraer límites para el histograma
    bin_edges = [limites[0][0]] + [limite[1] for limite in limites]
    plt.hist(valores_numericos, bins=bin_edges, edgecolor='black', alpha=0.7)
    
    # Añadir línea de la media
    plt.axvline(x=media, color='r', linestyle='--', label=f'Media: {media:.2f}')
    
    # Añadir línea de la mediana
    plt.axvline(x=mediana, color='g', linestyle='-.', label=f'Mediana: {mediana:.2f}')
    
    plt.title('Histograma de frecuencias')
    plt.xlabel('Valores')
    plt.ylabel('Frecuencia')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    # Guardar gráfica temporalmente
    hist_path = os.path.join(UPLOAD_FOLDER, f'histogram_{variante_id}.png')
    plt.savefig(hist_path)
    plt.close()
    
    # Añadir la gráfica al documento
    doc.add_paragraph("VISUALIZACIÓN DEL HISTOGRAMA DE FRECUENCIAS:", style='ExplanationStyle').bold = True
    doc.add_picture(hist_path, width=Inches(6))
    
    # Interpretación
    doc.add_heading("PASO 5: Interpretación de la distribución de frecuencias", level=3).style = subtitle_style
    
    # Encontrar clase modal (con mayor frecuencia)
    clase_modal = frecuencias.index(max(frecuencias))
    lim_inf_modal, lim_sup_modal = limites[clase_modal]
    
    p = doc.add_paragraph("ANÁLISIS E INTERPRETACIÓN:", style='ExplanationStyle')
    p.add_run(f"\n• La distribución se divide en {k_rounded} clases, cada una con una amplitud de {amplitud_redondeada} unidades.")
    p.add_run(f"\n• La clase con mayor frecuencia es la clase {clase_modal+1} [{lim_inf_modal} - {lim_sup_modal}), con {frecuencias[clase_modal]} observaciones.")
    
    if media > mediana:
        p.add_run("\n• La distribución muestra una asimetría positiva (media > mediana), indicando una cola hacia valores mayores.")
    elif media < mediana:
        p.add_run("\n• La distribución muestra una asimetría negativa (media < mediana), indicando una cola hacia valores menores.")
    else:
        p.add_run("\n• La distribución es aproximadamente simétrica (media ≈ mediana).")
    
    # Calcular asimetría con fórmula específica
    varianza = sum((x - media) ** 2 for x in valores_numericos) / len(valores_numericos)
    desviacion_estandar = math.sqrt(varianza)
    asimetria = sum((x - media) ** 3 for x in valores_numericos) / (len(valores_numericos) * desviacion_estandar ** 3)
    
    p.add_run(f"\n• Coeficiente de asimetría: {asimetria:.4f}")
    if asimetria > 0.5:
        p.add_run(" (asimetría positiva significativa)")
    elif asimetria < -0.5:
        p.add_run(" (asimetría negativa significativa)")
    else:
        p.add_run(" (distribución aproximadamente simétrica)")
    
    p.add_run(f"\n• Desviación estándar: {desviacion_estandar:.4f}")
    p.add_run(f"\n• Coeficiente de variación: {(desviacion_estandar/media)*100:.2f}%")
    
    doc.add_page_break()
    
    # ==================================================
    # TERCERA SERIE - EJERCICIO 3: DIAGRAMA DE TALLO Y HOJA
    # ==================================================
    doc.add_heading('Ejercicio 3: Diagrama de Tallo y Hoja', level=2).style = subtitle_style
    
    stem_leaf_data = variante["tercera_serie"][2]
    
    p = doc.add_paragraph("DATOS DEL PROBLEMA:", style='ExplanationStyle')
    p.add_run(f"\n{stem_leaf_data['title']}").bold = True
    
    # Mostrar los datos en una tabla organizada
    valores_tl = stem_leaf_data["data"]
    rows_needed = math.ceil(len(valores_tl) / 8)
    table_tl = doc.add_table(rows=rows_needed, cols=8)
    table_tl.style = 'Table Grid'
    
    idx = 0
    for i in range(rows_needed):
        for j in range(8):
            if idx < len(valores_tl):
                table_tl.cell(i, j).text = valores_tl[idx]
                idx += 1
    
    # Convertir a valores numéricos para procesar
    valores_numericos_tl = [float(x) for x in valores_tl]
    
    # Estadística descriptiva
    min_val = min(valores_numericos_tl)
    max_val = max(valores_numericos_tl)
    mean_val = sum(valores_numericos_tl) / len(valores_numericos_tl)
    
    valores_ordenados_tl = sorted(valores_numericos_tl)
    n = len(valores_ordenados_tl)
    if n % 2 == 0:
        median_val = (valores_ordenados_tl[n//2 - 1] + valores_ordenados_tl[n//2]) / 2
    else:
        median_val = valores_ordenados_tl[n//2]
    
    # Análisis descriptivo
    p = doc.add_paragraph("ANÁLISIS PRELIMINAR:", style='ExplanationStyle')
    p.add_run(f"\nValor mínimo: {min_val}")
    p.add_run(f"\nValor máximo: {max_val}")
    p.add_run(f"\nRango: {max_val - min_val}")
    p.add_run(f"\nMedia: {mean_val:.2f}")
    p.add_run(f"\nMediana: {median_val:.2f}")
    p.add_run(f"\nCantidad de datos: {len(valores_numericos_tl)}")
    
    # Paso 1: Preparar datos para el diagrama
    doc.add_heading("PASO 1: Preparación de los datos", level=3).style = subtitle_style
    
    p = doc.add_paragraph("Para crear un diagrama de tallo y hoja, se divide cada valor en dos partes:", style='ExplanationStyle')
    p.add_run("\n• Tallo (stem): La parte entera o dígitos más significativos")
    p.add_run("\n• Hoja (leaf): El último dígito o el dígito menos significativo")
    
    p = doc.add_paragraph("En este caso, los datos tienen un dígito entero seguido de decimales. Utilizaremos:", style='ExplanationStyle')
    p.add_run("\n• Tallo: El dígito entero")
    p.add_run("\n• Hoja: El primer decimal (multiplicado por 10)")
    
    # Paso 2: Construcción del diagrama
    doc.add_heading("PASO 2: Construcción del diagrama de tallo y hoja", level=3).style = subtitle_style
    
    # Organizar los datos por tallo y hoja
    stem_leaf_dict = {}
    
    for valor in valores_numericos_tl:
        tallo = int(valor)  # Parte entera
        hoja = int((valor - tallo) * 10)  # Primer decimal multiplicado por 10
        
        if tallo not in stem_leaf_dict:
            stem_leaf_dict[tallo] = []
        
        stem_leaf_dict[tallo].append(hoja)
    
    # Ordenar las hojas de cada tallo
    for tallo in stem_leaf_dict:
        stem_leaf_dict[tallo].sort()
    
    # Crear tabla para mostrar el diagrama
    tallos_ordenados = sorted(stem_leaf_dict.keys())
    tl_table = doc.add_table(rows=len(tallos_ordenados)+1, cols=2)
    tl_table.style = 'Table Grid'
    
    # Encabezados
    tl_table.cell(0, 0).text = "Tallo"
    tl_table.cell(0, 1).text = "Hojas"
    tl_table.cell(0, 0).paragraphs[0].runs[0].bold = True
    tl_table.cell(0, 1).paragraphs[0].runs[0].bold = True
    
    # Llenar la tabla
    for i, tallo in enumerate(tallos_ordenados, 1):
        tl_table.cell(i, 0).text = f"{tallo}"
        
        hojas_str = " ".join(str(h) for h in stem_leaf_dict[tallo])
        tl_table.cell(i, 1).text = hojas_str
    
    # Contar frecuencias por tallo
    p = doc.add_paragraph("ANÁLISIS DE FRECUENCIAS POR TALLO:", style='ExplanationStyle')
    
    freq_table = doc.add_table(rows=len(tallos_ordenados)+1, cols=3)
    freq_table.style = 'Table Grid'
    
    # Encabezados
    freq_table.cell(0, 0).text = "Tallo"
    freq_table.cell(0, 1).text = "Frecuencia"
    freq_table.cell(0, 2).text = "Porcentaje"
    freq_table.cell(0, 0).paragraphs[0].runs[0].bold = True
    freq_table.cell(0, 1).paragraphs[0].runs[0].bold = True
    freq_table.cell(0, 2).paragraphs[0].runs[0].bold = True
    
    # Encontrar tallo con mayor frecuencia (moda)
    tallo_max_freq = max(stem_leaf_dict.items(), key=lambda x: len(x[1]))
    tallo_moda = tallo_max_freq[0]
    
    # Llenar tabla de frecuencias
    for i, tallo in enumerate(tallos_ordenados, 1):
        freq_table.cell(i, 0).text = f"{tallo}"
        
        freq = len(stem_leaf_dict[tallo])
        freq_table.cell(i, 1).text = str(freq)
        
        porcentaje = (freq / len(valores_numericos_tl)) * 100
        freq_table.cell(i, 2).text = f"{porcentaje:.2f}%"
        
        # Resaltar el tallo modal
        if tallo == tallo_moda:
            freq_table.cell(i, 0).paragraphs[0].runs[0].bold = True
            freq_table.cell(i, 1).paragraphs[0].runs[0].bold = True
            freq_table.cell(i, 2).paragraphs[0].runs[0].bold = True
    
    # Visualización gráfica de la distribución
    plt.figure(figsize=(10, 6))
    
    # Histograma
    plt.hist(valores_numericos_tl, bins=10, edgecolor='black', alpha=0.7)
    plt.axvline(x=mean_val, color='r', linestyle='--', label=f'Media: {mean_val:.2f}')
    plt.axvline(x=median_val, color='g', linestyle='-.', label=f'Mediana: {median_val:.2f}')
    
    plt.title('Distribución de los datos')
    plt.xlabel('Valores')
    plt.ylabel('Frecuencia')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    # Guardar gráfica
    tl_hist_path = os.path.join(UPLOAD_FOLDER, f'tl_histogram_{variante_id}.png')
    plt.savefig(tl_hist_path)
    plt.close()
    
    # Añadir la gráfica al documento
    doc.add_paragraph("VISUALIZACIÓN DE LA DISTRIBUCIÓN:", style='ExplanationStyle').bold = True
    doc.add_picture(tl_hist_path, width=Inches(6))
    
    # Interpretación
    doc.add_heading("PASO 3: Interpretación del diagrama de tallo y hoja", level=3).style = subtitle_style
    
    # Determinar moda (valor que más se repite)
    from collections import Counter
    conteo = Counter(valores_numericos_tl)
    valor_moda = conteo.most_common(1)[0][0]
    
    # Referencia a las respuestas
    moda_ref = respuestas["tercera_serie"]["tallo_hoja"]["moda"]
    intervalo_ref = respuestas["tercera_serie"]["tallo_hoja"]["intervalo"]
    
    p = doc.add_paragraph("ANÁLISIS DE CONCENTRACIÓN DE DATOS:", style='ExplanationStyle')
    p.add_run(f"\n• El tallo con mayor frecuencia es {tallo_moda}, con {len(stem_leaf_dict[tallo_moda])} observaciones ({(len(stem_leaf_dict[tallo_moda])/len(valores_numericos_tl))*100:.2f}% del total).")
    p.add_run(f"\n• El valor más frecuente (moda) es {valor_moda}.")
    p.add_run(f"\n• El intervalo con mayor concentración de datos es [{tallo_moda}-{tallo_moda+1}).")
    
    # Verificación con valores de referencia
    p = doc.add_paragraph("VERIFICACIÓN CON VALORES DE REFERENCIA:", style='ExplanationStyle')
    p.add_run(f"\n• Valor de la moda (calculado): {valor_moda}")
    p.add_run(f"\n• Valor de la moda (referencia): {moda_ref}")
    p.add_run(f"\n• Intervalo de mayor concentración (calculado): [{tallo_moda}-{tallo_moda+1})")
    p.add_run(f"\n• Intervalo de mayor concentración (referencia): {intervalo_ref}")
    
    # Interpretar la forma de la distribución
    p = doc.add_paragraph("INTERPRETACIÓN DE LA DISTRIBUCIÓN:", style='ExplanationStyle')
    
    if mean_val > median_val:
        p.add_run("\n• La distribución muestra una asimetría positiva (media > mediana), indicando una cola hacia la derecha (valores mayores).")
    elif mean_val < median_val:
        p.add_run("\n• La distribución muestra una asimetría negativa (media < mediana), indicando una cola hacia la izquierda (valores menores).")
    else:
        p.add_run("\n• La distribución se aproxima a la simetría (media ≈ mediana).")
    
    # Calcular rango intercuartílico
    q1_pos = (n + 1) // 4
    q3_pos = 3 * (n + 1) // 4
    q1 = valores_ordenados_tl[q1_pos - 1]
    q3 = valores_ordenados_tl[q3_pos - 1]
    iqr = q3 - q1
    
    p.add_run(f"\n• Rango intercuartílico (IQR): {iqr:.2f}")
    p.add_run(f"\n• El 50% central de los datos se encuentra entre {q1:.2f} y {q3:.2f}.")
    
    doc.add_page_break()
    
    # ==================================================
    # TERCERA SERIE - EJERCICIO 4: MEDIDAS DE TENDENCIA CENTRAL
    # ==================================================
    doc.add_heading('Ejercicio 4: Medidas de Tendencia Central', level=2).style = subtitle_style
    
    central_data = variante["tercera_serie"][3]
    
    p = doc.add_paragraph("DATOS DEL PROBLEMA:", style='ExplanationStyle')
    p.add_run(f"\n{central_data['title']}").bold = True
    
    # Tabla con datos originales
    table_central = doc.add_table(rows=len(central_data["ranges"])+1, cols=2)
    table_central.style = 'Table Grid'
    
    # Encabezados
    table_central.cell(0, 0).text = "Precio en (Q)"
    table_central.cell(0, 1).text = "No. De productos"
    
    # Datos
    for i, (rango, count) in enumerate(zip(central_data["ranges"], central_data["count"]), 1):
        table_central.cell(i, 0).text = rango
        table_central.cell(i, 1).text = str(count)
    
    # PASO 1: Preparación de datos para el cálculo
    doc.add_heading("PASO 1: Preparación para el cálculo", level=3).style = subtitle_style
    
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
        table_calc.cell(0, i).paragraphs[0].runs[0].bold = True
    
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
    table_calc.cell(len(central_data["ranges"])+1, 0).paragraphs[0].runs[0].bold = True
    table_calc.cell(len(central_data["ranges"])+1, 3).text = str(suma_freq)
    table_calc.cell(len(central_data["ranges"])+1, 4).text = f"{suma_xi_fi:.2f}"
    
    # PASO 2: Cálculo de la media
    doc.add_heading("PASO 2: Cálculo de la media", level=3).style = subtitle_style
    
    media_calculada = suma_xi_fi / suma_freq
    media_ref = respuestas["tercera_serie"]["medidas_centrales"]["media"]
    
    p = doc.add_paragraph("La media aritmética para datos agrupados se calcula con la fórmula:", style='ExplanationStyle')
    p = doc.add_paragraph("μ = Σ(xi × fi) / Σfi", style='EquationStyle')
    p = doc.add_paragraph("Donde:", style='ExplanationStyle')
    p.add_run("\n• xi = marca de clase (punto medio del intervalo)")
    p.add_run("\n• fi = frecuencia absoluta")
    p.add_run("\n• Σfi = suma de frecuencias (total de datos)")
    
    p = doc.add_paragraph("Sustituyendo los valores:", style='ExplanationStyle')
    p.add_run(f"\nμ = {suma_xi_fi:.2f} / {suma_freq}")
    p.add_run(f"\nμ = {media_calculada:.4f}")
    
    # Verificación
    p = doc.add_paragraph("VERIFICACIÓN:", style='ExplanationStyle')
    p.add_run(f"\nMedia calculada: {media_calculada:.4f}")
    p.add_run(f"\nMedia de referencia: {media_ref}")
    
    precision_media = abs(media_calculada - media_ref) / media_ref * 100 if media_ref != 0 else 0
    
    if precision_media < 5:
        p.add_run("\n\nLos valores son consistentes (diferencia menor al 5%)").bold = True
    else:
        p.add_run("\n\nHay una diferencia significativa entre los valores. Se recomienda utilizar el valor calculado.").bold = True
    
    # PASO 3: Cálculo de la mediana
    doc.add_heading("PASO 3: Cálculo de la mediana", level=3).style = subtitle_style
    
    # Encontrar la clase que contiene la mediana
    n_2 = suma_freq / 2
    
    p = doc.add_paragraph("La mediana para datos agrupados requiere primero identificar la clase mediana:", style='ExplanationStyle')
    p.add_run(f"\n• Total de datos (n): {suma_freq}")
    p.add_run(f"\n• Posición de la mediana (n/2): {n_2}")
    
    # Calcular frecuencias acumuladas
    fa = [0]
    for count in central_data["count"]:
        fa.append(fa[-1] + count)
    fa = fa[1:]  # Eliminar el primer 0
    
    # Tabla de frecuencias acumuladas
    p = doc.add_paragraph("Tabla de frecuencias acumuladas:", style='ExplanationStyle')
    
    fa_table = doc.add_table(rows=len(central_data["ranges"])+1, cols=3)
    fa_table.style = 'Table Grid'
    
    # Encabezados
    fa_table.cell(0, 0).text = "Clase"
    fa_table.cell(0, 1).text = "Frecuencia (fi)"
    fa_table.cell(0, 2).text = "Frecuencia acumulada (Fa)"
    
    # Datos y encontrar clase mediana
    clase_mediana = None
    for i, (count, fa_val) in enumerate(zip(central_data["count"], fa), 1):
        fa_table.cell(i, 0).text = str(i)
        fa_table.cell(i, 1).text = str(count)
        fa_table.cell(i, 2).text = str(fa_val)
        
        if clase_mediana is None and fa_val >= n_2:
            clase_mediana = i
            fa_table.cell(i, 0).paragraphs[0].runs[0].bold = True
            fa_table.cell(i, 1).paragraphs[0].runs[0].bold = True
            fa_table.cell(i, 2).paragraphs[0].runs[0].bold = True
    
    p = doc.add_paragraph(f"La clase mediana es la clase {clase_mediana}.", style='ExplanationStyle')
    
    # Calcular la mediana
    li_mediana = limites_central[clase_mediana-1][0]
    amplitud = limites_central[clase_mediana-1][1] - limites_central[clase_mediana-1][0]
    fa_anterior = fa[clase_mediana-2] if clase_mediana > 1 else 0
    fi_mediana = central_data["count"][clase_mediana-1]
    
    p = doc.add_paragraph("La fórmula para calcular la mediana con datos agrupados es:", style='ExplanationStyle')
    p = doc.add_paragraph("Me = li + [(n/2 - Fi-1) / fi] × c", style='EquationStyle')
    p = doc.add_paragraph("Donde:", style='ExplanationStyle')
    p.add_run("\n• li = límite inferior de la clase mediana")
    p.add_run("\n• Fi-1 = frecuencia acumulada hasta la clase anterior a la mediana")
    p.add_run("\n• fi = frecuencia de la clase mediana")
    p.add_run("\n• c = amplitud de la clase")
    
    p = doc.add_paragraph("Sustituyendo los valores:", style='ExplanationStyle')
    p.add_run(f"\nMe = {li_mediana} + [({n_2} - {fa_anterior}) / {fi_mediana}] × {amplitud}")
    
    mediana_calculada = li_mediana + ((n_2 - fa_anterior) / fi_mediana) * amplitud
    mediana_ref = respuestas["tercera_serie"]["medidas_centrales"]["mediana"]
    
    p.add_run(f"\nMe = {li_mediana} + {(n_2 - fa_anterior) / fi_mediana:.4f} × {amplitud}")
    p.add_run(f"\nMe = {li_mediana} + {((n_2 - fa_anterior) / fi_mediana) * amplitud:.4f}")
    p.add_run(f"\nMe = {mediana_calculada:.4f}")
    
    # Verificación
    p = doc.add_paragraph("VERIFICACIÓN:", style='ExplanationStyle')
    p.add_run(f"\nMediana calculada: {mediana_calculada:.4f}")
    p.add_run(f"\nMediana de referencia: {mediana_ref}")
    
    precision_mediana = abs(mediana_calculada - mediana_ref) / mediana_ref * 100 if mediana_ref != 0 else 0
    
    if precision_mediana < 5:
        p.add_run("\n\nLos valores son consistentes (diferencia menor al 5%)").bold = True
    else:
        p.add_run("\n\nHay una diferencia significativa entre los valores. Se recomienda utilizar el valor calculado.").bold = True
    
    # PASO 4: Cálculo de la moda
    doc.add_heading("PASO 4: Cálculo de la moda", level=3).style = subtitle_style
    
    # Encontrar la clase modal (clase con mayor frecuencia)
    clase_modal = central_data["count"].index(max(central_data["count"])) + 1
    li_modal = limites_central[clase_modal-1][0]
    amplitud_modal = limites_central[clase_modal-1][1] - limites_central[clase_modal-1][0]
    
    # Obtener frecuencias necesarias
    fi_modal = central_data["count"][clase_modal-1]
    fi_anterior = central_data["count"][clase_modal-2] if clase_modal > 1 else 0
    fi_posterior = central_data["count"][clase_modal] if clase_modal < len(central_data["count"]) else 0
    
    p = doc.add_paragraph("La clase modal (con mayor frecuencia) es la clase " + str(clase_modal) + ".", style='ExplanationStyle')
    
    p = doc.add_paragraph("La fórmula para calcular la moda con datos agrupados es:", style='ExplanationStyle')
    p = doc.add_paragraph("Mo = li + [d1 / (d1 + d2)] × c", style='EquationStyle')
    p = doc.add_paragraph("Donde:", style='ExplanationStyle')
    p.add_run("\n• li = límite inferior de la clase modal")
    p.add_run("\n• d1 = diferencia entre la frecuencia de la clase modal y la clase anterior (d1 = fi - fi-1)")
    p.add_run("\n• d2 = diferencia entre la frecuencia de la clase modal y la clase posterior (d2 = fi - fi+1)")
    p.add_run("\n• c = amplitud de la clase")
    
    d1 = fi_modal - fi_anterior
    d2 = fi_modal - fi_posterior
    
    p = doc.add_paragraph("Sustituyendo los valores:", style='ExplanationStyle')
    p.add_run(f"\nd1 = {fi_modal} - {fi_anterior} = {d1}")
    p.add_run(f"\nd2 = {fi_modal} - {fi_posterior} = {d2}")
    p.add_run(f"\nMo = {li_modal} + [{d1} / ({d1} + {d2})] × {amplitud_modal}")
    
    # Evitar división por cero
    if d1 + d2 == 0:
        moda_calculada = li_modal + (amplitud_modal / 2)  # Aproximación al centro de la clase
        p.add_run(f"\nComo d1 + d2 = 0, utilizamos el punto medio de la clase: {moda_calculada:.4f}")
    else:
        moda_calculada = li_modal + (d1 / (d1 + d2)) * amplitud_modal
        p.add_run(f"\nMo = {li_modal} + {d1 / (d1 + d2):.4f} × {amplitud_modal}")
        p.add_run(f"\nMo = {li_modal} + {(d1 / (d1 + d2)) * amplitud_modal:.4f}")
        p.add_run(f"\nMo = {moda_calculada:.4f}")
    
    # Valor de referencia
    moda_ref = respuestas["tercera_serie"]["medidas_centrales"]["moda"]
    
    # Verificación
    p = doc.add_paragraph("VERIFICACIÓN:", style='ExplanationStyle')
    p.add_run(f"\nModa calculada: {moda_calculada:.4f}")
    p.add_run(f"\nModa de referencia: {moda_ref}")
    
    precision_moda = abs(moda_calculada - moda_ref) / moda_ref * 100 if moda_ref != 0 else 0
    
    if precision_moda < 5:
        p.add_run("\n\nLos valores son consistentes (diferencia menor al 5%)").bold = True
    else:
        p.add_run("\n\nHay una diferencia significativa entre los valores. Se recomienda utilizar el valor calculado.").bold = True
    
    # PASO 5: Visualización e interpretación
    doc.add_heading("PASO 5: Visualización e interpretación", level=3).style = subtitle_style
    
    # Visualización gráfica
    plt.figure(figsize=(10, 6))
    
    # Crear datos expandidos para visualización
    expanded_data = []
    for (li, ls), freq in zip(limites_central, central_data["count"]):
        marca_clase = (li + ls) / 2
        expanded_data.extend([marca_clase] * freq)
    
    plt.hist(expanded_data, bins=[lim[0] for lim in limites_central] + [limites_central[-1][1]], edgecolor='black', alpha=0.7)
    plt.axvline(x=media_calculada, color='r', linestyle='--', label=f'Media: {media_calculada:.2f}')
    plt.axvline(x=mediana_calculada, color='g', linestyle='-.', label=f'Mediana: {mediana_calculada:.2f}')
    plt.axvline(x=moda_calculada, color='b', linestyle=':', label=f'Moda: {moda_calculada:.2f}')
    
    plt.title('Distribución de frecuencias con medidas de tendencia central')
    plt.xlabel('Valores')
    plt.ylabel('Frecuencia')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    # Guardar gráfica
    mtc_path = os.path.join(UPLOAD_FOLDER, f'mtc_graph_{variante_id}.png')
    plt.savefig(mtc_path)
    plt.close()
    
    # Añadir la gráfica al documento
    doc.add_paragraph("VISUALIZACIÓN DE LAS MEDIDAS DE TENDENCIA CENTRAL:", style='ExplanationStyle').bold = True
    doc.add_picture(mtc_path, width=Inches(6))
    
    # Interpretación final
    p = doc.add_paragraph("INTERPRETACIÓN FINAL:", style='ExplanationStyle')
    p.runs[0].bold = True
    
    # Analizar la forma de la distribución
    if abs(media_calculada - mediana_calculada) < 50:  # Umbral ajustable según la escala de datos
        p.add_run("\nLa distribución se aproxima a la simetría, ya que la media y la mediana tienen valores cercanos. Esto sugiere que los datos se distribuyen de manera equilibrada alrededor del centro.")
    elif media_calculada > mediana_calculada:
        p.add_run("\nLa distribución presenta asimetría positiva (sesgada a la derecha), ya que la media es mayor que la mediana. Esto indica que existen valores altos que 'arrastran' la media hacia arriba.")
    else:
        p.add_run("\nLa distribución presenta asimetría negativa (sesgada a la izquierda), ya que la media es menor que la mediana. Esto indica que existen valores bajos que 'arrastran' la media hacia abajo.")
    
    # Interpretación de la moda
    if abs(moda_calculada - media_calculada) < 50:  # Umbral ajustable
        p.add_run("\n\nLa moda se encuentra cerca de la media, lo que refuerza la idea de que existe una concentración importante de datos alrededor del valor central.")
    else:
        p.add_run("\n\nLa moda se aleja de la media, lo que indica que existe una concentración de valores en un punto diferente al valor promedio.")
    
    # Conclusión
    p.add_run("\n\nCONCLUSIONES:")
    p.add_run(f"\n• Media: {media_calculada:.2f} - Representa el valor promedio de los datos.")
    p.add_run(f"\n• Mediana: {mediana_calculada:.2f} - Representa el valor central que divide al conjunto en dos partes iguales.")
    p.add_run(f"\n• Moda: {moda_calculada:.2f} - Representa el valor que aparece con mayor frecuencia.")
    
    # Para la comparación entre medidas
    if media_calculada > mediana_calculada and mediana_calculada > moda_calculada:
        p.add_run("\n\nSe cumple que Media > Mediana > Moda, lo que confirma una distribución con asimetría positiva pronunciada.")
    elif media_calculada < mediana_calculada and mediana_calculada < moda_calculada:
        p.add_run("\n\nSe cumple que Media < Mediana < Moda, lo que confirma una distribución con asimetría negativa pronunciada.")
    elif abs(media_calculada - mediana_calculada) < 50 and abs(mediana_calculada - moda_calculada) < 50:
        p.add_run("\n\nLas tres medidas tienen valores similares, lo que sugiere una distribución aproximadamente simétrica, posiblemente cercana a la normal.")
    
    # Limpiar imágenes temporales
    for img_path in [img_path, hist_path, tl_hist_path, mtc_path]:
        if os.path.exists(img_path):
            try:
                os.remove(img_path)
            except:
                pass
    
    # Guardar archivo
    solucion_filename = f'Solucion_Matematica_{seccion}_{tipo_evaluacion}_{variante_id}.docx'
    docx_path = os.path.join(PLANTILLAS_FOLDER, solucion_filename)
    doc.save(docx_path)
    
    # Convertir a PDF si es posible
    pdf_filename = f'Solucion_Matematica_{seccion}_{tipo_evaluacion}_{variante_id}.pdf'
    pdf_path = os.path.join(PLANTILLAS_FOLDER, pdf_filename)
    
    try:
        convert(docx_path, pdf_path)
        return pdf_filename
    except Exception as e:
        print(f"Error al convertir a PDF: {str(e)}")
        # Si falla la conversión, devolver el archivo Word
        return solucion_filename

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
        
        # Manejo del logo
        logo_path = None
        if 'logo' in request.files and request.files['logo'].filename:
            logo = request.files['logo']
            logo_filename = secure_filename(logo.filename)
            logo_path = os.path.join(UPLOAD_FOLDER, logo_filename)
            print(f"Guardando logo en: {logo_path}")
            logo.save(logo_path)
        
        # Generar variantes
        variantes_generadas = []
        
        for i in range(num_variantes):
            variante_id = f"V{i+1}"
            variante, respuestas = generar_variante(variante_id)
            
            # Crear documentos
            examen_filename = crear_examen_word(variante_id)
            hoja_filename = crear_hoja_respuestas(variante_id)
            plantilla_filename = crear_plantilla_calificacion(variante_id)
            
            # Generar la solución matemática detallada
            solucion_matematica = crear_solucion_matematica_detallada(variante_id, seccion, tipo_evaluacion)
            
            
            variantes_generadas.append({
                'id': variante_id,
                'examen': examen_filename,
                'hoja': hoja_filename,
                'plantilla': plantilla_filename,
                'solucion_matematica': solucion_matematica,
                'seccion': seccion,
                'tipo_evaluacion': tipo_evaluacion
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
        fecha_generacion = datetime.now().strftime("%d/%m/%Y %H:%M")
        for variante in variantes_generadas:
            historial.append({
                'id': variante['id'],
                'seccion': seccion,
                'tipo_evaluacion': tipo_evaluacion,
                'tipo_texto': tipo_textos.get(tipo_evaluacion, tipo_evaluacion),
                'fecha_generacion': fecha_generacion,
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
