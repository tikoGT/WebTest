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

app = Flask(__name__)
app.secret_key = "estadisticabasica2024"

# Directorios de almacenamiento
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
VARIANTES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'variantes')
EXAMENES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'examenes')
PLANTILLAS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plantillas')
HOJAS_RESPUESTA_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hojas_respuesta')

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
    
    memory_file.seek(0)
    
    # Guardar el archivo ZIP temporalmente
    zip_path = os.path.join(UPLOAD_FOLDER, f'examen_completo_{id_examen}.zip')
    with open(zip_path, 'wb') as f:
        f.write(memory_file.getvalue())
    
    return send_from_directory(UPLOAD_FOLDER, f'examen_completo_{id_examen}.zip', as_attachment=True)

# Generador de exámenes
def generar_variante(variante_id="V1"):
    # Crear variante de Primera Serie (barajar preguntas)
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
def crear_examen_word(variante_id):
    # Cargar la variante
    with open(os.path.join(VARIANTES_FOLDER, f'variante_{variante_id}.json'), 'r', encoding='utf-8') as f:
        variante = json.load(f)
    
    # Crear documento Word
    doc = Document()
    
    # Configurar márgenes
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Título Universidad
    title = doc.add_heading('Universidad Panamericana', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Facultad y curso
    for header_text in ['Facultad de Humanidades', 'Estadística Básica', 'Ing. Marco Antonio Jiménez', '2025']:
        header = doc.add_paragraph()
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header.add_run(header_text)
        run.bold = True
        run.font.size = Pt(12)
    
    # Título del examen
    exam_title = doc.add_heading(f'Evaluación Parcial ({variante_id})', 1)
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
    
    # Guardar el documento
    doc.save(os.path.join(EXAMENES_FOLDER, f'Examen_{variante_id}.docx'))
    
    return f'Examen_{variante_id}.docx'

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

# Rutas de la aplicación
@app.route('/')
def index():
    # Listar variantes existentes
    variantes = []
    if os.path.exists(VARIANTES_FOLDER):
        for archivo in os.listdir(VARIANTES_FOLDER):
            if archivo.startswith('variante_') and archivo.endswith('.json'):
                variante_id = archivo.replace('variante_', '').replace('.json', '')
                
                # Verificar si existen los demás archivos
                tiene_examen = os.path.exists(os.path.join(EXAMENES_FOLDER, f'Examen_{variante_id}.docx'))
                tiene_hoja = os.path.exists(os.path.join(HOJAS_RESPUESTA_FOLDER, f'HojaRespuestas_{variante_id}.pdf'))
                tiene_plantilla = os.path.exists(os.path.join(PLANTILLAS_FOLDER, f'Plantilla_{variante_id}.pdf'))
                
                variantes.append({
                    'id': variante_id,
                    'tiene_examen': tiene_examen,
                    'tiene_hoja': tiene_hoja,
                    'tiene_plantilla': tiene_plantilla
                })
    
    return render_template('index.html', variantes=variantes)

@app.route('/generar_examen', methods=['POST'])
def generar_examen():
    num_variantes = int(request.form.get('num_variantes', 1))
    
    variantes_generadas = []
    
    for i in range(num_variantes):
        variante_id = f"V{i+1}"
        variante, respuestas = generar_variante(variante_id)
        
        # Crear documentos
        examen_filename = crear_examen_word(variante_id)
        hoja_filename = crear_hoja_respuestas(variante_id)
        plantilla_filename = crear_plantilla_calificacion(variante_id)
        
        variantes_generadas.append({
            'id': variante_id,
            'examen': examen_filename,
            'hoja': hoja_filename,
            'plantilla': plantilla_filename
        })
    
    flash(f'Se han generado {num_variantes} variantes de examen', 'success')
    return redirect(url_for('index'))

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
