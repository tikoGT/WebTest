#!/usr/bin/env python3
"""
Script de inicio para el Sistema Generador de Exámenes de Estadística
Universidad Panamericana - Facultad de Humanidades
2024
"""

import os
import sys
import logging
from app import app

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

# Verificar directorios necesarios
directories = [
    'examenes',
    'hojas_respuesta',
    'plantillas',
    'variantes'
]

for directory in directories:
    if not os.path.exists(directory):
        os.makedirs(directory)
        logger.info(f"Directorio creado: {directory}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    
    logger.info(f"Iniciando Sistema Generador de Exámenes en el puerto {port}")
    logger.info("URL de acceso: http://localhost:5000")
    
    app.run(host='0.0.0.0', port=port, debug=True)
