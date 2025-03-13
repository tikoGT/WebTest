# Guía Rápida - Sistema Generador de Exámenes de Estadística

Esta guía le ayudará a utilizar el Sistema Generador de Exámenes de Estadística de manera eficiente. Siga estos pasos para generar sus exámenes, hojas de respuesta y plantillas de calificación.

## Índice

1. [Acceso al Sistema](#1-acceso-al-sistema)
2. [Generación de Exámenes](#2-generación-de-exámenes)
3. [Gestión de Variantes](#3-gestión-de-variantes)
4. [Impresión de Materiales](#4-impresión-de-materiales)
5. [Calificación de Exámenes](#5-calificación-de-exámenes)

## 1. Acceso al Sistema

1. Abra un navegador web (Chrome, Firefox, Safari, etc.)
2. Acceda a la URL: `http://localhost:5000` (si está ejecutando localmente)
3. Verá la página principal del sistema generador de exámenes.

## 2. Generación de Exámenes

Para generar nuevos exámenes:

1. En la sección "Generar Nuevos Exámenes", seleccione el número de variantes que desea crear (entre 1 y 10).
2. Haga clic en el botón "Generar Exámenes".
3. El sistema procesará la solicitud y creará las variantes con:
   - Preguntas aleatorias para la primera serie
   - Escenarios aleatorios para la segunda serie
   - Datos variados para los ejercicios prácticos

**Importante**: Cada vez que genera exámenes, se crean automáticamente los siguientes recursos para cada variante:
- Archivo Word del examen
- Hoja de respuestas en PDF
- Plantilla de calificación en PDF

## 3. Gestión de Variantes

Una vez generadas las variantes, aparecerán en la tabla "Variantes Generadas", donde podrá:

- **Descargar Examen**: Obtiene el archivo Word del examen para imprimir y distribuir a los estudiantes.
- **Descargar Hoja de Respuestas**: Obtiene el PDF con la hoja donde los estudiantes marcarán sus respuestas.
- **Descargar Plantilla de Calificación**: Obtiene el PDF con las respuestas correctas para facilitar la calificación.
- **Vista Previa**: Muestra el contenido del examen directamente en el navegador, incluyendo las respuestas correctas.
- **Descargar Todo**: Descarga un archivo ZIP con todos los recursos de la variante seleccionada.
- **Eliminar**: Borra la variante y todos sus archivos asociados.

## 4. Impresión de Materiales

Para imprimir los materiales:

1. **Exámenes**:
   - Descargue el archivo Word del examen.
   - Abra el archivo en Microsoft Word o una aplicación compatible.
   - Utilice la función de impresión del programa (recomendado: impresión a doble cara).

2. **Hojas de Respuesta**:
   - Descargue el archivo PDF de la hoja de respuestas.
   - Abra el archivo en Adobe Reader o cualquier visualizador de PDF.
   - Imprima el documento asegurándose de seleccionar "Tamaño real" o "100%" en las opciones de escala para mantener la alineación correcta de los círculos.

3. **Plantillas de Calificación**:
   - Descargue el archivo PDF de la plantilla de calificación.
   - Imprima solo la cantidad necesaria para el personal docente.
   - Guarde estas plantillas en un lugar seguro para mantener la integridad del examen.

## 5. Calificación de Exámenes

Para calificar los exámenes eficientemente:

1. Utilice la plantilla de calificación como guía.
2. Para la primera y segunda serie:
   - Las respuestas correctas están marcadas con círculos negros en la plantilla.
   - Compare las marcas del estudiante con la plantilla.
   - Cada respuesta correcta vale 4 puntos en la primera serie y 3 puntos en la segunda.

3. Para la tercera serie:
   - Revise los procedimientos y resultados de los estudiantes.
   - Compare los resultados finales con los valores indicados en la plantilla.
   - Cada ejercicio correcto vale 10 puntos.

4. La puntuación total es de 100 puntos:
   - Primera serie: 40 puntos (10 preguntas × 4 puntos)
   - Segunda serie: 20 puntos (6 escenarios × 3.33 puntos)
   - Tercera serie: 40 puntos (4 ejercicios × 10 puntos)

## Recomendaciones Adicionales

- **Respaldo**: Guarde copias de las variantes generadas y sus respuestas en caso de necesitar referencias futuras.
- **Variedad**: Genere nuevas variantes para cada grupo o para evaluaciones de recuperación.
- **Revisión**: Antes de distribuir, revise el contenido de los exámenes para asegurarse que cumplen con sus expectativas.

Si necesita asistencia adicional, consulte la documentación completa o contacte al soporte técnico.
