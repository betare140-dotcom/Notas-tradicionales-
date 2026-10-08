# Extractor de medios tradicionales

Aplicación Streamlit para convertir archivos XLSX en un Word de notas. No requiere Gemini, claves de API ni pagos por uso.

## Windows

1. Instala Python 3.11 o superior con la opción de agregar Python al PATH.
2. Descomprime el ZIP en una carpeta.
3. Ejecuta Iniciar_Windows.bat. La primera ejecución instala las dependencias y necesita internet.
4. Carga el Excel, elige la hoja y verifica la fila de encabezados.
5. Si tienes una columna Resumen, Extracto o Contenido, selecciónala; es opcional.
6. Pulsa Obtener extractos desde los enlaces para las notas que no tienen texto. Revisa y corrige la columna Resumen.
7. Descarga el Word.

Mantén abierta la ventana de comandos mientras uses la aplicación. Para detenerla, presiona Ctrl+C.

## Instalación manual

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS o Linux
# source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Publicar en Streamlit

Sube app.py, extractor.py y requirements.txt a un repositorio de GitHub. Selecciona app.py como archivo principal en Streamlit Community Cloud y usa Python 3.11 o superior. No se necesitan secretos. Esta entrega contiene los archivos para publicar; no incluye un despliegue realizado.

## Excel y salida

El archivo de ejemplo tiene la hoja Noticias, encabezados en la fila 5 y las columnas Fecha, Portal, Título descriptivo y URL. La app detecta esos encabezados y permite cambiarlos manualmente. Solo lee la hoja seleccionada. Admite enlaces en texto o hipervínculos de celda, y fechas de Excel, ISO, día/mes/año y fechas como 8 de octubre de 2026.

Todo el Word usa Verdana de 12 puntos e interlineado sencillo. Cada nota es un solo párrafo, con saltos manuales de línea equivalentes a Shift+Enter entre medio, título, extracto y URL. Hay separación entre notas. El Word muestra cada fecha una sola vez en azul. Debajo agrupa todas sus notas: medio en negritas, título normal, extracto normal y URL activa azul y subrayada, sin etiquetas. Conserva el título del Excel. Los datos faltantes se muestran para revisión y no se excluyen automáticamente. Las fechas no reconocidas mantienen el texto original y van al final si se ordena por fecha. El filtro de medios sirve para elegir qué fuentes incluir; no clasifica por sí mismo los canales de un archivo mixto.

Puedes filtrar medios y fechas, elegir el orden, quitar duplicados exactos opcionalmente y excluir filas individuales. Por defecto conserva todas las filas no vacías de las cuatro columnas seleccionadas. En Orden del Excel, las fechas se agrupan en el orden de su primera aparición, y las notas de cada fecha conservan su orden relativo.

## Extractos

La aplicación usa una columna de resumen o extracto del Excel si existe. Si eliges Contenido o Texto, selecciona hasta dos oraciones del texto, aproximadamente 75 palabras, relacionadas con el título. El botón de extracción consulta la noticia enlazada y busca el cuerpo del artículo o la descripción que publica el portal. Selecciona oraciones existentes, no redacta con IA. El resultado es una propuesta que debes revisar para confirmar el punto importante; la selección automática puede omitir matices.

Los extractos se pueden editar o pegar directamente en la tabla. El botón solo completa notas incluidas sin extracto, para conservar tus correcciones. Si un portal bloquea la lectura, exige acceso o no ofrece texto reconocible, se muestra el problema. No usa el título como sustituto del contenido. Al exportar una nota sin resumen se indica Extracto pendiente de completar. El estado técnico solo aparece en la interfaz.

La lectura de enlaces necesita internet. Las correcciones se mantienen durante la sesión actual de Streamlit; descarga el Word antes de cerrar o reiniciar la aplicación.

El archivo de ejemplo se genera con las 30 notas del Excel suministrado, en orden de fecha ascendente.
