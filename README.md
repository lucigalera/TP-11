# TP-11
TRABAJO PRACTICO

Análisis de Películas con Pandas
Este proyecto usa Python + Pandas para leer y analizar un archivo de películas dividido en hojas por décadas (1900s, 2000s y 2010s).

Requisitos
- pip install pandas openpyxl

Qué hace el código:
- Lee las hojas del Excel y las une en un solo DataFrame.
- Muestra datos básicos: cantidad de filas, primeras y últimas películas, directores, países.
- Ordena y filtra: mejores puntajes IMDB, menos votos, etc.
- Crea la columna Decade para agrupar por década.
- Limpia datos (quita nulos y rellena países vacíos).
- Guarda resultados en archivos Excel (movies_completo.xlsx, movies_basico.xlsx, analisis_por_pais.xlsx).

Resultados
- Con este script podés ver cuántas películas hay por país, director, década, cuáles tienen mejor puntaje IMDB y exportar todo a Excel.
