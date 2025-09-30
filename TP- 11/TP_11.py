import pandas as pd

# Ruta de tu archivo Excel
excel_file = r"movies.xls.ods" # <- ajustá según tu carpeta

# 1) Leer hojas del Excel
df_2000s = pd.read_excel(excel_file, sheet_name="2000s")  # xlrd para .xls, openpyxl para .xlsx
df_2010s = pd.read_excel(excel_file, sheet_name="2010s")
df_1900s = pd.read_excel(excel_file, sheet_name="1900s")

# Concatenar todas las décadas
df_movies = pd.concat([df_1900s, df_2000s, df_2010s], ignore_index=True)

print("Cantidad de filas 2000s:", len(df_2000s))
print("\nPrimeras filas de 2010s:")
print(df_2010s[["Title", "Year", "IMDB Score"]].head())
print("\nShape combinado:", df_movies.shape)

# 2) Primeras y últimas filas
print("\nPrimeras 5 filas de 2000s:")
print(df_2000s.head(5))

print("\nÚltimas 3 filas de 2010s:")
print(df_2010s.tail(3))

print("\nPrimeras 12 filas de 1900s:")
print(df_1900s.head(12))

# 3) Análisis por Country y Director
print("\nPelículas por país (2000s):")
print(df_2000s["Country"].value_counts())

print("\nTop 5 directores (2010s):")
print(df_2010s["Director"].value_counts().head(5))

print("\nPelículas por año (todas las décadas):")
print(df_movies["Year"].value_counts().sort_index())

# 4) Ordenar y filtrar
print("\nTop 5 por IMDB Score:")
print(df_movies.sort_values(by="IMDB Score", ascending=False).head(5))

print("\n10 con menos User Votes:")
print(df_movies.sort_values(by="User Votes", ascending=True).head(10))

print("\nOrdenado por Country y luego Year:")
print(df_movies.sort_values(by=["Country", "Year"], ascending=[True, True]).head(10))

# 5) Crear columna Decade
def get_decade(y):
    return f"{int(y)//10*10}s" if pd.notna(y) else "Unknown"

df_movies["Decade"] = df_movies["Year"].apply(get_decade)

print("\nPromedio de IMDB por Country:")
print(df_movies.groupby("Country")["IMDB Score"].mean().sort_values(ascending=False))

if {"Director", "Gross Earnings"}.issubset(df_movies.columns):
    print("\nTop 10 directores por Gross Earnings:")
    print(df_movies.groupby("Director")["Gross Earnings"].sum().sort_values(ascending=False).head(10))
else:
    print("\nTop 10 países por Gross Earnings:")
    print(df_movies.groupby("Country")["Gross Earnings"].sum().sort_values(ascending=False).head(10))

print("\nCantidad de títulos por Decade:")
print(df_movies.groupby("Decade")["Title"].count().sort_values(ascending=False))

# 6) Limpieza y estructura
print("\nEstructura general del DataFrame:")
print(df_movies.info())

df_clean = df_movies.dropna(subset=["IMDB Score"])
print("\nFilas originales vs limpias:", df_movies.shape, "=>", df_clean.shape)

df_clean["Country"] = df_clean["Country"].fillna("Unknown")
print("\nTop países después de fillna:")
print(df_clean["Country"].value_counts().head())

# 7) Guardar resultados en Excel
df_movies.to_excel("movies_completo.xlsx", index=False)

cols = [c for c in ["Title", "Year", "IMDB Score"] if c in df_movies.columns]
df_movies[cols].to_excel("movies_basico.xlsx", index=False)

prom_pais = df_movies.groupby("Country")["IMDB Score"].mean().reset_index().sort_values("IMDB Score", ascending=False)
prom_pais.to_excel("analisis_por_pais.xlsx", index=False)
