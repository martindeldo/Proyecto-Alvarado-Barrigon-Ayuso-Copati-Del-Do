import pandas as pd

df = pd.read_excel("productos.xlsx")

print("Columnas encontradas:")
print(df.columns)

print("\nDatos:")
print(df)