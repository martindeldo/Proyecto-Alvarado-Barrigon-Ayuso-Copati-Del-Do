import pandas as pd
import barcode
from barcode.writer import ImageWriter

datos = pd.read_excel("productos.xlsx")

codigo = barcode.get(
    "code128",
    "123456789",
    writer=ImageWriter()
)

codigo.save("codigos/codigo_barras")
