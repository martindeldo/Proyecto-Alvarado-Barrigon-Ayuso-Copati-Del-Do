import os
from openpyxl import load_workbook
import barcode
from barcode.writer import ImageWriter


# ============================================================
# CONFIGURACIÓN
# ============================================================

archivo_excel = "productos_con_barras.xlsx"

carpeta_codigos = "codigos"

codigo_inicial = 100000000001


# ============================================================
# CREAR CARPETA
# ============================================================

os.makedirs(
    carpeta_codigos,
    exist_ok=True
)


# ============================================================
# ABRIR EXCEL
# ============================================================

libro = load_workbook(
    archivo_excel
)

hoja = libro.active


# ============================================================
# BUSCAR FILA DE ENCABEZADOS
# ============================================================

fila_encabezados = None

for fila in range(
    1,
    hoja.max_row + 1
):

    valores = []

    for columna in range(
        1,
        hoja.max_column + 1
    ):

        valor = hoja.cell(
            row=fila,
            column=columna
        ).value

        if valor is not None:

            valores.append(
                str(valor).strip()
            )


    if fila <= 5:

        print(
            f"Fila {fila}: {valores}"
        )


    if (
        "Comidas" in valores
        and "Código de serie" in valores
    ):

        fila_encabezados = fila

        break


# ============================================================
# COMPROBAR
# ============================================================

if fila_encabezados is None:

    raise Exception(
        "No se encontró la fila de encabezados."
    )


print()
print(
    f"Fila de encabezados: "
    f"{fila_encabezados}"
)
print()


# ============================================================
# BUSCAR COLUMNAS
# ============================================================

columna_producto = None
columna_precio = None
columna_cantidad = None
columna_codigo = None


for columna in range(
    1,
    hoja.max_column + 1
):

    valor = hoja.cell(
        row=fila_encabezados,
        column=columna
    ).value


    if valor is None:

        continue


    nombre = str(
        valor
    ).strip()


    if nombre == "Comidas":

        columna_producto = columna


    elif nombre == "valor":

        columna_precio = columna


    elif nombre == "Cantidad":

        columna_cantidad = columna


    elif nombre == "Código de serie":

        columna_codigo = columna


# ============================================================
# MOSTRAR
# ============================================================

print(
    "Columnas detectadas:"
)

print(
    "Producto:",
    columna_producto
)

print(
    "Precio:",
    columna_precio
)

print(
    "Cantidad:",
    columna_cantidad
)

print(
    "Código:",
    columna_codigo
)

print()


# ============================================================
# COMPROBAR COLUMNAS
# ============================================================

if columna_producto is None:

    raise Exception(
        "No se encontró la columna 'Comidas'."
    )


if columna_precio is None:

    raise Exception(
        "No se encontró la columna 'valor'."
    )


if columna_codigo is None:

    raise Exception(
        "No se encontró la columna 'Código de serie'."
    )


# ============================================================
# BUSCAR ÚLTIMO CÓDIGO
# ============================================================

ultimo_codigo = (
    codigo_inicial - 1
)


for fila in range(
    fila_encabezados + 1,
    hoja.max_row + 1
):

    valor_codigo = hoja.cell(
        row=fila,
        column=columna_codigo
    ).value


    if valor_codigo is None:

        continue


    try:

        numero = int(
            float(valor_codigo)
        )


        if numero > ultimo_codigo:

            ultimo_codigo = numero


    except (
        ValueError,
        TypeError
    ):

        continue


print(
    "Último código encontrado:",
    ultimo_codigo
)

print()


# ============================================================
# PROCESAR PRODUCTOS
# ============================================================

productos_procesados = 0
codigos_nuevos = 0


for fila in range(
    fila_encabezados + 1,
    hoja.max_row + 1
):

    producto = hoja.cell(
        row=fila,
        column=columna_producto
    ).value


    if producto is None:

        continue


    producto = str(
        producto
    ).strip()


    if producto == "":

        continue


    # ========================================================
    # LAS CATEGORÍAS NO RECIBEN CÓDIGO
    #
    # Si no tienen precio, las consideramos categorías.
    # ========================================================

    if columna_precio is not None:

        precio = hoja.cell(
            row=fila,
            column=columna_precio
        ).value


        if precio is None:

            print(
                f"{producto} -> categoría, sin código"
            )

            continue


    # ========================================================
    # OBTENER CÓDIGO
    # ========================================================

    celda_codigo = hoja.cell(
        row=fila,
        column=columna_codigo
    )


    codigo_existente = (
        celda_codigo.value
    )


    # ========================================================
    # CÓDIGO EXISTENTE
    # ========================================================

    if (
        codigo_existente is not None
        and str(
            codigo_existente
        ).strip() != ""
    ):

        codigo = str(
            codigo_existente
        ).strip()


        print(
            f"{producto} -> "
            f"código existente: {codigo}"
        )


    # ========================================================
    # CÓDIGO NUEVO
    # ========================================================

    else:

        ultimo_codigo += 1

        codigo = str(
            ultimo_codigo
        )


        celda_codigo.value = codigo

        codigos_nuevos += 1


        print(
            f"{producto} -> "
            f"NUEVO código: {codigo}"
        )


    productos_procesados += 1


    # ========================================================
    # CREAR CÓDIGO DE BARRAS
    # ========================================================

    ruta_imagen = os.path.join(
        carpeta_codigos,
        f"{codigo}.png"
    )


    if not os.path.exists(
        ruta_imagen
    ):

        codigo_barras = barcode.get(
            "code128",
            codigo,
            writer=ImageWriter()
        )


        codigo_barras.save(
            os.path.join(
                carpeta_codigos,
                codigo
            )
        )


        print(
            f"    Código de barras creado: "
            f"{codigo}.png"
        )


# ============================================================
# GUARDAR EXCEL
# ============================================================

libro.save(
    archivo_excel
)


# ============================================================
# RESULTADO
# ============================================================

print()
print(
    "=========================================="
)

print(
    "       CÓDIGOS ACTUALIZADOS"
)

print(
    "=========================================="
)

print()

print(
    f"Productos procesados: "
    f"{productos_procesados}"
)

print(
    f"Códigos nuevos: "
    f"{codigos_nuevos}"
)

print(
    f"Excel: {archivo_excel}"
)

print(
    f"Carpeta de códigos: "
    f"{carpeta_codigos}"
)

print()

print(
    "Los códigos existentes fueron conservados."
)

print(
    "Las categorías no recibieron códigos."
)

print(
    "Los códigos de barras se guardaron como PNG."
)

print()