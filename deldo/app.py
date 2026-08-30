
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    jsonify
)

import stock
import smtplib

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# ============================================================
# FLASK
# ============================================================

app = Flask(__name__)

app.secret_key = "supermercado_clave_secreta"


# ============================================================
# CONFIGURACIÓN DEL CORREO
# ============================================================

CORREO_EMISOR = "alumno29.deldo.martin@ipm.edu.ar"

CONTRASENA_APP = "Tru:=356"


# ============================================================
# INICIO
# ============================================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ============================================================
# LOGIN
# ============================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        usuario = request.form.get(
            "usuario",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        ).strip()

        if (
            usuario == "martindeldo"
            and password == "1234"
        ):

            session["usuario"] = usuario

            return redirect(
                url_for("gestion")
            )

        return render_template(
            "login.html",
            error="Usuario o contraseña incorrectos"
        )

    return render_template(
        "login.html"
    )


# ============================================================
# GESTIÓN
# ============================================================

@app.route("/gestion")
def gestion():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    productos = stock.obtener_productos()

    return render_template(
        "gestion.html",
        productos=productos
    )


# ============================================================
# ENTRADA DE STOCK
# ============================================================

@app.route(
    "/entrada-stock",
    methods=["POST"]
)
def entrada_stock():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    codigo = request.form.get(
        "codigo",
        ""
    ).strip()

    cantidad = request.form.get(
        "cantidad",
        ""
    ).strip()

    resultado = stock.registrar_movimiento(
        codigo,
        "ENTRADA",
        cantidad
    )

    if not resultado["ok"]:

        return render_template(
            "gestion.html",
            productos=stock.obtener_productos(),
            error=resultado["mensaje"]
        )

    return redirect(
        url_for("gestion")
    )


# ============================================================
# SALIDA DE STOCK
# ============================================================

@app.route(
    "/salida-stock",
    methods=["POST"]
)
def salida_stock():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    codigo = request.form.get(
        "codigo",
        ""
    ).strip()

    cantidad = request.form.get(
        "cantidad",
        ""
    ).strip()

    resultado = stock.registrar_movimiento(
        codigo,
        "SALIDA",
        cantidad
    )

    if not resultado["ok"]:

        return render_template(
            "gestion.html",
            productos=stock.obtener_productos(),
            error=resultado["mensaje"]
        )

    return redirect(
        url_for("gestion")
    )


# ============================================================
# BUSCAR PRODUCTO
# ============================================================

@app.route("/buscar")
def buscar():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    codigo = request.args.get(
        "codigo",
        ""
    ).strip()

    if codigo:

        producto = stock.buscar_producto(
            codigo
        )

        if producto is not None:

            productos = [
                producto
            ]

        else:

            productos = []

    else:

        productos = stock.obtener_productos()

    return render_template(
        "gestion.html",
        productos=productos,
        busqueda=codigo
    )


# ============================================================
# STOCK BAJO
# ============================================================

@app.route("/stock-bajo")
def stock_bajo():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    productos = stock.obtener_stock_bajo()

    return render_template(
        "gestion.html",
        productos=productos,
        titulo="Productos con stock bajo"
    )


# ============================================================
# MOVIMIENTOS
# ============================================================

@app.route("/movimientos")
def movimientos():

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    movimientos_lista = (
        stock.obtener_movimientos()
    )

    return render_template(
        "movimientos.html",
        movimientos=movimientos_lista
    )


# ============================================================
# CAJA
# ============================================================

@app.route("/caja")
def caja():

    return render_template(
        "caja.html"
    )


# ============================================================
# API - BUSCAR PRODUCTO
# ============================================================

@app.route(
    "/api/producto/<codigo>"
)
def api_producto(codigo):

    producto = stock.buscar_producto(
        codigo
    )

    if producto is None:

        return jsonify({
            "error":
                "Producto no encontrado"
        }), 404

    return jsonify({

        "codigo":
            producto["codigo"],

        "nombre":
            producto["producto"],

        "precio":
            producto["precio"],

        "stock":
            producto["stock"]

    })


# ============================================================
# API - FINALIZAR VENTA
# ============================================================

@app.route(
    "/api/venta",
    methods=["POST"]
)
def api_venta():

    datos = request.get_json(
        silent=True
    )


    if datos is None:

        return jsonify({
            "error":
                "No se recibieron datos"
        }), 400


    productos = datos.get(
        "productos",
        []
    )


    correo = datos.get(
        "correo",
        ""
    ).strip()


    # ========================================================
    # COMPROBAR CORREO
    # ========================================================

    if correo == "":

        return jsonify({
            "error":
                "Debe ingresar un correo electrónico"
        }), 400


    # ========================================================
    # VALIDAR STOCK ANTES DE MODIFICARLO
    # ========================================================

    for producto in productos:

        codigo = str(
            producto.get(
                "codigo",
                ""
            )
        ).strip()


        try:

            cantidad = int(
                producto.get(
                    "cantidad",
                    1
                )
            )

        except (
            ValueError,
            TypeError
        ):

            return jsonify({
                "error":
                    "Cantidad inválida"
            }), 400


        if codigo == "":

            return jsonify({
                "error":
                    "Hay un producto sin código"
            }), 400


        producto_excel = stock.buscar_producto(
            codigo
        )


        if producto_excel is None:

            return jsonify({

                "error":
                    f"Producto {codigo} no encontrado"

            }), 404


        stock_actual = int(
            producto_excel["stock"]
        )


        if cantidad > stock_actual:

            return jsonify({

                "error":
                    f"Stock insuficiente para "
                    f"{producto_excel['producto']}. "
                    f"Disponible: {stock_actual}"

            }), 400


    # ========================================================
    # DESCONTAR STOCK Y REGISTRAR MOVIMIENTOS
    # ========================================================

    movimientos_realizados = []


    for producto in productos:

        codigo = str(
            producto.get(
                "codigo",
                ""
            )
        ).strip()


        cantidad = int(
            producto.get(
                "cantidad",
                1
            )
        )


        resultado = stock.registrar_movimiento(
            codigo,
            "SALIDA",
            cantidad
        )


        if not resultado["ok"]:

            return jsonify({

                "error":
                    resultado["mensaje"]

            }), 400


        movimientos_realizados.append(
            resultado
        )


    # ========================================================
    # CALCULAR TOTAL
    # ========================================================

    total = 0


    for producto in productos:

        precio = float(
            producto.get(
                "precio",
                0
            )
        )


        cantidad = int(
            producto.get(
                "cantidad",
                1
            )
        )


        total += (
            precio *
            cantidad
        )


    # ========================================================
    # CREAR RECIBO
    # ========================================================

    mensaje = """
RECIBO DE TU COMPRA
==============================

Supermercado

Gracias por realizar tu compra.

PRODUCTOS
------------------------------
"""


    if productos:

        for producto in productos:

            nombre = producto.get(
                "nombre",
                "Producto"
            )


            precio = float(
                producto.get(
                    "precio",
                    0
                )
            )


            cantidad = int(
                producto.get(
                    "cantidad",
                    1
                )
            )


            subtotal = (
                precio *
                cantidad
            )


            mensaje += (
                f"{nombre}\n"
                f"Cantidad: {cantidad}\n"
                f"Precio: ${precio:.2f}\n"
                f"Subtotal: ${subtotal:.2f}\n"
                f"------------------------------\n"
            )


    else:

        mensaje += (
            "No hay productos registrados.\n"
            "------------------------------\n"
        )


    mensaje += f"""
TOTAL DE TU COMPRA
==============================

${total:.2f}

Gracias por elegir nuestro supermercado.

¡Que tengas un buen día!
"""


    # ========================================================
    # CREAR EMAIL
    # ========================================================

    email = MIMEMultipart()


    email["From"] = CORREO_EMISOR


    email["To"] = correo


    email["Subject"] = (
        "Recibo de tu compra en el supermercado"
    )


    email.attach(
        MIMEText(
            mensaje,
            "plain",
            "utf-8"
        )
    )


    # ========================================================
    # ENVIAR EMAIL
    # ========================================================

    try:

        servidor = smtplib.SMTP(
            "smtp-mail.outlook.com",
            587
        )


        servidor.ehlo()


        servidor.starttls()


        servidor.ehlo()


        servidor.login(
            CORREO_EMISOR,
            CONTRASENA_APP
        )


        servidor.send_message(
            email
        )


        servidor.quit()


    except Exception as error:

        print(
            "ERROR AL ENVIAR EL CORREO:"
        )

        print(
            repr(error)
        )

        return jsonify({

            "error":
                "No se pudo enviar el recibo por correo"

        }), 500


    # ========================================================
    # RESPUESTA
    # ========================================================

    return jsonify({

        "ok":
            True,

        "mensaje":
            "Compra finalizada correctamente",

        "correo":
            correo,

        "total":
            total

    })


# ============================================================
# CERRAR SESIÓN
# ============================================================

@app.route("/logout")
def logout():

    session.pop(
        "usuario",
        None
    )

    return redirect(
        url_for("index")
    )


# ============================================================
# SERVIDOR
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=4000,
        debug=True
    )
