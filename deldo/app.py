from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session
)

import stock


# ============================================================
# FLASK
# ============================================================

app = Flask(__name__)

app.secret_key = "supermercado_clave_secreta"


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

    if "usuario" not in session:

        return redirect(
            url_for("login")
        )

    return render_template(
        "caja.html"
    )


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