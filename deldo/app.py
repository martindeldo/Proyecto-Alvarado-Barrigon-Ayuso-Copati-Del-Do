from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)

# Clave utilizada para proteger las sesiones
app.secret_key = "supermercado_clave_secreta"


@app.route("/")
def index():
    return render_template("index.html")


# =========================
# LOGIN
# =========================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        usuario = request.form["usuario"]
        password = request.form["password"]

        if usuario == "martindeldo" and password == "1234":

            session["usuario"] = usuario

            return redirect(url_for("gestion"))

        return render_template(
            "login.html",
            error="Usuario o contraseña incorrectos"
        )

    return render_template("login.html")


# =========================
# GESTIÓN DE PRODUCTOS
# =========================

@app.route("/gestion")
def gestion():

    # Si no inició sesión, lo mandamos al login
    if "usuario" not in session:
        return redirect(url_for("login"))

    return render_template("gestion.html")


# =========================
# CAJA
# =========================

@app.route("/caja")
def caja():
    return render_template("caja.html")


# =========================
# CERRAR SESIÓN
# =========================

@app.route("/logout")
def logout():

    session.pop("usuario", None)

    return redirect(url_for("index"))


# =========================
# INICIAR SERVIDOR
# =========================

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=4000,
        debug=True
    )