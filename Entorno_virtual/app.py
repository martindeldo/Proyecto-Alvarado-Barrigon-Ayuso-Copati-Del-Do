from flask import Flask, render_template

app = Flask(__name__)

@app.route("/index")
def index():
    return render_template("index.html")

@app.route("/caja")
def caja():
    return render_template("caja.html")

@app.route("/producto")
def producto():
    return render_template("producto.html")

@app.route("/stock")
def stock():
    return render_template("stock.html")

@app.route("/login")
def login():
    return render_template("login.html")

if __name__ == "__main__":
    app.run(debug=True,host="0.0.0.0", port=4000)