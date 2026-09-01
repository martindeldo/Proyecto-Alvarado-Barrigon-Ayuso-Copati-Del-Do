# 🛒 Proyecto-Alvaredo-Barrigon-Ayuso-Copati-Del-Do

# 🛒 Sistema de Gestión de Stock, Caja y Recibos

Sistema de gestión de productos, stock y caja registradora desarrollado con **Python, Excel, Flask, JavaScript y Raspberry Pi**.

El proyecto permite administrar productos, controlar entradas y salidas de stock, consultar productos, detectar productos con poco stock, registrar movimientos y realizar ventas mediante un sistema de Caja.

Además, el sistema puede utilizar un **lector de códigos de barras conectado por USB** para identificar productos rápidamente.

La información del sistema se almacena en un archivo de **Excel**, mientras que **Flask** permite utilizar las diferentes funciones desde una página web.

Como complemento, el sistema de Caja puede permitir que el cliente reciba el **recibo de su compra mediante Gmail**, enviándolo a la dirección de correo electrónico proporcionada durante la compra.

---

# 📋 Índice

* [¿Qué hace este proyecto?](#-qué-hace-este-proyecto)
* [¿Qué necesitamos?](#-qué-necesitamos)
* [¿Cómo funciona el sistema?](#-cómo-funciona-el-sistema)
* [Preparar Ubuntu](#-1-preparar-ubuntu)
* [Instalar Python](#-2-instalar-python)
* [Crear el entorno virtual](#-3-crear-el-entorno-virtual)
* [Activar el entorno virtual](#-4-activar-el-entorno-virtual)
* [Instalar las librerías](#-5-instalar-las-librerías)
* [Instalar Flask](#-6-instalar-flask)
* [Instalar Visual Studio Code](#-7-instalar-visual-studio-code)
* [Abrir y modificar el proyecto](#-8-abrir-y-modificar-el-proyecto)
* [Ejecutar el sistema](#-9-ejecutar-el-sistema)
* [Raspberry Pi](#-10-preparar-la-raspberry-pi)
* [Instalar el sistema en la Raspberry Pi](#-11-instalar-el-sistema-en-la-raspberry-pi)
* [Lector de códigos de barras](#-12-conectar-el-lector-de-códigos-de-barras)
* [FileZilla](#-13-usar-filezilla)
* [VNC Viewer](#-14-usar-vnc-viewer)
* [Sistema de Caja](#-15-sistema-de-caja)
* [¿Cómo funciona la Caja?](#-16-cómo-funciona-la-caja)
* [Lector de códigos en Caja](#-17-lector-de-códigos-en-caja)
* [Buscar un producto](#-18-buscar-un-producto)
* [Carrito de compras](#-19-carrito-de-compras)
* [Cantidad de productos](#-20-cantidad-de-productos)
* [Cálculo del total](#-21-cálculo-del-total)
* [Eliminar productos del carrito](#-22-eliminar-productos-del-carrito)
* [Finalizar una compra](#-23-finalizar-una-compra)
* [Descuento automático del stock](#-24-descuento-automático-del-stock)
* [Registro de la venta](#-25-registro-de-la-venta)
* [Backup del Excel](#-26-backup-del-excel)
* [Recibo de compra](#-27-recibo-de-compra)
* [Enviar el recibo por Gmail](#-28-enviar-el-recibo-por-gmail)
* [Configurar Gmail](#-29-configurar-gmail)
* [Contraseña de aplicación de Google](#-30-contraseña-de-aplicación-de-google)
* [Funcionamiento del envío del recibo](#-31-funcionamiento-del-envío-del-recibo)
* [API de la Caja](#-32-api-de-la-caja)
* [JavaScript de la Caja](#-33-javascript-de-la-caja)
* [Flask y la Caja](#-34-flask-y-la-caja)
* [Diferencia entre Gestión y Caja](#-35-diferencia-entre-gestión-y-caja)
* [Flujo completo de una venta](#-36-flujo-completo-de-una-venta)
* [¿Cómo se conectan Python, Excel y Flask?](#-37-cómo-se-conectan-python-excel-y-flask)
* [Flujo completo del sistema](#-38-flujo-completo-del-sistema)
* [Problemas frecuentes](#-39-problemas-frecuentes)
* [Resumen](#-40-resumen)

---

# 📌 ¿Qué hace este proyecto?

Este proyecto fue creado para administrar el stock de productos y realizar ventas de manera sencilla.

En lugar de tener que controlar todo manualmente, el programa permite:

* Agregar stock.
* Quitar stock.
* Buscar productos.
* Consultar la cantidad disponible.
* Detectar productos con poco stock.
* Registrar los movimientos realizados.
* Utilizar códigos de barras.
* Guardar toda la información en un archivo de Excel.
* Acceder al sistema desde una página web.
* Utilizar una caja registradora.
* Agregar productos a un carrito.
* Calcular automáticamente el total de una compra.
* Descontar productos vendidos del stock.
* Registrar las ventas como movimientos de salida.
* Crear backups del archivo Excel.
* Ejecutar el sistema en una Raspberry Pi.
* Generar un recibo de compra.
* Enviar el recibo al correo electrónico del cliente.

El objetivo es que una persona pueda descargar el proyecto, instalar lo necesario y ponerlo a funcionar sin tener conocimientos avanzados de programación.

---

# 🧰 ¿Qué necesitamos?

## 💻 Para preparar el proyecto

Necesitamos una computadora con:

* Ubuntu.
* Conexión a Internet.
* Python.
* Visual Studio Code.
* Git, si vamos a descargar el proyecto desde GitHub.
* Las librerías de Python utilizadas por el proyecto.

---

# 📦 Librerías utilizadas

El proyecto utiliza principalmente:

| Librería         | ¿Para qué sirve?                                 |
| ---------------- | ------------------------------------------------ |
| `pandas`         | Ayuda a trabajar con los datos de Excel          |
| `openpyxl`       | Permite leer y modificar archivos `.xlsx`        |
| `python-barcode` | Permite crear códigos de barras                  |
| `Pillow`         | Permite trabajar con imágenes                    |
| `Flask`          | Permite crear la página web del sistema          |
| `smtplib`        | Permite enviar correos electrónicos desde Python |
| `email`          | Permite construir el contenido del correo        |

> `smtplib` y `email` forman parte de la biblioteca estándar de Python, por lo que no es necesario instalarlas mediante `pip`.

---

# 🔌 Elementos físicos

Si queremos utilizar el sistema completo con Raspberry Pi y lector de códigos de barras, necesitamos:

* Raspberry Pi 3 Model B.
* Fuente de alimentación compatible con la Raspberry Pi.
* Tarjeta microSD.
* Lector de códigos de barras USB.
* Cable USB para conectar el lector.
* Monitor, teclado y mouse para la configuración inicial de la Raspberry Pi, si no vamos a configurarla de forma remota.
* Cable de red Ethernet o conexión Wi-Fi.
* Una computadora para preparar y administrar el sistema.

---

# 🧠 ¿Cómo funciona el sistema?

Antes de instalar nada, es importante entender la idea general.

El sistema tiene varias partes que trabajan juntas:

```text
                 ┌───────────────┐
                 │     USUARIO   │
                 └───────┬───────┘
                         │
                         ▼
                 ┌───────────────┐
                 │     FLASK     │
                 │  Página web   │
                 └───────┬───────┘
                         │
             ┌───────────┴───────────┐
             │                       │
             ▼                       ▼
       ┌─────────────┐        ┌─────────────┐
       │   GESTIÓN   │        │    CAJA     │
       │   DE STOCK  │        │    VENTAS   │
       └──────┬──────┘        └──────┬──────┘
              │                      │
              └──────────┬───────────┘
                         ▼
                  ┌─────────────┐
                  │   PYTHON    │
                  │  stock.py   │
                  └──────┬──────┘
                         │
                         ▼
                  ┌─────────────┐
                  │    EXCEL    │
                  │   STOCK     │
                  └─────────────┘
                         │
                         ▼
                  ┌─────────────┐
                  │    GMAIL    │
                  │   RECIBO    │
                  └─────────────┘
```

En palabras simples:

**Flask muestra la página.**

**JavaScript controla las acciones de la Caja.**

**Python realiza las operaciones.**

**`stock.py` se encarga de trabajar con el stock.**

**Excel guarda la información.**

**Gmail permite enviar el recibo al cliente.**

---

# 🐧 1. Preparar Ubuntu

Este proyecto está pensado principalmente para funcionar en **Ubuntu**.

Lo primero que debemos hacer es abrir la terminal.

Podemos hacerlo presionando:

```text
Ctrl + Alt + T
```

Se abrirá una ventana donde podemos escribir comandos.

No hay que preocuparse por esto.

Los comandos son simplemente instrucciones que le damos a Ubuntu.

---

# 🐍 2. Instalar Python

Primero comprobamos si Python ya está instalado.

Escribimos:

```bash
python3 --version
```

Si aparece algo parecido a:

```text
Python 3.x.x
```

Python ya está instalado.

Si no aparece, podemos instalarlo con:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

Ubuntu puede pedir nuestra contraseña.

La escribimos y presionamos Enter.

> Importante: cuando escribimos la contraseña en la terminal, normalmente no aparecen letras ni símbolos. Esto es normal.

---

# 📁 3. Crear el entorno virtual

Un entorno virtual es una carpeta especial donde vamos a guardar las librerías que necesita este proyecto.

Esto evita modificar las librerías generales de Ubuntu.

Primero entramos a la carpeta del proyecto.

Por ejemplo:

```bash
cd Proyecto
```

Después creamos el entorno:

```bash
python3 -m venv .venv
```

Esto crea una carpeta llamada:

```text
.venv
```

---

# ▶️ 4. Activar el entorno virtual

Cada vez que vayamos a trabajar con el proyecto debemos activar el entorno.

En Ubuntu:

```bash
source .venv/bin/activate
```

Si funcionó correctamente, veremos algo parecido a:

```text
(.venv) usuario@ubuntu:~/Proyecto$
```

La palabra:

```text
(.venv)
```

indica que el entorno está activado.

---

# 📚 5. Instalar las librerías

Con el entorno virtual activado:

```bash
pip install pandas
```

Después:

```bash
pip install openpyxl
```

Después:

```bash
pip install python-barcode
```

Después:

```bash
pip install Pillow
```

Y finalmente:

```bash
pip install Flask
```

También podemos instalar todo junto:

```bash
pip install pandas openpyxl python-barcode Pillow Flask
```

---

# 🔎 ¿Para qué sirve cada librería?

## pandas

`pandas` nos ayuda a trabajar con información organizada.

En nuestro proyecto puede utilizarse para leer y manejar los datos de los productos.

---

## openpyxl

`openpyxl` permite que Python pueda trabajar directamente con archivos de Excel.

Gracias a esta librería podemos:

* Abrir un Excel.
* Leer productos.
* Modificar cantidades.
* Crear hojas.
* Guardar movimientos.
* Guardar cambios.

---

## python-barcode

Permite generar códigos de barras desde Python.

Por ejemplo:

```text
100000000067
```

puede convertirse en una imagen de código de barras.

---

## Pillow

Permite trabajar con imágenes.

Es importante porque los códigos de barras pueden guardarse como imágenes.

---

## Flask

Flask permite convertir nuestro programa en una página web.

Gracias a Flask podemos tener una página desde la cual:

* Ver productos.
* Consultar stock.
* Registrar entradas.
* Registrar salidas.
* Consultar movimientos.
* Acceder a la Caja.
* Realizar ventas.
* Enviar recibos.

---

# 🌐 6. Instalar Flask

Si todavía no lo instalamos:

```bash
pip install Flask
```

Podemos comprobarlo con:

```bash
pip show Flask
```

---

# 📝 7. Instalar Visual Studio Code

Para modificar el proyecto recomendamos utilizar **Visual Studio Code**.

No debemos confundirlo con Visual Studio.

Para este proyecto utilizamos:

**Visual Studio Code**

Podemos descargarlo desde:

https://code.visualstudio.com/

En Ubuntu podemos descargar el archivo `.deb`.

Después:

```bash
cd ~/Descargas
```

Y:

```bash
sudo apt install ./code_*.deb
```

---

# 💻 8. Abrir y modificar el proyecto

Entramos a la carpeta:

```bash
cd ~/Proyecto
```

Después:

```bash
code .
```

El punto significa:

> "Abrir la carpeta en la que estoy".

---

# 📂 Estructura del proyecto

La estructura puede ser similar a:

```text
Proyecto/
│
├── backups/
│
├── static/
│   ├── css/
│   │   ├── caja.css
│   │   ├── gestion.css
│   │   └── index.css
│   │
│   ├── js/
│   │   └── caja.js
│   │
│   └── imagenes/
│
├── templates/
│   ├── index.html
│   ├── login.html
│   ├── gestion.html
│   ├── caja.html
│   └── movimientos.html
│
├── productos_corregido.xlsx
├── stock.py
├── app.py
│
└── .venv/
```

Los nombres pueden cambiar dependiendo de la versión del proyecto.

La carpeta `.venv` no debería modificarse manualmente.

---

# ▶️ 9. Ejecutar el sistema

Primero:

```bash
cd ~/Proyecto
```

Activamos el entorno:

```bash
source .venv/bin/activate
```

Después ejecutamos Flask:

```bash
python3 app.py
```

Si Flask está correctamente configurado, aparecerá un mensaje indicando que el servidor está funcionando.

---

# 🌐 Entrar a la página

Cuando Flask está funcionando, normalmente podemos acceder desde:

```text
http://127.0.0.1:4000
```

En este proyecto Flask utiliza el puerto:

```text
4000
```

Si estamos utilizando la Raspberry Pi desde otro dispositivo de la misma red, podemos utilizar la dirección IP de la Raspberry.

Por ejemplo:

```text
http://192.168.1.100:4000
```

La dirección será diferente en cada red.

---

# 🍓 10. Preparar la Raspberry Pi

La Raspberry Pi es una computadora pequeña.

En nuestro proyecto utilizamos una:

**Raspberry Pi 3 Model B**

La idea es que la Raspberry Pi pueda ejecutar el sistema sin necesidad de tener una computadora grande funcionando permanentemente.

---

# 💾 11. Instalar el sistema en la Raspberry Pi

Necesitamos:

* Raspberry Pi 3 Model B.
* MicroSD.
* Computadora.
* Lector de tarjetas microSD.
* Fuente de alimentación.

Para instalar el sistema operativo podemos utilizar **Raspberry Pi Imager**.

Podemos descargarlo desde:

https://www.raspberrypi.com/software/

---

## Paso 1 — Colocar la microSD

Introducimos la microSD en nuestra computadora.

---

## Paso 2 — Abrir Raspberry Pi Imager

Abrimos Raspberry Pi Imager.

---

## Paso 3 — Elegir la Raspberry

Seleccionamos:

```text
Raspberry Pi 3
```

---

## Paso 4 — Elegir el sistema operativo

Seleccionamos:

```text
Raspberry Pi OS
```

---

## Paso 5 — Elegir la microSD

Seleccionamos nuestra tarjeta.

**Mucho cuidado en este paso.**

La tarjeta será borrada.

---

## Paso 6 — Escribir el sistema

Presionamos el botón para comenzar.

Cuando termine, retiramos la tarjeta de forma segura.

---

# 🔌 12. Conectar el lector de códigos de barras

El lector se conecta mediante USB.

El funcionamiento es:

```text
Código de barras
       ↓
Lector USB
       ↓
Raspberry Pi
       ↓
Sistema de Caja
```

La mayoría de lectores USB funcionan como un dispositivo HID.

Esto significa que el lector envía los números como si fueran escritos mediante un teclado.

Por ejemplo:

```text
100000000067
```

Normalmente también envía:

```text
ENTER
```

---

# 📦 13. Usar FileZilla

FileZilla permite mover archivos entre computadoras.

En nuestro caso:

```text
COMPUTADORA
     │
     │ FileZilla
     ▼
RASPBERRY PI
```

Podemos utilizarlo para copiar:

* Archivos Python.
* HTML.
* CSS.
* JavaScript.
* Imágenes.
* Excel.
* Otros archivos.

---

# 🖥️ 14. Usar VNC Viewer

VNC permite controlar la Raspberry Pi de forma remota.

```text
Nuestra computadora
        │
        │ VNC
        ▼
 Raspberry Pi
```

Podemos utilizar la Raspberry Pi sin tener que conectar permanentemente monitor, teclado y mouse.

---

# 🛒 15. Sistema de Caja

La Caja permite realizar una venta utilizando el lector de códigos.

El funcionamiento general es:

```text
             🛒 CAJA
                │
                ▼
       Escanear producto
                │
                ▼
       Buscar en Excel
                │
                ▼
        Mostrar producto
                │
                ▼
       Agregar al carrito
                │
                ▼
        Calcular subtotal
                │
                ▼
          Calcular total
                │
                ▼
       Finalizar compra
                │
                ▼
       Descontar stock
                │
                ▼
      Registrar movimiento
                │
                ▼
        Generar recibo
                │
                ▼
       Enviar por Gmail
```

---

# 🏪 ¿Cómo entrar a Caja?

Desde la página principal encontramos:

```text
📦 Gestión de productos

🛒 Caja
```

La Gestión está pensada para administrar el sistema.

La Caja está pensada para realizar ventas.

---

# 🔐 Diferencia entre administrador y Caja

La Gestión necesita iniciar sesión.

La Caja está destinada al proceso de venta.

```text
Página principal
       │
       ├───────────────┐
       │               │
       ▼               ▼
   Gestión            Caja
       │               │
       ▼               ▼
    Login          caja.html
```

---

# 🔎 16. ¿Cómo funciona la Caja?

La Caja contiene un campo donde se coloca el código.

Por ejemplo:

```text
┌─────────────────────────────────┐
│ Escanear código de barras       │
│                                 │
│ [ 100000000067              ]   │
└─────────────────────────────────┘
```

El lector introduce automáticamente el código.

JavaScript detecta el código y realiza una solicitud a Flask.

El recorrido es:

```text
Lector
   ↓
caja.js
   ↓
/api/producto/<codigo>
   ↓
Flask
   ↓
stock.py
   ↓
Excel
```

---

# 🏷️ 17. Lector de códigos en Caja

Cuando el lector escanea:

```text
100000000067
```

JavaScript recibe el código.

Normalmente espera el:

```text
ENTER
```

Cuando lo recibe, busca el producto.

---

# 🔎 18. Buscar un producto

JavaScript puede realizar:

```javascript
fetch(`/api/producto/${codigo}`)
```

Por ejemplo:

```text
/api/producto/100000000067
```

Flask recibe la solicitud y busca el producto.

Puede devolver:

```text
Código
Nombre
Precio
Stock
```

Por ejemplo:

```text
Código: 100000000067
Producto: Salsa Golf
Precio: $1500
Stock: 20
```

---

# 🛍️ 19. Carrito de compras

Cada producto escaneado se agrega a un carrito.

Ejemplo:

```text
┌───────────────────────────────────────────────────────┐
│ Código       Producto      Precio   Cant.   Subtotal │
├───────────────────────────────────────────────────────┤
│ 100000000067 Salsa Golf   $1500      2       $3000   │
│ 100000000002 Pizza        $2000      1       $2000   │
└───────────────────────────────────────────────────────┘
```

---

# 🔢 20. Cantidad de productos

Si escaneamos el mismo producto varias veces, la cantidad aumenta.

```text
Primer escaneo → Cantidad: 1
Segundo escaneo → Cantidad: 2
Tercer escaneo → Cantidad: 3
```

---

# 💰 21. Cálculo del total

El sistema realiza:

```text
Precio × Cantidad = Subtotal
```

Por ejemplo:

```text
$1500 × 2 = $3000
```

Si tenemos:

```text
Salsa Golf       $3000
Pizza            $2000
Gaseosa          $1500
```

Entonces:

```text
TOTAL = $6500
```

---

# ❌ 22. Eliminar productos del carrito

Cada producto tiene un botón:

```text
❌
```

Al presionarlo, se elimina del carrito y el total vuelve a calcularse.

---

# 🧾 23. Finalizar una compra

Cuando todos los productos están cargados:

```text
🧾 Finalizar compra
```

El sistema toma los productos del carrito y los envía a Flask.

Por ejemplo:

```text
Producto: Salsa Golf
Código: 100000000067
Cantidad: 2
Precio: 1500
```

Después Flask procesa la venta.

---

# 📉 24. Descuento automático del stock

Antes de modificar Excel, se comprueba que haya suficiente stock.

Por ejemplo:

```text
Stock disponible: 20
Cantidad comprada: 5
```

Entonces:

```text
20 - 5 = 15
```

Nuevo stock:

```text
15
```

---

# ⚠️ Stock insuficiente

Si:

```text
Stock disponible: 3
Cantidad comprada: 5
```

la venta no debe realizarse.

El sistema mostrará:

```text
Stock insuficiente.
Disponible: 3
```

---

# 📋 25. Registro de la venta

Cuando se realiza una venta, se registra como una operación:

```text
SALIDA
```

La hoja:

```text
Movimientos
```

puede contener:

| Fecha      | Código       | Producto   | Operación | Cantidad | Stock anterior | Stock resultante |
| ---------- | ------------ | ---------- | --------- | -------: | -------------: | ---------------: |
| 01/09/2026 | 100000000067 | Salsa Golf | SALIDA    |        2 |             20 |               18 |

---

# 💾 26. Backup del Excel

El sistema crea backups antes de realizar modificaciones importantes.

Los backups se almacenan en:

```text
backups/
```

Ejemplo:

```text
backups/
│
├── productos_2026-09-01_08-31-18.xlsx
├── productos_2026-09-01_09-15-22.xlsx
└── ...
```

Esto permite recuperar información si ocurre algún problema.

---

# 🧾 27. Recibo de compra

Además de actualizar el stock, el sistema puede generar un recibo con la información de la compra.

El recibo puede contener:

```text
================================
          SUPERMERCADO
================================

Fecha: 01/09/2026
Hora: 10:25

--------------------------------
PRODUCTOS
--------------------------------

Salsa Golf
Cantidad: 2
Precio: $1500
Subtotal: $3000

Pizza
Cantidad: 1
Precio: $2000
Subtotal: $2000

--------------------------------

TOTAL: $5000

================================
       GRACIAS POR SU COMPRA
================================
```

El recibo funciona como comprobante de la operación realizada dentro del sistema.

---

# 📧 28. Enviar el recibo por Gmail

Una de las funciones adicionales del sistema es permitir que el cliente reciba el recibo de su compra mediante correo electrónico.

El funcionamiento sería:

```text
Cliente realiza compra
          ↓
     Finaliza venta
          ↓
    Se actualiza Excel
          ↓
   Se registra movimiento
          ↓
    Se genera recibo
          ↓
  Cliente coloca su Gmail
          ↓
       Python
          ↓
     Servidor Gmail
          ↓
   Recibo enviado
          ↓
   📧 Cliente recibe
```

Por ejemplo, después de finalizar la compra, la Caja puede solicitar:

```text
📧 ¿Desea recibir su recibo por correo?

Correo electrónico:

[____________________________]

[ Enviar recibo ]
```

El cliente introduce su dirección de correo.

Por ejemplo:

```text
cliente@gmail.com
```

El sistema prepara el recibo y lo envía.

---

# 📄 ¿Qué recibe el cliente?

El correo puede tener como asunto:

```text
Recibo de compra - Sistema de Caja
```

Y dentro:

```text
Hola.

Gracias por su compra.

Fecha: 01/09/2026
Hora: 10:25

Productos:

Salsa Golf x2 ........ $3000
Pizza x1 ............. $2000

TOTAL: $5000

Gracias por su compra.
```

También podemos configurar el sistema para enviar el recibo como **archivo adjunto**, por ejemplo:

```text
recibo_2026-09-01_10-25.pdf
```

Esto permite que el cliente tenga una copia del comprobante.

---

# 📧 29. Configurar Gmail

Para enviar correos desde Python utilizando Gmail, no debemos colocar nuestra contraseña normal de Gmail directamente dentro del programa.

Google dispone de mecanismos de seguridad para permitir que aplicaciones utilicen una cuenta de correo.

Una opción habitual es utilizar una **contraseña de aplicación**.

El concepto sería:

```text
Cuenta Gmail
     │
     ▼
Contraseña de aplicación
     │
     ▼
Python
     │
     ▼
Servidor SMTP de Gmail
     │
     ▼
Correo del cliente
```

---

# 🔐 30. Contraseña de aplicación de Google

Para utilizar una contraseña de aplicación, la cuenta de Google debe tener activada la **verificación en dos pasos**.

La contraseña de aplicación es diferente de la contraseña normal de Gmail.

Por seguridad:

**NO debemos escribir la contraseña de Gmail directamente en el código.**

Tampoco debemos subirla a GitHub.

Nunca debemos hacer algo como:

```python
PASSWORD = "mi_contraseña"
```

porque estaríamos exponiendo nuestras credenciales.

---

# 🔒 Variables de entorno

Una forma más segura de guardar los datos utilizados para enviar el correo es utilizar variables de entorno.

Por ejemplo:

```text
EMAIL_USUARIO
EMAIL_PASSWORD
```

El programa puede leerlas desde Python.

Conceptualmente:

```python
import os

correo = os.getenv("EMAIL_USUARIO")
password = os.getenv("EMAIL_PASSWORD")
```

De esta manera, la contraseña no tiene que aparecer directamente dentro del código.

---

# ⚠️ IMPORTANTE PARA GITHUB

Nunca debemos subir:

```text
Contraseñas
Contraseñas de aplicación
Claves privadas
Tokens
Credenciales
```

al repositorio público.

Si utilizamos un archivo `.env`, debemos incluirlo en:

```text
.gitignore
```

Por ejemplo:

```text
.env
```

Esto evita que Git intente subirlo al repositorio.

---

# 📤 31. Funcionamiento del envío del recibo

El proceso completo puede ser:

```text
                 CLIENTE
                    │
                    ▼
              Realiza compra
                    │
                    ▼
                 CAJA
                    │
                    ▼
             Finalizar venta
                    │
                    ▼
                 FLASK
                    │
                    ▼
                Python
                    │
             ┌──────┴──────┐
             │             │
             ▼             ▼
           EXCEL        RECIBO
             │             │
             ▼             ▼
         Stock/       Contenido del
       Movimientos       correo
                           │
                           ▼
                     Servidor Gmail
                           │
                           ▼
                    📧 CLIENTE
```

---

# 🧩 ¿Qué hace Python en el envío?

Python puede encargarse de:

1. Recibir los datos de la venta.
2. Comprobar el stock.
3. Actualizar Excel.
4. Registrar el movimiento.
5. Crear el contenido del recibo.
6. Obtener el correo electrónico del cliente.
7. Conectarse al servidor de correo.
8. Enviar el recibo.

---

# 📬 SMTP

Para enviar el correo se utiliza el protocolo:

```text
SMTP
```

SMTP significa:

```text
Simple Mail Transfer Protocol
```

Es el protocolo utilizado para enviar correos electrónicos.

El funcionamiento puede representarse como:

```text
Python
  │
  ▼
SMTP de Gmail
  │
  ▼
Cuenta del sistema
  │
  ▼
Correo del cliente
```

Python puede utilizar `smtplib` para establecer la conexión.

---

# 📦 Librerías para el correo

No es necesario instalar `smtplib` ni `email`.

Forman parte de Python.

Por ejemplo:

```python
import smtplib

from email.message import EmailMessage
```

Estas herramientas permiten preparar y enviar el correo.

---

# 🧾 Ejemplo conceptual

El sistema puede crear un mensaje similar a:

```python
mensaje = EmailMessage()

mensaje["Subject"] = "Recibo de compra"
mensaje["From"] = correo
mensaje["To"] = correo_cliente

mensaje.set_content("""
Gracias por su compra.

Total: $5000

Gracias por elegirnos.
""")
```

Después Python puede conectarse al servidor SMTP y enviar el mensaje.

---

# 🔐 Seguridad del correo

La cuenta utilizada para enviar los recibos debería ser una cuenta destinada al sistema.

Por ejemplo:

```text
supermercado.sistema@gmail.com
```

No es recomendable utilizar una cuenta personal.

Además:

* No compartir la contraseña.
* No subir credenciales a GitHub.
* No escribir contraseñas directamente en `app.py`.
* Utilizar variables de entorno.
* Utilizar contraseña de aplicación cuando corresponda.
* Revocar la contraseña de aplicación si deja de utilizarse.

---

# 📎 Recibo como archivo adjunto

Una versión más completa del sistema puede generar un archivo de recibo.

Por ejemplo:

```text
recibo_100000000067.pdf
```

Después Python puede adjuntar ese archivo al correo.

El cliente recibiría:

```text
📧 Recibo de compra

Adjunto:
📄 recibo.pdf
```

Esto es especialmente útil si queremos que el cliente pueda guardar o imprimir su recibo.

---

# 🔄 Flujo completo del recibo

```text
1. Cliente escanea productos
             ↓
2. Productos entran al carrito
             ↓
3. Se calcula el total
             ↓
4. Cliente finaliza la compra
             ↓
5. Se comprueba el stock
             ↓
6. Se descuenta el stock
             ↓
7. Se registra la SALIDA
             ↓
8. Se guarda Excel
             ↓
9. Se genera el recibo
             ↓
10. Cliente introduce su correo
             ↓
11. Python prepara el correo
             ↓
12. Gmail envía el mensaje
             ↓
13. Cliente recibe el recibo
```

---

# 🌐 32. API de la Caja

Las principales rutas utilizadas pueden ser:

```text
/api/producto/<codigo>
```

y:

```text
/api/venta
```

También puede existir una ruta específica para el envío del recibo, por ejemplo:

```text
/api/enviar-recibo
```

La ruta exacta dependerá de cómo esté implementada la aplicación.

---

# 🔎 `/api/producto/<codigo>`

Sirve para buscar un producto.

Ejemplo:

```text
/api/producto/100000000067
```

Funcionamiento:

```text
caja.js
   ↓
Flask
   ↓
stock.buscar_producto()
   ↓
productos_corregido.xlsx
```

---

# 🧾 `/api/venta`

Esta ruta se utiliza para finalizar la compra.

Recibe los productos y sus cantidades.

Después:

```text
Flask
   ↓
stock.py
   ↓
comprobar stock
   ↓
registrar SALIDA
   ↓
guardar Excel
```

---

# 📧 `/api/enviar-recibo`

Si el sistema implementa una ruta específica para enviar el recibo, su función puede ser:

```text
Caja
  ↓
correo del cliente
  ↓
Flask
  ↓
Python
  ↓
Gmail
  ↓
cliente
```

Esta ruta debería recibir únicamente la información necesaria para realizar el envío.

---

# 🟨 33. JavaScript de la Caja

El archivo:

```text
static/js/caja.js
```

se encarga de controlar la parte interactiva de la Caja.

Entre sus funciones se encuentran:

* Detectar el lector.
* Detectar ENTER.
* Buscar productos.
* Agregar productos al carrito.
* Aumentar cantidades.
* Calcular subtotales.
* Calcular el total.
* Eliminar productos.
* Finalizar la compra.
* Solicitar el correo electrónico.
* Enviar la información del recibo.
* Mostrar mensajes.

---

# 🐍 34. Flask y la Caja

Flask se encuentra entre el navegador y Python.

El recorrido es:

```text
NAVEGADOR
    │
    ▼
 caja.js
    │
    ▼
  FLASK
    │
    ▼
 stock.py
    │
    ▼
 EXCEL
```

Para el correo:

```text
NAVEGADOR
    │
    ▼
 caja.js
    │
    ▼
  FLASK
    │
    ▼
 PYTHON
    │
    ▼
  GMAIL
    │
    ▼
 CLIENTE
```

---

# 📦 35. Diferencia entre Gestión y Caja

## 📦 Gestión

Está pensada para el administrador.

Permite:

* Ver productos.
* Buscar productos.
* Agregar stock.
* Quitar stock.
* Ver stock bajo.
* Ver movimientos.

El acceso está protegido mediante login.

---

## 🛒 Caja

Está pensada para realizar ventas.

Permite:

* Escanear productos.
* Buscar productos automáticamente.
* Agregar productos al carrito.
* Modificar cantidades.
* Calcular subtotales.
* Calcular el total.
* Eliminar productos.
* Finalizar compras.
* Descontar stock.
* Registrar la salida.
* Generar el recibo.
* Enviar el recibo por correo.

---

# 🔄 36. Flujo completo de una venta

Supongamos:

```text
Producto:
Salsa Golf

Código:
100000000067

Precio:
$1500

Stock:
20
```

El cliente compra:

```text
2 unidades
```

---

## Paso 1 — Abrir Caja

Entramos a:

```text
🛒 Caja
```

---

## Paso 2 — Escanear

Escaneamos:

```text
100000000067
```

---

## Paso 3 — Buscar

JavaScript realiza:

```text
/api/producto/100000000067
```

---

## Paso 4 — Consultar Excel

Python encuentra:

```text
Salsa Golf
$1500
Stock: 20
```

---

## Paso 5 — Agregar al carrito

```text
Salsa Golf
Cantidad: 1
Subtotal: $1500
```

---

## Paso 6 — Segundo escaneo

```text
Cantidad: 2
Subtotal: $3000
```

---

## Paso 7 — Finalizar

Presionamos:

```text
🧾 Finalizar compra
```

---

## Paso 8 — Comprobar stock

```text
Stock disponible: 20
Compra: 2
```

Hay suficiente stock.

---

## Paso 9 — Descontar

```text
20 - 2 = 18
```

---

## Paso 10 — Guardar Excel

```text
Salsa Golf
Stock: 18
```

---

## Paso 11 — Registrar movimiento

```text
Código: 100000000067
Producto: Salsa Golf
Operación: SALIDA
Cantidad: 2
Stock anterior: 20
Stock resultante: 18
```

---

## Paso 12 — Generar recibo

El sistema prepara:

```text
Salsa Golf x2 ........ $3000

TOTAL: $3000
```

---

## Paso 13 — Solicitar correo

La Caja puede mostrar:

```text
📧 ¿Desea recibir su recibo?

Correo:
[________________________]

[Enviar recibo]
```

---

## Paso 14 — Enviar

Python utiliza la cuenta configurada para enviar el correo.

```text
Sistema
   ↓
Gmail
   ↓
cliente@gmail.com
```

---

## Paso 15 — Cliente recibe el recibo

El cliente recibe:

```text
📧 Recibo de compra
```

y puede conservarlo en su correo.

---

# 🧩 37. ¿Cómo se conectan Python, Excel y Flask?

Cada parte tiene un trabajo diferente.

---

## 🟢 Excel = guarda la información

Excel guarda los productos.

Por ejemplo:

```text
Código          Producto        Cantidad       Precio
100000000001    Hamburguesa     20             2500
100000000002    Pizza           15             3000
100000000003    Gaseosa         30             1500
```

---

## 🔵 Python = hace el trabajo

Python realiza las operaciones.

Por ejemplo:

```text
Stock:
20

Venta:
5
```

Python realiza:

```text
20 - 5 = 15
```

También puede encargarse del envío del recibo.

---

## 🟠 Flask = conecta la página con Python

```text
Navegador
    ↓
Flask
    ↓
Python
    ↓
Excel
```

---

## 🟣 JavaScript = controla la Caja

JavaScript permite:

```text
Escanear
   ↓
Detectar código
   ↓
Buscar producto
   ↓
Mostrar producto
   ↓
Agregar al carrito
   ↓
Finalizar
```

---

## 📧 Gmail = entrega el recibo

Cuando el cliente solicita el recibo:

```text
Python
   ↓
SMTP
   ↓
Gmail
   ↓
Cliente
```

---

# 🔄 38. Flujo completo del sistema

```text
                         USUARIO
                            │
                            ▼
                       PÁGINA WEB
                          FLASK
                            │
              ┌─────────────┴─────────────┐
              │                           │
              ▼                           ▼
          GESTIÓN                       CAJA
              │                           │
              │                      caja.js
              │                           │
              │                           ▼
              │                    LECTOR USB
              │                           │
              └─────────────┬─────────────┘
                            ▼
                         stock.py
                            │
                            ▼
                   productos_corregido.xlsx
                            │
                  ┌─────────┴─────────┐
                  │                   │
                  ▼                   ▼
                STOCK             MOVIMIENTOS
                                     
                            CAJA
                              │
                              ▼
                           VENTA
                              │
                              ▼
                           RECIBO
                              │
                              ▼
                           GMAIL
                              │
                              ▼
                           CLIENTE
```

---

# 🏷️ Generación de códigos de barras

El proyecto puede utilizar Python para crear códigos de barras.

La librería:

```text
python-barcode
```

permite generar la imagen.

Por ejemplo:

```text
100000000067
```

puede convertirse en una imagen.

---

# 📑 Excel y la hoja "Movimientos"

El sistema utiliza una hoja llamada:

```text
Movimientos
```

Esta hoja permite guardar un historial.

Por ejemplo:

| Fecha      | Código       | Producto   | Operación | Cantidad | Stock anterior | Stock actual |
| ---------- | ------------ | ---------- | --------- | -------: | -------------: | -----------: |
| 01/09/2026 | 100000000067 | Salsa Golf | SALIDA    |        5 |             20 |           15 |
| 01/09/2026 | 100000000002 | Pizza      | ENTRADA   |       10 |             15 |           25 |

Las ventas realizadas desde Caja también pueden registrarse como:

```text
SALIDA
```

---

# ⚠️ 39. Problemas frecuentes

## "python3: command not found"

Probar:

```bash
sudo apt update
sudo apt install python3
```

---

## "pip: command not found"

```bash
sudo apt install python3-pip
```

---

## No puedo crear el entorno virtual

```bash
sudo apt install python3-venv
```

Después:

```bash
python3 -m venv .venv
```

---

## No encuentro Flask

Activar primero:

```bash
source .venv/bin/activate
```

Después:

```bash
pip install Flask
```

---

## No encuentro pandas

```bash
pip install pandas
```

---

## No encuentro openpyxl

```bash
pip install openpyxl
```

---

## No encuentro python-barcode

```bash
pip install python-barcode
```

---

## No encuentro Pillow

```bash
pip install Pillow
```

---

# 🛒 Problemas de Caja

## ❌ Error de conexión

Comprobar:

1. Que Flask esté ejecutándose.
2. Que el navegador pueda acceder al servidor.
3. Que exista `/api/producto/<codigo>`.
4. Que `stock.py` esté disponible.
5. Que exista el Excel.
6. Que el código exista.

---

## ❌ Producto no encontrado

Comprobar:

* Código correcto.
* Código existente en Excel.
* Columna `Código de serie`.
* Ausencia de espacios.
* Producto correctamente registrado.

---

## ❌ La Caja no detecta el lector

Comprobar:

1. USB conectado.
2. Sistema operativo reconoce el lector.
3. Cursor dentro del campo.
4. Lector envía ENTER.
5. Probar código manualmente.

---

## ❌ Stock insuficiente

Significa que la cantidad solicitada supera el stock disponible.

---

# 📧 Problemas con el correo

## ❌ El correo no se envía

Comprobar:

1. Que exista conexión a Internet.
2. Que la cuenta de Gmail esté correctamente configurada.
3. Que la verificación en dos pasos esté configurada cuando sea necesaria.
4. Que la contraseña de aplicación sea correcta.
5. Que las variables de entorno estén correctamente configuradas.
6. Que la dirección del cliente sea válida.
7. Que Flask esté ejecutándose.
8. Que la ruta encargada del envío exista.
9. Que Python pueda conectarse al servidor SMTP.

---

## ❌ "Authentication failed"

Este error normalmente indica que Gmail rechazó la autenticación.

Comprobar:

```text
Correo utilizado
       ↓
Contraseña de aplicación
       ↓
Configuración de Google
```

No debemos colocar la contraseña normal de Gmail si estamos utilizando una contraseña de aplicación.

---

## ❌ El correo llega a Spam

El correo puede terminar en la carpeta de Spam dependiendo de la configuración del destinatario y de otros factores.

Debemos comprobar:

```text
Bandeja de entrada
Spam
Correo no deseado
```

---

## ❌ El recibo no llega

Comprobar:

* Dirección del cliente.
* Conexión a Internet.
* Configuración de Gmail.
* Spam.
* Mensaje de error de Flask.
* Configuración de SMTP.

---

# 🔐 Seguridad

Nunca debemos publicar:

```text
Contraseñas
Contraseñas de aplicación
Tokens
Claves privadas
Credenciales
```

en GitHub.

Si utilizamos:

```text
.env
```

debemos agregarlo a:

```text
.gitignore
```

Por ejemplo:

```text
.env
```

---

# ❌ No se actualiza el Excel

Comprobar:

* Que `productos_corregido.xlsx` exista.
* Que el archivo no esté abierto.
* Que Python tenga permisos.
* Que exista la carpeta `backups/`.
* Que no haya otro proceso utilizando el archivo.

---

# ❗ Error importante: "sudo"

En Ubuntu podemos encontrarnos con:

```text
is not in the sudoers file
```

o:

```text
Permission denied
```

Esto significa que el usuario actual no tiene permisos de administrador.

El comando:

```bash
sudo
```

permite ejecutar determinadas órdenes con permisos administrativos.

Si el usuario no tiene permisos de administrador, `sudo` puede fallar.

No recomendamos modificar manualmente los permisos del sistema si no sabemos exactamente qué estamos haciendo.

---

# 🔄 Cada vez que volvamos a trabajar

Normalmente:

```bash
cd ~/Proyecto
```

Después:

```bash
source .venv/bin/activate
```

Y:

```bash
python3 app.py
```

Cuando terminemos:

```bash
deactivate
```

---

# 📦 ¿Cómo saber qué librerías tenemos?

Con el entorno activado:

```bash
pip list
```

Deberían aparecer:

```text
Flask
Pillow
openpyxl
pandas
python-barcode
```

---

# 🧪 Comprobar que todo funciona

Activamos:

```bash
source .venv/bin/activate
```

Después:

```bash
python3
```

Y:

```python
import pandas
import openpyxl
import barcode
from PIL import Image
import flask
import smtplib
from email.message import EmailMessage
```

Si no aparece ningún error, las librerías están disponibles.

Para salir:

```python
exit()
```

---

# 🚀 Instalación rápida

Si ya tenemos Ubuntu y Python:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

Entramos al proyecto:

```bash
cd ~/Proyecto
```

Creamos el entorno:

```bash
python3 -m venv .venv
```

Activamos:

```bash
source .venv/bin/activate
```

Instalamos:

```bash
pip install pandas openpyxl python-barcode Pillow Flask
```

Ejecutamos:

```bash
python3 app.py
```

---

# 🧾 Sistema completo de Caja

```text
┌──────────────────────────────┐
│       🛒 SISTEMA DE CAJA     │
└──────────────┬───────────────┘
               │
               ▼
       ESCANEAR PRODUCTO
               │
               ▼
          CÓDIGO DE BARRAS
               │
               ▼
            caja.js
               │
               ▼
        /api/producto
               │
               ▼
             FLASK
               │
               ▼
            stock.py
               │
               ▼
             EXCEL
               │
               ▼
       PRODUCTO + PRECIO
               │
               ▼
           🛒 CARRITO
               │
               ▼
        CALCULAR TOTAL
               │
               ▼
       FINALIZAR COMPRA
               │
               ▼
        COMPROBAR STOCK
               │
               ▼
        DESCONTAR STOCK
               │
               ▼
       REGISTRAR SALIDA
               │
               ▼
          GENERAR RECIBO
               │
               ▼
        SOLICITAR GMAIL
               │
               ▼
          ENVIAR RECIBO
               │
               ▼
             CLIENTE
```

---

# 🏁 40. Resumen

Para poner en funcionamiento el proyecto completo:

## En la computadora

1. Instalar Ubuntu.
2. Instalar Python.
3. Descargar el proyecto.
4. Crear el entorno virtual.
5. Activar el entorno virtual.
6. Instalar las librerías.
7. Instalar Visual Studio Code.
8. Abrir el proyecto.
9. Revisar la configuración.
10. Configurar el correo si se utilizará el envío de recibos.
11. Ejecutar Flask.
12. Abrir la página desde el navegador.

---

## En la Raspberry Pi

1. Conseguir una Raspberry Pi 3 Model B.
2. Conseguir una microSD.
3. Instalar Raspberry Pi OS.
4. Configurar Internet.
5. Copiar el proyecto.
6. Instalar Python.
7. Crear el entorno virtual.
8. Instalar las librerías.
9. Configurar el correo.
10. Conectar el lector de códigos de barras.
11. Ejecutar Flask.
12. Abrir el sistema desde el navegador.

---

## Para utilizar Gestión

1. Abrir la página.
2. Entrar a Gestión.
3. Iniciar sesión.
4. Consultar productos.
5. Registrar entradas y salidas.
6. Consultar stock bajo.
7. Consultar movimientos.

---

## Para utilizar Caja

1. Abrir la página.
2. Entrar a Caja.
3. Escanear un producto.
4. Esperar a que aparezca.
5. Escanear nuevamente si se necesitan más unidades.
6. Revisar el carrito.
7. Revisar el total.
8. Finalizar la compra.
9. Comprobar el stock.
10. Comprobar el movimiento.
11. Generar el recibo.
12. Introducir el correo del cliente.
13. Enviar el recibo por Gmail.

---

# 💡 Idea principal del proyecto

La idea del sistema puede resumirse en:

> **El lector identifica el producto, JavaScript controla la Caja, Flask conecta la página con Python, `stock.py` realiza las operaciones, Excel guarda la información y Gmail permite enviar el recibo al cliente.**

De esta manera tenemos un sistema que combina **hardware, software, gestión de stock, caja registradora y comunicación por correo electrónico**.

El flujo completo es:

```text
LECTOR
   ↓
JAVASCRIPT
   ↓
FLASK
   ↓
PYTHON
   ↓
EXCEL
   ↓
VENTA
   ↓
RECIBO
   ↓
GMAIL
   ↓
CLIENTE
```

---

# 📚 Enlaces útiles

### Python

https://www.python.org/

### Flask

https://flask.palletsprojects.com/

### Visual Studio Code

https://code.visualstudio.com/

### Raspberry Pi

https://www.raspberrypi.com/

### Raspberry Pi Imager

https://www.raspberrypi.com/software/

### FileZilla

https://filezilla-project.org/

### RealVNC

https://www.realvnc.com/

### Google Account

https://myaccount.google.com/

---

# 👥 Para nuevos usuarios

No es necesario saber programación para comenzar.

La recomendación es seguir esta guía **en orden**, sin saltear pasos.

Si aparece un error:

1. Leer el mensaje que aparece.
2. Comprobar que el entorno virtual esté activado.
3. Comprobar que las librerías estén instaladas.
4. Comprobar que estamos dentro de la carpeta correcta.
5. Comprobar que Flask esté ejecutándose.
6. Revisar la sección de problemas frecuentes.
7. Si el problema continúa, consultar el mensaje completo del error.

No recomendamos borrar archivos del sistema ni ejecutar comandos que no entendamos.

---

# 🎓 Objetivo del proyecto

Este proyecto fue pensado para aprender y, al mismo tiempo, crear una herramienta que pueda utilizarse en una situación real.

La intención de este README es que cualquier persona pueda descargar el proyecto y seguir los pasos necesarios para ponerlo en funcionamiento, incluso si nunca trabajó anteriormente con:

* Python.
* Flask.
* Raspberry Pi.
* Excel.
* Códigos de barras.
* JavaScript.
* Sistemas de Caja.
* Envío de correos electrónicos.

El resultado es un sistema completo capaz de administrar productos, controlar el stock, realizar ventas, registrar movimientos, generar recibos y enviarlos al cliente por correo electrónico.
