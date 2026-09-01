# Proyecto-Alvaredo-Barrigon-Ayuso-Copati-Del-Do

# 🛒 Sistema de Gestión de Stock y Caja

Sistema de gestión de productos, stock y caja registradora desarrollado con **Python, Excel, Flask y Raspberry Pi**.

El proyecto permite administrar productos, controlar entradas y salidas de stock, consultar productos, detectar productos con poco stock, registrar los movimientos realizados y realizar ventas mediante un sistema de caja.

Además, el sistema puede utilizar un **lector de códigos de barras conectado por USB** para identificar productos rápidamente.

La información del sistema se almacena en un archivo de **Excel**, mientras que Flask permite utilizar las diferentes funciones desde una página web.

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
* [API de la Caja](#-27-api-de-la-caja)
* [JavaScript de la Caja](#-28-javascript-de-la-caja)
* [Flask y la Caja](#-29-flask-y-la-caja)
* [Diferencia entre Gestión y Caja](#-30-diferencia-entre-gestión-y-caja)
* [Flujo completo de una venta](#-31-flujo-completo-de-una-venta)
* [¿Cómo se conectan Python, Excel y Flask?](#-32-cómo-se-conectan-python-excel-y-flask)
* [Flujo completo del sistema](#-33-flujo-completo-del-sistema)
* [Problemas frecuentes](#-34-problemas-frecuentes)
* [Resumen](#-35-resumen)

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

## 📦 Librerías utilizadas

El proyecto utiliza principalmente:

| Librería         | ¿Para qué sirve?                          |
| ---------------- | ----------------------------------------- |
| `pandas`         | Ayuda a trabajar con los datos de Excel   |
| `openpyxl`       | Permite leer y modificar archivos `.xlsx` |
| `python-barcode` | Permite crear códigos de barras           |
| `Pillow`         | Permite trabajar con imágenes             |
| `Flask`          | Permite crear la página web del sistema   |

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
       │   DE STOCK  │        │  REGISTROS  │
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
```

En palabras simples:

**Flask muestra la página.**

**JavaScript controla las acciones de la Caja.**

**Python realiza las operaciones.**

**`stock.py` se encarga de trabajar con el stock.**

**Excel guarda la información.**

---

# 🐧 1. Preparar Ubuntu

Este proyecto está pensado principalmente para funcionar en **Ubuntu**.

Lo primero que debemos hacer es abrir la terminal.

Podemos hacerlo presionando:

```text
Ctrl + Alt + T
```

Se abrirá una ventana negra donde podemos escribir comandos.

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

Python recomienda utilizar entornos separados para cada proyecto, especialmente en Linux, para evitar que las librerías de un proyecto interfieran con otros proyectos.

---

# 📁 3. Crear el entorno virtual

Un entorno virtual es simplemente una carpeta especial donde vamos a guardar las librerías que necesita este proyecto.

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

Dentro de ella estarán las herramientas que necesita nuestro proyecto.

---

# ▶️ 4. Activar el entorno virtual

Cada vez que vayamos a trabajar con el proyecto debemos activar el entorno.

En Ubuntu escribimos:

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

Con el entorno virtual activado, podemos instalar todas las librerías necesarias.

Ejecutamos:

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

Esperamos a que termine la instalación.

---

# 🔎 ¿Para qué sirve cada librería?

## pandas

`pandas` nos ayuda a trabajar con información organizada.

En nuestro proyecto puede utilizarse para leer y manejar los datos de los productos.

Por ejemplo:

```text
Producto       Cantidad
Hamburguesa    10
Pizza          5
Gaseosa        20
```

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

Nuestro archivo `.xlsx` es el lugar donde se guarda la información del sistema.

---

## python-barcode

Esta librería permite generar códigos de barras desde Python.

Por ejemplo:

```text
100000000067
```

puede convertirse en una imagen de código de barras.

Después esa imagen puede guardarse y utilizarse en el sistema.

---

## Pillow

`Pillow` permite trabajar con imágenes.

En nuestro caso es importante porque los códigos de barras que generamos pueden guardarse como imágenes.

---

## Flask

Flask es la parte que permite convertir nuestro programa en una página web.

Gracias a Flask podemos tener una página desde la cual:

* Ver productos.
* Consultar stock.
* Registrar entradas.
* Registrar salidas.
* Consultar movimientos.
* Acceder a la Caja.
* Realizar ventas.
* Utilizar el sistema desde un navegador.

---

# 🌐 6. Instalar Flask

Si todavía no lo instalamos:

```bash
pip install Flask
```

Podemos comprobar que está instalado con:

```bash
pip show Flask
```

Si aparece información sobre Flask, la instalación funcionó.

---

# 📝 7. Instalar Visual Studio Code

Para modificar el proyecto recomendamos utilizar **Visual Studio Code**.

No debemos confundirlo con Visual Studio.

Para este proyecto utilizamos:

**Visual Studio Code**

Es un programa que nos permite abrir y modificar los archivos del proyecto de manera cómoda.

Podemos descargarlo desde la página oficial:

https://code.visualstudio.com/

En Ubuntu podemos descargar el archivo `.deb`.

Después de descargarlo, podemos instalarlo desde la terminal.

Por ejemplo, si el archivo está en la carpeta Descargas:

```bash
cd ~/Descargas
```

Luego:

```bash
sudo apt install ./code_*.deb
```

---

# 💻 8. Abrir y modificar el proyecto

Una vez instalado Visual Studio Code podemos abrir la carpeta del proyecto.

Desde la terminal:

```bash
cd ~/Proyecto
```

Y después:

```bash
code .
```

El punto significa:

> "Abrí la carpeta en la que estoy".

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

Primero abrimos una terminal.

Entramos en la carpeta del proyecto:

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

Podemos copiar la dirección y abrirla en el navegador.

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

La página oficial de Raspberry Pi ofrece Raspberry Pi OS y Raspberry Pi Imager.

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

Seleccionamos nuestra tarjeta microSD.

**Mucho cuidado en este paso.**

La tarjeta será borrada.

Debemos asegurarnos de seleccionar la tarjeta correcta.

---

## Paso 6 — Escribir el sistema

Presionamos el botón para comenzar la instalación.

Raspberry Pi Imager descargará el sistema y lo colocará en la microSD.

Cuando termine, retiramos la tarjeta de forma segura.

---

# 🔌 12. Conectar el lector de códigos de barras

El lector de códigos de barras se conecta mediante USB.

La idea es:

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

Por ejemplo, si el código es:

```text
100000000067
```

el lector envía:

```text
100000000067
```

y normalmente termina enviando un:

```text
ENTER
```

Esto permite que JavaScript detecte automáticamente que terminó el escaneo.

---

# 📦 13. Usar FileZilla

FileZilla es un programa que permite mover archivos entre dos computadoras.

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
* Archivos HTML.
* Archivos CSS.
* JavaScript.
* Imágenes.
* Archivos de Excel.
* Otros archivos del proyecto.

---

# 🖥️ 14. Usar VNC Viewer

VNC permite ver y controlar la pantalla de otra computadora desde nuestro propio equipo.

En nuestro caso:

```text
Nuestra computadora
        │
        │ VNC
        ▼
 Raspberry Pi
```

Esto significa que podemos controlar la Raspberry Pi como si estuviéramos frente a ella.

---

# 🛒 15. Sistema de Caja

El sistema de Caja permite realizar una venta utilizando el lector de códigos de barras.

A diferencia de la sección de Gestión, donde administramos el stock, la Caja está pensada para el momento en el que un cliente realiza una compra.

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
       Finalizar la compra
                │
                ▼
       Descontar el stock
                │
                ▼
      Registrar movimiento
                │
                ▼
              EXCEL
```

---

# 🏪 ¿Cómo entrar a Caja?

Desde la página principal encontramos dos opciones:

```text
📦 Gestión de productos

🛒 Caja
```

La opción:

```text
📦 Gestión de productos
```

requiere iniciar sesión como administrador.

La opción:

```text
🛒 Caja
```

puede entrar directamente a la sección de Caja.

Esto permite separar las tareas administrativas de las tareas de venta.

---

# 🔐 Diferencia entre administrador y Caja

El sistema utiliza una sesión para proteger las funciones administrativas.

La Gestión de productos necesita que el usuario esté autenticado.

Por ejemplo:

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
       │
       ▼
   Gestión de
    productos
```

De esta manera:

**Gestión = administración del sistema.**

**Caja = realización de ventas.**

---

# 🔎 16. ¿Cómo funciona la Caja?

La página de Caja contiene un campo donde se coloca el código del producto.

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

Cuando el lector escanea un producto, envía el código al campo correspondiente.

Por ejemplo:

```text
100000000067
```

El JavaScript espera el:

```text
ENTER
```

Cuando lo recibe, toma el código y busca el producto.

La función utilizada es:

```javascript
buscarProducto(codigo);
```

El sistema no necesita que el usuario escriba manualmente el código.

Esto hace que el proceso de venta sea mucho más rápido.

---

# 🔎 18. Buscar un producto

Cuando se escanea un código, `caja.js` realiza una solicitud:

```javascript
fetch(`/api/producto/${codigo}`)
```

Por ejemplo:

```text
/api/producto/100000000067
```

Flask recibe esa solicitud.

Después utiliza:

```python
stock.buscar_producto(codigo)
```

para buscar el producto en Excel.

Si encuentra el producto, devuelve información como:

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

# ❌ Producto no encontrado

Si el código no existe en Excel, la Caja muestra:

```text
❌ Producto no encontrado
```

Esto evita agregar al carrito un producto que no está registrado.

---

# 🛍️ 19. Carrito de compras

Cada producto escaneado se agrega a un carrito.

El carrito se mantiene en JavaScript.

Por ejemplo:

```text
┌───────────────────────────────────────────────────────┐
│ Código       Producto      Precio   Cant.   Subtotal │
├───────────────────────────────────────────────────────┤
│ 100000000067 Salsa Golf   $1500      2       $3000   │
│ 100000000002 Pizza        $2000      1       $2000   │
└───────────────────────────────────────────────────────┘
```

El carrito permite acumular varios productos antes de finalizar la venta.

---

# 🔢 20. Cantidad de productos

Si escaneamos el mismo producto más de una vez, el sistema aumenta automáticamente la cantidad.

Por ejemplo:

Primer escaneo:

```text
Salsa Golf
Cantidad: 1
```

Segundo escaneo:

```text
Salsa Golf
Cantidad: 2
```

Tercer escaneo:

```text
Salsa Golf
Cantidad: 3
```

El producto no se agrega como tres filas diferentes.

Se mantiene como un único producto con cantidad:

```text
Salsa Golf
Cantidad: 3
```

---

# 💰 21. Cálculo del total

Cada producto tiene:

```text
Precio × Cantidad = Subtotal
```

Por ejemplo:

```text
Salsa Golf
$1500 × 2 = $3000
```

Si tenemos:

```text
Salsa Golf       $3000
Pizza            $2000
Gaseosa          $1500
```

el sistema calcula:

```text
TOTAL = $6500
```

El total se actualiza automáticamente cada vez que agregamos o eliminamos un producto.

---

# ❌ 22. Eliminar productos del carrito

Cada producto tiene un botón para eliminarlo:

```text
❌
```

Al presionarlo, el producto se elimina del carrito.

Después el sistema vuelve a calcular automáticamente el total.

Por ejemplo:

Antes:

```text
Salsa Golf     $3000
Pizza          $2000

TOTAL: $5000
```

Eliminamos Pizza:

```text
Salsa Golf     $3000

TOTAL: $3000
```

---

# 🧾 23. Finalizar una compra

Cuando todos los productos fueron agregados al carrito, se presiona:

```text
🧾 Finalizar compra e imprimir ticket
```

El sistema toma todos los productos del carrito.

Después envía la información a Flask mediante:

```text
/api/venta
```

La información enviada contiene los productos y sus cantidades.

Por ejemplo:

```text
Producto:
Salsa Golf

Código:
100000000067

Cantidad:
2

Precio:
1500
```

---

# 📉 24. Descuento automático del stock

Antes de modificar el Excel, el sistema debe comprobar que exista suficiente stock.

Por ejemplo:

```text
Stock disponible: 20
Cantidad comprada: 5
```

Entonces:

```text
20 - 5 = 15
```

El nuevo stock será:

```text
15
```

---

## ⚠️ Stock insuficiente

Si tenemos:

```text
Stock disponible: 3
```

y el cliente intenta comprar:

```text
Cantidad: 5
```

el sistema no debe permitir la venta.

Mostrará un mensaje indicando:

```text
Stock insuficiente.
Disponible: 3
```

Esto evita que el stock quede en valores negativos.

---

# 📋 25. Registro de la venta

Cuando una venta se realiza, se utiliza la función:

```python
stock.registrar_movimiento()
```

con la operación:

```text
SALIDA
```

Esto permite que las ventas de Caja formen parte del historial de movimientos.

La hoja:

```text
Movimientos
```

puede contener información como:

| Fecha      | Código       | Producto   | Operación | Cantidad | Stock anterior | Stock resultante |
| ---------- | ------------ | ---------- | --------- | -------: | -------------: | ---------------: |
| 01/09/2026 | 100000000067 | Salsa Golf | SALIDA    |        2 |             20 |               18 |

De esta manera podemos saber qué productos fueron vendidos y cómo cambió el stock.

---

# 💾 26. Backup del Excel

El sistema crea backups antes de realizar modificaciones importantes en el archivo Excel.

Los backups se almacenan en:

```text
backups/
```

Por ejemplo:

```text
backups/
│
├── productos_2026-09-01_08-31-18.xlsx
├── productos_2026-09-01_09-15-22.xlsx
└── ...
```

Esto permite recuperar información en caso de que ocurra algún problema.

---

# 🌐 27. API de la Caja

La Caja se comunica con Flask mediante rutas especiales llamadas API.

Las principales rutas utilizadas son:

```text
/api/producto/<codigo>
```

y:

```text
/api/venta
```

---

## 🔎 `/api/producto/<codigo>`

Esta ruta sirve para buscar un producto.

Ejemplo:

```text
/api/producto/100000000067
```

El funcionamiento es:

```text
caja.js
   ↓
Flask
   ↓
stock.buscar_producto()
   ↓
productos_corregido.xlsx
```

Si el producto existe, Flask devuelve sus datos.

---

## 🧾 `/api/venta`

Esta ruta se utiliza para finalizar una compra.

Recibe:

```text
Productos
Código
Cantidad
```

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

# 🟨 28. JavaScript de la Caja

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
* Mostrar mensajes.

---

## 🔄 Ejemplo del funcionamiento

Cuando se escanea:

```text
100000000067
```

JavaScript hace:

```javascript
buscarProducto("100000000067");
```

Después realiza:

```javascript
fetch(
    `/api/producto/100000000067`
);
```

Flask recibe la solicitud.

---

# 🐍 29. Flask y la Caja

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

Flask recibe las solicitudes de JavaScript y decide qué función de Python debe ejecutar.

---

# 📦 30. Diferencia entre Gestión y Caja

El proyecto tiene dos partes principales.

## 📦 Gestión

La Gestión está pensada para el administrador.

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

La Caja está pensada para realizar ventas.

Permite:

* Escanear productos.
* Buscar productos automáticamente.
* Agregar productos al carrito.
* Modificar cantidades mediante nuevos escaneos.
* Calcular subtotales.
* Calcular el total.
* Eliminar productos.
* Finalizar compras.
* Descontar stock.
* Registrar la salida en Movimientos.

---

# 🔄 31. Flujo completo de una venta

Supongamos que tenemos:

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

Flask recibe el código.

---

## Paso 4 — Consultar Excel

Python busca:

```text
100000000067
```

Encuentra:

```text
Salsa Golf
$1500
Stock: 20
```

---

## Paso 5 — Agregar al carrito

El producto aparece:

```text
Salsa Golf
Precio: $1500
Cantidad: 1
Subtotal: $1500
```

---

## Paso 6 — Segundo escaneo

Escaneamos nuevamente:

```text
100000000067
```

La cantidad aumenta:

```text
Cantidad: 2
```

El subtotal pasa a:

```text
$1500 × 2 = $3000
```

---

## Paso 7 — Finalizar

Presionamos:

```text
🧾 Finalizar compra
```

---

## Paso 8 — Comprobar stock

El sistema comprueba:

```text
Stock disponible: 20
Compra: 2
```

Hay suficiente stock.

---

## Paso 9 — Descontar

Python realiza:

```text
20 - 2 = 18
```

---

## Paso 10 — Guardar Excel

Excel queda:

```text
Salsa Golf
Stock: 18
```

---

## Paso 11 — Registrar movimiento

La hoja `Movimientos` registra:

```text
Código: 100000000067
Producto: Salsa Golf
Operación: SALIDA
Cantidad: 2
Stock anterior: 20
Stock resultante: 18
```

---

# 🧩 32. ¿Cómo se conectan Python, Excel y Flask?

Esta es una de las partes más importantes del proyecto.

Podemos imaginar que cada uno tiene un trabajo diferente.

---

## 🟢 Excel = guarda la información

Excel funciona como el lugar donde guardamos nuestros productos.

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

---

## 🟠 Flask = conecta la página con Python

Flask recibe las acciones realizadas desde la página.

Por ejemplo:

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

## 🟣 JavaScript = controla la interacción de Caja

JavaScript permite que la Caja responda rápidamente a las acciones del usuario.

Por ejemplo:

```text
Escanear
   ↓
JavaScript detecta ENTER
   ↓
Busca producto
   ↓
Muestra producto
   ↓
Agrega al carrito
```

---

# 🧠 33. Flujo completo del sistema

Podemos resumir todo el sistema de esta manera:

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
```

---

# 🏷️ Generación de códigos de barras

El proyecto también puede utilizar Python para crear códigos de barras.

La librería:

```text
python-barcode
```

se encarga de generar la imagen.

Por ejemplo:

```text
100000000067
```

puede convertirse en una imagen de código de barras.

Después esa imagen puede utilizarse para identificar el producto.

---

# 📑 Excel y la hoja "Movimientos"

El sistema utiliza Excel no solamente para guardar el stock.

También utiliza una hoja llamada:

```text
Movimientos
```

Esta hoja permite guardar un historial.

Por ejemplo:

| Fecha      | Código       | Producto   | Operación | Cantidad | Stock anterior | Stock actual |
| ---------- | ------------ | ---------- | --------- | -------: | -------------: | -----------: |
| 01/09/2026 | 100000000067 | Salsa Golf | SALIDA    |        5 |             20 |           15 |
| 01/09/2026 | 100000000002 | Pizza      | ENTRADA   |       10 |             15 |           25 |

De esta manera podemos saber qué ocurrió con el stock.

Las ventas realizadas desde Caja también pueden registrarse como:

```text
SALIDA
```

---

# ⚠️ 34. Problemas frecuentes

## "python3: command not found"

Probar:

```bash
sudo apt update
sudo apt install python3
```

---

## "pip: command not found"

Probar:

```bash
sudo apt install python3-pip
```

---

## No puedo crear el entorno virtual

Probar:

```bash
sudo apt install python3-venv
```

Después:

```bash
python3 -m venv .venv
```

---

## No encuentro Flask

Activar primero el entorno:

```bash
source .venv/bin/activate
```

Después:

```bash
pip install Flask
```

---

## No encuentro pandas

Con el entorno activado:

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

## ❌ "Error de conexión"

Si Caja muestra:

```text
❌ Error de conexión
```

cuando escaneamos un producto, debemos comprobar:

1. Que Flask esté ejecutándose.
2. Que el navegador pueda acceder al servidor.
3. Que exista la ruta:

```text
/api/producto/<codigo>
```

4. Que `stock.py` esté disponible.
5. Que exista `productos_corregido.xlsx`.
6. Que el código exista en Excel.

---

## ❌ "Producto no encontrado"

Si aparece:

```text
❌ Producto no encontrado
```

debemos comprobar:

* Que el código esté escrito correctamente.
* Que el código exista en Excel.
* Que esté en la columna `Código de serie`.
* Que no tenga espacios adicionales.
* Que el producto esté correctamente registrado.

---

## ❌ La Caja no detecta el lector

Comprobar:

1. Que el lector esté conectado por USB.
2. Que Ubuntu/Raspberry Pi lo reconozca.
3. Que el cursor esté dentro del campo de código.
4. Que el lector envíe ENTER después del código.
5. Probar escribiendo manualmente un código en el campo.

---

## ❌ La Caja encuentra el producto pero no lo agrega

Comprobar que la respuesta de Flask contenga:

```text
codigo
nombre
precio
stock
```

El JavaScript utiliza estos datos para crear el producto del carrito.

---

## ❌ "Stock insuficiente"

Esto significa que el sistema detectó que la cantidad solicitada es mayor que la cantidad disponible.

Por ejemplo:

```text
Stock: 2
Compra: 5
```

No se permite realizar la operación.

---

## ❌ No se actualiza el Excel

Comprobar:

* Que `productos_corregido.xlsx` exista.
* Que el archivo no esté abierto en LibreOffice/Excel.
* Que Python tenga permisos para modificarlo.
* Que la carpeta de backups exista.
* Que no haya otro proceso utilizando el archivo.

---

# ❗ Error importante: "sudo"

En Ubuntu podemos encontrarnos con errores como:

```text
is not in the sudoers file
```

o:

```text
Permission denied
```

Esto significa que el usuario actual no tiene permiso para realizar determinadas tareas de administración.

Si ocurre esto, no debemos intentar modificar archivos importantes del sistema sin saber qué estamos haciendo.

Lo recomendable es utilizar una cuenta de Ubuntu que tenga permisos de administrador.

---

# 🔄 Cada vez que volvamos a trabajar en el proyecto

No necesitamos instalar todo nuevamente.

Normalmente hacemos:

```bash
cd ~/Proyecto
```

Después:

```bash
source .venv/bin/activate
```

Y ejecutamos:

```bash
python3 app.py
```

Cuando terminemos:

```bash
deactivate
```

Esto cierra el entorno virtual.

---

# 📦 ¿Cómo saber qué librerías tenemos instaladas?

Con el entorno activado:

```bash
pip list
```

Podemos comprobar que aparezcan:

```text
Flask
Pillow
openpyxl
pandas
python-barcode
```

---

# 🧪 Comprobar que todo funciona

Podemos hacer una prueba sencilla.

Primero activamos el entorno:

```bash
source .venv/bin/activate
```

Después:

```bash
python3
```

Y dentro de Python:

```python
import pandas
import openpyxl
import barcode
from PIL import Image
import flask
```

Si no aparece ningún error, las librerías están disponibles.

Para salir:

```python
exit()
```

---

# 🚀 Instalación rápida

Si ya tenemos Ubuntu y Python instalados, podemos resumir la instalación en:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

Entramos en el proyecto:

```bash
cd ~/Proyecto
```

Creamos el entorno:

```bash
python3 -m venv .venv
```

Lo activamos:

```bash
source .venv/bin/activate
```

Instalamos todo:

```bash
pip install pandas openpyxl python-barcode Pillow Flask
```

Y ya podemos ejecutar:

```bash
python3 app.py
```

---

# 🧾 Sistema completo de Caja

El funcionamiento de Caja puede resumirse en:

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
             EXCEL
```

---

# 🏁 35. Resumen

Para poner en funcionamiento el proyecto completo:

### En la computadora

1. Instalar Ubuntu.
2. Instalar Python.
3. Descargar el proyecto.
4. Crear el entorno virtual.
5. Activar el entorno virtual.
6. Instalar las librerías.
7. Instalar Visual Studio Code.
8. Abrir el proyecto.
9. Revisar la configuración.
10. Ejecutar Flask.
11. Abrir la página desde el navegador.

### En la Raspberry Pi

1. Conseguir una Raspberry Pi 3 Model B.
2. Conseguir una microSD.
3. Instalar Raspberry Pi OS.
4. Configurar Internet.
5. Copiar el proyecto.
6. Instalar Python.
7. Crear el entorno virtual.
8. Instalar las librerías.
9. Conectar el lector de códigos de barras.
10. Ejecutar Flask.
11. Abrir el sistema desde el navegador.

### Para utilizar Gestión

1. Abrir la página.
2. Entrar a Gestión.
3. Iniciar sesión.
4. Consultar productos.
5. Registrar entradas y salidas.
6. Consultar stock bajo.
7. Consultar movimientos.

### Para utilizar Caja

1. Abrir la página.
2. Entrar a Caja.
3. Escanear un producto.
4. Esperar a que aparezca el producto.
5. Escanear nuevamente si se necesitan más unidades.
6. Revisar el carrito.
7. Revisar el total.
8. Finalizar la compra.
9. Comprobar que el stock se haya actualizado.
10. Comprobar el movimiento registrado.

---

# 💡 Idea principal del proyecto

La idea del sistema puede resumirse en:

> **El lector identifica el producto, JavaScript controla la Caja, Flask conecta la página con Python, `stock.py` realiza las operaciones y Excel guarda la información.**

De esta manera tenemos un sistema que combina hardware y software para facilitar la administración del stock y la realización de ventas.

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

Este proyecto fue pensado para aprender y, al mismo tiempo, crear una herramienta que pueda utilizarse en una situación real.

La intención de este README es que cualquier persona pueda descargar el proyecto y seguir los pasos necesarios para ponerlo en funcionamiento, incluso si nunca trabajó anteriormente con Python, Flask, Raspberry Pi, Excel o códigos de barras.
