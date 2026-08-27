# Proyecto-Alvaredo-Barrigon-Ayuso-Copati-Del-Do

# 🛒 Sistema de Gestión de Stock

Sistema de gestión de productos y stock desarrollado con **Python, Excel, Flask y Raspberry Pi**.

El proyecto permite administrar productos, controlar entradas y salidas de stock, consultar productos, detectar productos con poco stock y registrar los movimientos realizados.

Además, el sistema puede utilizar un **lector de códigos de barras** conectado por USB para buscar productos de forma rápida.

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
* [¿Cómo se conectan Python, Excel y Flask?](#-15-cómo-funciona-python-excel-y-flask-juntos)
* [Flujo completo del sistema](#-16-flujo-completo-del-sistema)
* [Problemas frecuentes](#-17-problemas-frecuentes)
* [Resumen](#-18-resumen)

---

# 📌 ¿Qué hace este proyecto?

Este proyecto fue creado para administrar el stock de productos de manera sencilla.

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
* Lector de códigos de barras **Metrologic MS7120**.
* Cable USB para conectar el lector.
* Monitor, teclado y mouse para la configuración inicial de la Raspberry Pi, si no vamos a configurarla de forma remota.
* Cable de red Ethernet o conexión Wi-Fi.
* Una computadora para preparar y administrar el sistema.

---

# 🧠 ¿Cómo funciona el sistema?

Antes de instalar nada, es importante entender la idea general.

El sistema tiene tres partes principales:

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
                         ▼
                 ┌───────────────┐
                 │    PYTHON     │
                 │ Lógica del    │
                 │    sistema    │
                 └───────┬───────┘
                         │
                         ▼
                 ┌───────────────┐
                 │     EXCEL     │
                 │ Información   │
                 │    del stock  │
                 └───────────────┘
```

En palabras simples:

**Flask muestra la página.**

**Python hace el trabajo.**

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

Flask recomienda este método para mantener separadas las librerías de cada proyecto.

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

Flask actualmente requiere Python 3.9 o una versión posterior.

---

# 📝 7. Instalar Visual Studio Code

Para modificar el proyecto recomendamos utilizar **Visual Studio Code**.

No debemos confundirlo con Visual Studio.

Para este proyecto utilizamos:

**Visual Studio Code**

Es un programa que nos permite abrir y modificar los archivos del proyecto de manera cómoda.

Podemos descargarlo desde la página oficial:

[Descargar Visual Studio Code](https://code.visualstudio.com/download?utm_source=chatgpt.com)

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

La documentación oficial de Visual Studio Code recomienda el paquete `.deb` para Ubuntu y otras distribuciones basadas en Debian.

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
├── excel/
│   └── productos_corregido.xlsx
│
├── static/
│   ├── css/
│   ├── imagenes/
│   └── ...
│
├── templates/
│   ├── gestion.html
│   └── ...
│
├── stock.py
├── barritas.py
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

Después ejecutamos el programa correspondiente.

Por ejemplo:

```bash
python3 stock.py
```

Si utilizamos Flask, normalmente ejecutaremos el archivo principal de Flask.

Por ejemplo:

```bash
python3 app.py
```

El comando exacto dependerá del archivo principal incluido en el proyecto.

---

# 🌐 Entrar a la página

Cuando Flask está funcionando, normalmente aparece una dirección parecida a:

```text
http://127.0.0.1:5000
```

Podemos copiar esa dirección y abrirla en el navegador.

También puede aparecer una dirección de red, por ejemplo:

```text
http://192.168.1.100:5000
```

Esta segunda dirección puede permitir que otros dispositivos de la misma red accedan al sistema, dependiendo de cómo esté configurado Flask.

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

[Raspberry Pi — Software oficial](https://www.raspberrypi.com/software/?utm_source=chatgpt.com)

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

Para la Raspberry Pi 3 también existen versiones de 32 y 64 bits; la página oficial indica específicamente la compatibilidad de Raspberry Pi 3 con Raspberry Pi OS.

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

# 🔌 12. Conectar la Raspberry Pi

Colocamos la microSD en la Raspberry Pi.

Después conectamos:

```text
                 ┌───────────────────┐
                 │   Raspberry Pi 3  │
                 │                   │
                 │     microSD       │
                 │        ↓          │
                 │   Sistema         │
                 │                   │
                 └─────────┬─────────┘
                           │
                           ▼
                       Internet
```

La Raspberry necesita alimentación eléctrica.

Podemos conectarla mediante su puerto de alimentación correspondiente.

---

# 📡 Conectar a Internet

La Raspberry Pi puede conectarse mediante:

* Wi-Fi.
* Cable de red Ethernet.

Para este proyecto recomendamos utilizar una conexión estable.

Si vamos a acceder a la Raspberry desde otra computadora, necesitamos saber su dirección dentro de la red.

Por ejemplo:

```text
192.168.1.100
```

La dirección será diferente en cada red.

---

# 📦 13. Pasar el proyecto a la Raspberry

Una vez que tenemos la Raspberry funcionando, necesitamos copiar nuestro proyecto desde la computadora hacia ella.

Una opción sencilla es utilizar:

**FileZilla**

---

# 📁 ¿Qué es FileZilla?

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
* Imágenes.
* Archivos de Excel.
* Otros archivos del proyecto.

---

# 📥 Instalar FileZilla

Podemos descargar FileZilla desde su página oficial:

[FileZilla](https://filezilla-project.org/?utm_source=chatgpt.com)

Buscamos **FileZilla Client**.

---

# 🔐 Conectar FileZilla con la Raspberry

La forma recomendada es utilizar una conexión segura mediante SSH/SFTP.

En FileZilla tendremos que indicar los datos de nuestra Raspberry.

Normalmente necesitaremos:

```text
Servidor: dirección de la Raspberry
Usuario: usuario de la Raspberry
Contraseña: contraseña de la Raspberry
Puerto: 22
```

Por ejemplo:

```text
Servidor: 192.168.1.100
Usuario: usuario
Contraseña: ********
Puerto: 22
```

El usuario, contraseña y dirección dependen de la configuración que hayamos elegido.

---

# 📂 ¿Cómo se utiliza?

FileZilla normalmente muestra dos lados:

```text
┌──────────────────────┬──────────────────────┐
│   NUESTRA PC         │    RASPBERRY PI      │
│                      │                      │
│ Proyecto             │ home                 │
│ stock.py             │ Proyecto             │
│ app.py               │ stock.py             │
│ productos.xlsx       │ app.py               │
│                      │ productos.xlsx       │
└──────────────────────┴──────────────────────┘
```

Podemos arrastrar archivos de un lado hacia el otro.

Así podemos copiar nuestro proyecto a la Raspberry.

---

# 🖥️ 14. Usar VNC Viewer

Otra herramienta muy útil es **VNC Viewer**.

VNC permite ver y controlar la pantalla de otra computadora desde nuestro propio equipo.

En nuestro caso:

```text
Nuestra computadora
        │
        │ VNC
        ▼
 Raspberry Pi
```

Esto significa que podemos controlar la Raspberry Pi como si estuviéramos sentados frente a ella.

---

# 📥 Descargar VNC Viewer

Podemos descargarlo desde la página oficial de RealVNC:

[Descargar RealVNC Viewer](https://www.realvnc.com/es/connect/download/viewer/?utm_source=chatgpt.com)

RealVNC ofrece versiones para diferentes sistemas operativos y también herramientas para Raspberry Pi.

---

# ⚠️ Importante sobre VNC

Para poder controlar una Raspberry mediante VNC necesitamos tener configurado el acceso remoto en la Raspberry.

Actualmente RealVNC también ofrece **RealVNC Connect**, que combina las herramientas necesarias para establecer este tipo de conexión.

La idea general es:

```text
Raspberry Pi
     │
     │ VNC Server
     │
     ▼
   Internet
     │
     ▼
Nuestra PC
     │
     │ VNC Viewer
     ▼
Pantalla de Raspberry
```

---

# 📊 15. ¿Cómo funciona Python, Excel y Flask juntos?

Esta es una de las partes más importantes del proyecto.

Podemos imaginar que cada uno tiene un trabajo diferente.

---

## 🟢 Excel = guarda la información

Excel funciona como el lugar donde guardamos nuestros productos.

Por ejemplo:

```text
Código          Producto        Cantidad
100000000001    Hamburguesa     20
100000000002    Pizza           15
100000000003    Gaseosa         30
```

El archivo contiene la información que necesita el sistema.

---

## 🔵 Python = hace el trabajo

Python es quien se encarga de realizar las operaciones.

Por ejemplo:

Si tenemos:

```text
Hamburguesa
Stock: 20
```

y sacamos:

```text
5
```

Python realiza:

```text
20 - 5 = 15
```

Y después guarda el nuevo valor.

---

## 🟠 Flask = muestra todo en una página

Flask permite que podamos utilizar el sistema desde un navegador.

Por ejemplo:

```text
              PÁGINA WEB

        ┌─────────────────────┐
        │    Gestión Stock    │
        ├─────────────────────┤
        │                     │
        │ Código: 1000000001  │
        │                     │
        │ Cantidad: 5         │
        │                     │
        │ [ ENTRADA ]         │
        │ [ SALIDA  ]         │
        │                     │
        └─────────────────────┘
```

Cuando una persona presiona un botón, Flask recibe la acción y se comunica con Python.

---

# 🔄 Ejemplo completo

Supongamos que tenemos:

```text
Producto: Salsa Golf
Código: 100000000067
Stock: 20
```

Una persona pasa el código de barras por el lector.

El lector envía:

```text
100000000067
```

Python busca ese código en Excel.

Encuentra:

```text
Salsa Golf
Stock: 20
```

La persona indica:

```text
Cantidad que sale: 5
```

Python realiza:

```text
20 - 5 = 15
```

Finalmente Excel queda:

```text
Salsa Golf
Stock: 15
```

Y el movimiento puede quedar registrado:

```text
Producto: Salsa Golf
Código: 100000000067
Operación: SALIDA
Cantidad: 5
Stock anterior: 20
Stock actual: 15
Fecha: ...
```

---

# 🧩 16. Flujo completo del sistema

Podemos resumir todo el funcionamiento de esta manera:

```text
          LECTOR DE CÓDIGO DE BARRAS
                     │
                     ▼
                CÓDIGO
                     │
                     ▼
                  PYTHON
                     │
          ┌──────────┴──────────┐
          │                     │
          ▼                     ▼
       BUSCAR                 MODIFICAR
       PRODUCTO                STOCK
          │                     │
          └──────────┬──────────┘
                     ▼
                   EXCEL
                     │
                     ▼
                MOVIMIENTO
                     │
                     ▼
                  FLASK
                     │
                     ▼
                PÁGINA WEB
```

---

# 🏷️ ¿Cómo funciona el lector de códigos?

El **Metrologic MS7120** se conecta mediante USB.

Una vez conectado y configurado correctamente, normalmente funciona de una manera muy sencilla:

```text
Código de barras
       ↓
Metrologic MS7120
       ↓
      USB
       ↓
 Raspberry Pi
       ↓
     Python
```

El lector lee el código y envía los números al equipo.

Para el programa, es parecido a escribir el código utilizando un teclado.

Por ejemplo, si el código es:

```text
100000000067
```

el lector envía ese número.

Esto permite buscar el producto automáticamente.

---

# 🔢 Generación de códigos de barras

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

se convierte en una imagen similar a:

```text
|||| ||| |||| || |||| ||| ||
100000000067
```

La imagen puede almacenarse y posteriormente utilizarse para identificar el producto.

---

# 📑 Excel y la hoja "Movimientos"

El sistema utiliza Excel no solamente para guardar el stock.

También puede utilizar una hoja llamada:

```text
Movimientos
```

Esta hoja permite guardar un historial.

Por ejemplo:

| Fecha      | Código       | Producto   | Operación | Cantidad | Stock anterior | Stock actual |
| ---------- | ------------ | ---------- | --------- | -------: | -------------: | -----------: |
| 27/08/2026 | 100000000067 | Salsa Golf | SALIDA    |        5 |             20 |           15 |
| 27/08/2026 | 100000000002 | Pizza      | ENTRADA   |       10 |             15 |           25 |

De esta manera podemos saber qué ocurrió con el stock.

---

# ⚠️ 17. Problemas frecuentes

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

Y ya podemos trabajar con el proyecto.

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

Y ya podemos ejecutar el proyecto.

---

# 🏁 18. Resumen

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
10. Ejecutar el sistema.

### Para administrar la Raspberry

Podemos utilizar:

* **FileZilla** → para mover archivos entre nuestra computadora y la Raspberry.
* **VNC Viewer** → para ver y controlar la pantalla de la Raspberry desde otra computadora.

---

# 💡 Idea principal del proyecto

La idea del sistema puede resumirse en una sola frase:

> **El lector identifica el producto, Python realiza las operaciones, Excel guarda la información y Flask permite utilizar todo desde una página web.**

De esta manera tenemos un sistema que combina hardware y software para facilitar la administración del stock.

---

# 📚 Enlaces útiles

### Python

[Python — Página oficial](https://www.python.org/?utm_source=chatgpt.com)

### Flask

[Flask — Documentación oficial](https://flask.palletsprojects.com/?utm_source=chatgpt.com)

### Visual Studio Code

[Visual Studio Code — Página oficial](https://code.visualstudio.com/?utm_source=chatgpt.com)

### Raspberry Pi

[Raspberry Pi — Página oficial](https://www.raspberrypi.com/?utm_source=chatgpt.com)

### Raspberry Pi Imager

[Raspberry Pi Imager y sistemas operativos](https://www.raspberrypi.com/software/?utm_source=chatgpt.com)

### FileZilla

[FileZilla — Página oficial](https://filezilla-project.org/?utm_source=chatgpt.com)

### RealVNC

[RealVNC — Página oficial](https://www.realvnc.com/?utm_source=chatgpt.com)

---

# 👥 Para nuevos usuarios

No es necesario saber programación para comenzar.

La recomendación es seguir esta guía **en orden**, sin saltear pasos.

Si aparece un error:

1. Leer el mensaje que aparece.
2. Comprobar que el entorno virtual esté activado.
3. Comprobar que las librerías estén instaladas.
4. Comprobar que estamos dentro de la carpeta correcta.
5. Revisar la sección de problemas frecuentes.
6. Si el problema continúa, consultar el mensaje completo del error.

No recomendamos borrar archivos del sistema ni ejecutar comandos que no entendamos.

---

Este proyecto fue pensado para aprender y, al mismo tiempo, crear una herramienta que pueda utilizarse en una situación real.

La intención de este README es que cualquier persona pueda descargar el proyecto y seguir los pasos necesarios para ponerlo en funcionamiento, incluso si nunca trabajó anteriormente con Python, Flask, Raspberry Pi o códigos de barras.
