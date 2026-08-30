
const inputCodigo =
    document.getElementById("codigo");

const carrito = {};


// ==========================================
// DETECTAR 12 DÍGITOS AUTOMÁTICAMENTE
// ==========================================

inputCodigo.addEventListener(
    "input",
    function() {

        let codigo =
            inputCodigo.value.trim();


        // ==========================================
        // SOLO PERMITIR NÚMEROS
        // ==========================================

        if (!/^\d*$/.test(codigo)) {

            codigo =
                codigo.replace(/\D/g, "");

            inputCodigo.value =
                codigo;

        }


        // ==========================================
        // CUANDO HAYA EXACTAMENTE 12 NÚMEROS
        // ==========================================

        if (codigo.length === 12) {

            // Guardar el código
            const codigoEscaneado =
                codigo;


            // Limpiar el campo
            inputCodigo.value = "";


            // Buscar automáticamente
            buscarProducto(
                codigoEscaneado
            );

        }

    }
);


// ==========================================
// BUSCAR PRODUCTO
// ==========================================

async function buscarProducto(
    codigo
) {

    try {

        const respuesta =
            await fetch(
                `/api/producto/${codigo}`
            );


        const producto =
            await respuesta.json();


        // ==========================================
        // PRODUCTO NO ENCONTRADO
        // ==========================================

        if (!respuesta.ok) {

            mostrarMensaje(
                "❌ Producto no encontrado",
                "error"
            );


            inputCodigo.focus();


            return;
        }


        // ==========================================
        // AGREGAR PRODUCTO
        // ==========================================

        agregarProducto(
            producto
        );


        mostrarMensaje(
            `✅ ${producto.nombre}`,
            "ok"
        );


    } catch (
        error
    ) {

        console.error(
            error
        );


        mostrarMensaje(
            "❌ Error de conexión",
            "error"
        );

    }


    // Volver a enfocar el lector

    inputCodigo.focus();

}


// ==========================================
// AGREGAR AL CARRITO
// ==========================================

function agregarProducto(
    producto
) {

    const codigo =
        producto.codigo;


    // ==========================================
    // SI YA ESTÁ EN EL CARRITO
    // ==========================================

    if (
        carrito[codigo]
    ) {

        carrito[codigo].cantidad++;

    }


    // ==========================================
    // SI ES UN PRODUCTO NUEVO
    // ==========================================

    else {

        carrito[codigo] = {

            codigo:
                producto.codigo,

            nombre:
                producto.nombre,

            precio:
                Number(
                    producto.precio
                ),

            cantidad:
                1

        };

    }


    actualizarCarrito();

}


// ==========================================
// MOSTRAR CARRITO
// ==========================================

function actualizarCarrito() {

    const tabla =
        document.getElementById(
            "carrito"
        );


    tabla.innerHTML = "";


    let total = 0;


    // ==========================================
    // RECORRER PRODUCTOS
    // ==========================================

    for (
        const codigo in carrito
    ) {

        const producto =
            carrito[codigo];


        const subtotal =
            producto.precio *
            producto.cantidad;


        total +=
            subtotal;


        const fila =
            document.createElement(
                "tr"
            );


        fila.innerHTML = `

            <td>
                ${producto.nombre}
            </td>

            <td>
                $${producto.precio.toFixed(2)}
            </td>

            <td>
                ${producto.cantidad}
            </td>

            <td>
                $${subtotal.toFixed(2)}
            </td>

            <td>

                <button
                    type="button"
                    onclick="
                        eliminarDelCarrito(
                            '${codigo}'
                        )
                    "
                >
                    ❌
                </button>

            </td>

        `;


        tabla.appendChild(
            fila
        );

    }


    // ==========================================
    // MOSTRAR TOTAL
    // ==========================================

    document.getElementById(
        "total"
    ).textContent =
        total.toFixed(2);

}


// ==========================================
// ELIMINAR DEL CARRITO
// ==========================================

function eliminarDelCarrito(
    codigo
) {

    delete carrito[
        codigo
    ];


    actualizarCarrito();


    inputCodigo.focus();

}


// ==========================================
// FINALIZAR COMPRA
// ==========================================

async function finalizarCompra() {

    const productos =
        Object.values(
            carrito
        );


    // ==========================================
    // COMPROBAR CARRITO VACÍO
    // ==========================================

    if (
        productos.length === 0
    ) {

        mostrarMensaje(
            "⚠️ El carrito está vacío. Agregá al menos un producto.",
            "error"
        );


        inputCodigo.focus();


        return;
    }


    // ==========================================
    // PEDIR CORREO
    // ==========================================

    const correo =
        prompt(
            "Ingrese el correo electrónico para enviar el recibo:"
        );


    // ==========================================
    // COMPROBAR CORREO
    // ==========================================

    if (
        correo === null ||
        correo.trim() === ""
    ) {

        mostrarMensaje(
            "⚠️ Debés ingresar un correo electrónico.",
            "error"
        );


        inputCodigo.focus();


        return;
    }


    try {

        const respuesta =
            await fetch(
                "/api/venta",
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            productos:
                                productos,

                            correo:
                                correo.trim()

                        })

                }
            );


        const resultado =
            await respuesta.json();


        // ==========================================
        // ERROR
        // ==========================================

        if (
            !respuesta.ok
        ) {

            mostrarMensaje(
                `❌ ${resultado.error}`,
                "error"
            );


            inputCodigo.focus();


            return;
        }


        // ==========================================
        // COMPRA FINALIZADA
        // ==========================================

        mostrarMensaje(
            "✅ ¡Compra finalizada con éxito!",
            "ok"
        );


        // ==========================================
        // VACIAR CARRITO
        // ==========================================

        for (
            const codigo in carrito
        ) {

            delete carrito[
                codigo
            ];

        }


        actualizarCarrito();


        // ==========================================
        // VOLVER AL ESCÁNER
        // ==========================================

        inputCodigo.focus();


    } catch (
        error
    ) {

        console.error(
            error
        );


        mostrarMensaje(
            "❌ Error al finalizar la compra",
            "error"
        );


        inputCodigo.focus();

    }

}


// ==========================================
// MOSTRAR MENSAJES
// ==========================================

function mostrarMensaje(
    texto,
    tipo
) {

    const mensaje =
        document.getElementById(
            "mensaje"
        );


    mensaje.textContent =
        texto;


    mensaje.className =
        tipo;

}
