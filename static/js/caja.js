const inputCodigo =
    document.getElementById("codigo");

const carrito = {};


// ==========================================
// DETECTAR EL LECTOR
// ==========================================

inputCodigo.addEventListener(
    "keydown",
    function(event) {

        if (event.key !== "Enter") {
            return;
        }


        const codigo =
            inputCodigo.value.trim();


        if (codigo === "") {
            return;
        }


        inputCodigo.value = "";


        buscarProducto(codigo);

    }
);


// ==========================================
// BUSCAR PRODUCTO
// ==========================================

async function buscarProducto(codigo) {

    try {

        const respuesta = await fetch(
            `/api/producto/${codigo}`
        );


        const producto =
            await respuesta.json();


        if (!respuesta.ok) {

            mostrarMensaje(
                "❌ Producto no encontrado",
                "error"
            );

            return;
        }


        agregarProducto(producto);


        mostrarMensaje(
            `✅ ${producto.nombre}`,
            "ok"
        );


    } catch (error) {

        console.error(error);

        mostrarMensaje(
            "❌ Error de conexión",
            "error"
        );

    }

}


// ==========================================
// AGREGAR AL CARRITO
// ==========================================

function agregarProducto(producto) {

    const codigo =
        producto.codigo;


    if (carrito[codigo]) {

        carrito[codigo].cantidad++;

    } else {

        carrito[codigo] = {

            codigo:
                producto.codigo,

            nombre:
                producto.nombre,

            precio:
                Number(producto.precio),

            cantidad: 1

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


    for (
        const codigo in carrito
    ) {

        const producto =
            carrito[codigo];


        const subtotal =
            producto.precio *
            producto.cantidad;


        total += subtotal;


        const fila =
            document.createElement("tr");


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


        tabla.appendChild(fila);

    }


    document.getElementById(
        "total"
    ).textContent =
        total.toFixed(2);

}


// ==========================================
// ELIMINAR DEL CARRITO
// ==========================================

function eliminarDelCarrito(codigo) {

    delete carrito[codigo];

    actualizarCarrito();

}


// ==========================================
// FINALIZAR COMPRA
// ==========================================

async function finalizarCompra() {

    const productos =
        Object.values(carrito);


    if (productos.length === 0) {

        alert(
            "El carrito está vacío"
        );

        return;
    }


    try {

        const respuesta =
            await fetch(
                "/api/venta",
                {

                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body: JSON.stringify({
                        productos:
                            productos
                    })

                }
            );


        const resultado =
            await respuesta.json();


        if (!respuesta.ok) {

            alert(
                resultado.error
            );

            return;
        }


        alert(
            "Compra realizada correctamente"
        );


        // Vaciar carrito

        for (
            const codigo in carrito
        ) {

            delete carrito[codigo];

        }


        actualizarCarrito();


        // Volver a enfocar lector

        inputCodigo.focus();


    } catch (error) {

        console.error(error);

        alert(
            "Error al finalizar la compra"
        );

    }

}


// ==========================================
// MENSAJES
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