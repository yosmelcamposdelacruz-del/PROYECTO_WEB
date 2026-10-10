
const SPREADSHEET_ID = "1ZtLwY246DIXl5r_Qx-gFg4uq9S9_JI7NSNj84UaIJcI";
const CARPETA_FOTOS_ID = "1F1xVyfLAZnuloOtEvhxui2fWIybaw6Ly";

const NOMBRE_HOJA_CARRERAS = "Carreras";
const NOMBRE_HOJA_REGISTROS = "Nombres";

function obtenerHoja(ss, nombreHoja) {
    // Si se ejecuta a mano desde el editor llega sin datos: no hay nada que buscar.
    if (!ss || !nombreHoja) {
        return null;
    }

    const buscado = String(nombreHoja).trim().toLowerCase();
    const hojas = ss.getSheets();

    for (let i = 0; i < hojas.length; i++) {
        if (hojas[i].getName().trim().toLowerCase() === buscado) {
            return hojas[i];
        }
    }

    return null;
}

function doGet(e) {
    try {
        const parametros = e && e.parameter ? e.parameter : {};
        const accion = parametros.accion || "";

        if (accion === "ping") {
            return respuestaJSON({
                success: true,
                version: "registro-sin-duplicados-v3",
                mensaje: "Apps Script funciona correctamente"
            });
        }

        if (accion === "registrar") {
            return responder(parametros.callback, guardarRegistro(parametros));
        }

        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hoja = ss.getSheetByName(NOMBRE_HOJA_CARRERAS);

        if (!hoja) {
            throw new Error('No existe la pestaña "' + NOMBRE_HOJA_CARRERAS + '".');
        }

        const ultimaFila = hoja.getLastRow();

        if (ultimaFila < 2) {
            return responder(parametros.callback, []);
        }

        const rango = hoja.getRange(2, 1, ultimaFila - 1, 3);
        const valores = rango.getDisplayValues();
        const formulas = rango.getFormulas();
        const objetos = rango.getValues();
        const textoEnriquecido = rango.getRichTextValues();
        const fotosDrive = obtenerFotosDrive();

        const carreras = valores.map(function(fila, indice) {
            const nombre = String(fila[0] || "").trim();
            const descripcion = String(fila[1] || "").trim();
            let referenciaFoto = String(fila[2] || "").trim();
            const formulaFoto = String(formulas[indice][2] || "").trim();

            if (!nombre) {
                return null;
            }

            const coincidenciaFormula = formulaFoto.match(
                /^=\s*(?:IMAGE|IMAGEN|HYPERLINK|HIPERVINCULO)\(\s*"([^"]+)"/i
            );

            if (coincidenciaFormula) {
                referenciaFoto = coincidenciaFormula[1];
            }

            if (!referenciaFoto) {
                const valorCelda = objetos[indice][2];

                // Imagen insertada dentro de la celda (Insertar > Imagen > En la celda)
                if (valorCelda && typeof valorCelda.getUrl === "function") {
                    referenciaFoto = valorCelda.getUrl()
                        || valorCelda.getContentUrl()
                        || "";
                }
            }

            if (!referenciaFoto) {
                // Texto con enlace incrustado
                referenciaFoto = textoEnriquecido[indice][2].getLinkUrl() || "";
            }

            let imagen = resolverImagen(referenciaFoto, fotosDrive);

            // Respaldo: buscar en la carpeta de Drive un archivo con el nombre de la carrera
            if (!imagen) {
                const id = fotosDrive[normalizarNombre(nombre)];
                imagen = id ? crearUrlImagen(id) : "";
            }

            if (parametros.accion === "debug") {
                return {
                    nombre: nombre,
                    celdaTexto: fila[2],
                    celdaFormula: formulaFoto,
                    celdaTipo: objetos[indice][2] && objetos[indice][2].toString
                        ? Object.prototype.toString.call(objetos[indice][2])
                        : typeof objetos[indice][2],
                    celdaEnlace: textoEnriquecido[indice][2].getLinkUrl() || "",
                    archivosDrive: Object.keys(fotosDrive),
                    imagen: imagen
                };
            }

            return {
                nombre: nombre,
                descripcion: descripcion,
                imagen: imagen,
                foto: imagen
            };
        }).filter(function(carrera) {
            return carrera !== null;
        });

        return responder(parametros.callback, carreras);

    } catch (error) {
        console.error("Error en doGet: " + error.message);

        return responder(
            e && e.parameter ? e.parameter.callback : "",
            {
                success: false,
                mensaje: String(error.message || error)
            }
        );
    }
}

function obtenerFotosDrive() {
    const fotos = {};

    try {
        const carpeta = DriveApp.getFolderById(CARPETA_FOTOS_ID);
        const archivos = carpeta.getFiles();

        while (archivos.hasNext()) {
            const archivo = archivos.next();

            if (archivo.getMimeType().indexOf("image/") !== 0) {
                continue;
            }

            const nombreArchivo = archivo.getName().toLowerCase();
            const nombreBase = nombreArchivo.replace(/\.[^.]+$/, "");

            fotos[nombreArchivo] = archivo.getId();
            fotos[nombreBase] = archivo.getId();
            fotos[normalizarNombre(nombreBase)] = archivo.getId();
        }
    } catch (error) {
        // Sin permiso de Drive: se sigue usando solo URLs/IDs escritos en la hoja.
        console.error("No se pudo leer la carpeta de Drive: " + error.message);
        fotos["__error"] = error.message;
    }

    return fotos;
}

function normalizarNombre(texto) {
    return String(texto || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[̀-ͯ]/g, "")
        .replace(/[^a-z0-9]+/g, "");
}

function resolverImagen(referencia, fotos) {
    if (!referencia) {
        return "";
    }

    const valor = String(referencia).trim();

    if (/^https?:\/\//i.test(valor)) {
        const idDrive = extraerIdDrive(valor);

        if (idDrive) {
            return crearUrlImagen(idDrive);
        }

        return valor;
    }

    const id = extraerIdDrive(valor)
        || fotos[valor.toLowerCase()]
        || fotos[valor.toLowerCase().replace(/\.[^.]+$/, "")];

    return id ? crearUrlImagen(id) : "";
}

function crearUrlImagen(id) {
    return "https://drive.google.com/thumbnail?id="
        + encodeURIComponent(id)
        + "&sz=w1000";
}

function extraerIdDrive(valor) {
    const texto = String(valor || "").trim();

    if (/^[a-zA-Z0-9_-]{20,}$/.test(texto)) {
        return texto;
    }

    const coincidencia = texto.match(
        /(?:\/d\/|\/file\/d\/|[?&]id=)([a-zA-Z0-9_-]{20,})/
    );

    return coincidencia ? coincidencia[1] : "";
}

function guardarRegistro(datos) {
    try {
        const nombre = String(datos.nombre || "").trim();
        const correo = String(datos.correo || "").trim();
        const telefono = String(datos.telefono || "").trim();
        const carrera = String(datos.carrera || "").trim();

        if (!nombre || !correo || !telefono || !carrera) {
            return { success: false, mensaje: "Completa todos los campos." };
        }

        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) {
            return { success: false, mensaje: "El correo electrónico no es válido." };
        }

        const idRegistro = String(datos.idRegistro || "").trim();
        const claveDuplicado = "reg_" + (
            idRegistro || [correo, telefono, carrera].join("|").toLowerCase()
        );

        // Serializa las ejecuciones para que dos peticiones simultáneas no guarden doble.
        const candado = LockService.getScriptLock();
        candado.waitLock(15000);

        try {
            const cache = CacheService.getScriptCache();

            if (cache.get(claveDuplicado)) {
                return { success: true, mensaje: "Registro guardado correctamente." };
            }

            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            let hoja = obtenerHoja(ss, NOMBRE_HOJA_REGISTROS);

            if (!hoja) {
                hoja = ss.insertSheet(NOMBRE_HOJA_REGISTROS);
                hoja.appendRow([
                    "Fecha",
                    "Nombre",
                    "Correo",
                    "Teléfono",
                    "Carrera"
                ]);
            }

            // Respaldo: si ya hay una fila igual en los últimos 2 minutos, no se vuelve a guardar.
            if (existeRegistroReciente(hoja, correo, telefono, carrera)) {
                cache.put(claveDuplicado, "1", 120);
                return { success: true, mensaje: "Registro guardado correctamente." };
            }

            hoja.appendRow([
                new Date(),
                nombre,
                correo,
                "'" + telefono,
                carrera
            ]);

            // Recuerda este registro 2 minutos para ignorar reenvíos repetidos.
            cache.put(claveDuplicado, "1", 120);
        } finally {
            candado.releaseLock();
        }

        return { success: true, mensaje: "Registro guardado correctamente." };

    } catch (error) {
        console.error("Error guardando registro: " + error.message);

        return { success: false, mensaje: String(error.message || error) };
    }
}

function existeRegistroReciente(hoja, correo, telefono, carrera) {
    const ultimaFila = hoja.getLastRow();

    if (ultimaFila < 2) {
        return false;
    }

    const filaInicial = Math.max(2, ultimaFila - 9);
    const filas = hoja.getRange(filaInicial, 1, ultimaFila - filaInicial + 1, 5).getValues();
    const limite = Date.now() - 2 * 60 * 1000;
    const limpiarTelefono = function(valor) {
        return String(valor).replace(/^'/, "").trim();
    };

    return filas.some(function(fila) {
        const fecha = fila[0] instanceof Date ? fila[0].getTime() : 0;

        return fecha >= limite
            && String(fila[2]).trim().toLowerCase() === correo.toLowerCase()
            && limpiarTelefono(fila[3]) === telefono
            && String(fila[4]).trim() === carrera;
    });
}

function doPost(e) {
    return respuestaJSON(guardarRegistro(obtenerDatosPost(e)));
}

function obtenerDatosPost(e) {
    if (!e) {
        return {};
    }

    if (e.parameter && Object.keys(e.parameter).length > 0) {
        return e.parameter;
    }

    if (e.postData && e.postData.contents) {
        try {
            return JSON.parse(e.postData.contents);
        } catch (error) {
            return {};
        }
    }

    return {};
}

function responder(callback, datos) {
    if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
        return ContentService
            .createTextOutput(callback + "(" + JSON.stringify(datos) + ");")
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return respuestaJSON(datos);
}

function respuestaJSON(datos) {
    return ContentService
        .createTextOutput(JSON.stringify(datos))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Ejecuta esta función manualmente una sola vez (tras cambiar de cuenta).
 * Hace visibles las fotos para cualquier persona con el enlace, requisito
 * para que el navegador las muestre en la página web.
 */
function compartirCarpetaFotos() {
    const carpeta = DriveApp.getFolderById(CARPETA_FOTOS_ID);

    carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const archivos = carpeta.getFiles();
    let cantidad = 0;

    while (archivos.hasNext()) {
        const archivo = archivos.next();

        if (archivo.getMimeType().indexOf("image/") === 0) {
            archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            cantidad++;
        }
    }

    Logger.log("Carpeta '" + carpeta.getName() + "' compartida. Imágenes: " + cantidad);
}

/**
 * Ejecuta esta función manualmente una sola vez
 * para autorizar el acceso a Google Drive.
 */
function autorizarAccesoDrive() {
    const carpeta = DriveApp.getFolderById(CARPETA_FOTOS_ID);

    Logger.log("Acceso autorizado a la carpeta: " + carpeta.getName());

    const archivos = carpeta.getFiles();
    let cantidad = 0;

    while (archivos.hasNext()) {
        const archivo = archivos.next();

        if (archivo.getMimeType().indexOf("image/") === 0) {
            cantidad++;
        }
    }

    Logger.log("Cantidad de imágenes encontradas: " + cantidad);
}