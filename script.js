const archivo = document.getElementById("archivo");
const boton = document.getElementById("buscar");
const pregunta = document.getElementById("pregunta");
const respuesta = document.getElementById("respuesta");
const copiar = document.getElementById("copiar");

let documentoTexto = "";
let archivoActual = null;
let resultados = "";

const botonesNumeros = document.getElementById("botonesNumeros");
const botonesCantidad = document.getElementById("botonesCantidad");
let longitudActiva = 4;

for (let i = 3; i <= 9; i++) {
    const boton = document.createElement("button");
    boton.type = "button";
    boton.className = "boton-cantidad";
    boton.textContent = i;
    boton.dataset.cantidad = String(i);

    if (i === longitudActiva) {
        boton.classList.add("activo");
    }

    boton.addEventListener("click", () => {
        longitudActiva = Number(i);

        const cantidadBotones = botonesCantidad.querySelectorAll(".boton-cantidad");
        cantidadBotones.forEach(item => item.classList.toggle("activo", Number(item.dataset.cantidad) === longitudActiva));

        respuesta.innerHTML = "OTP'S <b>" + longitudActiva + "</b> DIGITOS activado.";
    });

    botonesCantidad.appendChild(boton);
}

for (let i = 0; i < 10; i++) {
    const boton = document.createElement("button");
    boton.type = "button";
    boton.className = "boton-numero";
    boton.textContent = i;
    boton.dataset.numero = String(i);

    boton.addEventListener("click", () => {
        buscarNumero(i);
    });

    botonesNumeros.appendChild(boton);
}

function escapeRegExp(texto) {
    return texto.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function limpiarHtml(texto) {
    return texto.replace(/<[^>]*>/g, "");
}

function buscarNumero(numero) {
    if(documentoTexto === ""){
        respuesta.innerHTML = "⚠️ Primero carga un documento.";
        return;
    }

    const base = String(numero);
    const patronBusqueda = base.repeat(longitudActiva);

    const lineas = documentoTexto.split(/\r?\n/);
    const encontrados = [];

    for(const linea of lineas){
        if(linea.includes(patronBusqueda)){
            const expresion = new RegExp(escapeRegExp(patronBusqueda), "gi");
            const resaltada = linea.replace(expresion, (match) => `<mark>${match}</mark>`);

            if(!encontrados.includes(linea)){
                encontrados.push(resaltada);
            }
        }
    }

    if(encontrados.length > 0){
        resultados = encontrados
            .map(r => "wa.me/" + limpiarHtml(r))
            .join("\n");

        respuesta.innerHTML =
            "<span class='etiqueta'>OTP'S <b>" + longitudActiva + "</b> DIGITOS</span><br><br>" +
            "Se encontraron <b>" + encontrados.length + "</b> coincidencia(s):<br><br>" +
            encontrados.join("<br><br>");
    } else {
        respuesta.innerHTML = "❌ No encontré información relacionada con <b>" + patronBusqueda + "</b>.";
        resultados = "";
    }
}

// Cargar documentos (TXT, Excel, Word)
archivo.addEventListener("change", async () => {

    const file = archivo.files[0];
    archivoActual = file;

    if(!file) return;


    // TXT
    if(file.name.endsWith(".txt")){

        const lector = new FileReader();

        lector.onload = function(e){

            documentoTexto = e.target.result;

const lineas = documentoTexto.split(/\r?\n/);

const peso = (archivoActual.size / 1024).toFixed(2);

const formato = archivoActual.name.split(".").pop().toUpperCase();


respuesta.innerHTML = `
🪷 <b>Archivo:</b> ${archivoActual.name}<br>
🪷 <b>Peso:</b> ${peso} KB<br>
🪷 <b>Formato:</b> ${formato}<br>
🪷 <b>Líneas:</b> ${lineas.length.toLocaleString()}<br><br>

OTP CARGADO ✅
`;

        };

        lector.readAsText(file);

    }


    // Excel (.xlsx / .xls)
    else if(
        file.name.endsWith(".xlsx") || 
        file.name.endsWith(".xls")
    ){

        const datos = await file.arrayBuffer();

        const libro = XLSX.read(datos);

        let texto = "";

        libro.SheetNames.forEach(nombre => {

            const hoja = libro.Sheets[nombre];

            texto += XLSX.utils.sheet_to_txt(hoja);

        });

        documentoTexto = texto;

const lineas = documentoTexto.split(/\r?\n/);

const peso = (archivoActual.size / 1024).toFixed(2);

const formato = archivoActual.name.split(".").pop().toUpperCase();


respuesta.innerHTML = `
🪷 <b>Archivo:</b> ${archivoActual.name}<br>
🪷 <b>Peso:</b> ${peso} KB<br>
🪷 <b>Formato:</b> ${formato}<br>
🪷 <b>Líneas:</b> ${lineas.length.toLocaleString()}<br><br>

OTP CARGADO ✅
`;

    }


    // Word (.docx)
    else if(file.name.endsWith(".docx")){

        const datos = await file.arrayBuffer();

        const resultado = await mammoth.extractRawText({
            arrayBuffer: datos
        });

        documentoTexto = resultado.value;

const lineas = documentoTexto.split(/\r?\n/);

const peso = (archivoActual.size / 1024).toFixed(2);

const formato = archivoActual.name.split(".").pop().toUpperCase();


respuesta.innerHTML = `
🪷 <b>Archivo:</b> ${archivoActual.name}<br>
🪷 <b>Peso:</b> ${peso} KB<br>
🪷 <b>Formato:</b> ${formato}<br>
🪷 <b>Líneas:</b> ${lineas.length.toLocaleString()}<br><br>

OTP CARGADO ✅
`;

    }


    else{

        respuesta.innerHTML = "❌ Formato no compatible";

    }

});

// Buscar información
boton.addEventListener("click", () => {

    if(documentoTexto === ""){
        respuesta.innerHTML = "⚠️ Primero carga un documento.";
        return;
    }

    const preguntaUsuario = pregunta.value.toLowerCase().trim();

    if(preguntaUsuario === ""){
        respuesta.innerHTML = "⚠️ Inserta los digitos, Ex: 4444 , 9999.";
        return;
    }

    const palabras = preguntaUsuario.split(/\s+/).filter(Boolean);

    const lineas = documentoTexto.split(/\r?\n/);

    let encontrados = [];

    for(let linea of lineas){

        let texto = linea.toLowerCase();

        for(let palabra of palabras){

            if(texto.includes(palabra)){

                // Evita repetir la misma línea
                if(!encontrados.includes(linea)){
                    const patron = new RegExp(escapeRegExp(palabra), "gi");
                    const resaltada = linea.replace(patron, (match) => `<mark>${match}</mark>`);

                    encontrados.push(resaltada);
                }

                break;
            }

        }

    }

    if(encontrados.length > 0){

        resultados = encontrados
            .map(r => "wa.me/" + limpiarHtml(r))
            .join("\n");

        respuesta.innerHTML =
            "OTP'S <b>" + encontrados.length + "</b> resultado(s):<br><br>" +
            encontrados.join("<br><br>");

    } else {

        respuesta.innerHTML =
            "❌ No encontré información relacionada.";

        resultados = "";

    }

});


// BOTÓN COPIAR
copiar.addEventListener("click", async () => {

    if(resultados === ""){
        alert("⚠️ No hay resultados para copiar.");
        return;
    }

    try {
        if (navigator.clipboard && window.isSecureContext) {
            await navigator.clipboard.writeText(resultados);
        } else {
            const area = document.createElement("textarea");
            area.value = resultados;
            area.setAttribute("readonly", "");
            area.style.position = "fixed";
            area.style.opacity = "0";
            document.body.appendChild(area);
            area.select();
            document.execCommand("copy");
            document.body.removeChild(area);
        }

        alert("✅ Resultados copiados.");
    } catch (error) {
        console.error(error);
        alert("❌ Error al copiar.");
    }

});