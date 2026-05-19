// URL de nuestro backend local
// URL relativa (busca la API en su propio servidor)
const API_BASE_URL = '/api/sheet';
let refreshInterval;

async function iniciarDashboard() {
    await cargarInfoIntegrada();
    // Auto-refresco en tiempo real cada 30 segundos sin recargar la página
    if(refreshInterval) clearInterval(refreshInterval);
    refreshInterval = setInterval(cargarInfoIntegrada, 30000);
}

async function cargarInfoIntegrada() {
    const fileSelect = document.getElementById('fileSelect');
    // Soporta la selección dinámica de múltiples planillas o usa el ID base por defecto
    const selectedSpreadsheetId = fileSelect ? fileSelect.value : '1ZAHVVzJTl7n6YebpXDxQtxiVYzs3Ymo1dC_v68_IX3A';

    const loader = document.getElementById('loader');
    const errorBox = document.getElementById('errorBox');
    const dataTable = document.getElementById('dataTable');
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    const lastUpdate = document.getElementById('lastUpdate');

    if (loader) loader.classList.remove('hidden');
    if (errorBox) errorBox.classList.add('hidden');

    try {
        // 1. Descarga paralela y asíncrona de ambas pestañas
        const [resHoja12, resRuteo] = await Promise.all([
            fetch(`${API_BASE_URL}/${selectedSpreadsheetId}/Hoja 12`),
            fetch(`${API_BASE_URL}/${selectedSpreadsheetId}/Ruteo`)
        ]);

        if (!resHoja12.ok || !resRuteo.ok) throw new Error("Error en la comunicación con el servidor Node.js");

        const jsonHoja12 = await resHoja12.json();
        const jsonRuteo = await resRuteo.json();

        // 2. RENDERIZADO DINÁMICO DE CABECERAS (Look original intacto)
        tableHead.innerHTML = '';
        tableBody.innerHTML = '';
        
        const headersHoja12 = jsonHoja12.headers || Object.keys(jsonHoja12.data[0] || {});

        headersHoja12.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText || '';
            tableHead.appendChild(th);
        });

        // Columnas de integración al final
        const thCronologia = document.createElement('th');
        thCronologia.textContent = "CRONOLOGÍA (RUTEO H)";
        thCronologia.style.backgroundColor = "#002244";
        thCronologia.style.color = "white";
        tableHead.appendChild(thCronologia);

        const thEstado = document.createElement('th');
        thEstado.textContent = "ESTADO VIAJE (RUTEO X)";
        thEstado.style.backgroundColor = "#002244";
        thEstado.style.color = "white";
        tableHead.appendChild(thEstado);

        // 3. RADAR CON MEMORIA (FORWARD-FILL) PARA LA PESTAÑA RUTEO
        const ruteoMap = new Map();
        let lastRuteoTractor = "";
        let lastRuteoTD = "";

        jsonRuteo.data.forEach(row => {
            const vals = Object.values(row);
            const rTractor = String(vals[0] || '').trim(); // Col A (0)
            const rTD = String(vals[11] || '').trim();     // Col L (11)

            // Memoria: Si la celda está vacía, mantiene el valor del conjunto superior
            if (rTractor) lastRuteoTractor = rTractor;
            if (rTD) lastRuteoTD = rTD;

            if (lastRuteoTractor && lastRuteoTD) {
                const cronologia = String(vals[7] || '').trim(); // Col H (7)
                const estado = String(vals[23] || '').trim();     // Col X (23)

                const key = `${lastRuteoTractor}-${lastRuteoTD}`;
                
                if (!ruteoMap.has(key)) {
                    ruteoMap.set(key, { cronologia: "", estado: "" });
                }
                
                const info = ruteoMap.get(key);
                
                // Acumulación Inteligente: Une los datos si el viaje se divide en varias filas
                if (cronologia && cronologia !== '---' && !info.cronologia.includes(cronologia)) {
                    info.cronologia = info.cronologia ? `${info.cronologia} / ${cronologia}` : cronologia;
                }
                if (estado && estado !== '---' && !info.estado.includes(estado)) {
                    info.estado = info.estado ? `${info.estado} / ${estado}` : estado;
                }
            }
        });

        // Índices clave en Hoja 12
        const idxTractor = headersHoja12.indexOf("VEHÍCULO");
        const idxTD = headersHoja12.indexOf("N° DE DESPACHO");

        // Variables de memoria para el renderizado de Hoja 12
        let lastH12Tractor = "";
        let lastH12TD = "";

        // 4. CONSTRUCCIÓN DE FILAS DE HOJA 12 (Sin alterar ni saltar nada)
        jsonHoja12.data.forEach(row => {
            if (!row || Object.keys(row).length === 0) return;
            const tr = document.createElement('tr');

            const vals = Object.values(row);
            
            // Detección automática de color de fila hexadecimal
            const colorHex = vals.find(val => val && val.toString().startsWith('#'));
            if (colorHex) tr.style.backgroundColor = colorHex;

            // Inyectar columnas nativas respetando las celdas vacías visuales del Sheet
            headersHoja12.forEach((header, index) => {
                const td = document.createElement('td');
                const valorCelda = row[header] !== undefined ? row[header] : (vals[index] || '');
                td.textContent = valorCelda;
                if (header === "VEHÍCULO" || header === "N° DE DESPACHO") td.style.fontWeight = "bold";
                tr.appendChild(td);
            });

            // Extraer llaves con memoria para buscar la integración correspondiente
            const currentTractor = String(row["VEHÍCULO"] || vals[idxTractor] || '').trim();
            const currentTD = String(row["N° DE DESPACHO"] || vals[idxTD] || '').trim();

            if (currentTractor) lastH12Tractor = currentTractor;
            if (currentTD) lastH12TD = currentTD;

            let cronologiaFinal = '---';
            let estadoFinal = '---';

            if (lastH12Tractor && lastH12TD) {
                const matchRuteo = ruteoMap.get(`${lastH12Tractor}-${lastH12TD}`);
                if (matchRuteo) {
                    cronologiaFinal = matchRuteo.cronologia || '---';
                    estadoFinal = matchRuteo.estado || '---';
                }
            }

            // Inyectar celda integrada H
            const tdCron = document.createElement('td');
            tdCron.textContent = cronologiaFinal;
            tdCron.style.backgroundColor = "rgba(255, 255, 255, 0.3)";
            tr.appendChild(tdCron);

            // Inyectar celda integrada X
            const tdEst = document.createElement('td');
            tdEst.textContent = estadoFinal;
            tdEst.style.backgroundColor = "rgba(255, 255, 255, 0.3)";
            tdEst.style.fontWeight = "bold";
            tr.appendChild(tdEst);

            tableBody.appendChild(tr);
        });

        if (dataTable) dataTable.classList.remove('hidden');
        const ahora = new Date();
        if (lastUpdate) lastUpdate.innerHTML = `Última actualización: ${ahora.toLocaleTimeString()} 🟢`;

    } catch (error) {
        console.error(error);
        if (errorBox) {
            errorBox.textContent = `❌ Error en el procesamiento de datos: ${error.message}`;
            errorBox.classList.remove('hidden');
        }
        if (lastUpdate) lastUpdate.innerHTML = `Desconectado 🔴`;
    } finally {
        if (loader) loader.classList.add('hidden');
    }
}

window.onload = iniciarDashboard;