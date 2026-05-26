import { initGoogleSheets } from '../config/googleSheets.js';

// 🛠️ FUNCIÓN MAGICA: Convierte el color RGB de Google a HEX (#FFFFFF)
const extraerHexDeGoogle = (color) => {
    if (!color) return null;
    const r = Math.round((color.red || 0) * 255);
    const g = Math.round((color.green || 0) * 255);
    const b = Math.round((color.blue || 0) * 255);
    
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
};

export const getSheetData = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        
        const sheet = doc.sheetsByTitle[sheetName];
        if (!sheet) {
            return res.status(404).json({ error: `La pestaña '${sheetName}' no existe.` });
        }

        // 1. Obtenemos los textos de las filas
        const rows = await sheet.getRows();
        
        // 👇 Solo cargamos la Columna H si NO estamos leyendo el Indice
        const requiereColores = sheetName.toLowerCase() !== 'indice';
        
        if (requiereColores && sheet.rowCount > 0) {
            try {
                await sheet.loadCells(`H1:H${sheet.rowCount}`);
            } catch (err) {
                console.log("Aviso: No se pudo cargar la columna H. Omitiendo colores.");
            }
        }

        const data = rows.map(row => {
            const filaCruda = [...(row._rawData || [])]; 

            // Solo intentamos extraer color si era una hoja operativa (Ruteo)
            if (requiereColores) {
                try {
                    const celdaH = sheet.getCell(row.rowNumber - 1, 7); // Columna H (índice 7)
                    const colorFondo = celdaH.backgroundColor;

                    if (colorFondo) {
                        const hex = extraerHexDeGoogle(colorFondo);
                        if (hex !== '#FFFFFF' && hex !== '#000000') {
                            filaCruda.push(hex); 
                        }
                    }
                } catch (error) {
                    // Silenciamos si la celda está vacía
                }
            }

            return filaCruda;
        });

        res.status(200).json({ 
            success: true, 
            headers: sheet.headerValues || [],
            data: data 
        });
    } catch (error) {
        console.error("❌ ERROR LEYENDO PESTAÑA:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};

export const writeTestLog = async (req, res) => {
    res.status(200).json({ message: "Log guardado" });
};

// 👇 MICROSERVICIO BFF ACTUALIZADO PARA REDISEÑO DE TARJETAS
export const getViajesIntegradosDia = async (req, res) => {
    try {
        const { spreadsheetId } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        
        const sheetRuteo = doc.sheetsByTitle['Ruteo'];
        const sheetH12 = doc.sheetsByTitle['Hoja 12'];
        
        if (!sheetRuteo) {
            return res.status(404).json({ success: false, error: 'Falta la pestaña Ruteo indispensable' });
        }

        const hasH12 = !!sheetH12;
        const rowsH12 = hasH12 ? await sheetH12.getRows() : [];

        // Carga selectiva de la columna H para conservar la memoria de la instancia en la nube
        if (sheetRuteo.rowCount > 0) {
            try {
                await sheetRuteo.loadCells({
                    startRowIndex: 0,
                    endRowIndex: sheetRuteo.rowCount,
                    startColumnIndex: 7,
                    endColumnIndex: 8
                });
            } catch (e) {
                console.warn("Aviso: No se cargó la columna H para colores históricos:", e.message);
            }
        }

        const rowsRuteo = await sheetRuteo.getRows();
        const headersRuteo = sheetRuteo.headerValues || [];
        const headersLimpio = headersRuteo.map(h => (h || "").toString().trim().toUpperCase());
        
        const regexTD = /^\d{7}$/;
        const regexFecha = /\d{1,2}\/\d{1,2}\/20\d{2}/;
        const regexHex = /^#[0-9A-Fa-f]{6}$/;

        // Mapeo dinámico de nuevos encabezados
        const idxUt = headersLimpio.indexOf("N° UT") >= 0 ? headersLimpio.indexOf("N° UT") : headersLimpio.indexOf("N UT");
        const idxSemi = headersLimpio.indexOf("SEMI");
        const idxChofer = headersLimpio.indexOf("CHOFER");
        const idxTracking = headersLimpio.indexOf("TRACKING");
        const idxHexA = headersLimpio.indexOf("HEXA");
        const idxHexHx = headersLimpio.indexOf("HEXHX");

        // Índices operacionales clásicos
        const idxDestino = headersLimpio.indexOf("DESTINO");
        const idxProducto = headersLimpio.indexOf("PRODUCTO") >= 0 ? headersLimpio.indexOf("PRODUCTO") : 24;
        const idxCantidad = headersLimpio.indexOf("CANTIDAD") >= 0 ? headersLimpio.indexOf("CANTIDAD") : 25;
        const idxCisternado = headersLimpio.indexOf("CISTERNADO") >= 0 ? headersLimpio.indexOf("CISTERNADO") : 29;
        const idxFiltroColX = idxDestino - 1;

        const idxAvisoVacio = headersLimpio.indexOf("AVISO DE VACIO");
        const idxLlegadaEta = headersLimpio.indexOf("LLEGADA(ETA)");
        const idxLlegada = headersLimpio.indexOf("LLEGADA");
        const limiteBusqueda = Math.max(idxAvisoVacio, idxLlegadaEta, idxLlegada, idxDestino - 6, 0);

        // Agrupamiento seguro de la Hoja 12
        const agrupadoH12 = {};
        for (let row of rowsH12) {
            const vals = row._rawData || [];
            const tdMatch = vals.find(v => regexTD.test((v || "").toString().trim()));
            const td = tdMatch ? tdMatch.toString().trim() : (vals[5] || "").toString().trim();
            
            if (td) {
                if (!agrupadoH12[td]) agrupadoH12[td] = [];
                agrupadoH12[td].push(vals);
            }
        }

        const viajesFinales = [];
        const procesados = new Set();

        for (let i = 0; i < rowsRuteo.length; i++) {
            const row = rowsRuteo[i];
            const vals = row._rawData || [];
            if (vals.length === 0) continue;

            const trRuteo = (vals[0] || "").toString().trim();
            const tdMatch = vals.find(v => regexTD.test((v || "").toString().trim()));
            const tdRuteo = tdMatch ? tdMatch.toString().trim() : "";

            // Filtro anti-basura robusto para depurar celdas rotas o vacías
            const isTractorValido = trRuteo.length > 0 && 
                                  trRuteo.length <= 12 && 
                                  !trRuteo.toUpperCase().includes("TERMINAL") &&
                                  !trRuteo.toUpperCase().includes("FECHA") &&
                                  !trRuteo.toUpperCase().includes("PRODUCTO") &&
                                  !trRuteo.startsWith("#") &&
                                  !trRuteo.startsWith(",");

            const isTdValido = tdRuteo.length > 0;

            if (isTractorValido || isTdValido) {
                const keyMerge = tdRuteo || `FALLBACK_${i}`;
                
                if (!procesados.has(keyMerge)) {
                    procesados.add(keyMerge);

                    // Color de legado de la columna H
                    let colorHexLegacy = null;
                    try {
                        const cell = sheetRuteo.getCell(row.rowNumber - 1, 7);
                        if (cell && cell.backgroundColor) {
                            colorHexLegacy = extraerHexDeGoogle(cell.backgroundColor);
                            if (colorHexLegacy === '#FFFFFF' || colorHexLegacy === '#000000') colorHexLegacy = null;
                        }
                    } catch (e) {}

                    // Lectura y validación estricta de nuevos datos de la fila
                    const numeroUt = idxUt >= 0 ? (vals[idxUt] || "").toString().trim() : "";
                    const semi = idxSemi >= 0 ? (vals[idxSemi] || "").toString().trim() : "";
                    const chofer = idxChofer >= 0 ? (vals[idxChofer] || "").toString().trim() : "";
                    
                    const rawTracking = idxTracking >= 0 ? (vals[idxTracking] || "").toString().trim() : "";
                    const ultimoTracking = rawTracking ? rawTracking.split('/')[0].trim() : "";

                    const colorHexAVal = idxHexA >= 0 ? (vals[idxHexA] || "").toString().trim() : "";
                    const colorHexHxVal = idxHexHx >= 0 ? (vals[idxHexHx] || "").toString().trim() : "";

                    const colorHexA = regexHex.test(colorHexAVal) ? colorHexAVal : null;
                    const colorHexHx = regexHex.test(colorHexHxVal) ? colorHexHxVal : null;

                    const isCompletado = idxFiltroColX >= 0 && (vals[idxFiltroColX] || "").toString().trim().length > 0;

                    let rawFecha = "";
                    if (idxDestino > 0) {
                        for (let j = idxDestino - 1; j > limiteBusqueda; j--) {
                            const valText = (vals[j] || "").toString();
                            if (regexFecha.test(valText)) {
                                rawFecha = valText;
                                break;
                            }
                        }
                    }
                    if (!rawFecha && headersLimpio.indexOf("FECHA PLANIFICADA") >= 0) {
                        rawFecha = vals[headersLimpio.indexOf("FECHA PLANIFICADA")] || "";
                    }
                    const fechaLimpia = rawFecha.split(" ")[0] || "Sin Fecha";

                    const filasH12 = agrupadoH12[tdRuteo] || [];
                    const paradas = [];
                    let terminalLimpia = "Sin Terminal";
                    let cisternadoReal = vals[idxCisternado] || "";

                    if (filasH12.length > 0) {
                        const textoCompleto = filasH12[0].join(" ").toUpperCase();
                        if (textoCompleto.includes("PLAZA HUINCUL") || textoCompleto.includes("TPH")) {
                            terminalLimpia = "Plaza Huincul";
                        } else if (textoCompleto.includes("DOCK SUD") || textoCompleto.includes("TDS")) {
                            terminalLimpia = "Dock Sud";
                        }
                        cisternadoReal = filasH12[0][13] || cisternadoReal;

                        for (let f12 of filasH12) {
                            paradas.push({
                                destino: f12[7] || "Sin Destino",
                                producto: f12[8] || "",
                                cantidad: f12[9] || "",
                                cisternado: f12[10] || ""
                            });
                        }
                    } else {
                        const textoRuteo = vals.join(" ").toUpperCase();
                        if (textoRuteo.includes("PLAZA HUINCUL") || textoRuteo.includes("TPH")) {
                            terminalLimpia = "Plaza Huincul";
                        } else if (textoRuteo.includes("DOCK SUD") || textoRuteo.includes("TDS")) {
                            terminalLimpia = "Dock Sud";
                        }

                        paradas.push({
                            destino: vals[idxDestino >= 0 ? idxDestino : 23] || "Validar Destino",
                            producto: vals[idxProducto] || "",
                            cantidad: vals[idxCantidad] || "",
                            cisternado: ""
                        });
                    }

                    viajesFinales.push({
                        idUnico: keyMerge,
                        tractor: trRuteo,
                        numDespacho: tdRuteo,
                        terminalOrigen: terminalLimpia,
                        fechaPlanificada: fechaLimpia,
                        cisternadoReal: cisternadoReal,
                        colorHex: colorHexLegacy,
                        isCompletado: isCompletado,
                        paradas: paradas,
                        // 👇 Nuevas propiedades para la vista principal de la App
                        numeroUt: numeroUt,
                        semi: semi,
                        chofer: chofer,
                        ultimoTracking: ultimoTracking,
                        colorHexA: colorHexA,
                        colorHexHx: colorHexHx
                    });
                }
            }
        }
        res.status(200).json({ success: true, data: viajesFinales });
    } catch (error) {
        console.error("❌ ERROR INTEGRANDO VIAJES EN BFF:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};
