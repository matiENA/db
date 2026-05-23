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

// 👇 EL NUEVO MICROSERVICIO BFF (Backend For Frontend)
export const getViajesIntegradosDia = async (req, res) => {
    try {
        const { spreadsheetId } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        
        const sheetRuteo = doc.sheetsByTitle['Ruteo'];
        const sheetH12 = doc.sheetsByTitle['Hoja 12'];
        
        if (!sheetRuteo || !sheetH12) {
            return res.status(404).json({ success: false, error: 'Faltan pestañas Ruteo u Hoja 12' });
        }

        // Cargamos SOLO la columna H para no saturar la memoria de Render
        if (sheetRuteo.rowCount > 0) {
            try { await sheetRuteo.loadCells(`H1:H${sheetRuteo.rowCount}`); } catch (e) {}
        }

        const rowsRuteo = await sheetRuteo.getRows();
        const headersRuteo = sheetRuteo.headerValues || [];
        const rowsH12 = await sheetH12.getRows();
        
        const regexTD = /^\d{7}$/;
        const regexFecha = /\d{1,2}\/\d{1,2}\/20\d{2}/;

        const idxDestino = headersRuteo.indexOf("DESTINO");
        const idxProducto = headersRuteo.indexOf("PRODUCTO") >= 0 ? headersRuteo.indexOf("PRODUCTO") : 24;
        const idxCantidad = headersRuteo.indexOf("CANTIDAD") >= 0 ? headersRuteo.indexOf("CANTIDAD") : 25;
        const idxCisternado = headersRuteo.indexOf("CISTERNADO") >= 0 ? headersRuteo.indexOf("CISTERNADO") : 29;
        const idxFiltroColX = idxDestino - 1;

        const idxAvisoVacio = headersRuteo.indexOf("Aviso de Vacio");
        const idxLlegadaEta = headersRuteo.indexOf("LLEGADA(ETA)");
        const idxLlegada = headersRuteo.indexOf("LLEGADA");
        const limiteBusqueda = Math.max(idxAvisoVacio, idxLlegadaEta, idxLlegada, idxDestino - 6, 0);

        // 1. AGRUPAR HOJA 12 EN LA MEMORIA DEL SERVIDOR
        const agrupadoH12 = {};
        for (let row of rowsH12) {
            const vals = row._rawData || [];
            const tdMatch = vals.find(v => regexTD.test((v||"").trim()));
            const td = tdMatch ? tdMatch.trim() : (vals[5] || "").trim();
            
            if (!agrupadoH12[td]) agrupadoH12[td] = [];
            agrupadoH12[td].push(vals);
        }

        const viajesFinales = [];
        const procesados = new Set();

        // 2. PROCESAR RUTEO Y APLICAR FILTRO ANTI-BASURA
        for (let i = 0; i < rowsRuteo.length; i++) {
            const row = rowsRuteo[i];
            const vals = row._rawData || [];
            
            const trRuteo = (vals[0] || "").trim();
            const tdMatch = vals.find(v => regexTD.test((v||"").trim()));
            const tdRuteo = tdMatch ? tdMatch.trim() : "";

            // 🛡️ EL FILTRO MAGNÍFICO (Rechaza títulos y basura de más de 12 letras)
            const isTractorValido = trRuteo.length > 0 && trRuteo.length <= 12 && 
                                  !trRuteo.toUpperCase().includes("TERMINAL") &&
                                  !trRuteo.toUpperCase().includes("FECHA");
            const isTdValido = tdRuteo.length > 0;

            if (isTractorValido || isTdValido) {
                const keyMerge = tdRuteo || `FALLBACK_${i}`;
                
                if (!procesados.has(keyMerge)) {
                    procesados.add(keyMerge);

                    let colorHex = null;
                    try {
                        const celdaH = sheetRuteo.getCell(row.rowNumber - 1, 7);
                        const bg = celdaH.backgroundColor;
                        if (bg) {
                            colorHex = extraerHexDeGoogle(bg);
                            if (colorHex === '#FFFFFF' || colorHex === '#000000') colorHex = null;
                        }
                    } catch (e) {}

                    const isCompletado = idxFiltroColX >= 0 && (vals[idxFiltroColX] || "").trim().length > 0;

                    let rawFecha = "";
                    if (idxDestino > 0) {
                        for (let j = idxDestino - 1; j > limiteBusqueda; j--) {
                            if (regexFecha.test(vals[j] || "")) {
                                rawFecha = vals[j];
                                break;
                            }
                        }
                    }
                    if (!rawFecha && headersRuteo.indexOf("FECHA PLANIFICADA") >= 0) {
                        rawFecha = vals[headersRuteo.indexOf("FECHA PLANIFICADA")] || "";
                    }
                    const fechaLimpia = rawFecha.split(" ")[0] || "Sin Fecha";

                    const filasH12 = agrupadoH12[tdRuteo] || [];
                    const paradas = [];
                    let terminalLimpia = "Sin Terminal";
                    let cisternadoReal = vals[idxCisternado] || "";

                    if (filasH12.length > 0) {
                        const textoCompleto = filasH12[0].join(" ").toUpperCase();
                        if (textoCompleto.includes("PLAZA HUINCUL")) terminalLimpia = "Plaza Huincul";
                        else if (textoCompleto.includes("DOCK SUD")) terminalLimpia = "Dock Sud";

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
                        if (textoRuteo.includes("PLAZA HUINCUL")) terminalLimpia = "Plaza Huincul";
                        else if (textoRuteo.includes("DOCK SUD")) terminalLimpia = "Dock Sud";

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
                        colorHex: colorHex,
                        isCompletado: isCompletado,
                        paradas: paradas
                    });
                }
            }
        }
        res.status(200).json({ success: true, data: viajesFinales });
    } catch (error) {
        console.error("❌ ERROR INTEGRANDO VIAJES:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};
