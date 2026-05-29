import { initGoogleSheets } from '../config/googleSheets.js';
import admin from 'firebase-admin';
import fs from 'fs';

// Inicialización de la Caché Global en Memoria del BFF [txt]
let cacheViajesConsolidados = []; 
let cacheDiasDisponibles = [];
let lastSyncTime = null;

const masterIndexSheetId = "1ny9yOftgyYWfzJFpQ9h8l2T_owDlyMV_HdEgeQ5Gm8E";

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

const extraerHexDeGoogle = (color) => {
    if (!color) return null;
    const r = Math.round((color.red || 0) * 255);
    const g = Math.round((color.green || 0) * 255);
    const b = Math.round((color.blue || 0) * 255);
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
};

const determinarTerminal = (texto) => {
    if (!texto) return "Sin Terminal";
    const t = texto.toUpperCase();
    if (t.includes("PLAZA HUINCUL") || t.includes("TPH") || t.includes("UTE")) return "Plaza Huincul";
    if (t.includes("DOCK SUD") || t.includes("TDS")) return "Dock Sud";
    return "Sin Terminal";
};

const secretsPath = '/etc/secrets/firebase-adminsdk.json';
const localPath = './firebase-adminsdk.json';
let firebaseInitialized = false;

try {
    const certPath = fs.existsSync(secretsPath) ? secretsPath : (fs.existsSync(localPath) ? localPath : null);
    if (certPath) {
        const serviceAccount = JSON.parse(fs.readFileSync(certPath, 'utf8'));
        admin.initializeApp({
            credential: admin.credential.cert(serviceAccount)
        });
        firebaseInitialized = true;
        console.log("🔥 [FIREBASE] Admin SDK inicializado exitosamente desde:", certPath);
    } else {
        console.warn("⚠️ [FIREBASE] Advertencia: No se encontró certificado Admin SDK. Notificaciones push desactivadas.");
    }
} catch (err) {
    console.error("❌ [FIREBASE] Falló la inicialización de Firebase Admin SDK:", err.message);
}

const enviarNotificacionPushUT = async (viaje) => {
    if (!firebaseInitialized) return;
    try {
        const topic = `ut_${viaje.numeroUt}`;
        
        const lEstado = (viaje.horarioVacio && viaje.horarioVacio.trim().length > 0)
            ? `VACIO: ${viaje.horarioVacio.trim()}`
            : (viaje.estadoUt ? viaje.estadoUt.trim().toUpperCase() : "PENDIENTE");

        // 👇 PAYLOAD CONSOLIDAD CON CANAL ASOCIADO Y ALTA PRIORIDAD [1.1.6]
        const message = {
            notification: {
                title: `[UT: ${viaje.numeroUt || "S/D"}] ${viaje.chofer || "Chofer S/D"}`,
                body: `${viaje.tractor || "S/D"} | ${viaje.semi || "S/D"} | Viaje: ${viaje.nViaje || "S/D"}\nTD: ${viaje.numDespacho || "S/N"}\n${lEstado}`
            },
            data: {
                idUnico: (viaje.idUnico || "").toString(),
                tractor: (viaje.tractor || "").toString(),
                numDespacho: (viaje.numDespacho || "").toString(),
                terminalOrigen: (viaje.terminalOrigen || "").toString(),
                fechaPlanificada: (viaje.fechaPlanificada || "").toString(),
                cisternadoReal: (viaje.cisternadoReal || "").toString(),
                isCompletado: (viaje.isCompletado || false).toString(),
                numeroUt: (viaje.numeroUt || "").toString(),
                semi: (viaje.semi || "").toString(),
                chofer: (viaje.chofer || "").toString(),
                ultimoTracking: (viaje.ultimoTracking || "").toString(),
                nViaje: (viaje.nViaje || "").toString(),
                llegadaPlanta: (viaje.llegadaPlanta || "").toString(),
                horarioVacio: (viaje.horarioVacio || "").toString(),
                estadoUt: (viaje.estadoUt || "").toString()
            },
            topic: topic,
            // 👇 ACTUALIZADO: Bloque Android mapea el canal e indica el sonido [1.1.6]
            android: {
                priority: "high",
                notification: {
                    channelId: "canal_estados_criticos", // 👈 REQUERIDO: Enruta la notificación en segundo plano [1.1.4, 1.1.6]
                    sound: "default"
                }
            }
        };

        const response = await admin.messaging().send(message);
        console.log(`📡 [PUSH] Alerta mixta enviada con éxito al tópico [${topic}] para UT ${viaje.numeroUt}. ID:`, response);
    } catch (error) {
        console.error(`❌ [PUSH ERROR] Error enviando alerta para UT ${viaje.numeroUt}:`, error.message);
    }
};
// Procesamiento de planillas con optimización de memoria
const obtenerViajesDePlanillaInterno = async (spreadsheetId) => {
    const doc = await initGoogleSheets(spreadsheetId);
    
    const sheetRuteo = doc.sheetsByTitle['Ruteo'];
    const sheetH12 = doc.sheetsByTitle['Hoja 12'];
    
    if (!sheetRuteo) return [];

    const hasH12 = !!sheetH12;
    const rowsH12 = hasH12 ? await sheetH12.getRows() : [];

    if (sheetRuteo.rowCount > 0) {
        try {
            await sheetRuteo.loadCells({
                startRowIndex: 0,
                endRowIndex: sheetRuteo.rowCount,
                startColumnIndex: 7,
                endColumnIndex: 8
            });
        } catch (e) {}
    }

    const rowsRuteo = await sheetRuteo.getRows();
    const headersRuteo = sheetRuteo.headerValues || [];
    const headersLimpio = headersRuteo.map(h => (h || "").toString().trim().toUpperCase());
    
    const regexTD = /^\d{7}$/;
    const regexFecha = /\d{1,2}\/\d{1,2}\/20\d{2}/;
    const regexHex = /^#[0-9A-Fa-f]{6}$/;

    const idxUt = headersLimpio.indexOf("N° UT") >= 0 ? headersLimpio.indexOf("N° UT") : headersLimpio.indexOf("N UT");
    const idxSemi = headersLimpio.indexOf("SEMI");
    const idxChofer = headersLimpio.indexOf("CHOFER");
    const idxTracking = headersLimpio.indexOf("TRACKING");
    const idxHexA = headersLimpio.indexOf("HEXA");
    const idxHexHx = headersLimpio.indexOf("HEXHX");
    
    const idxNViaje = headersLimpio.indexOf("N VIAJE");
    const idxLlegadaPlanta = headersLimpio.indexOf("LLEGADA A PLANTA");
    const idxVacio = headersLimpio.indexOf("VACIO"); 
    const idxDireccion = headersLimpio.indexOf("DIRECCION");
    const idxEstadoUt = headersLimpio.indexOf("ESTADOUT"); 

    const idxDestino = headersLimpio.indexOf("DESTINO");
    const idxProducto = headersLimpio.indexOf("PRODUCTO") >= 0 ? headersLimpio.indexOf("PRODUCTO") : 24;
    const idxCantidad = headersLimpio.indexOf("CANTIDAD") >= 0 ? headersLimpio.indexOf("CANTIDAD") : 25;
    const idxCisternado = headersLimpio.indexOf("CISTERNADO") >= 0 ? headersLimpio.indexOf("CISTERNADO") : 29;

    const idxAvisoVacio = headersLimpio.indexOf("AVISO DE VACIO");
    const idxLlegadaEta = headersLimpio.indexOf("LLEGADA(ETA)");
    const idxLlegada = headersLimpio.indexOf("LLEGADA");
    const limiteBusqueda = Math.max(idxAvisoVacio, idxLlegadaEta, idxLlegada, idxDestino - 6, 0);

    const agrupadoRuteo = {};
    for (const row of rowsRuteo) {
        const rVals = row._rawData || [];
        const rTdMatch = rVals.find(v => regexTD.test((v || "").toString().trim()));
        const rTd = rTdMatch ? rTdMatch.toString().trim() : "";
        if (rTd) {
            if (!agrupadoRuteo[rTd]) agrupadoRuteo[rTd] = [];
            agrupadoRuteo[rTd].push(rVals);
        }
    }

    const agrupadoH12 = {};
    for (const row of rowsH12) {
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

        const isTractorValido = trRuteo.length > 0 && 
                              trRuteo.length <= 12 && 
                              !trRuteo.toUpperCase().includes("TERMINAL") &&
                              !trRuteo.toUpperCase().includes("FECHA") &&
                              !trRuteo.toUpperCase().includes("PRODUCTO") &&
                              !trRuteo.startsWith("#") &&
                              !trRuteo.startsWith(",");

        const isTdValido = tdRuteo.length > 0;

        if (isTractorValido || isTdValido) {
            const keyMerge = tdRuteo ? tdRuteo.trim() : `FALLBACK_${spreadsheetId.substring(0, 8)}_${i}`;
            
            if (!procesados.has(keyMerge)) {
                procesados.add(keyMerge);

                let colorHexLegacy = null;
                try {
                    const bg = sheetRuteo.getCell(row.rowNumber - 1, 7).backgroundColor;
                    if (bg) {
                        colorHexLegacy = extraerHexDeGoogle(bg);
                        if (colorHexLegacy === '#FFFFFF' || colorHexLegacy === '#000000') colorHexLegacy = null;
                    }
                } catch (e) {}

                const numeroUt = idxUt >= 0 ? (vals[idxUt] || "").toString().trim() : "";
                const semi = idxSemi >= 0 ? (vals[idxSemi] || "").toString().trim() : "";
                const chofer = idxChofer >= 0 ? (vals[idxChofer] || "").toString().trim() : "";
                
                const rawTracking = idxTracking >= 0 ? (vals[idxTracking] || "").toString().trim() : "";
                const ultimoTracking = rawTracking ? rawTracking.split('/')[0].trim() : "";

                const colorHexAVal = idxHexA >= 0 ? (vals[idxHexA] || "").toString().trim() : "";
                const colorHexHxVal = idxHexHx >= 0 ? (vals[idxHexHx] || "").toString().trim() : "";

                const colorHexA = regexHex.test(colorHexAVal) ? colorHexAVal : null;
                const colorHexHx = regexHex.test(colorHexHxVal) ? colorHexHxVal : null;

                const nViaje = idxNViaje >= 0 ? (vals[idxNViaje] || "").toString().trim() : "";
                const llegadaPlanta = idxLlegadaPlanta >= 0 ? (vals[idxLlegadaPlanta] || "").toString().trim() : "";
                const horarioVacio = idxVacio >= 0 ? (vals[idxVacio] || "").toString().trim() : "";

                const isCompletado = idxVacio >= 0 && (vals[idxVacio] || "").toString().trim().length > 0;

                let estadoUt = idxEstadoUt >= 0 ? (vals[idxEstadoUt] || "").toString().trim() : "";
                if (isCompletado) {
                    estadoUt = "VACIO";
                }

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
                const filasRuteoDeEsteTD = agrupadoRuteo[tdRuteo] || [];
                
                const mapaDirecciones = {};
                filasRuteoDeEsteTD.forEach(rVals => {
                    const dest = (rVals[idxDestino] || "").toString().trim();
                    const addr = idxDireccion >= 0 ? (rVals[idxDireccion] || "").toString().trim() : "";
                    if (dest) mapaDirecciones[dest] = addr;
                });

                const paradas = [];
                let terminalLimpia = "Sin Terminal";
                let cisternadoReal = vals[idxCisternado] || "";

                if (filasH12.length > 0) {
                    const textoCompleto = filasH12[0].join(" ").toUpperCase();
                    terminalLimpia = determinarTerminal(textoCompleto);
                    cisternadoReal = filasH12[0][13] || cisternadoReal;

                    for (let f12 of filasH12) {
                        const destH12 = f12[7] || "Sin Destino";
                        const dir = mapaDirecciones[destH12] || ""; 
                        paradas.push({
                            destino: destH12,
                            producto: f12[8] || "",
                            amount: f12[9] || "",
                            cantidad: f12[9] || "",
                            cisternado: f12[10] || "",
                            direccion: dir
                        });
                    }
                } else {
                    const textoRuteo = vals.join(" ").toUpperCase();
                    terminalLimpia = determinarTerminal(textoRuteo);
                    const dir = idxDireccion >= 0 ? (vals[idxDireccion] || "").toString().trim() : "";

                    paradas.push({
                        destino: vals[idxDestino >= 0 ? idxDestino : 23] || "Validar Destino",
                        producto: vals[idxProducto] || "",
                        amount: vals[idxCantidad] || "",
                        cantidad: vals[idxCantidad] || "",
                        cisternado: "",
                        direccion: dir
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
                    numeroUt: numeroUt,
                    semi: semi,
                    chofer: chofer,
                    ultimoTracking: ultimoTracking,
                    colorHexA: colorHexA,
                    colorHexHx: colorHexHx,
                    nViaje: nViaje,
                    llegadaPlanta: llegadaPlanta,
                    horarioVacio: horarioVacio,
                    estadoUt: estadoUt
                });
            }
        }
    }
    return viajesFinales;
};

export const sincronizarDatosGoogleSheets = async () => {
    try {
        console.log("[BFF] 🔄 Iniciando ciclo de sincronización secuencial de Google Sheets...");
        
        const docMaster = await initGoogleSheets(masterIndexSheetId);
        const sheetIndice = docMaster.sheetsByTitle['Indice'];
        
        if (!sheetIndice) {
            console.error("❌ Sincronizador: No se encontró la pestaña 'Indice' en la planilla maestra.");
            return;
        }

        const rowsIndice = await sheetIndice.getRows();
        
        const parseFecha = (str) => {
            try {
                const parts = str.split("/");
                if (parts.length === 3) {
                    const d = parts[0].padStart(2, '0');
                    const m = parts[1].padStart(2, '0');
                    const y = parts[2].split(" ")[0].trim();
                    return parseInt(`${y}${m}${d}`, 10);
                }
            } catch (e) {}
            return 0;
        };

        const listaDiasCompleta = rowsIndice.map(row => {
            const vals = row._rawData || [];
            return {
                fecha: vals[0] || "Sin fecha",
                sheetId: vals[1] || ""
            };
        })
        .filter(d => d.sheetId.trim().length > 0)
        .sort((a, b) => parseFecha(b.fecha) - parseFecha(a.fecha));

        const listaDiasLimitada = listaDiasCompleta.slice(0, 10);
        const resultados = [];
        
        for (let j = 0; j < listaDiasLimitada.length; j++) {
            const dia = listaDiasLimitada[j];
            try {
                const viajesDia = await obtenerViajesDePlanillaInterno(dia.sheetId);
                resultados.push(viajesDia);
                await delay(1500); 
            } catch (err) {
                console.error(`[BFF ERROR] Falló la descarga de ${dia.fecha}:`, err.message);
            }
        }

        const todosLosViajes = resultados.flat();

        const vistos = new Set();
        const viajesConsolidados = [];
        for (const v of todosLosViajes) {
            if (!vistos.has(v.idUnico)) {
                vistos.add(v.idUnico);
                viajesConsolidados.push(v);
            }
        }
        
        viajesConsolidados.sort((a, b) => parseFecha(b.fechaPlanificada) - parseFecha(a.fechaPlanificada));

        if (cacheViajesConsolidados.length > 0 && firebaseInitialized) {
            for (const nuevo of viajesConsolidados) {
                const viejo = cacheViajesConsolidados.find(v => v.idUnico === nuevo.idUnico);
                if (viejo) {
                    const cambioVacio = viejo.horarioVacio !== nuevo.horarioVacio;
                    const cambioEstadoUt = viejo.estadoUt !== nuevo.estadoUt;
                    
                    if ((cambioVacio || cambioEstadoUt) && nuevo.numeroUt) {
                        enviarNotificacionPushUT(nuevo);
                    }
                }
            }
        }

        cacheViajesConsolidados = viajesConsolidados;
        cacheDiasDisponibles = listaDiasLimitada;
        lastSyncTime = new Date();

        console.log(`[BFF CACHE] ✅ Sincronización finalizada: ${cacheViajesConsolidados.length} viajes listos en memoria.`);
    } catch (error) {
        console.error("❌ ERROR CRÍTICO EN TRABAJO DE SEGUNDO PLANO DEL BFF. SE CONSERVA LA CACHÉ ANTERIOR:", error.message);
    }
};

setTimeout(() => {
    console.log("🚀 Iniciando primera carga del caché BFF en memoria (Modo Throttling)...");
    sincronizarDatosGoogleSheets();
}, 1000);

setInterval(sincronizarDatosGoogleSheets, 60000);

// =================================================================================================
// 👇 EXPORTS DE ENDPOINTS ULTRA-RÁPIDOS (< 5ms)
// =================================================================================================

export const getSheetData = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        const sheet = doc.sheetsByTitle[sheetName];
        if (!sheet) return res.status(404).json({ success: false, error: `La pestaña '${sheetName}' no existe.` });

        const rows = await sheet.getRows();
        const data = rows.map(row => [...(row._rawData || [])]);

        res.status(200).json({ success: true, headers: sheet.headerValues || [], data: data });
    } catch (error) {
        console.error("❌ ERROR LEYENDO PESTAÑA:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};

export const getViajesIntegradosDia = async (req, res) => {
    try {
        const { spreadsheetId } = req.params;
        const viajes = await obtenerViajesDePlanillaInterno(spreadsheetId);
        res.status(200).json({ success: true, data: viajes });
    } catch (error) {
        console.error("❌ ERROR INTEGRANDO VIAJES INDIVIDUALES:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};

export const getViajesRecientesAgregados = async (req, res) => {
    if (cacheViajesConsolidados.length === 0) {
        return res.status(202).json({
            success: true,
            message: "Sincronizador inicializando caché en segundo plano.",
            diasDisponibles: [],
            data: []
        });
    }

    res.status(200).json({
        success: true,
        diasDisponibles: cacheDiasDisponibles,
        data: cacheViajesConsolidados,
        cachedAt: lastSyncTime
    });
};

export const writeTestLog = async (req, res) => {
    res.status(200).json({ message: "Log guardado" });
};
