import { initGoogleSheets } from '../config/googleSheets.js';
import admin from 'firebase-admin';
import fs from 'fs';

// Inicialización de la Caché Global en Memoria del BFF
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
        const lineaEstado = (viaje.horarioVacio && viaje.horarioVacio.trim().length > 0)
            ? `VACIO: ${viaje.horarioVacio.trim()}`
            : (viaje.estadoUt ? viaje.estadoUt.trim().toUpperCase() : "PENDIENTE");

        const message = {
            notification: {
                title: `[UT: ${viaje.numeroUt || "S/D"}] ${viaje.chofer || "Chofer S/D"}`,
                body: `${viaje.tractor || "S/D"} | ${viaje.semi || "S/D"} | Viaje: ${viaje.nViaje || "S/D"}\nTD: ${viaje.numDespacho || "S/N"}\n${lineaEstado}`
            },
            data: {
                idUnico: (viaje.idUnico || "").toString(),
                tractor: (viaje.tractor || "").toString(),
                numDespacho: (viaje.numDespacho || "").toString(),
                terminalOrigen: (viaje.terminalOrigen || "").toString(),
                fechaPlanificada: (viaje.fechaPlanificada || "").toString(),
                numeroUt: (viaje.numeroUt || "").toString(),
                estadoUt: (viaje.estadoUt || "").toString()
            },
            topic: topic,
            android: { priority: "high" }
        };

        const response = await admin.messaging().send(message);
        console.log(`📡 [PUSH] Alerta mixta enviada con éxito al tópico [${topic}] para UT ${viaje.numeroUt}. ID:`, response);
    } catch (error) {
        console.error(`❌ [PUSH ERROR] Error enviando alerta para UT ${viaje.numeroUt}:`, error.message);
    }
};

const obtenerViajesDePlanillaInterno = async (spreadsheetId) => {
    const doc = await initGoogleSheets(spreadsheetId);
    
    // Si la hoja se llama diferente en tu archivo real, asegúrate de que esto coincida
    const sheetRuteo = doc.sheetsByTitle['Ruteo'] || doc.sheetsByIndex[0];
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
    // 👇 AQUÍ ATRAPAMOS LA COLUMNA CON LOS NOMBRES LIMPIOS 👇
    const idxDestinosAR = headersLimpio.indexOf("DESTINOS") >= 0 ? headersLimpio.indexOf("DESTINOS") : headersLimpio.indexOf("DESTINOAR");
    
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

        if (isTractorValido || tdRuteo.length > 0) {
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
                if (isCompletado) estadoUt = "VACIO";

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

                // 👇 EXTRAEMOS EL NOMBRE LIMPIO A NIVEL VIAJE 👇
                const destinoLimpio = idxDestinosAR >= 0 ? (vals[idxDestinosAR] || "").toString().trim() : "";

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
                            destinoAR: destinoLimpio, // Inyectado
                            producto: f12[8] || "",
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
                        destinoAR: destinoLimpio, // Inyectado
                        producto: vals[idxProducto] || "",
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
                    colorHex
