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
        console.log("🔥 [FIREBASE] Admin SDK inicializado exitosamente.");
    } else {
        console.warn("⚠️ [FIREBASE] No se encontró certificado. Notificaciones push desactivadas.");
    }
} catch (err) {
    console.error("❌ [FIREBASE] Falló inicialización:", err.message);
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
            android: { priority: "high" }
        };

        await admin.messaging().send(message);
    } catch (error) {
        console.error(`❌ [PUSH ERROR] UT ${viaje.numeroUt}:`, error.message);
    }
};

const obtenerViajesDePlanillaInterno = async (spreadsheetId) => {
    const doc = await initGoogleSheets(spreadsheetId);
    
    const sheetRuteo = doc.sheetsByTitle['Ruteo'];
    const sheetH12 = doc.sheetsByTitle['Hoja 12'];
    if (!sheetRuteo) return [];

    const rowsH12 = !!sheetH12 ? await sheetH12.getRows() : [];

    if (sheetRuteo.rowCount > 0) {
        try {
            await sheetRuteo.loadCells({ startRowIndex: 0, endRowIndex: sheetRuteo.rowCount, startColumnIndex: 7, endColumnIndex: 8 });
        } catch (e) {}
    }

    const rowsRuteo = await sheetRuteo.getRows();
    const headersLimpio = (sheetRuteo.headerValues || []).map(h => (h || "").toString().trim().toUpperCase());
    
    const regexTD = /^\d{7}$/;
    const regexFecha = /\d{1,2}\/\d{1,2}\/20\d{2}/;
    const regexHex = /^#[0-9A-Fa-f]{6}$/;

    // Índices de columnas
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

    // 1. Agrupación Hoja 12
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

    // 2. AGRUPACIÓN TOTAL DE RUTEO (El corazón de la solución para fusionar viajes 5, 6, etc.)
    const viajesAgrupadosMap = new Map();
    let currentFallbackGroup = 0;
    let lastFallbackUt = "";

    for (let i = 0; i < rowsRuteo.length; i++) {
        const row = rowsRuteo[i];
        const vals = row._rawData || [];
        if (vals.length === 0) continue;

        const trRuteo = (vals[0] || "").toString().trim();
        const isTractorValido = trRuteo.length > 0 && trRuteo.length <= 12 && !trRuteo.toUpperCase().includes("TERMINAL") && !trRuteo.startsWith("#");
        
        const tdMatch = vals.find(v => regexTD.test((v || "").toString().trim()));
        const tdRuteo = tdMatch ? tdMatch.toString().trim() : "";
        if (!isTractorValido && !tdRuteo) continue;

        const numeroUt = idxUt >= 0 ? (vals[idxUt] || "").toString().trim() : "";
        const chofer = idxChofer >= 0 ? (vals[idxChofer] || "").toString().trim() : "";
        const nViaje = idxNViaje >= 0 ? (vals[idxNViaje] || "").toString().trim() : "";

        if (chofer || nViaje || (numeroUt && numeroUt !== lastFallbackUt)) {
            currentFallbackGroup++;
            if (numeroUt) lastFallbackUt = numeroUt;
        }

        const keyMerge = tdRuteo ? tdRuteo.trim() : `FALLBACK_${numeroUt}_${currentFallbackGroup}`;

        if (!viajesAgrupadosMap.has(keyMerge)) {
            viajesAgrupadosMap.set(keyMerge, { rows: [], cellRowIndex: row.rowNumber - 1 });
        }
        viajesAgrupadosMap.get(keyMerge).rows.push(vals);
    }

    // 3. CONSTRUCCIÓN DE VIAJES INTEGADOS
    const viajesFinales = [];

    for (const [keyMerge, dataGroup] of viajesAgrupadosMap.entries()) {
        const filasRuteoDeEsteTD = dataGroup.rows;
        const valsPrincipales = filasRuteoDeEsteTD[0]; // Fila madre
        
        const trRuteo = (valsPrincipales[0] || "").toString().trim();
        const tdMatch = valsPrincipales.find(v => regexTD.test((v || "").toString().trim()));
        const tdRuteo = tdMatch ? tdMatch.toString().trim() : "";

        let colorHexLegacy = null;
        try {
            const bg = sheetRuteo.getCell(dataGroup.cellRowIndex, 7).backgroundColor;
            if (bg) {
                colorHexLegacy = extraerHexDeGoogle(bg);
                if (colorHexLegacy === '#FFFFFF' || colorHexLegacy === '#000000') colorHexLegacy = null;
            }
        } catch (e) {}

        const numeroUt = idxUt >= 0 ? (valsPrincipales[idxUt] || "").toString().trim() : "";
        const semi = idxSemi >= 0 ? (valsPrincipales[idxSemi] || "").toString().trim() : "";
        const chofer = idxChofer >= 0 ? (valsPrincipales[idxChofer] || "").toString().trim() : "";
        
        const rawTracking = idxTracking >= 0 ? (valsPrincipales[idxTracking] || "").toString().trim() : "";
        const ultimoTracking = rawTracking ? rawTracking.split('/')[0].trim() : "";

        const colorHexAVal = idxHexA >= 0 ? (valsPrincipales[idxHexA] || "").toString().trim() : "";
        const colorHexHxVal = idxHexHx >= 0 ? (valsPrincipales[idxHexHx] || "").toString().trim() : "";
        const nViaje = idxNViaje >= 0 ? (valsPrincipales[idxNViaje] || "").toString().trim() : "";
        const llegadaPlanta = idxLlegadaPlanta >= 0 ? (valsPrincipales[idxLlegadaPlanta] || "").toString().trim() : "";

        // Escaneo profundo: Si alguna fila de este viaje marcó vacío, todo el viaje es completado
        let isCompletado = false;
        let horarioVacio = "";
        filasRuteoDeEsteTD.forEach(rVals => {
            const rowCompletado = idxVacio >= 0 && (rVals[idxVacio] || "").toString().trim().length > 0;
            if (rowCompletado) {
                isCompletado = true;
                if (!horarioVacio) horarioVacio = (rVals[idxVacio] || "").toString().trim();
            }
        });

        let estadoUt = idxEstadoUt >= 0 ? (valsPrincipales[idxEstadoUt] || "").toString().trim() : "";
        if (isCompletado) estadoUt = "VACIO";

        let rawFecha = "";
        if (idxDestino > 0) {
            for (let j = idxDestino - 1; j > limiteBusqueda; j--) {
                const valText = (valsPrincipales[j] || "").toString();
                if (regexFecha.test(valText)) {
                    rawFecha = valText;
                    break;
                }
            }
        }
        if (!rawFecha && headersLimpio.indexOf("FECHA PLANIFICADA") >= 0) {
            rawFecha = valsPrincipales[headersLimpio.indexOf("FECHA PLANIFICADA")] || "";
        }

        const filasH12 = tdRuteo ? (agrupadoH12[tdRuteo] || []) : [];
        const mapaDirecciones = {};
        filasRuteoDeEsteTD.forEach(rVals => {
            const dest = (rVals[idxDestino] || "").toString().trim();
            const addr = idxDireccion >= 0 ? (rVals[idxDireccion] || "").toString().trim() : "";
            if (dest) mapaDirecciones[dest] = addr;
        });

        const paradas = [];
        let terminalLimpia = "Sin Terminal";
        let cisternadoReal = valsPrincipales[idxCisternado] || "";

        if (filasH12.length > 0) {
            const textoCompleto = filasH12[0].join(" ").toUpperCase();
            terminalLimpia = determinarTerminal(textoCompleto);
            cisternadoReal = filasH12[0][13] || cisternadoReal;

            for (let f12 of filasH12) {
                const destH12 = f12[7] || "Sin Destino";
                paradas.push({
                    destino: destH12,
                    producto: f12[8] || "",
                    cantidad: f12[9] || "",
                    cisternado: f12[10] || "",
                    direccion: mapaDirecciones[destH12] || ""
                });
            }
        } else {
            terminalLimpia = determinarTerminal(valsPrincipales.join(" ").toUpperCase());
            // Map de las filas secundarias reales (Resuelve el bug 5/6)
            filasRuteoDeEsteTD.forEach(rVals => {
                paradas.push({
                    destino: rVals[idxDestino >= 0 ? idxDestino : 23] || "Validar Destino",
                    producto: rVals[idxProducto] || "",
                    cantidad: rVals[idxCantidad] || "",
                    cisternado: "",
                    direccion: idxDireccion >= 0 ? (rVals[idxDireccion] || "").toString().trim() : ""
                });
            });
        }

        viajesFinales.push({
            idUnico: keyMerge,
            tractor: trRuteo,
            numDespacho: tdRuteo,
            terminalOrigen: terminalLimpia,
            fechaPlanificada: rawFecha.split(" ")[0] || "Sin Fecha",
            cisternadoReal: cisternadoReal,
            colorHex: colorHexLegacy,
            isCompletado: isCompletado,
            paradas: paradas,
            numeroUt: numeroUt,
            semi: semi,
            chofer: chofer,
            ultimoTracking: ultimoTracking,
            colorHexA: regexHex.test(colorHexAVal) ? colorHexAVal : null,
            colorHexHx: regexHex.test(colorHexHxVal) ? colorHexHxVal : null,
            nViaje: nViaje,
            llegadaPlanta: llegadaPlanta,
            horarioVacio: horarioVacio,
            estadoUt: estadoUt
        });
    }

    return viajesFinales;
};

export const sincronizarDatosGoogleSheets = async () => {
    try {
        console.log("[BFF] 🔄 Iniciando ciclo de sincronización...");
        const docMaster = await initGoogleSheets(masterIndexSheetId);
        const sheetIndice = docMaster.sheetsByTitle['Indice'];
        
        if (!sheetIndice) return;

        const rowsIndice = await sheetIndice.getRows();
        const parseFecha = (str) => {
            try {
                const parts = str.split("/");
                if (parts.length === 3) return parseInt(`${parts[2].split(" ")[0].trim()}${parts[1].padStart(2, '0')}${parts[0].padStart(2, '0')}`, 10);
            } catch (e) {} return 0;
        };

        const listaDiasLimitada = rowsIndice.map(row => ({ fecha: row._rawData[0] || "", sheetId: row._rawData[1] || "" }))
            .filter(d => d.sheetId.trim().length > 0).sort((a, b) => parseFecha(b.fecha) - parseFecha(a.fecha)).slice(0, 10);

        const resultados = [];
        for (const dia of listaDiasLimitada) {
            try {
                resultados.push(await obtenerViajesDePlanillaInterno(dia.sheetId));
                await delay(1500); 
            } catch (err) { console.error(`[BFF ERROR] ${dia.fecha}:`, err.message); }
        }

        const viajesConsolidados = [];
        const vistos = new Set();
        for (const v of resultados.flat()) {
            if (!vistos.has(v.idUnico)) { vistos.add(v.idUnico); viajesConsolidados.push(v); }
        }
        viajesConsolidados.sort((a, b) => parseFecha(b.fechaPlanificada) - parseFecha(a.fechaPlanificada));

        if (cacheViajesConsolidados.length > 0 && firebaseInitialized) {
            for (const nuevo of viajesConsolidados) {
                const viejo = cacheViajesConsolidados.find(v => v.idUnico === nuevo.idUnico);
                if (viejo && ((viejo.horarioVacio !== nuevo.horarioVacio) || (viejo.estadoUt !== nuevo.estadoUt)) && nuevo.numeroUt) {
                    enviarNotificacionPushUT(nuevo);
                }
            }
        }

        cacheViajesConsolidados = viajesConsolidados;
        cacheDiasDisponibles = listaDiasLimitada;
        lastSyncTime = new Date();
        console.log(`[BFF CACHE] ✅ Sincronización finalizada: ${cacheViajesConsolidados.length} viajes.`);
    } catch (error) { console.error("❌ ERROR CRÍTICO EN TRABAJO DE SEGUNDO PLANO:", error.message); }
};

const iniciarWorkerLogistico = async () => {
    await delay(3000); 
    while (true) {
        await sincronizarDatosGoogleSheets();
        await delay(60000); 
    }
};
iniciarWorkerLogistico();

export const getSheetData = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        const sheet = doc.sheetsByTitle[sheetName];
        if (!sheet) return res.status(404).json({ success: false, error: `La pestaña '${sheetName}' no existe.` });
        res.status(200).json({ success: true, headers: sheet.headerValues || [], data: (await sheet.getRows()).map(row => [...(row._rawData || [])]) });
    } catch (error) { res.status(500).json({ success: false, error: error.message }); }
};

export const getViajesIntegradosDia = async (req, res) => {
    try {
        res.status(200).json({ success: true, data: await obtenerViajesDePlanillaInterno(req.params.spreadsheetId) });
    } catch (error) { res.status(500).json({ success: false, error: error.message }); }
};

export const getViajesRecientesAgregados = async (req, res) => {
    if (cacheViajesConsolidados.length === 0) {
        return res.status(202).json({ success: true, message: "Inicializando caché", diasDisponibles: [], data: [] });
    }
    
    // Devolvemos el array original de `data` para ruteo1, y también `activos` y `completados` para ruteoADI
    res.status(200).json({
        success: true,
        diasDisponibles: cacheDiasDisponibles,
        data: cacheViajesConsolidados, 
        activos: cacheViajesConsolidados.filter(v => !v.isCompletado),
        completados: cacheViajesConsolidados.filter(v => v.isCompletado),
        cachedAt: lastSyncTime
    });
};

export const writeTestLog = async (req, res) => { res.status(200).json({ message: "Log guardado" }); };
