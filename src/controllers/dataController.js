// 👇 EXPORT AGREGADOR CON FILTRADO DINÁMICO DE CARGA PROGRESIVA (LAZY FETCHING) [txt]
export const getViajesRecientesAgregados = async (req, res) => {
    try {
        const { masterIndexSheetId } = req.params;
        
        // Mapeo dinámico del query parameter de límites
        const limitParam = parseInt(req.query.limit, 10);
        const limit = !isNaN(limitParam) && limitParam > 0 ? limitParam : 10;

        const docMaster = await initGoogleSheets(masterIndexSheetId);
        const sheetIndice = docMaster.sheetsByTitle['Indice'];
        
        if (!sheetIndice) {
            return res.status(404).json({ success: false, error: "No se encontró la pestaña 'Indice' en la planilla maestra." });
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

        // 👇 MANTENER ARQUITECTURA CRÍTICA: Extrae la Columna B (sheetId) de forma normal [txt]
        const listaDiasCompleta = rowsIndice.map(row => {
            const vals = row._rawData || [];
            return {
                fecha: vals[0] || "Sin fecha",
                sheetId: vals[1] || ""
            };
        })
        .filter(d => d.sheetId.trim().length > 0)
        .sort((a, b) => parseFecha(b.fecha) - parseFecha(a.fecha));

        // 👇 LAZY FETCHING: Corte y delimitación dinámico en el backend [txt]
        const listaDiasLimitada = listaDiasCompleta.slice(0, limit);

        // Procesamiento en paralelo de los días delimitados
        const promesasViajes = listaDiasLimitada.map(async (dia) => {
            try {
                return await obtenerViajesDePlanillaInterno(dia.sheetId);
            } catch (err) {
                console.error(`Error procesando planilla ID ${dia.sheetId}:`, err.message);
                return [];
            }
        });

        const resultados = await Promise.all(promesasViajes);
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

        res.status(200).json({
            success: true,
            diasDisponibles: listaDiasLimitada,
            data: viajesConsolidados
        });
    } catch (error) {
        console.error("❌ ERROR EN EL AGREGADOR BFF CON LIMIT:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};
