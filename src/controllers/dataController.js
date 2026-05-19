import { initGoogleSheets } from '../config/googleSheets.js';

export const getSheetData = async (req, res) => {
    try {
        const { spreadsheetId, sheetName } = req.params;
        const doc = await initGoogleSheets(spreadsheetId);
        
        const sheet = doc.sheetsByTitle[sheetName];
        if (!sheet) {
            return res.status(404).json({ error: `La pestaña '${sheetName}' no existe.` });
        }

        const rows = await sheet.getRows();
        
        // SOLUCIÓN ESTRUCTURAL: Extraemos el array crudo exacto de cada fila.
        // Esto elimina cualquier desfasaje vertical u horizontal causado por celdas vacías.
        const data = rows.map(row => row._rawData || []);

        res.status(200).json({ 
            success: true, 
            headers: sheet.headerValues || [],
            data: data 
        });
    } catch (error) {
        console.error("❌ ERROR LEYENDO PESTAÑA:", error);
        res.status(500).json({ error: error.message });
    }
};

export const writeTestLog = async (req, res) => {
    res.status(200).json({ message: "Log guardado" });
};