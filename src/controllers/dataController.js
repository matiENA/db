import { initGoogleSheets } from '../config/googleSheets.js';

// 🛠️ FUNCIÓN MAGICA: Convierte el color RGB de Google a HEX (#FFFFFF)
const extraerHexDeGoogle = (color) => {
    if (!color) return null;
    // Multiplicamos por 255 porque Google envía valores entre 0 y 1
    const r = Math.round((color.red || 0) * 255);
    const g = Math.round((color.green || 0) * 255);
    const b = Math.round((color.blue || 0) * 255);
    
    // Matemática binaria para convertir a Hexadecimal
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
        
        // 2. CARGAMOS LA ESTRUCTURA FÍSICA SOLO DE LA COLUMNA "H" (Para no saturar memoria)
        // La Columna H es la 8va letra.
        await sheet.loadCells(`H1:H${sheet.rowCount}`);

        const data = rows.map(row => {
            // Clonamos la fila cruda de textos
            const filaCruda = [...(row._rawData || [])]; 

            try {
                // sheet.getCell requiere un índice desde 0. 
                // row.rowNumber es la fila real de Excel (ej: 2). Columna H es índice 7.
                const celdaH = sheet.getCell(row.rowNumber - 1, 7);
                
                // Extraemos el color de fondo
                const colorFondo = celdaH.backgroundColor;

                if (colorFondo) {
                    const hex = extraerHexDeGoogle(colorFondo);
                    
                    // Si el color no es blanco puro (#FFFFFF) o negro vacío, lo inyectamos
                    if (hex !== '#FFFFFF' && hex !== '#000000') {
                        filaCruda.push(hex); // Se añade al final del array de esa fila
                    }
                }
            } catch (error) {
                // Silenciamos si la celda está vacía o sin formato
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
        res.status(500).json({ error: error.message });
    }
};

export const writeTestLog = async (req, res) => {
    res.status(200).json({ message: "Log guardado" });
};
