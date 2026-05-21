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
        
        // 👇 NUEVA LÓGICA: Solo cargamos la Columna H si NO estamos leyendo el Indice
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
