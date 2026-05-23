import { Router } from 'express';
import { getSheetData, writeTestLog, getViajesIntegradosDia } from '../controllers/dataController.js';

const router = Router();

// RUTA CLÁSICA (Para el Índice)
router.get('/sheet/:spreadsheetId/:sheetName', getSheetData);
router.post('/logs', writeTestLog);

// 👇 LA NUEVA RUTA PARA ANDROID (Microservicio BFF)
router.get('/viajes-integrados/:spreadsheetId', getViajesIntegradosDia);

// SOLO DEBE HABER UN EXPORT AL FINAL
export default router;
