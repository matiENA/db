import { Router } from 'express';
import { getSheetData, writeTestLog, getViajesIntegradosDia, getViajesRecientesAgregados } from '../controllers/dataController.js';

const router = Router();

router.get('/sheet/:spreadsheetId/:sheetName', getSheetData);
router.post('/logs', writeTestLog);
router.get('/viajes-integrados/:spreadsheetId', getViajesIntegradosDia);

// 👇 NUEVO ENDPOINT AGREGADOR DE VIAJES RECIENTES
router.get('/viajes-recientes/:masterIndexSheetId', getViajesRecientesAgregados);

export default router;
