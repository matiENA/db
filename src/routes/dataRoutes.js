import { Router } from 'express';
import { getSheetData, writeTestLog, getViajesIntegradosDia } from '../controllers/dataController.js';

const router = Router();

router.get('/sheet/:spreadsheetId/:sheetName', getSheetData);
router.post('/logs', writeTestLog);

// 👇 LA NUEVA RUTA PARA ANDROID
router.get('/viajes-integrados/:spreadsheetId', getViajesIntegradosDia);

export default router;
