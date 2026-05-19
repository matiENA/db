import { Router } from 'express';
import { getSheetData, writeTestLog } from '../controllers/dataController.js';

const router = Router();

// NUEVA RUTA ESCALABLE: Acepta el ID de la planilla y la pestaña
router.get('/sheet/:spreadsheetId/:sheetName', getSheetData);
router.post('/logs', writeTestLog);

export default router;