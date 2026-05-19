import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const serviceAccountAuth = new JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

// AHORA RECIBE EL ID DINÁMICAMENTE COMO PARÁMETRO
export const initGoogleSheets = async (spreadsheetId) => {
    try {
        const doc = new GoogleSpreadsheet(spreadsheetId, serviceAccountAuth);
        await doc.loadInfo();
        return doc;
    } catch (error) {
        console.error('❌ Error conectando al archivo:', error.message);
        throw error;
    }
};