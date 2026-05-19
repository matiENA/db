import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import dataRoutes from './routes/dataRoutes.js';

dotenv.config();

const app = express();

// Permisos generales
app.use(cors({ origin: '*' }));
app.use(express.json());

// 1. PRIMERO: Servimos el Frontend (Nuestra web unificada)
app.use(express.static('public'));

// 2. SEGUNDO: Rutas del Backend (El motor de Excel)
app.use('/api', dataRoutes);

// 3. TERCERO: Escudo de seguridad 404 (Si no existe, bloquea)
app.use((req, res) => {
    res.status(404).json({ success: false, message: 'El endpoint solicitado no existe.' });
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n======================================================`);
    console.log(`🚀 SISTEMA RUTEO UNIFICADO ENCENDIDO`);
    console.log(`🌍 Dashboard Web: http://localhost:${PORT}`);
    console.log(`======================================================\n`);
});