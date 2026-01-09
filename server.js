const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');
const { createCanvas } = require('canvas');
const d3Cloud = require('d3-cloud');

const app = express();
const PORT = process.env.PORT || 3000;

// --- CONFIGURACIÓN DE SEGURIDAD (CORS) ---
// Permite que tu página web se comunique con este servidor
app.use(cors({
    origin: '*', // En producción, idealmente pon aquí tu dominio: 'https://tu-pagina.com'
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type']
}));

app.use(express.json());

// --- CONFIGURACIÓN DE CARGA DE ARCHIVOS ---
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// --- LISTA DE PALABRAS A IGNORAR EN NUBES ---
const STOPWORDS = [
    'de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como',
    'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde',
    'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola',
    'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal'
];

// --- FUNCIONES DE AYUDA ---

// 1. Limpiar y separar palabras
function getWordsFromString(text) {
    if (!text || typeof text !== 'string') return [];
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, "") // Quitar puntuación
        .match(/\b(\w+)\b/g) // Buscar palabras
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

// 2. Calcular NPS
function calculateNPS(promotores, pasivos, detractores) {
    const total = promotores + pasivos + detractores;
    if (total === 0) return 0;
    return parseFloat((((promotores - detractores) / total) * 100).toFixed(1));
}

// 3. Detectar Fecha desde Excel
function parseExcelDate(value) {
    if (!value) return null;
    // Si es objeto fecha de JS
    if (value instanceof Date) return value;
    // Si es string "dd/mm/yyyy"
    if (typeof value === 'string') {
        const parts = value.split('/');
        if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    // Si es número serial de Excel (raro pero posible)
    return new Date(value);
}

// 4. Generar Nube de Palabras (Devuelve Base64)
function generarNubeImagen(wordList, colorHex) {
    if (!wordList || wordList.length === 0) return Promise.resolve(null);

    return new Promise(resolve => {
        const width = 800;
        const height = 400;
        const canvas = createCanvas(width, height);
        const ctx = canvas.getContext('2d');

        // Normalizar tamaños
        const maxWeight = Math.max(...wordList.map(w => w[1]));
        const minWeight = Math.min(...wordList.map(w => w[1]));

        const words = wordList.slice(0, 60).map(item => ({
            text: item[0],
            size: 20 + 60 * ((item[1] - minWeight) / ((maxWeight - minWeight) || 1))
        }));

        const layout = d3Cloud()
            .size([width, height])
            .canvas(() => createCanvas(1, 1))
            .words(words)
            .padding(5)
            .rotate(() => (Math.random() > 0.5 ? 0 : 90))
            .font('sans-serif')
            .fontSize(d => d.size)
            .on('end', (drawnWords) => {
                ctx.fillStyle = '#ffffff'; // Fondo blanco
                ctx.fillRect(0, 0, width, height);
                
                ctx.translate(width / 2, height / 2);
                drawnWords.forEach(w => {
                    ctx.save();
                    ctx.translate(w.x, w.y);
                    ctx.rotate(w.rotate * Math.PI / 180);
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillStyle = colorHex; // Color pasado por parámetro
                    ctx.font = `${w.size}px sans-serif`;
                    ctx.fillText(w.text, 0, 0);
                    ctx.restore();
                });
                
                resolve(canvas.toDataURL().split(',')[1]); // Devuelve solo el string base64
            });

        layout.start();
    });
}

// --- RUTA PRINCIPAL ---
app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        console.log("Recibiendo solicitud de procesamiento anual...");

        // 1. Validar Archivos y Datos
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'Falta el archivo Excel' });
        }
        
        let datosManuales;
        try {
            datosManuales = JSON.parse(req.body.datosManuales || '{}');
        } catch (e) {
            datosManuales = { enero: {}, febrero: {} };
        }

        // 2. Leer Excel
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        // 3. Mapear Columnas (Buscar donde está "Sector", "Fecha", etc.)
        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            if (cell.value) {
                const val = cell.value.toString().toLowerCase().trim();
                if (val.includes('fecha')) colMap.fecha = colNumber;
                if (val.includes('hora')) colMap.hora = colNumber;
                if (val.includes('sector')) colMap.sector = colNumber;
                if (val.includes('calificaci')) colMap.calificacion = colNumber; // "calificacion" o "calificación"
                if (val.includes('comentario')) colMap.comentario = colNumber;
            }
        });

        if (!colMap.fecha || !colMap.calificacion) {
            return res.status(400).json({ success: false, message: 'El Excel no tiene columnas de "Fecha" o "Calificación".' });
        }

        // 4. Estructura de Datos para almacenar todo
        // sectorsData = { "Cajas": { 0: {mp:1, p:0...}, 1: {...}, palabrasPos: [], palabrasNeg: [], horas: [] } }
        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return; // Saltar cabecera

            const rawDate = row.getCell(colMap.fecha).value;
            const date = parseExcelDate(rawDate);
            if (!date) return;

            const monthIndex = date.getMonth(); // 0 = Enero, 11 = Diciembre
            // Solo procesamos de Marzo (2) en adelante desde el Excel, ya que Ene/Feb son manuales
            if (monthIndex < 2) return; 

            const sectorName = colMap.sector ? (row.getCell(colMap.sector).value || 'General').toString().trim() : 'General';
            const calificacion = (row.getCell(colMap.calificacion).value || '').toString().toLowerCase();
            const comentario = colMap.comentario ? (row.getCell(colMap.comentario).value || '').toString() : '';
            
            // Inicializar sector si no existe
            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    palabrasPositivas: [],
                    palabrasNegativas: [],
                    horasCriticas: Array(24).fill(0)
                };
            }

            const stats = sectorsData[sectorName].meses[monthIndex];
            stats.total++;

            // Clasificar
            if (calificacion.includes('muy positiva')) {
                stats.mp++;
                if (comentario) sectorsData[sectorName].palabrasPositivas.push(...getWordsFromString(comentario));
            } else if (calificacion.includes('positiva')) {
                stats.p++;
                if (comentario) sectorsData[sectorName].palabrasPositivas.push(...getWordsFromString(comentario));
            } else if (calificacion.includes('muy negativa')) {
                stats.mn++;
                if (comentario) sectorsData[sectorName].palabrasNegativas.push(...getWordsFromString(comentario));
                // Registrar hora para detectar horas críticas
                if (colMap.hora) {
                    const rawTime = row.getCell(colMap.hora).value;
                    let hour = 12; // Default
                    if (rawTime instanceof Date) hour = rawTime.getUTCHours();
                    else if (typeof rawTime === 'string') hour = parseInt(rawTime.split(':')[0]) || 12;
                    if(hour >= 0 && hour < 24) sectorsData[sectorName].horasCriticas[hour]++;
                }
            } else if (calificacion.includes('negativa')) {
                stats.n++;
                if (comentario) sectorsData[sectorName].palabrasNegativas.push(...getWordsFromString(comentario));
            }
        });

        // 5. Procesar Datos Finales y Unir con Manuales
        const resultadoFinal = [];

        for (const [nombreSector, data] of Object.entries(sectorsData)) {
            
            // A. Insertar Datos Manuales (Enero y Febrero)
            // Se asume que los datos manuales ingresados aplican al sector que estamos procesando
            if (datosManuales.enero) {
                data.meses[0] = {
                    mp: datosManuales.enero.muy_positivas || 0,
                    p: datosManuales.enero.positivas || 0,
                    n: datosManuales.enero.negativas || 0,
                    mn: datosManuales.enero.muy_negativas || 0,
                    total: datosManuales.enero.total || 0
                };
            }
            if (datosManuales.febrero) {
                data.meses[1] = {
                    mp: datosManuales.febrero.muy_positivas || 0,
                    p: datosManuales.febrero.positivas || 0,
                    n: datosManuales.febrero.negativas || 0,
                    mn: datosManuales.febrero.muy_negativas || 0,
                    total: datosManuales.febrero.total || 0
                };
            }

            // B. Calcular NPS por mes
            const mesesFormateados = data.meses.map((m, i) => {
                const promotores = m.mp;
                const detractores = m.n + m.mn;
                const pasivos = m.p; // Consideramos "positiva" como pasiva o promotora? 
                // ESTÁNDAR NPS: Muy Positiva (9-10) = Promotor, Positiva (7-8) = Pasivo, Neg/MuyNeg (0-6) = Detractor
                // Ajuste según tu lógica de negocio:
                // Si consideras "Positiva" como buena, el cálculo varía. Aquí uso estándar estricto:
                // Asumiendo: MP=Promotor, P=Pasivo, N/MN=Detractor
                
                return {
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    nps: calculateNPS(m.mp, m.p, m.n + m.mn),
                    total: m.total
                };
            });

            // C. Generar Nubes
            // Contar frecuencia de palabras
            const contar = (arr) => {
                const counts = {};
                arr.forEach(w => counts[w] = (counts[w] || 0) + 1);
                return Object.entries(counts).sort((a, b) => b[1] - a[1]);
            };

            const topPositivas = contar(data.palabrasPositivas);
            const topNegativas = contar(data.palabrasNegativas);

            const imgPos = await generarNubeImagen(topPositivas, '#43a047'); // Verde
            const imgNeg = await generarNubeImagen(topNegativas, '#d32f2f'); // Rojo

            // D. Calcular Hora Crítica (Hora con más quejas)
            let maxQuejas = 0;
            let horaCritica = null;
            data.horasCriticas.forEach((count, h) => {
                if (count > maxQuejas) {
                    maxQuejas = count;
                    horaCritica = h;
                }
            });

            // E. Métricas Anuales
            const totalAnual = data.meses.reduce((acc, curr) => acc + curr.total, 0);
            const npsPromedio = (mesesFormateados.reduce((acc, curr) => acc + curr.nps, 0) / 12).toFixed(1);

            resultadoFinal.push({
                nombre: nombreSector,
                meses: mesesFormateados,
                nubePositivaB64: imgPos,
                nubeNegativaB64: imgNeg,
                horaCritica: horaCritica,
                palabrasNegativas: topNegativas.slice(0, 5).map(x => x[0]), // Top 5 palabras para la IA
                totalAnual: totalAnual,
                npsPromedio: npsPromedio
            });
        }

        res.json({ success: true, data: { sectores: resultadoFinal } });

    } catch (error) {
        console.error('Error en servidor:', error);
        res.status(500).json({ success: false, message: 'Error interno: ' + error.message });
    }
});

// --- RUTA DE SALUD (Para que Render sepa que está vivo) ---
app.get('/health', (req, res) => res.send('OK'));

app.listen(PORT, () => {
    console.log(`✅ Servidor escuchando en puerto ${PORT}`);
});