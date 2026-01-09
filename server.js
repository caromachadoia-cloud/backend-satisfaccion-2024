const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');
const { createCanvas } = require('canvas');
const d3Cloud = require('d3-cloud');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: '*', methods: ['GET', 'POST'], allowedHeaders: ['Content-Type'] }));
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Palabras a ignorar (Stopwords)
const STOPWORDS = [
    'de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como',
    'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde',
    'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola',
    'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'yyyy', 'todo', 
    'nada', 'nadie', 'gente', 'cosas', 'porque', 'estan'
];

function getWordsFromString(text) {
    if (!text || typeof text !== 'string') return [];
    // Limpieza más agresiva para evitar basura
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, "")
        .replace(/\s{2,}/g, " ")
        .match(/\b(\w+)\b/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

function calculateNPS(promotores, pasivos, detractores) {
    const total = promotores + pasivos + detractores;
    if (total === 0) return 0;
    const nps = ((promotores - detractores) / total) * 100;
    return parseFloat(nps.toFixed(1));
}

// Generador de Nube Robusto (Usa fuente del sistema)
function generarNubeImagen(wordList, colorHex) {
    if (!wordList || wordList.length === 0) return Promise.resolve(null);
    
    return new Promise(resolve => {
        const width = 600; 
        const height = 400;
        const canvas = createCanvas(width, height);
        const ctx = canvas.getContext('2d');
        
        // Fondo blanco para que se vea en PDF
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);

        // Normalización de tamaños
        const maxWeight = Math.max(...wordList.map(w => w[1]));
        const minWeight = Math.min(...wordList.map(w => w[1]));
        
        // Tomamos hasta 60 palabras
        const words = wordList.slice(0, 60).map(item => ({
            text: item[0],
            size: 20 + 60 * ((item[1] - minWeight) / ((maxWeight - minWeight) || 1))
        }));

        d3Cloud()
            .size([width, height])
            .canvas(() => createCanvas(1, 1))
            .words(words)
            .padding(10)
            .rotate(() => 0) // Sin rotación para evitar problemas de renderizado
            .font('sans-serif') // FUENTE SEGURA PARA LINUX/RENDER
            .fontSize(d => d.size)
            .on('end', (drawnWords) => {
                ctx.translate(width / 2, height / 2);
                drawnWords.forEach(w => {
                    ctx.save();
                    ctx.translate(w.x, w.y);
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillStyle = colorHex;
                    ctx.font = `bold ${w.size}px sans-serif`;
                    ctx.fillText(w.text, 0, 0);
                    ctx.restore();
                });
                resolve(canvas.toDataURL().split(',')[1]);
            })
            .start();
    });
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        console.log("Procesando archivo...");
        if (!req.file) return res.status(400).json({ success: false, message: 'Falta archivo' });
        
        let datosManuales = { enero: {}, febrero: {} };
        try {
            datosManuales = JSON.parse(req.body.datosManuales || '{}');
        } catch (e) {}

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        // Mapeo inteligente de columnas
        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim().replace(/ /g, '_') || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
            
            // Prioridad a descripcion textual
            if (val === 'calificacion_desc' || val === 'calificacion_descripcion') colMap.calificacion = colNumber;
            else if (!colMap.calificacion && val.includes('calific') && val.includes('desc')) colMap.calificacion = colNumber;
            
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        // Fallback manual si no encuentra calificacion
        if (!colMap.calificacion) {
             worksheet.getRow(1).eachCell((cell, colNumber) => {
                 const val = cell.value?.toString().toLowerCase() || '';
                 if (val.includes('desc') && val.includes('calific')) colMap.calificacion = colNumber;
             });
        }

        if (!colMap.fecha || !colMap.calificacion) {
            return res.status(400).json({ success: false, message: 'Columnas Fecha o Calificación no encontradas' });
        }

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            
            const rawDate = row.getCell(colMap.fecha).value;
            let date;
            if (rawDate instanceof Date) date = rawDate;
            else if (typeof rawDate === 'string') {
                const parts = rawDate.split('/');
                if(parts.length === 3) date = new Date(parts[2], parts[1]-1, parts[0]);
            }
            if (!date || isNaN(date.getTime())) return;

            const monthIndex = date.getMonth(); 
            const sectorName = colMap.sector ? (row.getCell(colMap.sector).value || 'General').toString().trim() : 'General';
            const ubicacionName = colMap.ubicacion ? (row.getCell(colMap.ubicacion).value || 'General').toString().trim() : 'General';
            
            const califVal = row.getCell(colMap.calificacion).value;
            const calif = califVal ? califVal.toString().toLowerCase() : '';
            const comment = colMap.comentario ? (row.getCell(colMap.comentario).value || '').toString() : '';

            // Estructura de datos
            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {},
                    palabrasPos: [], 
                    palabrasNeg: [], 
                    horasCriticas: Array(24).fill(0)
                };
            }
            
            // Estructura por ubicación
            if (!sectorsData[sectorName].ubicaciones[ubicacionName]) {
                sectorsData[sectorName].ubicaciones[ubicacionName] = { 
                    mp:0, p:0, n:0, mn:0, total:0, 
                    horas: Array(24).fill(0) 
                };
            }

            const sSector = sectorsData[sectorName];
            const sMes = sSector.meses[monthIndex];
            const sUbic = sSector.ubicaciones[ubicacionName];

            // Hora
            let hour = 12;
            if (colMap.hora) {
                const rawTime = row.getCell(colMap.hora).value;
                if (rawTime instanceof Date) hour = rawTime.getUTCHours();
                else if (typeof rawTime === 'string') {
                    const parts = rawTime.split(':');
                    if(parts.length > 0) hour = parseInt(parts[0]);
                }
            }

            sMes.total++;
            sUbic.total++;

            // Clasificación
            if (calif.includes('muy positiva')) {
                sMes.mp++; sUbic.mp++;
                if (comment) sSector.palabrasPos.push(...getWordsFromString(comment));
            } else if (calif.includes('positiva')) {
                sMes.p++; sUbic.p++;
                if (comment) sSector.palabrasPos.push(...getWordsFromString(comment));
            } else if (calif.includes('muy negativa')) {
                sMes.mn++; sUbic.mn++;
                if (comment) sSector.palabrasNeg.push(...getWordsFromString(comment));
                // Hora crítica solo cuenta si es muy negativa
                if (hour >= 0 && hour < 24) {
                    sSector.horasCriticas[hour]++;
                    sUbic.horas[hour]++;
                }
            } else if (calif.includes('negativa')) {
                sMes.n++; sUbic.n++;
                if (comment) sSector.palabrasNeg.push(...getWordsFromString(comment));
            }
        });

        const resultado = [];

        for (const [nombre, data] of Object.entries(sectorsData)) {
            // Unir Manuales (Solo al total mensual, no se puede asignar ubicación)
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i].mp += datosManuales[mes].muy_positivas || 0;
                    data.meses[i].p += datosManuales[mes].positivas || 0;
                    data.meses[i].n += datosManuales[mes].negativas || 0;
                    data.meses[i].mn += datosManuales[mes].muy_negativas || 0;
                    data.meses[i].total += datosManuales[mes].total || 0;
                }
            });

            const mesesFinal = data.meses.map((m, i) => ({
                nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                nps: calculateNPS(m.mp, m.p, m.n + m.mn),
                total: m.total
            }));

            // Procesar Ubicaciones
            const rankingUbicaciones = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, critH = null;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; critH = h; } });
                return {
                    nombre: key,
                    total: u.total,
                    nps: calculateNPS(u.mp, u.p, u.n + u.mn),
                    horaCritica: critH
                };
            }).sort((a,b) => b.nps - a.nps);

            // Generar Nubes
            const contar = (arr) => {
                const counts = {};
                arr.forEach(w => counts[w] = (counts[w] || 0) + 1);
                return Object.entries(counts).sort((a, b) => b[1] - a[1]);
            };

            const topPos = contar(data.palabrasPos);
            const topNeg = contar(data.palabrasNeg);

            // Generación de imágenes
            const imgPos = await generarNubeImagen(topPos, '#2e7d32'); // Verde oscuro
            const imgNeg = await generarNubeImagen(topNeg, '#c62828'); // Rojo oscuro

            // Hora crítica global
            let maxGlob = 0, hGlob = null;
            data.horasCriticas.forEach((c, h) => { if(c > maxGlob) { maxGlob = c; hGlob = h; } });

            const totalAnual = data.meses.reduce((sum, m) => sum + m.total, 0);
            const npsPromedio = (mesesFinal.reduce((sum, m) => sum + m.nps, 0) / 12).toFixed(1);

            resultado.push({
                nombre: nombre,
                meses: mesesFinal,
                ubicaciones: rankingUbicaciones,
                nubePositivaB64: imgPos,
                nubeNegativaB64: imgNeg,
                palabrasNegativas: topNeg.slice(0, 5).map(x => x[0]),
                horaCritica: hGlob,
                totalAnual, npsPromedio
            });
        }

        res.json({ success: true, data: { sectores: resultado } });

    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, message: e.message });
    }
});

app.get('/health', (req, res) => res.send('OK'));
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
