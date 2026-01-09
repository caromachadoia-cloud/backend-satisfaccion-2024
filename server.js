const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: '*', methods: ['GET', 'POST'], allowedHeaders: ['Content-Type'] }));
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'yyyy', 'todo', 'nada', 'nadie', 'gente', 'fueron', 'tener', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba'];

// --- FILTRO INTELIGENTE DE COMENTARIOS ---
function limpiarYValidarComentario(text) {
    if (!text || text.length < 15) return null; // Ignorar muy cortos
    
    // 1. Eliminar emojis y caracteres especiales raros
    let limpio = text.replace(/([\u2700-\u27BF]|[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD10-\uDDFF])/g, '');
    
    // 2. Detectar si es basura (ej: "aaaaaaaaaa" o "asdasdasdasd")
    if (/(.)\1{3,}/.test(limpio.toLowerCase())) return null; // 4 letras iguales seguidas
    
    return limpio.trim();
}

function puntuarComentario(text, tipo) {
    let score = text.length; // Base por longitud
    const keysPos = ['atencion', 'amable', 'rapido', 'limpio', 'excelente', 'ayudo', 'recomiendo'];
    const keysNeg = ['demora', 'sucio', 'espera', 'maquina', 'atencion', 'lento', 'olor', 'baño', 'caja'];
    
    const keywords = tipo === 'pos' ? keysPos : keysNeg;
    keywords.forEach(key => {
        if (text.toLowerCase().includes(key)) score += 50; // Bonus por palabra clave
    });
    return score;
}

function getWordsFromString(text) {
    if (!text) return [];
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ")
        .match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

function calculateSatisfaction(promotores, pasivos, detractores) {
    const total = promotores + pasivos + detractores;
    if (total === 0) return 0;
    return parseFloat((((promotores - detractores) / total) * 100).toFixed(1));
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'Falta archivo' });
        
        let datosManuales = JSON.parse(req.body.datosManuales || '{}');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim().replace(/ /g, '_') || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
            if (val.includes('desc')) colMap.calificacion = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const dateVal = row.getCell(colMap.fecha).value;
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (isNaN(date.getTime())) return;

            const monthIndex = date.getMonth();
            const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion).value || 'General').toString().trim();
            const calif = (row.getCell(colMap.calificacion).value || '').toString().toLowerCase();
            const rawComment = (row.getCell(colMap.comentario).value || '').toString();

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {},
                    palabrasPos: [], palabrasNeg: [], horasNeg: Array(24).fill(0),
                    comentariosPos: [], comentariosNeg: []
                };
            }

            if (!sectorsData[sectorName].ubicaciones[ubicName]) {
                sectorsData[sectorName].ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0, horas: Array(24).fill(0) };
            }

            const sSector = sectorsData[sectorName];
            const statsMes = sSector.meses[monthIndex];
            const statsUbic = sSector.ubicaciones[ubicName];

            statsMes.total++; statsUbic.total++;

            // Procesar calificación y comentario
            const cleanComment = limpiarYValidarComentario(rawComment);

            if (calif.includes('muy positiva')) {
                statsMes.mp++; statsUbic.mp++;
                if (cleanComment) {
                    sSector.palabrasPos.push(...getWordsFromString(cleanComment));
                    sSector.comentariosPos.push({ text: cleanComment, score: puntuarComentario(cleanComment, 'pos'), date });
                }
            } else if (calif.includes('muy negativa') || calif.includes('negativa')) {
                const isVeryNeg = calif.includes('muy');
                isVeryNeg ? statsMes.mn++ : statsMes.n++;
                isVeryNeg ? statsUbic.mn++ : statsUbic.n++;

                let hour = 12;
                if (colMap.hora) {
                    const hVal = row.getCell(colMap.hora).value;
                    hour = (hVal instanceof Date) ? hVal.getUTCHours() : parseInt(hVal?.toString().split(':')[0]) || 12;
                }
                sSector.horasNeg[hour]++;
                statsUbic.horas[hour]++;

                if (cleanComment) {
                    sSector.palabrasNeg.push(...getWordsFromString(cleanComment));
                    sSector.comentariosNeg.push({ text: cleanComment, score: puntuarComentario(cleanComment, 'neg'), date });
                }
            }
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            // Unir manuales
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i].mp += datosManuales[mes].muy_positivas || 0;
                    data.meses[i].total += datosManuales[mes].total || 0;
                }
            });

            const mesesFinal = data.meses.map((m, i) => ({
                nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                sat: calculateSatisfaction(m.mp, m.p, m.n + m.mn),
                total: m.total
            }));

            // Ranking Ubicaciones
            const ranking = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, hCrit = 0;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; hCrit = h; } });
                return { nombre: key, total: u.total, sat: calculateSatisfaction(u.mp, u.p, u.n + u.mn), horaCritica: hCrit };
            }).sort((a,b) => b.sat - a.sat);

            // Mejores comentarios (Top 3 por score)
            const getTop = (arr) => arr.sort((a,b) => b.score - a.score).slice(0, 3).map(c => ({
                texto: c.text, 
                meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:${c.date.getUTCMinutes().toString().padStart(2,'0')}hs`
            }));

            // Palabras para nube (limitado a top 60 para evitar amontonamiento)
            const contar = (arr) => Object.entries(arr.reduce((a,c)=>(a[c]=(a[c]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 60);

            let maxG = 0, hG = 0;
            data.horasNeg.forEach((c, h) => { if(c > maxG) { maxG = c; hG = h; } });

            return {
                nombre, meses: mesesFinal, ubicaciones: ranking,
                palabrasPos: contar(data.palabrasPos),
                palabrasNeg: contar(data.palabrasNeg),
                horaCritica: hG,
                comentarios: { pos: getTop(data.comentariosPos), neg: getTop(data.comentariosNeg) },
                totalAnual: data.meses.reduce((s, m) => s + m.total, 0),
                satPromedio: (mesesFinal.reduce((s, m) => s + m.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.listen(PORT, () => console.log(`Server running on ${PORT}`));
