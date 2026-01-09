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

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'fueron', 'tener', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba'];

function getWords(text) {
    if (!text || text.length < 5) return [];
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ")
        .match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
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
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
            if (val === 'calificacion' || val === 'calificación') colMap.val = colNumber; 
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.val).value);
            const dateVal = row.getCell(colMap.fecha).value;
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (isNaN(date.getTime()) || isNaN(rating)) return;

            const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion).value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario).value || '').toString().trim();

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {},
                    // Guardamos palabras por cada nota específica
                    palabrasPorNota: { 1: [], 2: [], 3: [], 4: [] },
                    horasNeg: Array(24).fill(0),
                    comsPos: [], comsNeg: []
                };
            }

            const s = sectorsData[sectorName];
            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0, horas: Array(24).fill(0) };
            
            const statsMes = s.meses[date.getMonth()];
            const statsUbic = s.ubicaciones[ubicName];
            statsMes.total++; statsUbic.total++;

            if (rating >= 3) {
                rating === 4 ? statsMes.mp++ : statsMes.p++;
                rating === 4 ? statsUbic.mp++ : statsUbic.p++;
                const words = getWords(comment);
                words.forEach(w => s.palabrasPorNota[rating].push(w));
                if (comment.length > 20 && rating === 4) s.comsPos.push({ texto: comment, len: comment.length, date });
            } else {
                rating === 1 ? statsMes.mn++ : statsMes.n++;
                rating === 1 ? statsUbic.mn++ : statsUbic.n++;
                const words = getWords(comment);
                words.forEach(w => s.palabrasPorNota[rating].push(w));
                
                let hVal = row.getCell(colMap.hora).value;
                let hour = (hVal instanceof Date) ? hVal.getUTCHours() : parseInt(hVal?.toString().split(':')[0]) || 12;
                s.horasNeg[hour]++; statsUbic.horas[hour]++;
                if (comment.length > 20 && rating === 1) s.comsNeg.push({ texto: comment, len: comment.length, date });
            }
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i].mp += datosManuales[mes].muy_positivas || 0;
                    data.meses[i].total += datosManuales[mes].total || 0;
                }
            });

            const mesesFinal = data.meses.map((m, i) => ({
                nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                sat: m.total > 0 ? parseFloat((((m.mp - (m.n + m.mn)) / m.total) * 100).toFixed(1)) : 0,
                total: m.total
            }));

            // Función para contar palabras y mantener la etiqueta de la nota
            const procesarPalabras = (notas) => {
                let conteoGlobal = {};
                notas.forEach(n => {
                    data.palabrasPorNota[n].forEach(w => {
                        if (!conteoGlobal[w]) conteoGlobal[w] = { freq: 0, nota: n };
                        conteoGlobal[w].freq++;
                    });
                });
                return Object.entries(conteoGlobal)
                    .map(([word, info]) => [word, info.freq, info.nota])
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 40); // Top 40 para que no se amontone
            };

            const ranking = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, hC = 0;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; hC = h; } });
                return { nombre: key, total: u.total, sat: u.total > 0 ? parseFloat((((u.mp - (u.n + u.mn)) / u.total) * 100).toFixed(1)) : 0, hCrit: hC };
            }).sort((a,b) => b.sat - a.sat);

            return {
                nombre, meses: mesesFinal, ubicaciones: ranking,
                nubePos: procesarPalabras([3, 4]),
                nubeNeg: procesarPalabras([1, 2]),
                horaCritica: data.horasNeg.indexOf(Math.max(...data.horasNeg)),
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.len-a.len).slice(0,3).map(c=>({texto: c.texto, meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:00hs`})),
                    neg: data.coms1.sort((a,b)=>b.len-a.len).slice(0,3).map(c=>({texto: c.texto, meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:00hs`}))
                },
                satPromedio: (mesesFinal.reduce((s, m) => s + m.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server Ready`));
