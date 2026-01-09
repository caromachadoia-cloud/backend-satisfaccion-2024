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

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'tener', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba', 'fueron', 'todo', 'esta', 'estos', 'ami', 'estuvo', 'estuvieron', 'hacia', 'para', 'pero', 'tiene'];

function getWords(text) {
    if (!text || text.length < 4) return [];
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
            if (val === 'calificacion' || val === 'calificación') colMap.val = colNumber; // Columna numérica G
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.val).value);
            if (isNaN(rating)) return;

            const dateVal = row.getCell(colMap.fecha).value;
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion).value || 'General').toString().trim();
            const rawComment = (row.getCell(colMap.comentario).value || '').toString().trim();

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {}, palabrasRating4: [], palabrasRating1: [], horasCriticas: Array(24).fill(0),
                    commentsRating4: [], commentsRating1: []
                };
            }

            const s = sectorsData[sectorName];
            if (!s.ubicaciones[ubicName]) s.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0, horas: Array(24).fill(0) };
            
            const statsMes = s.meses[date.getMonth()];
            const statsUbic = s.ubicaciones[ubicName];
            statsMes.total++; statsUbic.total++;

            let hVal = row.getCell(colMap.hora).value;
            let hour = (hVal instanceof Date) ? hVal.getUTCHours() : parseInt(hVal?.toString().split(':')[0]) || 12;

            // FILTRO ESTRICTO: Solo tomamos palabras y comentarios de los extremos
            if (rating === 4) {
                statsMes.mp++; statsUbic.mp++;
                s.palabrasRating4.push(...getWords(rawComment));
                if (rawComment.length > 20) s.commentsRating4.push({ texto: rawComment, len: rawComment.length, date });
            } else if (rating === 1) {
                statsMes.mn++; statsUbic.mn++;
                s.horasCriticas[hour]++;
                statsUbic.horas[hour]++;
                s.palabrasRating1.push(...getWords(rawComment));
                if (rawComment.length > 20) s.commentsRating1.push({ texto: rawComment, len: rawComment.length, date });
            }
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            // Unir manuales (Solo mp y total)
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

            const ranking = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, hC = 0;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; hC = h; } });
                return { nombre: key, total: u.total, sat: u.total > 0 ? parseFloat((((u.mp - (u.n + u.mn)) / u.total) * 100).toFixed(1)) : 0, hCrit: hC };
            }).sort((a,b) => b.sat - a.sat);

            const contarTop20 = (arr) => Object.entries(arr.reduce((a,c)=>(a[c]=(a[c]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 20);

            const fmt = (arr) => arr.sort((a,b)=>b.len-a.len).slice(0,3).map(c=>({texto: c.texto, meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:00hs`}));

            return {
                nombre, meses: mesesFinal, ubicaciones: ranking,
                top20Pos: contarTop20(data.palabrasRating4),
                top20Neg: contarTop20(data.palabrasRating1),
                horaCritica: data.horasCriticas.indexOf(Math.max(...data.horasCriticas)),
                comentarios: { pos: fmt(data.commentsRating4), neg: fmt(data.commentsRating1) },
                satPromedio: (mesesFinal.reduce((s, m) => s + m.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Backend Listo`));
