const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba', 'fueron', 'todo', 'estuvo', 'pueden', 'ser', 'solo', 'tenía', 'nada', 'esto'];

function getWords(text) {
    if (!text || text.length < 5) return [];
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'No se recibió archivo' });

        let manual = JSON.parse(req.body.datosManuales || '{}');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
            if (val.includes('calificacion')) colMap.rating = colNumber;
        });

        const sectores = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.rating)?.value);
            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal || isNaN(rating)) return;

            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (isNaN(date.getTime())) return;

            let hVal = row.getCell(colMap.hora)?.value;
            let hora = (hVal instanceof Date) ? hVal.getHours() : (typeof hVal === 'number' ? Math.floor(hVal * 24) : 12);
            
            const sector = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sector]) {
                sectores[sector] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    comsPos: [], comsNeg: [], palabrasPos: [], palabrasNeg: []
                };
            }

            const s = sectores[sector];
            const stats = s.meses[date.getMonth()];

            stats.total++; // Cada fila es una respuesta
            if (rating === 4) stats.mp++;
            if (rating === 3) stats.p++;
            if (rating === 2) stats.n++;
            if (rating === 1) stats.mn++;

            const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${hora}:00hs` };
            if (comment.length > 10) {
                if (rating >= 3) { s.comsPos.push(info); s.palabrasPos.push(...getWords(comment)); }
                else { s.comsNeg.push(info); s.palabrasNeg.push(...getWords(comment)); }
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => {
                if (manual[m]) data.meses[i] = manual[m];
            });

            const mesesFinal = data.meses.map((m, i) => {
                const f = m.total / 100;
                // FORMULA SOLICITADA: (E/(A/100)) - ((B-C)/(A/100))
                const val = m.total > 0 ? ( (m.mp / f) - ((m.mn - m.n) / f) ).toFixed(1) : 0;
                return {
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: parseFloat(val),
                    total: m.total
                };
            });

            const contar = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 25);

            return {
                nombre, meses: mesesFinal,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) 
                },
                nubePos: contar(data.palabrasPos),
                nubeNeg: contar(data.palabrasNeg),
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON`));
