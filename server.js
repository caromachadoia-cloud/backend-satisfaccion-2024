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

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'esta', 'estaba', 'fueron', 'estuvo', 'para', 'pero', 'hace', 'solo', 'tenía', 'nada'];

function getWords(text) {
    if (!text || text.length < 5) return [];
    return text.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ").match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'No se recibió archivo' });

        let datosManuales = JSON.parse(req.body.datosManuales || '{}');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.worksheets[0];

        const colMap = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (val.includes('respuestas') || val === 'total') colMap.total = colNumber;
            if (val.includes('muy negativas')) colMap.mn = colNumber;
            if (val.includes('negativas') && !val.includes('muy')) colMap.n = colNumber;
            if (val.includes('muy positivas')) colMap.mp = colNumber;
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            try {
                const totalRow = parseInt(row.getCell(colMap.total)?.value) || 0;
                const dateVal = row.getCell(colMap.fecha)?.value;
                if (!dateVal || totalRow === 0) return;

                let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
                if (isNaN(date.getTime())) return;

                let hVal = row.getCell(colMap.hora)?.value;
                let horaReal = (hVal instanceof Date) ? hVal.getHours() : (typeof hVal === 'number' ? Math.floor(hVal * 24) : 12);

                const sectorName = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
                const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

                if (!sectorsData[sectorName]) {
                    sectorsData[sectorName] = {
                        meses: Array.from({length: 12}, () => ({ mp:0, mn:0, n:0, total:0 })),
                        palabrasPos: [], palabrasNeg: [], comsPos: [], comsNeg: []
                    };
                }

                const s = sectorsData[sectorName];
                const statsMes = s.meses[date.getMonth()];
                const mp = parseInt(row.getCell(colMap.mp)?.value) || 0;
                const mn = parseInt(row.getCell(colMap.mn)?.value) || 0;
                const n = parseInt(row.getCell(colMap.n)?.value) || 0;

                statsMes.total += totalRow;
                statsMes.mp += mp;
                statsMes.mn += mn;
                statsMes.n += n;

                const infoCom = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
                if (comment.length > 10) {
                    if (mp > 0) { s.palabrasPos.push(...getWords(comment)); s.comsPos.push(infoCom); }
                    else if (mn > 0 || n > 0) { s.palabrasNeg.push(...getWords(comment)); s.comsNeg.push(infoCom); }
                }
            } catch (err) {}
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i] = {
                        mp: datosManuales[mes].muy_positivas || 0,
                        mn: datosManuales[mes].muy_negativas || 0,
                        n: datosManuales[mes].negativas || 0,
                        total: datosManuales[mes].total || 0
                    };
                }
            });

            const mesesFinal = data.meses.map((m, i) => {
                // FORMULA: (E / (A/100)) - ((B-C) / (A/100))
                const factor = m.total / 100;
                const satValue = m.total > 0 ? ( (m.mp / factor) - ((m.mn - m.n) / factor) ).toFixed(1) : 0;
                return {
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: parseFloat(satValue),
                    total: m.total
                };
            });

            const contar = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 25);

            return {
                nombre, meses: mesesFinal,
                comentarios: { pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) },
                nubePos: contar(data.palabrasPos), nubeNeg: contar(data.palabrasNeg),
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server running`));
