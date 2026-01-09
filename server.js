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
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ")
        .match(/[a-záéíóúñü]+/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

function esComentarioValido(texto) {
    if (!texto) return false;
    const limpio = texto.trim();
    return limpio.length > 15 && /[a-zA-Záéíóúñ]{3,}/.test(limpio);
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
            if (val.includes('respuestas')) colMap.total = colNumber; // Columna A
            if (val.includes('muy negativas')) colMap.mn = colNumber;  // Columna B
            if (val.includes('negativas') && !val.includes('muy')) colMap.n = colNumber; // Columna C
            if (val.includes('muy positivas')) colMap.mp = colNumber; // Columna E
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            try {
                const totalRow = parseInt(row.getCell(colMap.total).value) || 0;
                const dateVal = row.getCell(colMap.fecha).value;
                if (!dateVal || totalRow === 0) return;

                let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
                let hVal = row.getCell(colMap.hora).value;
                let horaReal = 0;
                if (hVal instanceof Date) horaReal = hVal.getHours();
                else if (typeof hVal === 'number') horaReal = Math.floor(hVal * 24);

                const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
                const comment = (row.getCell(colMap.comentario).value || '').toString().trim();

                if (!sectorsData[sectorName]) {
                    sectorsData[sectorName] = {
                        meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                        palabras4: [], palabras1: [], horasNeg: Array(24).fill(0), coms4: [], coms1: []
                    };
                }

                const s = sectorsData[sectorName];
                const statsMes = s.meses[date.getMonth()];

                const mp = parseInt(row.getCell(colMap.mp).value) || 0;
                const mn = parseInt(row.getCell(colMap.mn).value) || 0;
                const n = parseInt(row.getCell(colMap.n).value) || 0;

                statsMes.total += totalRow;
                statsMes.mp += mp;
                statsMes.mn += mn;
                statsMes.n += n;

                const infoCom = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };

                if (comment.length > 5) {
                    if (mp > 0) {
                        s.palabras4.push(...getWords(comment));
                        if (esComentarioValido(comment)) s.coms4.push(infoCom);
                    } else if (mn > 0 || n > 0) {
                        s.horasNeg[horaReal]++;
                        s.palabras1.push(...getWords(comment));
                        if (esComentarioValido(comment)) s.coms1.push(infoCom);
                    }
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
                // FÓRMULA SOLICITADA: (E2/(A2/100)) - ((B2-C2)/(A2/100))
                const satValue = m.total > 0 ? ( (m.mp / (m.total/100)) - ((m.mn - m.n) / (m.total/100)) ).toFixed(1) : 0;
                return {
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: parseFloat(satValue),
                    total: m.total
                };
            });

            const fmtComs = (arr) => arr.sort((a,b) => b.texto.length - a.texto.length).slice(0, 3);
            const contar = (arr) => Object.entries(arr.reduce((a,w)=>(a[w]=(a[w]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 25);

            return {
                nombre, meses: mesesFinal,
                comentarios: { pos: fmtComs(data.coms4), neg: fmtComs(data.coms1) },
                nubePos: contar(data.palabras4), nubeNeg: contar(data.palabras1),
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Servidor Corriendo`));
