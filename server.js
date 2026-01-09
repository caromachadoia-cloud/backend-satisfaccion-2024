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

// Función para validar si un comentario tiene "sentido"
function esComentarioValido(texto) {
    if (!texto) return false;
    const limpio = texto.trim();
    // Debe tener al menos 15 caracteres y contener letras (evita solo emojis o números)
    const tieneLetras = /[a-zA-Záéíóúñ]{3,}/.test(limpio);
    return limpio.length > 15 && tieneLetras;
}

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'atencion', 'servicio', 'excelente', 'buena', 'mala', 'bien', 'mal', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba', 'fueron', 'todo', 'estuvo', 'para', 'pero'];

function getWords(text) {
    if (!text || text.length < 5) return [];
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, " ")
        .match(/[a-záéíóúñü]+/g)
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
            if (val.includes('fecha')) colMap.fecha = colNumber;
            if (val.includes('hora')) colMap.hora = colNumber;
            if (val.includes('sector')) colMap.sector = colNumber;
            if (val.includes('ubicacion')) colMap.ubicacion = colNumber;
            if (val.includes('calificacion') && !val.includes('desc')) colMap.rating = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            try {
                const rVal = row.getCell(colMap.rating).value;
                const rating = parseInt(rVal);
                const dateVal = row.getCell(colMap.fecha).value;
                if (!dateVal || isNaN(rating)) return;

                let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
                const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
                const comment = (row.getCell(colMap.comentario).value || '').toString().trim();

                if (!sectorsData[sectorName]) {
                    sectorsData[sectorName] = {
                        meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                        ubicaciones: {}, palabras4: [], palabras1: [], horasNeg: Array(24).fill(0),
                        comsPos: [], comsNeg: []
                    };
                }

                const s = sectorsData[sectorName];
                const statsMes = s.meses[date.getMonth()];
                statsMes.total++;

                if (rating === 4) {
                    statsMes.mp++;
                    s.palabras4.push(...getWords(comment));
                    if (esComentarioValido(comment)) {
                        s.comsPos.push({ texto: comment, date });
                    }
                } else if (rating === 1) {
                    statsMes.mn++;
                    s.palabras1.push(...getWords(comment));
                    if (esComentarioValido(comment)) {
                        s.comsNeg.push({ texto: comment, date });
                    }
                } else if (rating === 2) { statsMes.n++; }
                else if (rating === 3) { statsMes.p++; }
            } catch (err) {}
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            // Formatear comentarios: Los más largos suelen ser los más descriptivos
            const formatComs = (arr) => arr
                .sort((a,b) => b.texto.length - a.texto.length)
                .slice(0, 3)
                .map(c => ({
                    texto: c.texto,
                    meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:00hs`
                }));

            return {
                nombre,
                meses: data.meses.map((m, i) => ({
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: m.total > 0 ? parseFloat((((m.mp + m.p - (m.n + m.mn)) / m.total) * 100).toFixed(1)) : 0,
                    total: m.total
                })),
                comentarios: {
                    pos: formatComs(data.comsPos),
                    neg: formatComs(data.comsNeg)
                },
                nubePos: Object.entries(data.palabras4.reduce((acc, w) => (acc[w] = (acc[w] || 0) + 1, acc), {})).sort((a,b)=>b[1]-a[1]).slice(0, 25),
                nubeNeg: Object.entries(data.palabras1.reduce((acc, w) => (acc[w] = (acc[w] || 0) + 1, acc), {})).sort((a,b)=>b[1]-a[1]).slice(0, 25),
                satPromedio: (data.meses.reduce((sum, m) => sum + (m.total > 0 ? ((m.mp+m.p-(m.n+m.mn))/m.total)*100 : 0), 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
