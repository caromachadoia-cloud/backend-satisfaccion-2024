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

// STOPWORDS ampliadas para limpiar mejor la nube
const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'hola', 'gracias', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'atencion', 'servicio', 'estaba', 'esta', 'estuvo', 'fueron', 'tiene', 'tienen'];

function getWords(text) {
    if (!text || text.length < 3) return [];
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
            const rating = parseInt(row.getCell(colMap.rating).value);
            const dateVal = row.getCell(colMap.fecha).value;
            if (!dateVal || isNaN(rating)) return;

            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario).value || '').toString().trim();

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, mn:0, total:0 })),
                    palabrasPos: [], palabrasNeg: [], horasNeg: Array(24).fill(0), coms: {pos:[], neg:[]}
                };
            }

            const s = sectorsData[sectorName];
            const mesIdx = date.getMonth();
            s.meses[mesIdx].total++;

            if (rating >= 4) {
                s.meses[mesIdx].mp++;
                if (comment.length > 5) {
                    s.palabrasPos.push(...getWords(comment));
                    s.coms.pos.push({ texto: comment, date });
                }
            } else if (rating <= 2) {
                s.meses[mesIdx].mn++;
                s.palabrasNeg.push(...getWords(comment));
                s.coms.neg.push({ texto: comment, date });
            }
        });

        const resultado = Object.entries(sectorsData).map(([nombre, data]) => {
            // Inyectar datos manuales (Enero/Febrero)
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i].mp = datosManuales[mes].muy_positivas || 0;
                    data.meses[i].total = datosManuales[mes].total || 0;
                }
            });

            const contar = (arr) => {
                let counts = {};
                arr.forEach(w => counts[w] = (counts[w] || 0) + 1);
                return Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0, 25);
            };

            return {
                nombre,
                meses: data.meses.map((m, i) => ({
                    nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                    sat: m.total > 0 ? ((m.mp / m.total) * 100).toFixed(1) : 0,
                    total: m.total
                })),
                nubePos: contar(data.palabrasPos),
                nubeNeg: contar(data.palabrasNeg),
                comentarios: { 
                    pos: data.coms.pos.slice(0, 3), 
                    neg: data.coms.neg.slice(0, 3) 
                },
                satAnual: (data.meses.reduce((acc, m) => acc + (m.total > 0 ? (m.mp/m.total)*100 : 0), 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
