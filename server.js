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

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ success: false, message: 'No hay archivo' });

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
            if (val.includes('calificacion') && !val.includes('desc')) colMap.rating = colNumber;
        });

        const sectores = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rating = parseInt(row.getCell(colMap.rating)?.value);
            const dateVal = row.getCell(colMap.fecha)?.value;
            if (!dateVal || isNaN(rating)) return;

            // Extraer hora real
            let hVal = row.getCell(colMap.hora)?.value;
            let horaReal = 12;
            if (hVal instanceof Date) horaReal = hVal.getHours();
            else if (typeof hVal === 'number') horaReal = Math.floor(hVal * 24);

            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            const sector = (row.getCell(colMap.sector)?.value || 'General').toString().trim();
            const comment = (row.getCell(colMap.comentario)?.value || '').toString().trim();

            if (!sectores[sector]) {
                sectores[sector] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    // Nueva estructura para análisis horario
                    statsHoras: Array.from({length: 24}, () => ({ total: 0, neg: 0 })),
                    comsPos: [], comsNeg: []
                };
            }

            const s = sectores[sector];
            const statsMes = s.meses[date.getMonth()];
            const statsH = s.statsHoras[horaReal];

            // Acumular datos generales
            statsMes.total++;
            statsH.total++;

            if (rating === 4) statsMes.mp++;
            if (rating === 3) statsMes.p++;
            if (rating === 2) { statsMes.n++; statsH.neg++; }
            if (rating === 1) { statsMes.mn++; statsH.neg++; }

            if (comment.length > 20) {
                const info = { texto: comment, meta: `${date.getDate()}/${date.getMonth()+1} ${horaReal}:00hs` };
                if (rating >= 3) s.comsPos.push(info);
                else s.comsNeg.push(info);
            }
        });

        const final = Object.entries(sectores).map(([nombre, data]) => {
            ['enero', 'febrero'].forEach((m, i) => {
                if (manual[m]) data.meses[i] = manual[m];
            });

            // Calcular Satisfacción por mes
            const mesesFinal = data.meses.map((m) => {
                const f = m.total / 100;
                const val = m.total > 0 ? ( (m.mp / f) - ((m.mn + m.n) / f) ).toFixed(1) : 0;
                return { sat: parseFloat(val), total: m.total };
            });

            // ENCONTRAR HORA CRÍTICA
            let horaCritica = 0;
            let maxNegatividad = -1;
            let porcentajeCritico = 0;

            data.statsHoras.forEach((h, index) => {
                if (h.total >= 5) { // Filtro para evitar horas con 1 sola respuesta
                    const porc = (h.neg / h.total) * 100;
                    if (porc > maxNegatividad) {
                        maxNegatividad = porc;
                        horaCritica = index;
                        porcentajeCritico = porc.toFixed(0);
                    }
                }
            });

            return {
                nombre, 
                meses: mesesFinal,
                comentarios: { 
                    pos: data.comsPos.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3), 
                    neg: data.comsNeg.sort((a,b)=>b.texto.length-a.texto.length).slice(0,3) 
                },
                satAnual: (mesesFinal.reduce((a, b) => a + b.sat, 0) / 12).toFixed(1),
                infoHora: {
                    hora: horaCritica.toString().padStart(2, '0') + ':00',
                    porcentaje: porcentajeCritico
                }
            };
        });

        res.json({ success: true, data: { sectores: final } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server ON`));
