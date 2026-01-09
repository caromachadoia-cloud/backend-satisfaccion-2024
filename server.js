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

const STOPWORDS = ['de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola', 'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'fueron', 'tener', 'hace', 'falta', 'mucha', 'mucho', 'esta', 'estos', 'estaba', 'todo'];

function limpiarTexto(text) {
    if (!text || text.length < 15) return null;
    let limpio = text.replace(/([\u2700-\u27BF]|[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD10-\uDDFF])/g, '');
    if (/(.)\1{3,}/.test(limpio.toLowerCase())) return null; 
    return limpio.trim();
}

function getWords(text) {
    if (!text) return [];
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
            if (val.includes('desc')) colMap.calificacion = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const dateVal = row.getCell(colMap.fecha).value;
            let date = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
            if (isNaN(date.getTime())) return;

            const sectorName = (row.getCell(colMap.sector).value || 'General').toString().trim();
            const ubicName = (row.getCell(colMap.ubicacion).value || 'General').toString().trim();
            const calif = (row.getCell(colMap.calificacion).value || '').toString().toLowerCase();
            const rawComment = (row.getCell(colMap.comentario).value || '').toString();

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {}, palabrasPos: [], palabrasNeg: [], horasNeg: Array(24).fill(0),
                    comentariosPos: [], comentariosNeg: []
                };
            }

            const sSector = sectorsData[sectorName];
            if (!sSector.ubicaciones[ubicName]) sSector.ubicaciones[ubicName] = { mp:0, p:0, n:0, mn:0, total:0, horas: Array(24).fill(0) };
            
            const statsMes = sSector.meses[date.getMonth()];
            const statsUbic = sSector.ubicaciones[ubicName];
            statsMes.total++; statsUbic.total++;

            const cleanComment = limpiarTexto(rawComment);

            if (calif.includes('muy positiva')) {
                statsMes.mp++; statsUbic.mp++;
                if (cleanComment) {
                    sSector.palabrasPos.push(...getWords(cleanComment));
                    sSector.comentariosPos.push({ text: cleanComment, len: cleanComment.length, date });
                }
            } else if (calif.includes('negativa')) {
                const isVery = calif.includes('muy');
                isVery ? statsMes.mn++ : statsMes.n++;
                isVery ? statsUbic.mn++ : statsUbic.n++;
                
                let hVal = row.getCell(colMap.hora).value;
                let hour = (hVal instanceof Date) ? hVal.getUTCHours() : parseInt(hVal?.toString().split(':')[0]) || 12;
                sSector.horasNeg[hour]++;
                statsUbic.horas[hour]++;

                if (cleanComment) {
                    sSector.palabrasNeg.push(...getWords(cleanComment));
                    sSector.comentariosNeg.push({ text: cleanComment, len: cleanComment.length, date });
                }
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

            const ranking = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, hCrit = 0;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; hCrit = h; } });
                return { nombre: key, total: u.total, sat: u.total > 0 ? parseFloat((((u.mp - (u.n + u.mn)) / u.total) * 100).toFixed(1)) : 0, horaCritica: hCrit };
            }).sort((a,b) => b.sat - a.sat);

            const contar = (arr) => Object.entries(arr.reduce((a,c)=>(a[c]=(a[c]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]).slice(0, 50);

            const fmtCom = (arr) => arr.sort((a,b)=>b.len-a.len).slice(0,3).map(c=>({texto: c.text, meta: `${c.date.getUTCDate()}/${c.date.getUTCMonth()+1} ${c.date.getUTCHours()}:${c.date.getUTCMinutes().toString().padStart(2,'0')}hs`}));

            return {
                nombre, meses: mesesFinal, ubicaciones: ranking,
                palabrasPos: contar(data.palabrasPos),
                palabrasNeg: contar(data.palabrasNeg),
                horaCritica: data.horasNeg.indexOf(Math.max(...data.horasNeg)),
                comentarios: { pos: fmtCom(data.comentariosPos), neg: fmtCom(data.comentariosNeg) },
                satPromedio: (mesesFinal.reduce((s, m) => s + m.sat, 0) / 12).toFixed(1)
            };
        });

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) { res.status(500).json({ success: false, message: e.message }); }
});

app.listen(PORT, () => console.log(`Server running`));
