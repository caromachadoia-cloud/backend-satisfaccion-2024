const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');
const { createCanvas } = require('canvas');
const d3Cloud = require('d3-cloud');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: '*', methods: ['GET', 'POST'], allowedHeaders: ['Content-Type'] }));
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const STOPWORDS = [
    'de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un', 'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como',
    'más', 'pero', 'sus', 'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'muy', 'cuando', 'también', 'hasta', 'hay', 'donde',
    'quien', 'desde', 'todo', 'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son', 'fue', 'gracias', 'hola',
    'buen', 'dia', 'tarde', 'noche', 'lugar', 'servicio', 'atencion', 'excelente', 'buena', 'mala', 'regular', 'bien', 'mal', 'yyyy', 'todo', 
    'nada', 'nadie', 'gente', 'cosas', 'porque', 'estan', 'estaba', 'fueron'
];

function getWordsFromString(text) {
    if (!text || typeof text !== 'string') return [];
    return text.toLowerCase()
        .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, "")
        .replace(/\s{2,}/g, " ")
        .match(/\b(\w+)\b/g)
        ?.filter(word => !STOPWORDS.includes(word) && word.length > 3) || [];
}

function calculateNPS(promotores, pasivos, detractores) {
    const total = promotores + pasivos + detractores;
    if (total === 0) return 0;
    const nps = ((promotores - detractores) / total) * 100;
    return parseFloat(nps.toFixed(1));
}

function formatDateShort(date) {
    if (!date) return '';
    const d = new Date(date);
    const day = d.getUTCDate().toString().padStart(2, '0');
    const month = (d.getUTCMonth() + 1).toString().padStart(2, '0');
    const hours = d.getUTCHours().toString().padStart(2, '0');
    const mins = d.getUTCMinutes().toString().padStart(2, '0');
    return `${day}/${month} ${hours}:${mins}hs`;
}

// AUMENTAMOS RESOLUCIÓN PARA PÁGINA COMPLETA
function generarNubeImagen(wordList, colorHex) {
    if (!wordList || wordList.length === 0) return Promise.resolve(null);
    
    return new Promise(resolve => {
        const width = 1000; // MÁS GRANDE
        const height = 500; // MÁS GRANDE
        const canvas = createCanvas(width, height);
        const ctx = canvas.getContext('2d');
        
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);

        const maxWeight = Math.max(...wordList.map(w => w[1]));
        const minWeight = Math.min(...wordList.map(w => w[1]));
        
        const words = wordList.slice(0, 60).map(item => ({
            text: item[0],
            size: 30 + 70 * ((item[1] - minWeight) / ((maxWeight - minWeight) || 1)) // Letras más grandes
        }));

        d3Cloud()
            .size([width, height])
            .canvas(() => createCanvas(1, 1))
            .words(words)
            .padding(10)
            .rotate(() => 0)
            .font('sans-serif')
            .fontSize(d => d.size)
            .on('end', (drawnWords) => {
                ctx.translate(width / 2, height / 2);
                drawnWords.forEach(w => {
                    ctx.save();
                    ctx.translate(w.x, w.y);
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillStyle = colorHex;
                    ctx.font = `bold ${w.size}px sans-serif`;
                    ctx.fillText(w.text, 0, 0);
                    ctx.restore();
                });
                resolve(canvas.toDataURL().split(',')[1]);
            })
            .start();
    });
}

app.post('/procesar-anual', upload.single('archivoExcel'), async (req, res) => {
    try {
        console.log("Procesando...");
        if (!req.file) return res.status(400).json({ success: false, message: 'Falta archivo' });
        
        let datosManuales = {};
        try { datosManuales = JSON.parse(req.body.datosManuales || '{}'); } catch (e) {}

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
            if (val === 'calificacion_desc' || val === 'calificacion_descripcion') colMap.calificacion = colNumber;
            else if (!colMap.calificacion && val.includes('calific') && val.includes('desc')) colMap.calificacion = colNumber;
            if (val.includes('comentario')) colMap.comentario = colNumber;
        });

        if (!colMap.calificacion) {
             worksheet.getRow(1).eachCell((cell, colNumber) => {
                 const val = cell.value?.toString().toLowerCase() || '';
                 if (val.includes('desc') && val.includes('calific')) colMap.calificacion = colNumber;
             });
        }

        if (!colMap.fecha || !colMap.calificacion) return res.status(400).json({ success: false, message: 'Faltan columnas clave.' });

        const sectorsData = {};

        worksheet.eachRow((row, rowNum) => {
            if (rowNum === 1) return;
            const rawDate = row.getCell(colMap.fecha).value;
            let date;
            if (rawDate instanceof Date) date = rawDate;
            else if (typeof rawDate === 'string') {
                const parts = rawDate.split('/');
                if(parts.length === 3) date = new Date(parts[2], parts[1]-1, parts[0]);
            }
            if (!date || isNaN(date.getTime())) return;

            const monthIndex = date.getMonth(); 
            const sectorName = colMap.sector ? (row.getCell(colMap.sector).value || 'General').toString().trim() : 'General';
            const ubicacionName = colMap.ubicacion ? (row.getCell(colMap.ubicacion).value || 'General').toString().trim() : 'General';
            const califVal = row.getCell(colMap.calificacion).value;
            const calif = califVal ? califVal.toString().toLowerCase() : '';
            const comment = colMap.comentario ? (row.getCell(colMap.comentario).value || '').toString().trim() : '';

            let hour = 12;
            let fullDate = new Date(date); 
            if (colMap.hora) {
                const rawTime = row.getCell(colMap.hora).value;
                if (rawTime instanceof Date) {
                    hour = rawTime.getUTCHours();
                    fullDate.setHours(hour, rawTime.getUTCMinutes());
                } else if (typeof rawTime === 'string') {
                    const parts = rawTime.split(':');
                    if(parts.length > 0) hour = parseInt(parts[0]);
                    fullDate.setHours(hour, parts[1] || 0);
                }
            }

            if (!sectorsData[sectorName]) {
                sectorsData[sectorName] = {
                    meses: Array.from({length: 12}, () => ({ mp:0, p:0, n:0, mn:0, total:0 })),
                    ubicaciones: {},
                    palabrasPos: [], palabrasNeg: [], horasDetractoras: Array(24).fill(0),
                    listadoComentarios: { positiva: [], negativa: [] }
                };
            }
            
            if (!sectorsData[sectorName].ubicaciones[ubicacionName]) {
                sectorsData[sectorName].ubicaciones[ubicacionName] = { mp:0, p:0, n:0, mn:0, total:0, horas: Array(24).fill(0) };
            }

            const sSector = sectorsData[sectorName];
            const sMes = sSector.meses[monthIndex];
            const sUbic = sSector.ubicaciones[ubicacionName];

            sMes.total++; sUbic.total++;

            if (calif.includes('muy positiva')) {
                sMes.mp++; sUbic.mp++;
                if (comment.length > 3) {
                    sSector.palabrasPos.push(...getWordsFromString(comment));
                    sSector.listadoComentarios.positiva.push({ text: comment, date: fullDate });
                }
            } else if (calif.includes('positiva')) {
                sMes.p++; sUbic.p++;
                if (comment.length > 3) sSector.palabrasPos.push(...getWordsFromString(comment));
            } else if (calif.includes('muy negativa')) {
                sMes.mn++; sUbic.mn++;
                if (hour >= 0 && hour < 24) { sSector.horasDetractoras[hour]++; sUbic.horas[hour]++; }
                if (comment.length > 3) {
                    sSector.palabrasNeg.push(...getWordsFromString(comment));
                    sSector.listadoComentarios.negativa.push({ text: comment, date: fullDate });
                }
            } else if (calif.includes('negativa')) {
                sMes.n++; sUbic.n++;
                if (hour >= 0 && hour < 24) { sSector.horasDetractoras[hour]++; sUbic.horas[hour]++; }
                if (comment.length > 3) {
                    sSector.palabrasNeg.push(...getWordsFromString(comment));
                    sSector.listadoComentarios.negativa.push({ text: comment, date: fullDate });
                }
            }
        });

        const resultado = [];

        for (const [nombre, data] of Object.entries(sectorsData)) {
            ['enero', 'febrero'].forEach((mes, i) => {
                if (datosManuales[mes]) {
                    data.meses[i].mp += datosManuales[mes].muy_positivas || 0;
                    data.meses[i].p += datosManuales[mes].positivas || 0;
                    data.meses[i].n += datosManuales[mes].negativas || 0;
                    data.meses[i].mn += datosManuales[mes].muy_negativas || 0;
                    data.meses[i].total += datosManuales[mes].total || 0;
                }
            });

            const mesesFinal = data.meses.map((m, i) => ({
                nombre: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][i],
                nps: calculateNPS(m.mp, m.p, m.n + m.mn),
                total: m.total
            }));

            const rankingUbicaciones = Object.entries(data.ubicaciones).map(([key, u]) => {
                let maxH = 0, critH = null;
                u.horas.forEach((c, h) => { if(c > maxH) { maxH = c; critH = h; } });
                return { nombre: key, total: u.total, nps: calculateNPS(u.mp, u.p, u.n + u.mn), horaCritica: critH };
            }).sort((a,b) => b.nps - a.nps);

            const contar = (arr) => Object.entries(arr.reduce((a,c)=>(a[c]=(a[c]||0)+1,a),{})).sort((a,b)=>b[1]-a[1]);
            const imgPos = await generarNubeImagen(contar(data.palabrasPos), '#2e7d32');
            const imgNeg = await generarNubeImagen(contar(data.palabrasNeg), '#c62828');

            let maxGlob = 0, hGlob = null;
            data.horasDetractoras.forEach((c, h) => { if(c > maxGlob) { maxGlob = c; hGlob = h; } });

            // SORT COMMENTS: Más largos primero
            const sortComments = (arr) => arr.sort((a, b) => b.text.length - a.text.length).slice(0, 3).map(c => ({
                texto: c.text, fecha: formatDateShort(c.date)
            }));

            resultado.push({
                nombre: nombre,
                meses: mesesFinal,
                ubicaciones: rankingUbicaciones,
                nubePositivaB64: imgPos,
                nubeNegativaB64: imgNeg,
                horaCritica: hGlob,
                palabrasNegativas: contar(data.palabrasNeg).slice(0, 5).map(x => x[0]),
                comentariosReales: {
                    positivos: sortComments(data.listadoComentarios.positiva),
                    negativos: sortComments(data.listadoComentarios.negativa)
                },
                totalAnual: data.meses.reduce((s, m) => s + m.total, 0),
                npsPromedio: (mesesFinal.reduce((s, m) => s + m.nps, 0) / 12).toFixed(1)
            });
        }

        res.json({ success: true, data: { sectores: resultado } });
    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, message: e.message });
    }
});

app.get('/health', (req, res) => res.send('OK'));
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
